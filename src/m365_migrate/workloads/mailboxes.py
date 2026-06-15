"""Mailbox workload: discover, plan, and migrate Exchange Online mailbox settings.

Like the other workloads, this is staged so a human can review before any write:

    discover  -> read mailbox settings for mailbox-enabled source users
    plan      -> map source -> target user (UPN rewrite); decide settings | skip
    migrate   -> PATCH the writable mailboxSettings onto the target mailbox

**Scope.** This migrates mailbox *settings* (time zone, language, working hours,
automatic replies, date/time formats, delegate options) via Graph's
``/users/{id}/mailboxSettings`` endpoint. Migrating mailbox *content* (mail,
calendar, contacts) is not a Graph REST operation tenant-to-tenant: it requires a
native cross-tenant mailbox move (MRS / migration endpoints) or a third-party
tool, and is tracked separately on the roadmap.
"""

from __future__ import annotations

from m365_migrate.config import Config
from m365_migrate.graph_client import GraphClient, GraphError
from m365_migrate.mapping import rewrite_upn
from m365_migrate.models import PlannedMailbox, SourceMailbox
from m365_migrate.workloads.users import get_target_user_ids


def _rewrite(upn: str, config: Config) -> str:
    """Apply the configured UPN domain rewrite to a raw UPN string."""
    if config.options.rewrite_upn_domain:
        return rewrite_upn(upn, config.source.primary_domain, config.target.primary_domain)
    return upn


def discover_mailboxes(client: GraphClient) -> list[SourceMailbox]:
    """Read mailbox settings for every mailbox-enabled user in the source tenant.

    Users without a mailbox return ``404`` from the mailboxSettings endpoint and
    are skipped, so the result contains only genuinely mailbox-enabled users.
    """
    users = client.get_all(
        "/users", params={"$select": "id,userPrincipalName,mail", "$top": 999}
    )
    mailboxes: list[SourceMailbox] = []
    for user in users:
        upn = user.get("userPrincipalName")
        # No SMTP address almost always means no Exchange mailbox; skip the probe.
        if not upn or not user.get("mail"):
            continue
        try:
            data = client.get(f"/users/{user['id']}/mailboxSettings")
        except GraphError as exc:
            if exc.status_code == 404:
                continue  # mailbox not provisioned for this user
            raise
        mailboxes.append(SourceMailbox.from_graph(upn, data))
    return mailboxes


def plan_mailboxes(
    mailboxes: list[SourceMailbox],
    config: Config,
    target_upns: set[str] | None = None,
) -> list[PlannedMailbox]:
    """Build a mailbox-settings migration plan without writing anything.

    ``target_upns`` (the UPNs that exist in the target tenant) lets the planner
    skip source mailboxes whose target account is not present yet.
    """
    existing = {u.lower() for u in (target_upns or set())}
    planned: list[PlannedMailbox] = []
    for mb in mailboxes:
        target_upn = _rewrite(mb.user_principal_name, config)

        if target_upn.lower() not in existing:
            action, reason = "skip", "target account not found"
        elif not mb.settings:
            action, reason = "skip", "no settings to migrate"
        else:
            action, reason = "settings", None

        planned.append(
            PlannedMailbox(
                source_upn=mb.user_principal_name,
                target_upn=target_upn,
                action=action,
                reason=reason,
                settings=mb.settings,
            )
        )
    return planned


def migrate_mailboxes(
    client: GraphClient,
    planned: list[PlannedMailbox],
    config: Config,
    *,
    dry_run: bool = True,
) -> list[dict]:
    """Apply planned mailbox settings to the target tenant.

    Resolves each target UPN to its user id and PATCHes the writable
    mailboxSettings. With ``dry_run=True`` (default) nothing is written; each
    actionable mailbox is reported as ``would-update``.
    """
    target_ids = get_target_user_ids(client)
    results: list[dict] = []

    for p in planned:
        record: dict = {"target_upn": p.target_upn}

        if p.action != "settings":
            record["status"] = f"skipped:{p.action}"
            record["reason"] = p.reason
            results.append(record)
            continue

        target_id = target_ids.get(p.target_upn.lower())
        if not target_id:
            record["status"] = "skipped:not-in-target"
            results.append(record)
            continue

        if dry_run:
            record["status"] = "would-update"
            results.append(record)
            continue

        try:
            client.patch(f"/users/{target_id}/mailboxSettings", json=p.settings)
            record["status"] = "updated"
        except GraphError as exc:
            record["status"] = "error"
            record["reason"] = f"{exc.status_code}"
        results.append(record)

    return results
