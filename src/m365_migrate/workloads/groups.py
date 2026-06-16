"""Groups workload: discover, plan, and sync groups + their membership.

Like the users workload, this is staged so a human can review before any write:

    discover  -> read groups from the source tenant (with user members + owners)
    plan      -> match each source group to the target by mailNickname, classify
                 it, and resolve members + owners through the user UPN mapping
    sync      -> provision missing groups and reconcile membership + ownership

Only **security** and **Microsoft 365** groups can be provisioned through Graph.
Mail-enabled security groups and distribution lists require Exchange Online and
are surfaced as ``skip`` for a later workload. Membership and ownership are
resolved through the same source->target UPN rewrite used by the users workload,
so a member/owner is only added once the corresponding target user exists.
Reconciling owners matters beyond access control: an M365 group must have an owner
before it can be Teams-enabled by the teams workload.
"""

from __future__ import annotations

from m365_migrate.config import Config
from m365_migrate.graph_client import GraphClient, GraphError
from m365_migrate.mapping import rewrite_upn
from m365_migrate.models import GROUP_EXPAND, GROUP_SELECT_FIELDS, PlannedGroup, SourceGroup
from m365_migrate.workloads.users import get_target_user_ids

# Group kinds the tool can create directly via Graph.
CREATABLE_KINDS = {"security", "microsoft365"}


def discover_groups(client: GraphClient) -> list[SourceGroup]:
    """Read all groups from the source tenant, including their members + owners."""
    select = ",".join(GROUP_SELECT_FIELDS)
    raw = client.get_all(
        "/groups", params={"$select": select, "$expand": GROUP_EXPAND, "$top": 999}
    )
    return [SourceGroup.from_graph(item) for item in raw]


def _rewrite(upn: str, config: Config) -> str:
    """Apply the configured UPN domain rewrite to a raw UPN string."""
    if config.options.rewrite_upn_domain:
        return rewrite_upn(upn, config.source.primary_domain, config.target.primary_domain)
    return upn


def get_target_group_ids(client: GraphClient) -> dict[str, str]:
    """Return a mapping of lowercased target group mailNickname -> group id."""
    raw = client.get_all(
        "/groups", params={"$select": "id,mailNickname", "$top": 999}
    )
    return {
        item["mailNickname"].lower(): item["id"]
        for item in raw
        if item.get("mailNickname") and item.get("id")
    }


def get_group_member_ids(client: GraphClient, group_id: str) -> set[str]:
    """Return the set of directory-object ids already members of a target group."""
    raw = client.get_all(
        f"/groups/{group_id}/members", params={"$select": "id", "$top": 999}
    )
    return {item["id"] for item in raw if item.get("id")}


def get_group_owner_ids(client: GraphClient, group_id: str) -> set[str]:
    """Return the set of directory-object ids already owners of a target group."""
    raw = client.get_all(
        f"/groups/{group_id}/owners", params={"$select": "id", "$top": 999}
    )
    return {item["id"] for item in raw if item.get("id")}


def plan_groups(
    groups: list[SourceGroup],
    config: Config,
    existing_target_groups: dict[str, str] | None = None,
) -> list[PlannedGroup]:
    """Build a group reconciliation plan without writing anything.

    ``existing_target_groups`` maps lowercased target mailNickname -> group id and
    lets the planner decide whether a group must be created (``create``) or
    already exists in the target (``exists``). Group kinds that cannot be
    provisioned through Graph are marked ``skip``. Member and owner UPNs are
    rewritten to the target domain so the plan reflects what would be reconciled.
    Dynamic groups carry their membership rule instead of assigned members.
    """
    existing = existing_target_groups or {}
    planned: list[PlannedGroup] = []
    for group in groups:
        kind = group.kind
        nickname = group.mail_nickname

        if kind not in CREATABLE_KINDS:
            action, reason = "skip", f"{kind} group not provisionable via Graph"
        elif not nickname:
            action, reason = "skip", "group has no mailNickname"
        elif nickname.lower() in existing:
            action, reason = "exists", None
        else:
            action, reason = "create", None

        planned.append(
            PlannedGroup(
                source_id=group.id,
                mail_nickname=nickname,
                display_name=group.display_name,
                kind=kind,
                action=action,
                reason=reason,
                description=group.description,
                target_member_upns=[_rewrite(m, config) for m in group.member_upns],
                target_owner_upns=[_rewrite(o, config) for o in group.owner_upns],
                is_dynamic=group.is_dynamic,
                membership_rule=group.membership_rule,
            )
        )
    return planned


def _directory_object_ref(client: GraphClient, object_id: str) -> str:
    """Build an @odata.id reference to a directory object on this client's base URL."""
    return f"{client.base_url}/directoryObjects/{object_id}"


def sync_groups(
    client: GraphClient,
    planned: list[PlannedGroup],
    config: Config,
    *,
    dry_run: bool = True,
) -> list[dict]:
    """Provision missing groups and reconcile membership + ownership in the target.

    For each ``create``/``exists`` plan entry this resolves (or creates) the
    target group, then adds the source members and owners that resolve to existing
    target users and are not already present. Dynamic groups are created with their
    membership rule and skip manual member assignment. With ``dry_run=True``
    (default) nothing is written; planned actions are reported with a ``would-``
    prefix.
    """
    target_groups = get_target_group_ids(client)
    target_users = get_target_user_ids(client)
    results: list[dict] = []

    for p in planned:
        record: dict = {"mail_nickname": p.mail_nickname, "kind": p.kind}

        if p.action == "skip":
            record["status"] = f"skipped:{p.reason}"
            results.append(record)
            continue

        # --- resolve or create the target group ---
        group_id = target_groups.get((p.mail_nickname or "").lower())
        existing_members: set[str] = set()
        if p.action == "create" and not group_id:
            if dry_run:
                record["group"] = "would-create"
            else:
                try:
                    created = client.post("/groups", json=p.to_graph_body())
                    group_id = created.get("id")
                    record["group"] = "created"
                except GraphError as exc:
                    record["group"] = f"error:{exc.status_code}"
                    record["status"] = "error"
                    results.append(record)
                    continue
        else:
            record["group"] = "exists"
            if group_id and not p.is_dynamic:
                existing_members = get_group_member_ids(client, group_id)

        # --- reconcile membership ---
        # Dynamic groups are populated by their membership rule (migrated in the
        # create body), so manual member assignment is skipped.
        if p.is_dynamic:
            record["members"] = "dynamic-rule"
        else:
            resolved, unresolved = [], 0
            for upn in p.target_member_upns:
                uid = target_users.get(upn.lower())
                if uid:
                    resolved.append(uid)
                else:
                    unresolved += 1

            to_add = [uid for uid in resolved if uid not in existing_members]
            if unresolved:
                record["members_unresolved"] = unresolved

            if not to_add:
                record["members"] = "none"
            elif dry_run or group_id is None:
                # group_id is None only in a dry run create (group not yet provisioned).
                record["members"] = f"would-add:{len(to_add)}"
            else:
                added, errors = 0, 0
                for uid in to_add:
                    try:
                        client.post(
                            f"/groups/{group_id}/members/$ref",
                            json={"@odata.id": _directory_object_ref(client, uid)},
                        )
                        added += 1
                    except GraphError:
                        errors += 1
                record["members"] = f"added:{added}"
                if errors:
                    record["members_errors"] = errors

        # --- reconcile owners (a group needs an owner before it can be teamified) ---
        if p.target_owner_upns:
            existing_owners = get_group_owner_ids(client, group_id) if group_id else set()
            resolved_owners, unresolved_owners = [], 0
            for upn in p.target_owner_upns:
                uid = target_users.get(upn.lower())
                if uid:
                    resolved_owners.append(uid)
                else:
                    unresolved_owners += 1

            to_add_owners = [uid for uid in resolved_owners if uid not in existing_owners]
            if unresolved_owners:
                record["owners_unresolved"] = unresolved_owners

            if not to_add_owners:
                record["owners"] = "none"
            elif dry_run or group_id is None:
                record["owners"] = f"would-add:{len(to_add_owners)}"
            else:
                added, errors = 0, 0
                for uid in to_add_owners:
                    try:
                        client.post(
                            f"/groups/{group_id}/owners/$ref",
                            json={"@odata.id": _directory_object_ref(client, uid)},
                        )
                        added += 1
                    except GraphError:
                        errors += 1
                record["owners"] = f"added:{added}"
                if errors:
                    record["owners_errors"] = errors

        record["status"] = "ok"
        results.append(record)

    return results
