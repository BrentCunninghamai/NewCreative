"""Users / Identities workload: discover, plan, and migrate user objects.

The flow is deliberately three-staged so a human can inspect the plan before
anything is written to the target tenant:

    discover  -> read users from the source tenant
    plan      -> decide target UPNs, filter, detect conflicts (read-only)
    migrate   -> create the planned users in the target tenant
"""

from __future__ import annotations

import secrets
import string

from m365_migrate.config import Config
from m365_migrate.graph_client import GraphClient, GraphError
from m365_migrate.mapping import rewrite_upn
from m365_migrate.models import USER_SELECT_FIELDS, PlannedUser, SourceUser


def discover_users(client: GraphClient) -> list[SourceUser]:
    """Read all users from the source tenant."""
    select = ",".join(USER_SELECT_FIELDS)
    raw = client.get_all("/users", params={"$select": select, "$top": 999})
    return [SourceUser.from_graph(item) for item in raw]


def _target_upn(user: SourceUser, config: Config) -> str:
    if config.options.rewrite_upn_domain:
        return rewrite_upn(
            user.user_principal_name,
            config.source.primary_domain,
            config.target.primary_domain,
        )
    return user.user_principal_name


def plan_users(
    users: list[SourceUser],
    config: Config,
    existing_target_upns: set[str] | None = None,
) -> list[PlannedUser]:
    """Build a migration plan without writing anything.

    ``existing_target_upns`` (lowercased) lets the planner flag users that
    already exist in the target tenant as ``conflict`` rather than ``create``.
    """
    existing = {u.lower() for u in (existing_target_upns or set())}
    planned: list[PlannedUser] = []
    for user in users:
        target_upn = _target_upn(user, config)

        if config.options.skip_guests and user.user_type.lower() == "guest":
            action, reason = "skip", "guest user"
        elif target_upn.lower() in existing:
            action, reason = "conflict", "target UPN already exists"
        else:
            action, reason = "create", None

        planned.append(
            PlannedUser(
                source_id=user.id,
                source_upn=user.user_principal_name,
                target_upn=target_upn,
                display_name=user.display_name,
                action=action,
                reason=reason,
            )
        )
    return planned


def discover_target_upns(client: GraphClient) -> set[str]:
    """Read existing userPrincipalNames from the target tenant for conflict checks."""
    raw = client.get_all("/users", params={"$select": "userPrincipalName", "$top": 999})
    return {item["userPrincipalName"] for item in raw if item.get("userPrincipalName")}


def _generate_password(length: int = 20) -> str:
    """Generate a strong random initial password (user must reset on first sign-in)."""
    alphabet = string.ascii_letters + string.digits + "!@#$%^&*-_"
    # Ensure complexity requirements are comfortably met.
    return "".join(secrets.choice(alphabet) for _ in range(length))


def migrate_users(
    client: GraphClient,
    planned: list[PlannedUser],
    config: Config,
    *,
    dry_run: bool = True,
) -> list[dict]:
    """Create the planned users in the target tenant.

    Returns a result record per planned user. With ``dry_run=True`` (the
    default) nothing is written — each creatable user is reported as ``would-create``.
    """
    results: list[dict] = []
    for p in planned:
        if p.action != "create":
            results.append({"target_upn": p.target_upn, "status": f"skipped:{p.action}", "reason": p.reason})
            continue

        if dry_run:
            results.append({"target_upn": p.target_upn, "status": "would-create"})
            continue

        body = p.to_graph_body(default_usage_location=None)
        body["passwordProfile"]["password"] = _generate_password()
        try:
            created = client.post("/users", json=body)
            results.append(
                {"target_upn": p.target_upn, "status": "created", "target_id": created.get("id")}
            )
        except GraphError as exc:
            results.append({"target_upn": p.target_upn, "status": "error", "reason": str(exc)})
    return results
