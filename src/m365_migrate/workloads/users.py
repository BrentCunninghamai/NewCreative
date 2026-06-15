"""Users / Identities workload: discover, plan, migrate, and enrich user objects.

The flow is deliberately staged so a human can inspect the plan before anything
is written to the target tenant:

    discover  -> read users from the source tenant (with manager + licenses)
    plan      -> decide target UPNs, filter, detect conflicts (read-only)
    migrate   -> create the planned users in the target tenant
    enrich    -> set manager relationships and assign licenses on target users

``enrich`` is a separate step because it depends on the target accounts already
existing (created by ``migrate``) and on the source->target user mapping.
"""

from __future__ import annotations

import secrets
import string

from m365_migrate.config import Config
from m365_migrate.graph_client import GraphClient, GraphError
from m365_migrate.mapping import rewrite_upn
from m365_migrate.models import USER_EXPAND, USER_SELECT_FIELDS, PlannedUser, SourceUser


def discover_users(client: GraphClient) -> list[SourceUser]:
    """Read all users from the source tenant, including manager and licenses."""
    select = ",".join(USER_SELECT_FIELDS)
    raw = client.get_all(
        "/users", params={"$select": select, "$expand": USER_EXPAND, "$top": 999}
    )
    return [SourceUser.from_graph(item) for item in raw]


def _rewrite(upn: str, config: Config) -> str:
    """Apply the configured UPN domain rewrite to a raw UPN string."""
    if config.options.rewrite_upn_domain:
        return rewrite_upn(upn, config.source.primary_domain, config.target.primary_domain)
    return upn


def _target_upn(user: SourceUser, config: Config) -> str:
    return _rewrite(user.user_principal_name, config)


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


def get_target_user_ids(client: GraphClient) -> dict[str, str]:
    """Return a mapping of lowercased target UPN -> target user id."""
    raw = client.get_all(
        "/users", params={"$select": "id,userPrincipalName", "$top": 999}
    )
    return {
        item["userPrincipalName"].lower(): item["id"]
        for item in raw
        if item.get("userPrincipalName") and item.get("id")
    }


def get_target_sku_ids(client: GraphClient) -> set[str]:
    """Return the set of license SKU IDs available in the target tenant."""
    raw = client.get_all("/subscribedSkus")
    return {item["skuId"] for item in raw if item.get("skuId")}


def _graph_user_ref(client: GraphClient, user_id: str) -> str:
    """Build an @odata.id reference to a user resource on this client's base URL."""
    return f"{client.base_url}/users/{user_id}"


def enrich_users(
    client: GraphClient,
    users: list[SourceUser],
    config: Config,
    *,
    dry_run: bool = True,
) -> list[dict]:
    """Set manager relationships and assign licenses on already-migrated users.

    Resolves each source user (and their manager) to the corresponding target
    account via the configured UPN rewrite, then:

    - sets the manager reference when both user and manager exist in the target;
    - assigns the source user's license SKUs that are available in the target.

    With ``dry_run=True`` (default) nothing is written; planned actions are
    reported with a ``would-`` prefix.
    """
    target_ids = get_target_user_ids(client)
    available_skus = get_target_sku_ids(client)
    results: list[dict] = []

    for user in users:
        target_upn = _target_upn(user, config)
        target_id = target_ids.get(target_upn.lower())
        record: dict = {"target_upn": target_upn}

        if not target_id:
            record["status"] = "skipped:not-in-target"
            results.append(record)
            continue

        # --- manager ---
        if user.manager_upn:
            mgr_target_upn = _rewrite(user.manager_upn, config)
            mgr_id = target_ids.get(mgr_target_upn.lower())
            if not mgr_id:
                record["manager"] = "skipped:manager-not-in-target"
            elif dry_run:
                record["manager"] = "would-set"
            else:
                try:
                    client.put(
                        f"/users/{target_id}/manager/$ref",
                        json={"@odata.id": _graph_user_ref(client, mgr_id)},
                    )
                    record["manager"] = "set"
                except GraphError as exc:
                    record["manager"] = f"error:{exc.status_code}"

        # --- licenses ---
        skus = sorted(s for s in user.assigned_sku_ids if s in available_skus)
        if skus:
            if dry_run:
                record["licenses"] = f"would-assign:{len(skus)}"
            else:
                try:
                    client.post(
                        f"/users/{target_id}/assignLicense",
                        json={
                            "addLicenses": [{"skuId": s} for s in skus],
                            "removeLicenses": [],
                        },
                    )
                    record["licenses"] = f"assigned:{len(skus)}"
                except GraphError as exc:
                    record["licenses"] = f"error:{exc.status_code}"

        record["status"] = "ok"
        results.append(record)

    return results
