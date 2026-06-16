"""Files workload: discover, plan, and copy OneDrive / SharePoint drive content.

OneDrive personal storage and SharePoint document libraries are both exposed by
Graph as *drives* of *driveItems*, so a single set of primitives covers both. As
with the other workloads the flow is staged:

    discover  -> walk a source drive: folders, files, and direct user grants
    plan      -> classify each item (copy | skip); rewrite grantee UPNs (read-only)
    migrate   -> recreate folders, upload file content, reapply resolvable grants

**Scope.** Content is copied by download → upload: small files in a single
``PUT .../content``, and files above ``SIMPLE_UPLOAD_LIMIT`` via a resumable
*upload session* (chunked), so large files are migrated rather than skipped. Each
file is buffered in memory between download and upload. Only **direct user**
permission grants are reapplied, and only when the grantee resolves to a target
account; sharing links and group/external grants are skipped. The high-fidelity
SharePoint Migration API (version history, full metadata) remains on the roadmap.
"""

from __future__ import annotations

from m365_migrate.config import Config
from m365_migrate.graph_client import GraphClient, GraphError
from m365_migrate.mapping import rewrite_upn
from m365_migrate.models import DriveGrant, DriveItem, PlannedDriveItem
from m365_migrate.workloads.users import get_target_user_ids

# Graph's simple-upload ceiling for ``PUT .../content`` is 250 MiB, but staying
# well under it keeps single requests reliable; larger files need upload sessions.
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024  # 4 MiB


def _rewrite(upn: str, config: Config) -> str:
    """Apply the configured UPN domain rewrite to a raw UPN string."""
    if config.options.rewrite_upn_domain:
        return rewrite_upn(upn, config.source.primary_domain, config.target.primary_domain)
    return upn


def resolve_drive_roots(
    config: Config,
    *,
    user: str | None = None,
    site: str | None = None,
    target_site: str | None = None,
) -> tuple[str, str]:
    """Return the (source, target) drive root paths for a OneDrive user or a site.

    For ``user`` (a source UPN) the target is that user's rewritten OneDrive. For
    ``site`` the target site id must be supplied explicitly, since sites do not
    map by UPN.
    """
    if user:
        return f"/users/{user}/drive", f"/users/{_rewrite(user, config)}/drive"
    if site:
        if not target_site:
            raise ValueError("--site requires a corresponding --target-site")
        return f"/sites/{site}/drive", f"/sites/{target_site}/drive"
    raise ValueError("specify either a user (OneDrive) or a site (SharePoint)")


def discover_drive_items(client: GraphClient, drive_root: str) -> list[DriveItem]:
    """Walk a drive breadth-first, returning folders before their children.

    The ordering guarantees a parent folder always appears before the items it
    contains, so a later copy can create folders before populating them.
    """
    items: list[DriveItem] = []
    queue: list[str] = ["root/children"]
    while queue:
        segment = queue.pop(0)
        for raw in client.get_all(
            f"{drive_root}/{segment}", params={"$expand": "permissions", "$top": 200}
        ):
            item = DriveItem.from_graph(raw)
            items.append(item)
            if item.is_folder:
                queue.append(f"items/{item.id}/children")
    return items


def plan_drive_items(
    items: list[DriveItem],
    config: Config,
) -> list[PlannedDriveItem]:
    """Build a copy plan without writing anything.

    Folders and files are all ``copy``; the migrate phase picks a simple upload or
    a chunked upload session based on each file's size. Grantee UPNs are rewritten
    to the target domain so the plan reflects what would be reapplied.
    """
    planned: list[PlannedDriveItem] = []
    for item in items:
        target_grants = [
            DriveGrant(upn=_rewrite(g.upn, config), roles=g.roles) for g in item.grants
        ]
        planned.append(
            PlannedDriveItem(
                source_id=item.id,
                name=item.name,
                relative_path=item.relative_path,
                is_folder=item.is_folder,
                size=item.size,
                action="copy",
                reason=None,
                target_grants=target_grants,
            )
        )
    return planned


def _parent_segment(relative_path: str) -> str:
    """Return the path-addressed children endpoint segment for an item's parent."""
    parent, _, _ = relative_path.rpartition("/")
    return f"root:/{parent}:/children" if parent else "root/children"


def _grant_role(roles: list[str]) -> str:
    """Collapse Graph permission roles to a single invite role (write or read)."""
    return "write" if any(r in ("write", "owner") for r in roles) else "read"


def migrate_drive_items(
    source: GraphClient,
    target: GraphClient,
    planned: list[PlannedDriveItem],
    source_root: str,
    target_root: str,
    config: Config,
    *,
    dry_run: bool = True,
) -> list[dict]:
    """Recreate folders, upload file content, and reapply resolvable grants.

    Reads from ``source`` and writes to ``target`` (different tenants / drives).
    With ``dry_run=True`` (default) nothing is written; planned actions are
    reported with a ``would-`` prefix.
    """
    target_users = get_target_user_ids(target)
    results: list[dict] = []

    for p in planned:
        record: dict = {"path": p.relative_path, "type": "folder" if p.is_folder else "file"}

        if p.action != "copy":
            record["status"] = f"skipped:{p.action}"
            record["reason"] = p.reason
            results.append(record)
            continue

        try:
            if p.is_folder:
                if dry_run:
                    record["status"] = "would-create-folder"
                else:
                    target.post(
                        f"{target_root}/{_parent_segment(p.relative_path)}",
                        json={
                            "name": p.name,
                            "folder": {},
                            "@microsoft.graph.conflictBehavior": "replace",
                        },
                    )
                    record["status"] = "folder-created"
            else:
                large = p.size > SIMPLE_UPLOAD_LIMIT
                if dry_run:
                    record["status"] = "would-upload-session" if large else "would-upload"
                else:
                    data = source.get_content(f"{source_root}/items/{p.source_id}/content")
                    if large:
                        target.upload_large_file(
                            f"{target_root}/root:/{p.relative_path}:/createUploadSession", data
                        )
                        record["status"] = "uploaded-session"
                    else:
                        target.put_content(
                            f"{target_root}/root:/{p.relative_path}:/content", data
                        )
                        record["status"] = "uploaded"
        except GraphError as exc:
            record["status"] = "error"
            record["reason"] = f"{exc.status_code}"
            results.append(record)
            continue

        # --- reapply direct user grants that resolve to a target account ---
        resolvable = [g for g in p.target_grants if g.upn.lower() in target_users]
        if resolvable:
            if dry_run:
                record["grants"] = f"would-grant:{len(resolvable)}"
            else:
                granted = 0
                for g in resolvable:
                    try:
                        target.post(
                            f"{target_root}/root:/{p.relative_path}:/invite",
                            json={
                                "recipients": [{"email": g.upn}],
                                "roles": [_grant_role(g.roles)],
                                "requireSignIn": True,
                                "sendInvitation": False,
                            },
                        )
                        granted += 1
                    except GraphError:
                        pass
                record["grants"] = f"granted:{granted}"

        results.append(record)

    return results
