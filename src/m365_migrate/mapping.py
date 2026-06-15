"""UPN rewriting and source→target user mapping persistence.

The mapping is the durable record that links a source user to the user created
in the target tenant. It is written as CSV so it can be reviewed in a
spreadsheet and re-used by later workloads (mailbox, OneDrive, etc.).
"""

from __future__ import annotations

import csv
from pathlib import Path

from m365_migrate.models import PlannedUser


def rewrite_upn(upn: str, source_domain: str, target_domain: str) -> str:
    """Rewrite the domain portion of a UPN from source to target.

    Only rewrites when the UPN's domain matches ``source_domain`` (case
    -insensitive). UPNs on other (e.g. already-vanity) domains are left intact.
    """
    local, _, domain = upn.partition("@")
    if not domain:
        return upn
    if domain.lower() == source_domain.lower():
        return f"{local}@{target_domain}"
    return upn


def write_mapping(path: str | Path, planned: list[PlannedUser]) -> Path:
    """Write the planned users to a CSV mapping file and return its path."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(
            ["source_id", "source_upn", "target_upn", "display_name", "action", "reason"]
        )
        for p in planned:
            writer.writerow(
                [p.source_id, p.source_upn, p.target_upn, p.display_name or "", p.action, p.reason or ""]
            )
    return path


def read_mapping(path: str | Path) -> list[PlannedUser]:
    """Read a CSV mapping file back into :class:`PlannedUser` records."""
    path = Path(path)
    rows: list[PlannedUser] = []
    with path.open(newline="", encoding="utf-8") as fh:
        for row in csv.DictReader(fh):
            rows.append(
                PlannedUser(
                    source_id=row["source_id"],
                    source_upn=row["source_upn"],
                    target_upn=row["target_upn"],
                    display_name=row.get("display_name") or None,
                    action=row["action"],
                    reason=row.get("reason") or None,
                )
            )
    return rows
