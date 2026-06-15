"""Command-line interface for m365-migrate.

Examples
--------
    m365-migrate auth-check
    m365-migrate users discover
    m365-migrate users plan
    m365-migrate users migrate            # dry run (no writes)
    m365-migrate users migrate --execute  # actually create users
    m365-migrate groups plan              # plan group + membership reconciliation
    m365-migrate groups sync --execute    # create groups and add members
    m365-migrate mailbox plan             # plan mailbox-settings migration
    m365-migrate mailbox migrate --execute  # apply mailbox settings to target
    m365-migrate files plan --user jane@src.com         # plan a OneDrive copy
    m365-migrate files migrate --user jane@src.com --execute  # copy the drive
"""

from __future__ import annotations

import json
from pathlib import Path

import typer
from rich.console import Console
from rich.table import Table

from m365_migrate.auth import build_token_provider
from m365_migrate.config import Config, ConfigError, load_config
from m365_migrate.graph_client import GraphClient
from m365_migrate.mapping import write_mapping
from m365_migrate.workloads import files as files_workload
from m365_migrate.workloads import groups as groups_workload
from m365_migrate.workloads import mailboxes as mailboxes_workload
from m365_migrate.workloads import users as users_workload

app = typer.Typer(help="Microsoft 365 tenant-to-tenant migration tool.", no_args_is_help=True)
users_app = typer.Typer(help="Users / identities workload.", no_args_is_help=True)
app.add_typer(users_app, name="users")
groups_app = typer.Typer(help="Groups + membership workload.", no_args_is_help=True)
app.add_typer(groups_app, name="groups")
mailbox_app = typer.Typer(help="Exchange Online mailbox-settings workload.", no_args_is_help=True)
app.add_typer(mailbox_app, name="mailbox")
files_app = typer.Typer(help="OneDrive / SharePoint files workload.", no_args_is_help=True)
app.add_typer(files_app, name="files")

console = Console()
err_console = Console(stderr=True)

CONFIG_OPTION = typer.Option("config.yaml", "--config", "-c", help="Path to config file.")


def _load(config_path: str) -> Config:
    try:
        return load_config(config_path)
    except ConfigError as exc:
        err_console.print(f"[red]Config error:[/red] {exc}")
        raise typer.Exit(code=2)


def _client(config: Config, which: str) -> GraphClient:
    tenant = config.source if which == "source" else config.target
    return GraphClient(build_token_provider(tenant))


@app.command("auth-check")
def auth_check(config_path: str = CONFIG_OPTION) -> None:
    """Verify that both tenants authenticate and Graph is reachable."""
    config = _load(config_path)
    for which in ("source", "target"):
        try:
            with _client(config, which) as client:
                me = client.get("/organization")
                org = me.get("value", [{}])[0].get("displayName", "(unknown)")
            console.print(f"[green]OK[/green] {which}: authenticated to '{org}'")
        except Exception as exc:  # noqa: BLE001 - surface any auth/network failure
            err_console.print(f"[red]FAIL[/red] {which}: {exc}")
            raise typer.Exit(code=1)


@users_app.command("discover")
def users_discover(config_path: str = CONFIG_OPTION) -> None:
    """Read users from the source tenant and write them to the output dir."""
    config = _load(config_path)
    with _client(config, "source") as client:
        found = users_workload.discover_users(client)
    out = Path(config.options.output_dir) / "source_users.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps([u.model_dump() for u in found], indent=2))
    console.print(f"Discovered [bold]{len(found)}[/bold] users -> {out}")


@users_app.command("plan")
def users_plan(config_path: str = CONFIG_OPTION) -> None:
    """Produce a migration plan (read-only) and write the mapping CSV."""
    config = _load(config_path)
    with _client(config, "source") as source:
        found = users_workload.discover_users(source)
    with _client(config, "target") as target:
        existing = users_workload.discover_target_upns(target)

    planned = users_workload.plan_users(found, config, existing_target_upns=existing)
    mapping_path = write_mapping(Path(config.options.output_dir) / "user_mapping.csv", planned)
    _print_plan(planned)
    console.print(f"\nMapping written to [bold]{mapping_path}[/bold]")


@users_app.command("migrate")
def users_migrate(
    config_path: str = CONFIG_OPTION,
    execute: bool = typer.Option(
        False, "--execute", help="Actually create users. Without this flag it is a dry run."
    ),
) -> None:
    """Create planned users in the target tenant (dry run unless --execute)."""
    config = _load(config_path)
    with _client(config, "source") as source:
        found = users_workload.discover_users(source)
    with _client(config, "target") as target:
        existing = users_workload.discover_target_upns(target)
        planned = users_workload.plan_users(found, config, existing_target_upns=existing)
        results = users_workload.migrate_users(target, planned, config, dry_run=not execute)

    mode = "EXECUTE" if execute else "DRY RUN"
    console.print(f"[bold]{mode}[/bold] — {len(results)} users processed")
    counts: dict[str, int] = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    for status, count in sorted(counts.items()):
        console.print(f"  {status}: {count}")
    if not execute:
        console.print("\n[yellow]No changes were made.[/yellow] Re-run with --execute to apply.")


@users_app.command("enrich")
def users_enrich(
    config_path: str = CONFIG_OPTION,
    execute: bool = typer.Option(
        False, "--execute", help="Actually set managers/licenses. Without this flag it is a dry run."
    ),
) -> None:
    """Set manager links and assign licenses on migrated users (dry run unless --execute).

    Run this after `users migrate`; it operates on accounts that already exist
    in the target tenant.
    """
    config = _load(config_path)
    with _client(config, "source") as source:
        found = users_workload.discover_users(source)
    with _client(config, "target") as target:
        results = users_workload.enrich_users(target, found, config, dry_run=not execute)

    mode = "EXECUTE" if execute else "DRY RUN"
    console.print(f"[bold]{mode}[/bold] — {len(results)} users processed")
    counts: dict[str, int] = {}
    for r in results:
        for key in ("status", "manager", "licenses"):
            value = r.get(key)
            if value:
                counts[f"{key}:{value}"] = counts.get(f"{key}:{value}", 0) + 1
    for label, count in sorted(counts.items()):
        console.print(f"  {label}: {count}")
    if not execute:
        console.print("\n[yellow]No changes were made.[/yellow] Re-run with --execute to apply.")


@groups_app.command("discover")
def groups_discover(config_path: str = CONFIG_OPTION) -> None:
    """Read groups (with members) from the source tenant to the output dir."""
    config = _load(config_path)
    with _client(config, "source") as client:
        found = groups_workload.discover_groups(client)
    out = Path(config.options.output_dir) / "source_groups.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps([g.model_dump() for g in found], indent=2))
    console.print(f"Discovered [bold]{len(found)}[/bold] groups -> {out}")


@groups_app.command("plan")
def groups_plan(config_path: str = CONFIG_OPTION) -> None:
    """Produce a group reconciliation plan (read-only)."""
    config = _load(config_path)
    with _client(config, "source") as source:
        found = groups_workload.discover_groups(source)
    with _client(config, "target") as target:
        existing = groups_workload.get_target_group_ids(target)
    planned = groups_workload.plan_groups(found, config, existing_target_groups=existing)
    _print_group_plan(planned)


@groups_app.command("sync")
def groups_sync(
    config_path: str = CONFIG_OPTION,
    execute: bool = typer.Option(
        False, "--execute", help="Actually create groups/add members. Without this flag it is a dry run."
    ),
) -> None:
    """Create missing groups and reconcile membership (dry run unless --execute).

    Run this after `users migrate` so member UPNs resolve to existing target
    accounts.
    """
    config = _load(config_path)
    with _client(config, "source") as source:
        found = groups_workload.discover_groups(source)
    with _client(config, "target") as target:
        existing = groups_workload.get_target_group_ids(target)
        planned = groups_workload.plan_groups(found, config, existing_target_groups=existing)
        results = groups_workload.sync_groups(target, planned, config, dry_run=not execute)

    mode = "EXECUTE" if execute else "DRY RUN"
    console.print(f"[bold]{mode}[/bold] — {len(results)} groups processed")
    counts: dict[str, int] = {}
    for r in results:
        for key in ("status", "group", "members"):
            value = r.get(key)
            if value:
                counts[f"{key}:{value}"] = counts.get(f"{key}:{value}", 0) + 1
    for label, count in sorted(counts.items()):
        console.print(f"  {label}: {count}")
    if not execute:
        console.print("\n[yellow]No changes were made.[/yellow] Re-run with --execute to apply.")


@mailbox_app.command("discover")
def mailbox_discover(config_path: str = CONFIG_OPTION) -> None:
    """Read mailbox settings for mailbox-enabled source users to the output dir."""
    config = _load(config_path)
    with _client(config, "source") as client:
        found = mailboxes_workload.discover_mailboxes(client)
    out = Path(config.options.output_dir) / "source_mailboxes.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps([m.model_dump() for m in found], indent=2))
    console.print(f"Discovered [bold]{len(found)}[/bold] mailboxes -> {out}")


@mailbox_app.command("plan")
def mailbox_plan(config_path: str = CONFIG_OPTION) -> None:
    """Produce a mailbox-settings migration plan (read-only)."""
    config = _load(config_path)
    with _client(config, "source") as source:
        found = mailboxes_workload.discover_mailboxes(source)
    with _client(config, "target") as target:
        existing = users_workload.discover_target_upns(target)
    planned = mailboxes_workload.plan_mailboxes(found, config, target_upns=existing)
    _print_mailbox_plan(planned)


@mailbox_app.command("migrate")
def mailbox_migrate(
    config_path: str = CONFIG_OPTION,
    execute: bool = typer.Option(
        False, "--execute", help="Actually apply settings. Without this flag it is a dry run."
    ),
) -> None:
    """Apply mailbox settings to the target tenant (dry run unless --execute).

    Run this after `users migrate` so target mailboxes exist to receive settings.
    """
    config = _load(config_path)
    with _client(config, "source") as source:
        found = mailboxes_workload.discover_mailboxes(source)
    with _client(config, "target") as target:
        existing = users_workload.discover_target_upns(target)
        planned = mailboxes_workload.plan_mailboxes(found, config, target_upns=existing)
        results = mailboxes_workload.migrate_mailboxes(target, planned, config, dry_run=not execute)

    mode = "EXECUTE" if execute else "DRY RUN"
    console.print(f"[bold]{mode}[/bold] — {len(results)} mailboxes processed")
    counts: dict[str, int] = {}
    for r in results:
        counts[r["status"]] = counts.get(r["status"], 0) + 1
    for status, count in sorted(counts.items()):
        console.print(f"  {status}: {count}")
    if not execute:
        console.print("\n[yellow]No changes were made.[/yellow] Re-run with --execute to apply.")


USER_OPTION = typer.Option(None, "--user", help="Source user UPN (migrate their OneDrive).")
SITE_OPTION = typer.Option(None, "--site", help="Source SharePoint site id.")
TARGET_SITE_OPTION = typer.Option(None, "--target-site", help="Target SharePoint site id (with --site).")


def _drive_roots(config: Config, user: str | None, site: str | None, target_site: str | None) -> tuple[str, str]:
    try:
        return files_workload.resolve_drive_roots(config, user=user, site=site, target_site=target_site)
    except ValueError as exc:
        err_console.print(f"[red]Error:[/red] {exc}")
        raise typer.Exit(code=2)


@files_app.command("discover")
def files_discover(
    config_path: str = CONFIG_OPTION,
    user: str = USER_OPTION,
    site: str = SITE_OPTION,
    target_site: str = TARGET_SITE_OPTION,
) -> None:
    """Walk a source drive (OneDrive user or SharePoint site) to the output dir."""
    config = _load(config_path)
    source_root, _ = _drive_roots(config, user, site, target_site)
    with _client(config, "source") as client:
        found = files_workload.discover_drive_items(client, source_root)
    label = user or site or "drive"
    out = Path(config.options.output_dir) / f"source_files_{label.replace('@', '_at_')}.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps([i.model_dump() for i in found], indent=2))
    console.print(f"Discovered [bold]{len(found)}[/bold] items -> {out}")


@files_app.command("plan")
def files_plan(
    config_path: str = CONFIG_OPTION,
    user: str = USER_OPTION,
    site: str = SITE_OPTION,
    target_site: str = TARGET_SITE_OPTION,
) -> None:
    """Produce a drive copy plan (read-only)."""
    config = _load(config_path)
    source_root, _ = _drive_roots(config, user, site, target_site)
    with _client(config, "source") as source:
        found = files_workload.discover_drive_items(source, source_root)
    planned = files_workload.plan_drive_items(found, config)
    _print_files_plan(planned)


@files_app.command("migrate")
def files_migrate(
    config_path: str = CONFIG_OPTION,
    user: str = USER_OPTION,
    site: str = SITE_OPTION,
    target_site: str = TARGET_SITE_OPTION,
    execute: bool = typer.Option(
        False, "--execute", help="Actually copy files. Without this flag it is a dry run."
    ),
) -> None:
    """Copy a source drive to the target (dry run unless --execute)."""
    config = _load(config_path)
    source_root, target_root = _drive_roots(config, user, site, target_site)
    with _client(config, "source") as source:
        found = files_workload.discover_drive_items(source, source_root)
        planned = files_workload.plan_drive_items(found, config)
        with _client(config, "target") as target:
            results = files_workload.migrate_drive_items(
                source, target, planned, source_root, target_root, config, dry_run=not execute
            )

    mode = "EXECUTE" if execute else "DRY RUN"
    console.print(f"[bold]{mode}[/bold] — {len(results)} items processed")
    counts: dict[str, int] = {}
    for r in results:
        for key in ("status", "grants"):
            value = r.get(key)
            if value:
                counts[f"{key}:{value}"] = counts.get(f"{key}:{value}", 0) + 1
    for label, count in sorted(counts.items()):
        console.print(f"  {label}: {count}")
    if not execute:
        console.print("\n[yellow]No changes were made.[/yellow] Re-run with --execute to apply.")


def _print_files_plan(planned: list) -> None:
    table = Table(title="Drive copy plan")
    table.add_column("Path")
    table.add_column("Type")
    table.add_column("Size")
    table.add_column("Action")
    table.add_column("Grants")
    table.add_column("Reason")
    for p in planned:
        color = {"copy": "green", "skip": "yellow"}.get(p.action, "white")
        table.add_row(
            p.relative_path,
            "folder" if p.is_folder else "file",
            str(p.size),
            f"[{color}]{p.action}[/{color}]",
            str(len(p.target_grants)),
            p.reason or "",
        )
    console.print(table)


def _print_mailbox_plan(planned: list) -> None:
    table = Table(title="Mailbox-settings migration plan")
    table.add_column("Source UPN")
    table.add_column("Target UPN")
    table.add_column("Action")
    table.add_column("Settings")
    table.add_column("Reason")
    for p in planned:
        color = {"settings": "green", "skip": "yellow"}.get(p.action, "white")
        table.add_row(
            p.source_upn,
            p.target_upn,
            f"[{color}]{p.action}[/{color}]",
            str(len(p.settings)),
            p.reason or "",
        )
    console.print(table)


def _print_group_plan(planned: list) -> None:
    table = Table(title="Group reconciliation plan")
    table.add_column("mailNickname")
    table.add_column("Display name")
    table.add_column("Kind")
    table.add_column("Action")
    table.add_column("Members")
    table.add_column("Reason")
    for p in planned:
        color = {"create": "green", "exists": "cyan", "skip": "yellow"}.get(p.action, "white")
        table.add_row(
            p.mail_nickname or "",
            p.display_name or "",
            p.kind,
            f"[{color}]{p.action}[/{color}]",
            str(len(p.target_member_upns)),
            p.reason or "",
        )
    console.print(table)


def _print_plan(planned: list) -> None:
    table = Table(title="User migration plan")
    table.add_column("Source UPN")
    table.add_column("Target UPN")
    table.add_column("Action")
    table.add_column("Reason")
    for p in planned:
        color = {"create": "green", "skip": "yellow", "conflict": "red"}.get(p.action, "white")
        table.add_row(p.source_upn, p.target_upn, f"[{color}]{p.action}[/{color}]", p.reason or "")
    console.print(table)


if __name__ == "__main__":  # pragma: no cover
    app()
