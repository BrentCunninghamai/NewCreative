"""Command-line interface for m365-migrate.

Examples
--------
    m365-migrate auth-check
    m365-migrate users discover
    m365-migrate users plan
    m365-migrate users migrate            # dry run (no writes)
    m365-migrate users migrate --execute  # actually create users
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
from m365_migrate.workloads import users as users_workload

app = typer.Typer(help="Microsoft 365 tenant-to-tenant migration tool.", no_args_is_help=True)
users_app = typer.Typer(help="Users / identities workload.", no_args_is_help=True)
app.add_typer(users_app, name="users")

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
