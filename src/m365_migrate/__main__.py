"""Allow running the CLI via `python -m m365_migrate`."""

from m365_migrate.cli import app

if __name__ == "__main__":  # pragma: no cover
    app()
