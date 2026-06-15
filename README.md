# NewCreative — m365-migrate

A Microsoft 365 **tenant-to-tenant (T2T) migration** tool.

This is an early, working foundation. The first workload — **Users / Identities** —
is implemented end-to-end (discover → plan → migrate) with a safe dry-run
default. Other workloads (Exchange mailboxes, OneDrive/SharePoint, Teams) are
planned and will build on the source→target user mapping this workload produces.

## Why users first?

A T2T migration is several distinct jobs (identities, mail, files, Teams), each
with its own Graph/Exchange APIs and throttling behavior. Every other workload
needs to know *which target user* each source object belongs to — so the user
mapping is the foundation everything else keys off of.

## Requirements

- Python 3.10+
- An Entra ID (Azure AD) **app registration in each tenant** with
  admin-consented application permissions:
  - Source app: `User.Read.All`
  - Target app: `User.ReadWrite.All`

## Install

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
```

## Configure

```bash
cp config.example.yaml config.yaml   # config.yaml is git-ignored
# edit config.yaml; export secrets referenced as ${ENV:...}
export SOURCE_CLIENT_SECRET=...       # if you use the ${ENV:...} syntax
export TARGET_CLIENT_SECRET=...
```

## Use

```bash
m365-migrate auth-check                 # verify both tenants authenticate
m365-migrate users discover             # dump source users to out/source_users.json
m365-migrate users plan                 # write out/user_mapping.csv + print plan
m365-migrate users migrate              # DRY RUN — shows what would happen
m365-migrate users migrate --execute    # actually create users in the target
```

`migrate` is a **dry run unless you pass `--execute`.** Created users get a
strong random password and are flagged to reset it on first sign-in.

## Develop

```bash
pytest                                  # full suite, no real tenants needed
```

Tests mock Microsoft Graph (via `respx`), so the suite runs offline.

## Layout

```
src/m365_migrate/
  config.py          # YAML config + ${ENV:...} secret resolution
  auth.py            # per-tenant client-credentials token providers
  graph_client.py    # Graph REST wrapper: paging + 429/5xx retry
  models.py          # SourceUser / PlannedUser domain models
  mapping.py         # UPN rewriting + mapping CSV read/write
  workloads/users.py # discover / plan / migrate the Users workload
  cli.py             # Typer CLI
tests/               # offline tests (Graph mocked)
docs/architecture.md # design + roadmap
```

## License

Public domain (The Unlicense). See `LICENSE`.
