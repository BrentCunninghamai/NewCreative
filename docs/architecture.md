# Architecture & roadmap

## Goal

A tool to migrate one Microsoft 365 tenant's contents into another
(tenant-to-tenant / T2T): identities first, then mail, files, and Teams.

## Design principles

1. **Workloads are independent modules.** Each (users, mailbox, files, Teams)
   lives under `src/m365_migrate/workloads/` and shares the auth + Graph client
   layers. They can run independently and in a sensible order.
2. **Plan before write.** Every workload separates a read-only *plan* phase from
   a *migrate* phase, and `migrate` is a dry run unless explicitly executed.
   This makes runs reviewable and auditable.
3. **The user mapping is the spine.** `out/user_mapping.csv` links each source
   user to its target user. Later workloads (mailbox, OneDrive) resolve owners
   through this mapping rather than re-deriving them.
4. **Credentials stay out of the repo.** Config is YAML; secrets are pulled from
   environment variables via `${ENV:NAME}`. `config.yaml` is git-ignored.
5. **Offline-testable.** The Graph client takes an injectable HTTP client, so
   the whole suite runs against mocked Graph responses (no real tenants).

## Components

| Module | Responsibility |
| --- | --- |
| `config.py` | Load/validate YAML config; resolve `${ENV:...}` secrets. |
| `auth.py` | Per-tenant client-credentials token providers (`azure-identity`). |
| `graph_client.py` | Graph REST wrapper: bearer auth, `@odata.nextLink` paging, retry on 429/5xx with `Retry-After`. |
| `models.py` | `SourceUser` / `PlannedUser` domain models + Graph (de)serialization. |
| `mapping.py` | UPN domain rewriting; mapping CSV read/write. |
| `workloads/users.py` | `discover_users` → `plan_users` → `migrate_users`. |
| `cli.py` | Typer CLI surface. |

## Authentication

App-only (OAuth2 client credentials) per tenant. Each tenant has its own Entra
ID app registration with admin-consented application permissions. The source app
needs read access; the target app needs write access.

## Users workload flow

```
discover  GET /users (source, $expand=manager) -> SourceUser[] (incl. licenses)
plan      rewrite UPN domain, filter guests,   -> PlannedUser[] (create|skip|conflict)
          diff against target /users
migrate   POST /users (target) for "create"    -> results (dry-run by default)
enrich    resolve target ids via UPN mapping,  -> results (dry-run by default)
          PUT manager/$ref + POST assignLicense
```

Conflicts (target UPN already exists) are surfaced, never silently overwritten.
``enrich`` runs after ``migrate``: it sets each user's manager (when both the
user and manager exist in the target) and assigns the source user's license
SKUs that are available in the target tenant.

## Roadmap

- [x] Users / Identities: discover, plan (with conflict detection), migrate.
- [x] Users: license assignment + manager links (`enrich`).
- [ ] Users: group membership.
- [ ] Exchange Online mailboxes (mail, calendar, contacts).
- [ ] OneDrive / SharePoint (files, libraries, permissions).
- [ ] Teams (teams, channels, membership, files).
- [ ] Resumable runs + structured run logs / reporting.
- [ ] Concurrency with per-tenant throttling budgets.

## Known limitations (today)

- Only the Users workload exists; it provisions accounts, manager links, and
  licenses, but not yet group membership.
- License assignment requires the matching SKU to exist in the target tenant;
  unavailable SKUs are skipped (not purchased automatically).
- New users get a random password and must reset on first sign-in; there is no
  password/identity federation handoff.
- No incremental/delta sync yet — `plan` is a full comparison each run.
