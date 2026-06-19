# m365-migrate (.NET)

The packaged **Windows desktop application** for Microsoft 365 tenant-to-tenant
migration. This is the product direction: a ShareGate-style native app built on
.NET 8 / WPF, with a reusable cross-platform engine underneath.

> **New here? See [SETUP.md](SETUP.md)** for step-by-step Azure AD app
> registration, permissions/consent, downloading the exe, and running a migration.

> The original Python implementation (repo root, `src/m365_migrate/`) is the
> proven reference that this engine is ported from. It remains in the repo while
> the port completes, then will be retired.

## Layout

```
dotnet/
  src/
    M365Migrate.Core/   net8.0 class library — the engine (auth, Graph client,
                        models, workloads). Cross-platform; builds/tests on CI.
    M365Migrate.App/    net8.0-windows WPF desktop UI (Windows-only build).
  tests/
    M365Migrate.Core.Tests/  xUnit tests for the engine (offline; fake HttpClient).
```

## Engine status

Ported from the Python engine, staged discover → plan → migrate with a dry-run
default. **All workloads are now ported** (full parity with the Python engine):

- **User mapping (preview)** — scan both tenants and build a source→target identity
  map (CSV override → cross-tenant/`#EXT#` → mail/SMTP → UPN rewrite), exported as an
  editable CSV. Read-only; the backbone for bulk, ShareGate-style migrations.
- **Users** — discover, plan (conflict + guest filtering), migrate.
- **Groups** — discover, plan, sync membership + ownership; dynamic groups
  recreated with their membership rule.
- **Mailboxes** — discover, plan, migrate Exchange Online mailbox *settings*.
- **Mail (content)** — copy a user's mail folders + messages source→target,
  full-fidelity via MIME; idempotent (skips messages whose internetMessageId
  already exists in the target folder).
- **Calendar & Contacts (content)** — copy a user's calendar events and contacts;
  best-effort idempotent (skips events matching subject+start+end, contacts
  matching display name + primary email).
- **Bulk Mail / Bulk OneDrive (mapped users)** — build the user mapping, then run the
  content workload for **every matched user** in one pass, one result row per user,
  with live per-user progress. Dry run unless Execute; resumable (mail dedup / drive
  re-upload), and per-user errors (e.g. no OneDrive) don't stop the batch. **Auto-sync**
  re-runs the delta on an interval to pre-seed to ~100% before cutover. Bulk OneDrive
  **auto-handles multi-geo**: if `/users/{id}/drive` returns `notSupported`, it falls back
  to the user's OneDrive site URL (derived from the tenant OneDrive host + UPN) — no manual URLs.

There is also a companion CLI, **`M365Migrate.ProfileSwap.exe`** (ProfWiz-style local
profile takeover for cutover) — see [PROFILESWAP.md](PROFILESWAP.md).
- **Files** — OneDrive/SharePoint discover (BFS) → plan → copy: small files via
  simple upload, large files via resumable upload session; reapply direct grants.
- **SharePoint (site)** — resolve a source + target **site by URL**, match document
  libraries by name, and copy each library with the delta-aware files engine (also
  moves a Team's files, which live in its SharePoint site).
- **Teams** — discover teams + channels, plan, enable Teams on the migrated M365
  group, and recreate standard channels.
- **Teams (messages)** — import channel **message history** via Graph migration
  mode: create a fresh migration-mode team, recreate channels, import each message
  with original author + timestamp, completeMigration, then add the source team's
  **owners + members** (mapped to target accounts). Imports top-level messages **and
  their threaded replies**. Run once per team (not idempotent).

## Build & test

The cross-platform engine and its tests build anywhere .NET 8 is installed:

```bash
dotnet test dotnet/tests/M365Migrate.Core.Tests/M365Migrate.Core.Tests.csproj
```

CI (`.github/workflows/dotnet.yml`) runs this on every push/PR.

## Desktop app

`M365Migrate.App` is a single-window WPF UI: enter both tenants' app-registration
details (tenant id, client id, secret, primary domain), **Test connections** to
confirm auth + permissions, pick a workload, then **Discover & Plan** and
**Migrate** (dry run unless *Execute* is ticked). Results stream into a grid. A
**Cancel** button stops a long or mistaken run mid-flight. All workloads (Users,
Groups, Mailboxes, Mail content, Files, Teams) are wired up. Files and Mail use
the **Scope** field (a source user UPN).

### Prerequisites

Each tenant needs an **Azure AD app registration** (client id + secret) with
admin-consented Graph *application* permissions for the workloads you run, e.g.
`User.ReadWrite.All`, `Group.ReadWrite.All`, `MailboxSettings.ReadWrite`,
`Files.ReadWrite.All`, `Sites.ReadWrite.All`, `Team.Create`, `Channel.Create`.
Use **Test connections** first — it calls `/organization` on both tenants and
reports the org names, so credential/permission problems surface before a run.

### Reports & logs

Every plan and migrate writes a timestamped CSV to
`%LOCALAPPDATA%\m365-migrate\reports` (open it from the app with **Open reports
folder**), and the app appends actions/errors to
`%LOCALAPPDATA%\m365-migrate\logs\app-YYYYMMDD.log`. If something goes wrong,
that log is the fastest way to diagnose it.

## Packaging (Windows)

The app publishes to a single self-contained executable:

```bash
dotnet publish dotnet/src/M365Migrate.App -c Release -r win-x64 \
  --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

This produces a downloadable `M365Migrate.exe` that runs on a Windows machine with
no prerequisites. CI (`.github/workflows/dotnet.yml`) builds it on a
`windows-latest` runner and:

- uploads it as the **`m365-migrate-windows`** artifact on every push/PR, and
- on every push to `main`, attaches it to a rolling **`latest`** GitHub Release —
  so the newest build is always at `…/releases/latest` for a one-click download.

## Design notes

- **Engine vs. UI split.** All migration logic lives in `M365Migrate.Core` so it
  is unit-tested on Linux CI; WPF is a thin Windows-only shell over it.
- **Injectable auth + HttpClient.** `GraphClient` takes a `TokenProvider` and an
  `HttpClient`, so tests mock Graph with a fake handler and never touch the
  network — exactly as the Python engine injects its client/token provider.
- **Smart, not sneaky.** The tool maximizes what is achievable through supported
  Microsoft Graph mechanisms (throttling backoff, batching, delta) and guides the
  operator through legitimate setup (app registrations, admin consent, cross-tenant
  access). It does not attempt to bypass authentication, security, or licensing
  controls; genuinely native-only operations (mailbox content moves, distribution
  lists) are surfaced honestly rather than faked.
```
