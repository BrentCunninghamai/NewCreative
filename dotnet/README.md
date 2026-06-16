# m365-migrate (.NET)

The packaged **Windows desktop application** for Microsoft 365 tenant-to-tenant
migration. This is the product direction: a ShareGate-style native app built on
.NET 8 / WPF, with a reusable cross-platform engine underneath.

> The original Python implementation (repo root, `src/m365_migrate/`) is the
> proven reference that this engine is ported from. It remains in the repo while
> the port completes, then will be retired.

## Layout

```
dotnet/
  src/
    M365Migrate.Core/   net8.0 class library — the engine (auth, Graph client,
                        models, workloads). Cross-platform; builds/tests on CI.
    M365Migrate.App/    net8.0-windows WPF desktop UI (Windows-only build). [stage 2]
  tests/
    M365Migrate.Core.Tests/  xUnit tests for the engine (offline; fake HttpClient).
```

## Engine status

Ported from the Python engine, staged discover → plan → migrate with a dry-run
default:

- **Users** — discover, plan (conflict + guest filtering), migrate.
- **Groups** — discover, plan, sync membership + ownership; dynamic groups
  recreated with their membership rule.

Remaining workloads (mailbox settings, OneDrive/SharePoint files incl. large-file
upload sessions, Teams) are ported next, mirroring the Python originals.

## Build & test

The cross-platform engine and its tests build anywhere .NET 8 is installed:

```bash
dotnet test dotnet/tests/M365Migrate.Core.Tests/M365Migrate.Core.Tests.csproj
```

CI (`.github/workflows/dotnet.yml`) runs this on every push/PR.

## Packaging (Windows)

The WPF app (stage 2) publishes to a single self-contained executable:

```bash
dotnet publish dotnet/src/M365Migrate.App -c Release -r win-x64 \
  --self-contained -p:PublishSingleFile=true
```

This produces a downloadable `.exe` that runs on a Windows machine with no
prerequisites. CI will build it on a `windows-latest` runner and attach it as an
artifact.

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
