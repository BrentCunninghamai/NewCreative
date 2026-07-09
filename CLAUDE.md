# CLAUDE.md

Guidance for AI assistants (Claude Code and others) working in this repository.

## What this project is

**NewCreative / m365-migrate** is a Microsoft 365 **tenant-to-tenant (T2T)
migration** tool: identities, groups, mail, OneDrive/SharePoint files, and
Teams from one M365 tenant into another, driven by Microsoft Graph.

**The product is a packaged Windows desktop app** (ShareGate-style UI) built on
**.NET 8 / WPF**, under `dotnet/`. The original **Python package**
(`src/m365_migrate/`) is the proven reference the engine was ported from — the
port is **complete (full workload parity)** — and it stays in the repo as the
behavioral spec until it is retired. **New product work goes in the .NET
solution;** consult the Python code when porting or checking intended behavior.

## Repository layout

```
dotnet/                      .NET solution — THE PRODUCT
  src/M365Migrate.Core/      net8.0 class library: the engine (auth, Graph
                             client, models, workloads). Cross-platform.
  src/M365Migrate.App/       net8.0-windows WPF desktop UI (single window,
                             MVVM: ViewModels/MainViewModel.cs drives it).
  src/M365Migrate.ProfileSwap/  Companion CLI: ProfWiz-style local Windows
                             profile takeover for cutover (elevated, HKLM).
  tests/M365Migrate.Core.Tests/  xUnit tests for the engine (offline; fake
                             HttpMessageHandler — never hits the network).
  README.md                  Engine/app status, build & packaging details.
  SETUP.md                   End-user guide: app registrations, consent, usage.
  SIGNING.md                 Azure Trusted Signing / SmartScreen notes.
  PROFILESWAP.md             Profile swapper guide.
  Directory.Build.props      Shared props: latest C#, Nullable + ImplicitUsings.
src/m365_migrate/            Python reference implementation (Typer CLI).
tests/                       Python tests (pytest; Graph mocked via respx).
docs/architecture.md         Design + workload flows (written for the Python
                             engine; the .NET engine mirrors it).
config.example.yaml          Example config for the Python CLI.
.github/workflows/ci.yml     Python CI: pytest on 3.10/3.11/3.12.
.github/workflows/dotnet.yml .NET CI: engine tests on Linux; Windows job
                             publishes single-file exes, optionally signs them,
                             and updates the rolling "latest" GitHub Release.
.claude/                     Session hooks (see "Dev environment" below).
```

## Commands

### .NET (primary)

```bash
# Build + run the engine test suite (works on Linux/macOS/Windows):
dotnet test dotnet/tests/M365Migrate.Core.Tests/M365Migrate.Core.Tests.csproj

# Publish the Windows app (Windows only — WPF targets net8.0-windows):
dotnet publish dotnet/src/M365Migrate.App -c Release -r win-x64 \
  --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

The WPF app and ProfileSwap **cannot be built on Linux**; CI's `windows-latest`
job does that. On Linux, validate changes through the Core test suite.

### Python (reference)

```bash
pip install -e ".[dev]"       # install (the SessionStart hook does this into .venv)
pytest                        # full suite, offline, no real tenants needed
m365-migrate --help           # Typer CLI (also: python -m m365_migrate)
```

## Workloads (engine capabilities)

All staged **discover → plan → migrate**, dry-run by default (see conventions).
Details in `dotnet/README.md`:

- **User mapping (preview)** — source→target identity map (CSV override →
  same-UPN native → cross-tenant/`#EXT#` → mail/SMTP alias → UPN rewrite),
  exported as editable CSV; flags guest-only targets; `ContentReady` gates bulk runs.
- **Users** — discover/plan/migrate; license assignment by SKU part number.
- **Groups** — membership + owners; dynamic groups recreated with their rule.
- **Mailboxes** — Exchange Online mailbox *settings*.
- **Mail (content)** — folders + messages via MIME; idempotent by internetMessageId.
- **Calendar & Contacts (content)** — best-effort idempotent copy.
- **Files** — OneDrive/SharePoint BFS discover → copy; large files via upload
  sessions; multi-geo OneDrive fallback by site URL.
- **SharePoint (site)** — site-by-URL, library-matched, delta-aware copy.
- **Teams** — enable Teams + recreate channels; **Teams messages** imported via
  Graph migration mode with original author/timestamp incl. threaded replies
  (run once per team — not idempotent).
- **Bulk Mail / Bulk OneDrive** — run a content workload for every mapped
  content-ready user; resumable; auto-sync delta re-runs before cutover.

## Conventions specific to this project

- **Plan before write.** Every workload separates a read-only plan phase from
  the migrate phase; `migrate` is a **dry run unless `--execute`** (CLI) or the
  *Execute* checkbox (app) is set. Preserve this in anything you add.
- **Never commit credentials.** Real config lives in git-ignored `config.yaml`;
  secrets resolve from env via `${ENV:NAME}`. Only `config.example.yaml` is
  committed. The app encrypts saved profile secrets with DPAPI.
- **Keep tests offline.** Inject the HTTP client / token provider so tests mock
  Graph (`respx` in Python, `FakeHttpMessageHandler` in .NET) rather than hit
  the network. New engine behavior ships with tests.
- **Engine vs. UI split.** All migration logic belongs in `M365Migrate.Core`
  (unit-testable on Linux CI); WPF stays a thin shell. Don't put Graph calls in
  the ViewModel.
- **Smart, not sneaky.** Use supported Graph mechanisms (throttling backoff,
  batching, delta, migration mode). Never bypass authentication, security, or
  licensing controls; surface genuinely native-only gaps honestly.
- **Update the docs with the change.** New workload or structural change →
  update this file, `dotnet/README.md` (and `SETUP.md` if user-facing) in the
  same PR so they stay an accurate map.

## Dev environment (Claude Code sessions)

`.claude/settings.json` wires two hooks:

- **SessionStart** runs `.claude/setup.sh`: creates `.venv` and installs the
  Python package with dev deps (idempotent).
- **PostToolUse** on Write/Edit of any `*.py` file runs the Python test suite
  and blocks on failure — expect fast feedback when touching Python code.

There is no equivalent hook for .NET; run `dotnet test` yourself after engine
changes.

## License: public domain

This project is released under The Unlicense. Practical implications:

- Only add code/content that is public-domain-compatible. Do **not** paste in
  code under restrictive licenses (GPL, proprietary, etc.) or content imposing
  attribution/copyleft obligations.
- Prefer permissively licensed dependencies (MIT, BSD, Apache-2.0, ISC,
  Unlicense, CC0) and note any constraints a dependency carries.

## Git workflow & conventions

- **Default branch:** `main`.
- **Do not commit directly to `main`.** Develop on a feature branch and open a
  pull request. When a session specifies a designated working branch, use
  exactly that branch.
- **Pushing:** `git push -u origin <branch-name>`. After pushing, open a PR
  (ready for review, not a draft) if one doesn't already exist for the branch.
- **Commit messages:** imperative mood, concise subject, explain the *why* in
  the body when it isn't obvious.
- **Keep changes focused.** One logical change per PR where practical.
- **CI must pass:** both workflows run on every PR (Python tests ×3 versions;
  .NET engine tests + Windows publish).

## Working agreement for AI assistants

- **Verify, don't assume.** Check the actual files on disk before acting on
  assumptions about structure or tooling — the project moves fast and this
  file may lag.
- **Be honest about state.** If something doesn't exist or a run failed, say so
  rather than inventing it.
- **Ask when genuinely ambiguous.** Product decisions (workload semantics,
  UX direction) belong to the maintainer; surface choices instead of picking
  silently.
- **Leave the repo better documented.** Reflect structural changes here.

---
*Maintainers: keep this file current. When in doubt, the working tree is the
source of truth — this document should describe it accurately.*
