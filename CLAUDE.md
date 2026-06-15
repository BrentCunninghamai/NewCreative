# CLAUDE.md

Guidance for AI assistants (Claude Code and others) working in this repository.

## Project status: early development

**NewCreative** is a Microsoft 365 **tenant-to-tenant (T2T) migration** tool,
distributed as the Python package `m365-migrate`. The Users / Identities and
Groups workloads are implemented end-to-end, plus Exchange mailbox **settings**
migration, OneDrive / SharePoint **file** copy, and **Teams** (enable Teams +
recreate channels); remaining workloads are on the roadmap (see
`docs/architecture.md`).

### Stack & layout

- **Language:** Python 3.10+ (CLI built with Typer).
- **Key deps:** `httpx`, `azure-identity`, `pydantic`, `pyyaml`, `rich`,
  `tenacity`. Dev: `pytest`, `respx`.
- **Source:** `src/m365_migrate/`. **Tests:** `tests/`. **Docs:**
  `docs/architecture.md`. **Manifest:** `pyproject.toml`.

```
src/m365_migrate/
  config.py          # YAML config + ${ENV:...} secret resolution
  auth.py            # per-tenant client-credentials token providers
  graph_client.py    # Graph REST wrapper: paging + 429/5xx retry
  models.py          # SourceUser/PlannedUser + Group + Mailbox + DriveItem + Team models
  mapping.py         # UPN rewriting + mapping CSV I/O
  workloads/users.py # Users workload: discover / plan / migrate / enrich
  workloads/groups.py # Groups workload: discover / plan / sync (membership)
  workloads/mailboxes.py # Mailbox settings: discover / plan / migrate
  workloads/files.py # OneDrive/SharePoint: discover / plan / migrate (file copy)
  workloads/teams.py # Teams: discover / plan / migrate (enable Teams + channels)
  cli.py             # Typer CLI entry point (`m365-migrate`)
```

### Commands

- **Install (dev):** `pip install -e ".[dev]"`
- **Test:** `pytest` (mocks Graph via `respx`; needs no real tenants)
- **Run:** `m365-migrate --help` (or `python -m m365_migrate --help`)

### Conventions specific to this project

- **Plan before write.** Each workload separates a read-only plan phase from the
  migrate phase; `migrate` is a dry run unless `--execute` is passed.
- **Never commit credentials.** Real config goes in git-ignored `config.yaml`;
  secrets resolve from env via `${ENV:NAME}`. Only `config.example.yaml` is
  committed.
- **Keep tests offline.** Inject the HTTP client / token provider so tests mock
  Graph rather than hitting the network.

> When you add a new workload or change structure, **update this file and
> `docs/architecture.md` in the same change** so they stay an accurate map.

## License: public domain

This project is released into the public domain under The Unlicense. Practical
implications for contributions:

- Only add code/content that is itself public-domain-compatible. Do **not**
  paste in code under restrictive licenses (GPL, proprietary, etc.) or content
  that would impose attribution or copyleft obligations.
- Be cautious with third-party dependencies — prefer permissively licensed ones
  (MIT, BSD, Apache-2.0, ISC, Unlicense, CC0) and note any license constraints
  the dependency carries.

## Git workflow & conventions

- **Default branch:** `main`.
- **Do not commit directly to `main`.** Develop on a feature branch and open a
  pull request for review. When a session specifies a designated working
  branch, use exactly that branch.
- **Pushing:** use `git push -u origin <branch-name>`. After pushing, open a PR
  (ready for review, not a draft) if one does not already exist for the branch.
- **Commit messages:** write clear, descriptive messages in the imperative mood
  (e.g. "Add user auth module", not "added stuff"). Keep the subject line
  concise and explain the *why* in the body when it isn't obvious.
- **Keep changes focused.** One logical change per PR where practical.

## Conventions to follow as the project grows

These are the defaults to apply once code is introduced. Update this section
with concrete commands and paths the moment the toolchain is chosen.

1. **Pick and document the stack.** When the first language/framework is added,
   record here: how to install dependencies, how to build, how to run, and how
   to test. Include the exact commands.
2. **Match existing style.** Once a codebase exists, mirror its formatting,
   naming, and structural patterns rather than importing outside conventions.
   Prefer the project's configured formatter/linter over manual styling.
3. **Add tests with features.** Establish a test directory and runner early;
   new behavior should ship with tests.
4. **Keep the root tidy.** Source under a conventional directory (e.g. `src/`),
   tests alongside or under `tests/`, docs under `docs/`. Add a `.gitignore`
   appropriate to the chosen stack before committing build artifacts or
   dependencies.

## Working agreement for AI assistants

- **Verify, don't assume.** This file describes a near-empty repo. Before
  acting on any assumption about structure or tooling, check the actual files
  on disk — the project may have grown since this was written.
- **Be honest about state.** If something doesn't exist yet, say so rather than
  inventing it. Don't fabricate build/test commands that haven't been set up.
- **Ask when genuinely ambiguous.** Foundational decisions (language, framework,
  architecture) belong to the maintainer. Surface the choice instead of picking
  silently.
- **Leave the repo better documented.** Whenever you make a structural change,
  reflect it here.

---
*Maintainers: keep this file current. When in doubt, the working tree is the
source of truth — this document should describe it accurately.*
