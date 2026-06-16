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
| `models.py` | `SourceUser`/`PlannedUser`, `SourceGroup`/`PlannedGroup`, `SourceMailbox`/`PlannedMailbox` models + Graph (de)serialization. |
| `mapping.py` | UPN domain rewriting; mapping CSV read/write. |
| `workloads/users.py` | `discover_users` → `plan_users` → `migrate_users` → `enrich_users`. |
| `workloads/groups.py` | `discover_groups` → `plan_groups` → `sync_groups`. |
| `workloads/mailboxes.py` | `discover_mailboxes` → `plan_mailboxes` → `migrate_mailboxes` (settings). |
| `workloads/files.py` | `discover_drive_items` → `plan_drive_items` → `migrate_drive_items` (OneDrive/SharePoint). |
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

## Groups workload flow

```
discover  GET /groups (source, $expand=members,owners) -> SourceGroup[] (by kind)
plan      match target by mailNickname,                 -> PlannedGroup[] (create|exists|skip)
          classify kind, rewrite member + owner UPNs
sync      POST /groups for "create",                    -> results (dry-run by default)
          POST members/$ref + owners/$ref for resolved users
```

Only **security** and **Microsoft 365** (Unified) groups are provisioned via
Graph; mail-enabled security groups and distribution lists are marked ``skip``
(they need Exchange Online, a later workload). Both **membership** and
**ownership** are reconciled through the same source→target UPN rewrite as users,
so a member/owner is added only once its target account exists; users with no
target account are reported as unresolved. Reconciling owners is also a
prerequisite for the teams workload — an M365 group must have an owner before it
can be Teams-enabled.

## Mailbox workload flow

```
discover  GET /users + /users/{id}/mailboxSettings -> SourceMailbox[] (settings only)
plan      rewrite UPN, diff against target /users   -> PlannedMailbox[] (settings|skip)
migrate   PATCH /users/{id}/mailboxSettings (target) -> results (dry-run by default)
```

This workload migrates mailbox **settings** (time zone, language, working hours,
automatic replies, date/time formats, delegate options) — the writable subset of
`mailboxSettings`. Read-only fields (e.g. `userPurpose`) are dropped. Users with
no mailbox (404 from the endpoint) are skipped at discovery.

> **Mailbox content is out of scope for the Graph layer.** Moving mail, calendar,
> and contact *items* tenant-to-tenant is not a Graph REST operation; it requires
> a native cross-tenant mailbox move (MRS / migration endpoints) or a third-party
> tool. That orchestration is tracked separately on the roadmap.

## Files workload flow (OneDrive / SharePoint)

OneDrive personal storage and SharePoint document libraries are both Graph
*drives* of *driveItems*, so one set of primitives serves both. A drive root is
either `/users/{upn}/drive` (OneDrive) or `/sites/{id}/drive` (SharePoint).

```
discover  walk source drive (BFS) + $expand=permissions -> DriveItem[] (folders first)
plan      mark folders/files copy, rewrite grantee UPNs  -> PlannedDriveItem[]
migrate   POST folders, upload file content (download->upload), -> results (dry-run default)
          re-invite resolvable direct user grants
```

Breadth-first discovery guarantees a parent folder precedes its children, so the
copy can create folders before populating them. Reads come from the source
client and writes go to the target client (different tenants). Files at or below
`SIMPLE_UPLOAD_LIMIT` upload in a single `PUT .../content`; larger files use a
resumable **upload session** (`createUploadSession` + chunked PUTs of 3.2 MiB),
so large files are migrated, not skipped. Each file is buffered in memory between
download and upload.

> **Fidelity caveats.** The high-fidelity **SharePoint Migration API** (Azure
> blob + manifests, version history, full item metadata) is still future work.
> Only **direct user** grants are reapplied, and only when the grantee resolves to
> a target account; sharing links and group/external grants are skipped.

## Teams workload flow

A team is a Teams-enabled Microsoft 365 group, so this workload composes with the
others instead of duplicating them: team **membership** is the backing M365
group's membership (migrated by the groups workload), and channel **files** live
in the team's SharePoint library (migrated by the files workload). This workload
owns the Teams-specific layer — enabling Teams and recreating channels.

```
discover  /groups (resourceProvisioningOptions has "Team") + /teams/{id}/channels -> SourceTeam[]
plan      match team to target group by mailNickname; classify team + channels    -> PlannedTeam[]
migrate   PUT /groups/{id}/team (enable), POST standard channels                  -> results (dry-run default)
```

A Teams-enabled group shares its id with its team, so the matched target group id
doubles as the team id. Run `groups sync` first so the backing group exists.

> **Scope.** Only **standard** channels are recreated; the default *General*
> channel is created automatically with the team, and **private / shared**
> channels (which need channel-scoped membership) are surfaced as `skip`. Tabs,
> apps, and channel-level settings are future work. Re-runs are idempotent —
> already-enabled teams and existing channels are left in place.

## Roadmap

- [x] Users / Identities: discover, plan (with conflict detection), migrate.
- [x] Users: license assignment + manager links (`enrich`).
- [x] Groups: provision security/M365 groups + reconcile membership & ownership
      (`groups sync`).
- [ ] Groups: mail-enabled security groups + distribution lists (via Exchange).
- [ ] Groups: dynamic membership rules, nested groups.
- [x] Exchange Online mailboxes: settings migration (`mailbox migrate`).
- [ ] Exchange Online mailboxes: content move (mail/calendar/contacts) via native
      cross-tenant mailbox migration.
- [x] OneDrive / SharePoint: copy files/folders (simple + chunked upload sessions
      for large files) + reapply direct user grants (`files migrate`).
- [ ] OneDrive / SharePoint: SharePoint Migration API (version history, full
      metadata); sharing links and group/external grants.
- [x] Teams: enable Teams on migrated M365 groups + recreate standard channels
      (`teams migrate`). Membership rides on the group; channel files on SharePoint.
- [ ] Teams: private/shared channels, channel membership, tabs, apps, and settings.
- [ ] Resumable runs + structured run logs / reporting.
- [ ] Concurrency with per-tenant throttling budgets.

## Known limitations (today)

- Users, Groups, Mailbox-settings, Files, and Teams workloads exist; mailbox
  content does not yet.
- License assignment requires the matching SKU to exist in the target tenant;
  unavailable SKUs are skipped (not purchased automatically).
- Groups: only security and Microsoft 365 groups are provisioned. Mail-enabled
  security groups and distribution lists are skipped, and dynamic membership rules
  and nested (group-in-group) members are not migrated yet. Members and owners are
  reconciled (users only — non-user owners such as service principals are ignored).
- Mailbox: only settings are migrated, not mail/calendar/contact content (which
  needs a native cross-tenant mailbox move). Target mailboxes must already exist
  to receive settings.
- Files: small files use simple upload and large files a chunked upload session,
  but each file is buffered in memory (no streaming), and version history and most
  item metadata are not preserved; only direct user grants that resolve in the
  target are reapplied.
- Teams: a team's backing M365 group must already exist in the target (run
  `groups sync` first) so it can be Teams-enabled. Enabling Teams requires the
  group to have an owner, which `groups sync` now reconciles — but only owners
  with an existing target account are added, so a team whose owners haven't been
  migrated as users will still fail to teamify. Only standard channels are
  recreated — private/shared channels, channel membership, tabs, apps, and team
  settings are not migrated yet.
- New users get a random password and must reset on first sign-in; there is no
  password/identity federation handoff.
- No incremental/delta sync yet — `plan` is a full comparison each run.
