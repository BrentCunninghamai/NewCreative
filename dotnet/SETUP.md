# Setup & usage guide — m365-migrate (Windows app)

This walks you from zero to a working tenant-to-tenant migration. The most
error-prone part is the **Azure AD app registrations and admin consent** — do
those carefully and the rest is straightforward.

You need: **Global Administrator** (or Privileged Role Admin for consent) on
**both** the source and target Microsoft 365 tenants.

---

## 1. Register an app in EACH tenant

Do this **twice** — once in the source tenant, once in the target tenant.

1. Sign in to the [Microsoft Entra admin center](https://entra.microsoft.com) for
   the tenant.
2. **Identity → Applications → App registrations → New registration**.
   - Name: `m365-migrate`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: leave blank. **Register.**
3. On the app's **Overview**, copy the **Application (client) ID** and the
   **Directory (tenant) ID** — you'll paste these into the app.
4. **Certificates & secrets → New client secret** → copy the secret **Value**
   immediately (it's only shown once).
5. **API permissions → Add a permission → Microsoft Graph → Application
   permissions** → add the permissions for the workloads you'll run (see table
   below), then **Grant admin consent for &lt;tenant&gt;** (this step is
   essential — without it every call returns 403).

> **Fast path (recommended):** instead of adding permissions one by one, run the
> app, fill in both tenants' IDs, and click **“App setup (permissions & 1-click
> consent)”**. It writes a file containing (a) a **manifest block** to paste into
> each app registration’s **Manage → Manifest** (`requiredResourceAccess`) so all
> permissions are added at once, and (b) a **one-click admin-consent link** per
> tenant — a Global Admin opens it, signs in, and approves everything in one go.

### Graph application permissions by workload

| Workload   | Minimum Graph application permissions |
|------------|----------------------------------------|
| Users      | `User.ReadWrite.All` (+ `Organization.Read.All` for Test connections) |
| Users enrich (licenses/manager) | `User.ReadWrite.All`, `Directory.ReadWrite.All` |
| Groups     | `Group.ReadWrite.All`, `User.Read.All` |
| Mailboxes  | `MailboxSettings.ReadWrite`, `User.Read.All` |
| Mail (content) | `Mail.ReadWrite`, `User.Read.All` |
| Calendar & Contacts (content) | `Calendars.ReadWrite`, `Contacts.ReadWrite`, `User.Read.All` |
| Files      | `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `User.Read.All` |
| Teams      | `Group.ReadWrite.All`, `Team.Create`, `Channel.ReadBasic.All` |
| Teams (messages) | `Teamwork.Migrate.All`, `User.Read.All` |
| MTO confirmation (optional) | `MultiTenantOrganization.Read.All` (on the **target** app) |

`Organization.Read.All` is what **Test connections** reads, so add it to both
apps. Granting the broad set above on both tenants is simplest; tighten later.

`MultiTenantOrganization.Read.All` is **optional**: with it, **Test connections**
also reads the target's Multi-Tenant Organization and tells you whether the source
tenant is a member (so you know some source users are *expected* to already exist in
the target via cross-tenant sync, and the plan will flag them as `conflict`). Without
it, the tool still detects those already-present users from their `#EXT#` identities —
this permission just makes the relationship explicit instead of inferred. Add it to
the **target** app registration only.

---

## 2. Get the app

Download **`M365Migrate.exe`** from the repository's
[**latest release**](../../releases/tag/latest). It is a single self-contained
executable — no .NET install required.

> **Windows SmartScreen:** the download is **not code-signed**, so Windows may
> warn "Windows protected your PC". Click **More info → Run anyway**. To remove
> this prompt permanently, set up signing per [SIGNING.md](SIGNING.md) — the CI
> pipeline is already wired for Azure Trusted Signing and just needs credentials.

---

## 3. Run a migration

1. Launch `M365Migrate.exe`.
2. Fill in **Source tenant** and **Target tenant**: Tenant ID, App (client) ID,
   Client secret, and Primary domain (e.g. `contoso.onmicrosoft.com` and
   `fabrikam.onmicrosoft.com`).
3. Click **Test connections** — you should see both organizations' names. If not,
   fix the registration/consent before continuing (see Troubleshooting).
4. Pick a **workload**. For **Files (OneDrive)**, also fill the **Files scope**
   with the source user's UPN.
5. Click **Discover & Plan** and review the grid (this never writes anything).
6. Click **Migrate**. It is a **dry run** until you tick **Execute**. Tick
   **Execute** and migrate again to apply.

### Start with User mapping (preview)

Before migrating content, run the **User mapping (preview)** workload (it's the first
in the list). It scans **both** tenants and builds a **source → target** identity map,
choosing the best match per user by precedence:

1. an explicit **CSV override** you provide,
2. **cross-tenant / B2B** (`#EXT#`) identity already in the target,
3. **primary mail / SMTP proxy** address match,
4. **UPN domain rewrite** (source domain → target domain).

The grid shows each source user, the **method** used, and the matched target (or
`unmatched`). It also exports an editable `mapping-*.csv` to the reports folder. This is
read-only — nothing is written. Use it to confirm who maps to whom (and fix the messy
hybrid/orchestrator cases) before running the actual content workloads.

### Bulk content for all mapped users

Once the mapping looks right, the **Bulk Mail (mapped users)** and **Bulk OneDrive
(mapped users)** workloads run that content for **every matched user** in a single pass —
this is the ShareGate-style "migrate everyone" step. **Discover & Plan** lists the queued
users; **Migrate** runs them all (dry run until you tick **Execute**), one result row per
user with live progress. It's resumable (mail uses mailbox-wide dedup; OneDrive re-upload
replaces), and a per-user problem (e.g. a user with no OneDrive, or a mailbox still
mid-move) is recorded and skipped without stopping the batch. **Cancel** stops between
users; already-processed users stay done.

### Recommended order (later workloads depend on earlier ones)

1. **Users** — so accounts exist in the target.
2. **Groups** — membership/owners resolve to the migrated users.
3. **Mailboxes**, **Files**, **Teams** — these rely on the target users/groups
   existing. (Teams needs Groups run first, since a team rides on its M365 group,
   and the group needs an owner to be Teams-enabled.)

---

## Merging multiple source tenants into one target

Use **one target app registration** (single-tenant, in the target) and **one app
registration per source tenant**. Run the flow once per source: keep the Target
fields the same and change only the Source fields each time.

Because several sources land in the same target, identities can collide (two
sources each with `john@…`, or a `sales` group in both). Set a per-source **Name
prefix** (or suffix) — e.g. `contoso-` for one source, `northwind-` for the next.
It tags each migrated user's UPN local-part (`contoso-john@target`) and each group
mailNickname (`contoso-sales`), so nothing overwrites another source. Leave it
blank for a single-source migration. Optionally set a **Display-name suffix**
(e.g. `(Contoso)`) so merged users are also distinguishable in the target GAL
(`Jane Doe (Contoso)`).

## Taking over a user already (partly) migrated by Microsoft

Common case: a user was started with **Microsoft's native cross-tenant migration** —
Entra **cross-tenant sync** for identity plus an Exchange **cross-tenant mailbox move**
(the "T2T … CatchUp" batches in EAC, often with Microsoft **cross-tenant licenses**) —
and you want this tool to **finish** the job (fill mail gaps, and do **OneDrive** and
**Teams**, which the mailbox move does *not* cover).

Two things make this work:

1. **Mailbox-wide mail dedup.** The mail workload skips any message already in the
   target by `internetMessageId` (regardless of folder), so it only copies what
   Microsoft's move hasn't already landed. Safe to re-run.
2. **Target user UPN override.** These users' target identity is usually **not** a clean
   domain rewrite of the source (e.g. source `paule@net1.com`, but the target mailbox
   is `paul.encarnacao@target.onmicrosoft.com`, with the M365 *username* something else
   again). Domain-rewrite would target the wrong (or a non-existent) account. So fill:
   - **Scope** = the user's **source** UPN (the net1.com mailbox you're reading from).
   - **Target user UPN** = the user's **actual UPN in the target** — i.e. the value shown
     as **Username** in the target's M365 admin *Manage username and email* (not an alias,
     not necessarily the primary SMTP). This is used verbatim as the target mailbox/OneDrive.

   Leave **Target user UPN** blank for normal users whose target *is* a domain rewrite.
   The **Target user UPN** field also accepts the user's **object ID (GUID)** — handy when
   the UPN is ambiguous in these hybrid setups.

After **Discover & Plan** for a per-user content workload (Mail / Files / Calendar), the
grid's first **Identity** row shows exactly who was resolved on each side, e.g.
*Paul Encarnacao (PaulE@net1.com) → Paul Encarnacao (paul.encarnacao@…onmicrosoft.com)*.
Confirm that's the right person **before** you tick Execute. If it reads `NOT FOUND`, fix
the Scope (source) or set the **Target user UPN** (UPN or object ID) and re-plan.

> **Timing — don't race an in-flight mailbox move.** While Microsoft's EAC batch shows
> the mailbox as **Synced** (not **Completed**), it is an *active* move target. Let that
> user's batch **Complete** (or remove it) **before** running this tool's **Mail** workload
> on the same mailbox, so the two aren't writing to it at once. **OneDrive** and **Teams**
> are not part of the mailbox move, so you can run those at any time without conflict.

## 4. Troubleshooting

| Symptom | Cause / fix |
|---------|-------------|
| Test connections fails with **401** | Wrong tenant id / client id / secret, or secret expired. Re-copy the secret value. |
| Test connections fails with **403** | Admin consent not granted, or missing permission. Re-run **Grant admin consent**. |
| Plan shows users as **conflict** | The target UPN already exists — expected when re-running or for shared identities. |
| Plan shows **Name = Detail** (UPN unchanged) | The Source primary domain doesn't match the users' actual UPN domain, so no rewrite happened. Use the real UPN domain (e.g. `contoso.onmicrosoft.com`), **not** the `…mail.onmicrosoft.com` routing domain. |
| Lots of `…#EXT#@…` rows | Those are external/B2B guest identities; with **Skip guest users** ticked they're skipped (they can't be recreated as normal users). |
| User shows **conflict — already in target (cross-tenant/B2B sync)** | The user already exists in the target via cross-tenant sync / a Multi-Tenant Org (matched by their decoded #EXT# identity or mail), so the tool won't create a native duplicate. Review whether to keep the synced identity or convert it. |
| Mailbox users **skipped** | The user has no Exchange mailbox, or the target account doesn't exist yet (run Users first). |
| Teams **skipped: target M365 group missing** | Run **Groups** first so the backing group exists. |
| Teams **error** on enable | The target group has no owner — ensure Groups sync added owners (owners must themselves exist as migrated users). |
| Files **error** on a huge file | Very large files use an upload session; transient failures can be re-run (re-upload is idempotent/replace). |
| Files shows **OneDrive — unavailable** (404) | The source user genuinely has **no provisioned OneDrive**. Nothing to copy — focus on Mail/Teams. |
| Files shows **OneDrive — error** / `notSupported` | The user **has** a OneDrive but Graph refused the read. Almost always the **source app** is missing `Files.ReadWrite.All` + `Sites.ReadWrite.All` (Application) **with admin consent** — add and consent both, then re-plan. If the tenant is **multi-geo**, the drive may live in another geo. |
| Identity row shows **NOT FOUND** for the target | The Target user UPN doesn't resolve — common when the source domain isn't verified in the target (so `user@olddomain` isn't the target UPN). Put the user's **object ID (GUID)** or real target UPN in **Target user UPN or object ID**, then re-plan. |
| Mail **slow** / 429s | Expected for big mailboxes — Graph throttles; the client backs off and retries. Re-running is safe: dedup is **mailbox-wide** by `internetMessageId`, so any message already in the target (from a prior run, coexistence sync, or another migration tool — even if filed in a different folder) is skipped, not duplicated. |
| Mail folders look different | Folders are matched by display name; tenants in different languages may not match well-known folders (Inbox, etc.). |

---

## What this tool does and does not do

**Does:** create users; provision security/M365 groups with membership, owners,
and dynamic rules; migrate mailbox *settings*; copy **mail content** (folders +
messages, full-fidelity MIME, idempotent with **mailbox-wide** dedup so mail already
present from another tool/sync isn't duplicated) and **calendar + contacts** per user;
copy OneDrive/SharePoint files and folders (small + large via upload sessions) and
reapply direct user sharing; enable Teams and recreate standard channels; and
import Teams channel **message history** (migration mode — fresh team, original
authors + timestamps, top-level messages, run once).

**Does not yet:** Teams threaded replies, Teams team membership for migration-mode
teams (add after migration), and 1:1/group chats; distribution lists /
mail-enabled security groups (need Exchange Online); private/shared channels, tabs,
apps; file version history and full metadata (needs the SharePoint Migration API).
These are surfaced honestly rather than silently skipped.

> **Teams (messages) note:** this path creates a **new** migration-mode team per
> source team (separate from the structure-only Teams workload, which enables Teams
> on the already-migrated M365 group). Use one or the other per team. It's **not
> idempotent** — re-running creates a duplicate team, so run it once.
