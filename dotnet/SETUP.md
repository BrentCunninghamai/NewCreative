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

### Graph application permissions by workload

| Workload   | Minimum Graph application permissions |
|------------|----------------------------------------|
| Users      | `User.ReadWrite.All` (+ `Organization.Read.All` for Test connections) |
| Users enrich (licenses/manager) | `User.ReadWrite.All`, `Directory.ReadWrite.All` |
| Groups     | `Group.ReadWrite.All`, `User.Read.All` |
| Mailboxes  | `MailboxSettings.ReadWrite`, `User.Read.All` |
| Files      | `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `User.Read.All` |
| Teams      | `Group.ReadWrite.All`, `Team.Create`, `Channel.ReadBasic.All` |

`Organization.Read.All` is what **Test connections** reads, so add it to both
apps. Granting the broad set above on both tenants is simplest; tighten later.

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

### Recommended order (later workloads depend on earlier ones)

1. **Users** — so accounts exist in the target.
2. **Groups** — membership/owners resolve to the migrated users.
3. **Mailboxes**, **Files**, **Teams** — these rely on the target users/groups
   existing. (Teams needs Groups run first, since a team rides on its M365 group,
   and the group needs an owner to be Teams-enabled.)

---

## 4. Troubleshooting

| Symptom | Cause / fix |
|---------|-------------|
| Test connections fails with **401** | Wrong tenant id / client id / secret, or secret expired. Re-copy the secret value. |
| Test connections fails with **403** | Admin consent not granted, or missing permission. Re-run **Grant admin consent**. |
| Plan shows users as **conflict** | The target UPN already exists — expected when re-running or for shared identities. |
| Mailbox users **skipped** | The user has no Exchange mailbox, or the target account doesn't exist yet (run Users first). |
| Teams **skipped: target M365 group missing** | Run **Groups** first so the backing group exists. |
| Teams **error** on enable | The target group has no owner — ensure Groups sync added owners (owners must themselves exist as migrated users). |
| Files **error** on a huge file | Very large files use an upload session; transient failures can be re-run (re-upload is idempotent/replace). |

---

## What this tool does and does not do

**Does:** create users; provision security/M365 groups with membership, owners,
and dynamic rules; migrate mailbox *settings*; copy OneDrive/SharePoint files and
folders (small + large via upload sessions) and reapply direct user sharing;
enable Teams and recreate standard channels.

**Does not (by design / platform limits):** move mailbox *content*
(mail/calendar/contacts — needs a native cross-tenant mailbox move); distribution
lists / mail-enabled security groups (need Exchange Online); private/shared
channels, tabs, apps; file version history and full metadata (needs the
SharePoint Migration API). These are surfaced honestly rather than silently
skipped.
