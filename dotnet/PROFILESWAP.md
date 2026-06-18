# Profile swapper (`M365Migrate.ProfileSwap.exe`)

A ForensiT-ProfWiz-style companion for the **cutover** step: after a user signs into their
**new tenant** account on their PC, re-point that account to the **existing local Windows
profile** (desktop, documents, app data, Outlook cache, browser data, …) instead of a fresh
empty one — so they keep working as before.

It is a **separate command-line tool**, runs **on each PC**, and must run **as
Administrator** (it edits HKLM and NTFS permissions).

> **Important — validate on a test PC first.** This tool changes the registry and file
> permissions. It is **safe by default** (read-only `list`, dry-run `swap`) and backs up the
> `ProfileList` registry key before any change, but the live repoint has **not** been
> validated in this build's CI (CI only compiles it). Test it on a throwaway machine and
> confirm sign-in works before using it across a fleet. Keep the printed backup `.reg`.

## Usage

Open an **elevated** Command Prompt / PowerShell.

```
:: 1. See the local profiles and their SIDs
M365Migrate.ProfileSwap.exe list

:: 2. Dry run — shows exactly what would change, writes nothing
M365Migrate.ProfileSwap.exe swap --old "OLDDOMAIN\jdoe" --new "AzureAD\jane@target.com"

:: 3. Apply it (backs up ProfileList first)
M365Migrate.ProfileSwap.exe swap --old "OLDDOMAIN\jdoe" --new "AzureAD\jane@target.com" --execute --yes
```

- `--old` / `--new` accept an **account name** (`DOMAIN\user`, `AzureAD\user@domain`) or a
  raw **SID** (`S-1-5-21-…`).
- The **new account must have signed into the PC once** so its SID exists locally; otherwise
  `--new` can't be resolved (the tool tells you).
- `--backup <dir>` overrides the backup location (default
  `%ProgramData%\m365-migrate\profileswap`).

## What it does

1. Verifies the **new account already has a local profile** (it has signed in once); if not,
   it stops with a clear message and changes nothing.
2. Backs up `HKLM\…\ProfileList` to a `.reg` file.
3. Grants the new account **Full Control across the whole profile tree** via `icacls /T /C`
   (so child items with their own/non-inheriting permissions are updated too; locked files and
   junctions are skipped and reported, not fatal).
4. Sets the new SID's existing `ProfileImagePath` to the old profile folder, so the next
   sign-in loads the migrated profile.

Then the user signs out and back in to the new account.

## Limits / not yet done

- Does not unjoin/join the device to Entra — do that with your normal process first.
- Does not migrate the old SID's loaded registry hive ACLs beyond the folder grant; complex
  multi-profile machines may need extra handling.
- No bulk/remote execution yet (run per device, e.g. via your RMM/Intune as a packaged exe).
