using System.Diagnostics;
using System.Runtime.Versioning;
using System.Security.AccessControl;
using System.Security.Principal;
using Microsoft.Win32;

namespace M365Migrate.ProfileSwap;

/// <summary>A local Windows user profile as recorded under HKLM ProfileList.</summary>
public sealed record LocalProfile(string Sid, string? Account, string ProfilePath)
{
    public bool IsSystem => Sid is "S-1-5-18" or "S-1-5-19" or "S-1-5-20" || Sid.StartsWith("S-1-5-80");
}

/// <summary>The outcome of an applied swap.</summary>
public sealed record SwapResult(string BackupPath, string AclSummary);

/// <summary>The change a swap would make, for review before anything is written.</summary>
public sealed record SwapPlan(LocalProfile OldProfile, string NewSid, string? NewAccount)
{
    public string Describe() =>
        $"Re-point profile '{OldProfile.ProfilePath}' (was {OldProfile.Account ?? OldProfile.Sid})\n" +
        $"  -> new owner {NewAccount ?? NewSid} [{NewSid}]\n" +
        $"  Actions: grant {NewAccount ?? NewSid} Full Control on the profile folder;\n" +
        $"           set ProfileList\\{NewSid}\\ProfileImagePath = {OldProfile.ProfilePath}.";
}

/// <summary>
/// ForensiT-ProfWiz-style local profile takeover for tenant-to-tenant cutover: after a user
/// signs into their NEW tenant account on the PC, re-point that account to the EXISTING local
/// profile (desktop, documents, app data, Outlook cache, ...) instead of a fresh empty one.
/// Windows-only and must run elevated. Destructive steps are dry-run by default and back up
/// the ProfileList registry key first.
/// </summary>
[SupportedOSPlatform("windows")]
public static class ProfileManager
{
    private const string ProfileListKey =
        @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList";

    public static bool IsElevated()
    {
        using var identity = WindowsIdentity.GetCurrent();
        return new WindowsPrincipal(identity).IsInRole(WindowsBuiltInRole.Administrator);
    }

    public static bool LooksLikeSid(string value) =>
        value.StartsWith("S-1-", StringComparison.OrdinalIgnoreCase);

    /// <summary>Enumerate local profiles from the registry, resolving account names where possible.</summary>
    public static List<LocalProfile> ListProfiles()
    {
        var profiles = new List<LocalProfile>();
        using var key = Registry.LocalMachine.OpenSubKey(ProfileListKey);
        if (key is null) return profiles;
        foreach (var sid in key.GetSubKeyNames())
        {
            using var sub = key.OpenSubKey(sid);
            var path = sub?.GetValue("ProfileImagePath") as string ?? "";
            profiles.Add(new LocalProfile(sid, TryResolveAccount(sid), path));
        }
        return profiles;
    }

    /// <summary>Resolve an account name or SID string to a SID. Returns null if unresolvable.</summary>
    public static string? ResolveSid(string accountOrSid)
    {
        if (LooksLikeSid(accountOrSid))
            return accountOrSid;
        try
        {
            var sid = (SecurityIdentifier)new NTAccount(accountOrSid).Translate(typeof(SecurityIdentifier));
            return sid.Value;
        }
        catch
        {
            return null;
        }
    }

    private static string? TryResolveAccount(string sid)
    {
        try
        {
            return new SecurityIdentifier(sid).Translate(typeof(NTAccount)).Value;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>Build a swap plan; throws with a clear message if either side can't be resolved.</summary>
    public static SwapPlan Plan(string oldAccountOrSid, string newAccountOrSid)
    {
        var oldSid = ResolveSid(oldAccountOrSid)
            ?? throw new InvalidOperationException($"Could not resolve old account '{oldAccountOrSid}'.");
        var profile = ListProfiles().FirstOrDefault(p => string.Equals(p.Sid, oldSid, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"No local profile found for '{oldAccountOrSid}' ({oldSid}).");

        var newSid = ResolveSid(newAccountOrSid)
            ?? throw new InvalidOperationException(
                $"Could not resolve new account '{newAccountOrSid}'. Have the user sign into the new account once " +
                "on this PC so its SID exists locally, then retry.");

        // The new account must already have a profile (i.e. it has signed in once). Without
        // this its ProfileList key is absent; creating one would leave a bogus/incomplete
        // entry. Enforce the precondition before anything is changed.
        if (!ProfileKeyExists(newSid))
            throw new InvalidOperationException(
                $"The new account '{newAccountOrSid}' ({newSid}) has no local profile yet. Sign into it once on " +
                "this PC (a fresh profile is created), then run the swap to re-point it to the old profile.");

        return new SwapPlan(profile, newSid, LooksLikeSid(newAccountOrSid) ? TryResolveAccount(newSid) : newAccountOrSid);
    }

    private static bool ProfileKeyExists(string sid)
    {
        using var k = Registry.LocalMachine.OpenSubKey($@"{ProfileListKey}\{sid}");
        return k is not null;
    }

    /// <summary>
    /// Apply the swap: back up ProfileList, grant the new account Full Control across the whole
    /// profile tree, and point the new SID's existing ProfileImagePath at the old profile.
    /// Caller must confirm; this is destructive.
    /// </summary>
    public static SwapResult Execute(SwapPlan plan, string backupDirectory)
    {
        if (!IsElevated())
            throw new InvalidOperationException("Run as Administrator — this changes HKLM and NTFS permissions.");

        // Open the new SID's existing ProfileList key for write. Never CreateSubKey here: if
        // it's absent the account hasn't signed in, and creating a bogus entry (after we'd
        // already changed ACLs) would corrupt ProfileList. Fail before any change instead.
        using var key = Registry.LocalMachine.OpenSubKey($@"{ProfileListKey}\{plan.NewSid}", writable: true)
            ?? throw new InvalidOperationException(
                $"The new account ({plan.NewSid}) has no ProfileList entry — sign into it once first. Nothing changed.");

        Directory.CreateDirectory(backupDirectory);
        var backup = Path.Combine(backupDirectory, $"ProfileList-backup-{DateTime.Now:yyyyMMdd-HHmmss}.reg");
        BackupProfileList(backup);

        var aclSummary = GrantFullControl(plan.OldProfile.ProfilePath, plan.NewSid);
        key.SetValue("ProfileImagePath", plan.OldProfile.ProfilePath, RegistryValueKind.ExpandString);

        return new SwapResult(backup, aclSummary);
    }

    private static void BackupProfileList(string file)
    {
        var psi = new ProcessStartInfo("reg.exe",
            $"export \"HKLM\\{ProfileListKey}\" \"{file}\" /y")
        {
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        using var p = Process.Start(psi);
        p?.WaitForExit();
        if (p is null || p.ExitCode != 0)
            throw new InvalidOperationException("Failed to back up the ProfileList registry key — aborting.");
    }

    /// <summary>
    /// Grant the SID Full Control across the entire profile tree using icacls, so child items
    /// with protected/non-inheriting ACLs are also updated (a root-only inheritable ACE would
    /// miss them). <c>/T</c> recurses, <c>/C</c> continues past per-file failures and reparse
    /// points (legacy junctions like "Application Data"). Returns icacls' summary line.
    /// </summary>
    private static string GrantFullControl(string profilePath, string sid)
    {
        if (!Directory.Exists(profilePath))
            throw new DirectoryNotFoundException($"Profile folder not found: {profilePath}");

        var psi = new ProcessStartInfo("icacls.exe",
            $"\"{profilePath}\" /grant *{sid}:(OI)(CI)F /T /C")
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };
        using var p = Process.Start(psi) ?? throw new InvalidOperationException("Could not start icacls.exe.");
        var stdout = p.StandardOutput.ReadToEnd();
        var stderr = p.StandardError.ReadToEnd();
        p.WaitForExit();

        var summary = stdout.Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(l => l.Trim()).LastOrDefault(l => l.Length > 0) ?? "";
        // With /C icacls returns non-zero if any item failed; that's expected (locked files).
        // Only treat it as fatal if nothing was processed at all.
        if (p.ExitCode != 0 && !summary.StartsWith("Successfully processed", StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("icacls failed: " + (stderr.Trim().Length > 0 ? stderr.Trim() : summary));
        return summary;
    }
}
