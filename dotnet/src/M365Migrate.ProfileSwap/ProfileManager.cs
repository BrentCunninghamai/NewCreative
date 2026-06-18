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

        return new SwapPlan(profile, newSid, LooksLikeSid(newAccountOrSid) ? TryResolveAccount(newSid) : newAccountOrSid);
    }

    /// <summary>
    /// Apply the swap: back up ProfileList, grant the new account Full Control on the profile
    /// folder, and point the new SID's ProfileImagePath at the existing profile. Returns the
    /// backup file path. Caller must confirm; this is destructive.
    /// </summary>
    public static string Execute(SwapPlan plan, string backupDirectory)
    {
        if (!IsElevated())
            throw new InvalidOperationException("Run as Administrator — this changes HKLM and NTFS permissions.");

        Directory.CreateDirectory(backupDirectory);
        var backup = Path.Combine(backupDirectory, $"ProfileList-backup-{DateTime.Now:yyyyMMdd-HHmmss}.reg");
        BackupProfileList(backup);

        GrantFullControl(plan.OldProfile.ProfilePath, plan.NewSid);

        using var key = Registry.LocalMachine.CreateSubKey($@"{ProfileListKey}\{plan.NewSid}")
            ?? throw new InvalidOperationException("Could not open the new SID's ProfileList key.");
        key.SetValue("ProfileImagePath", plan.OldProfile.ProfilePath, RegistryValueKind.ExpandString);

        return backup;
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

    private static void GrantFullControl(string profilePath, string sid)
    {
        if (!Directory.Exists(profilePath))
            throw new DirectoryNotFoundException($"Profile folder not found: {profilePath}");

        var di = new DirectoryInfo(profilePath);
        var sec = di.GetAccessControl();
        sec.AddAccessRule(new FileSystemAccessRule(
            new SecurityIdentifier(sid),
            FileSystemRights.FullControl,
            InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
            PropagationFlags.None,
            AccessControlType.Allow));
        di.SetAccessControl(sec);
    }
}
