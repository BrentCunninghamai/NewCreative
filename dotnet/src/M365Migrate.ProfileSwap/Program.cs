using System.Runtime.Versioning;
using M365Migrate.ProfileSwap;

// ProfWiz-style local profile takeover for tenant-to-tenant cutover. Windows-only,
// elevated. Safe by default: `list` and a dry-run `swap` write nothing; the actual
// repoint requires `--execute --yes`.

if (!OperatingSystem.IsWindows())
{
    Console.Error.WriteLine("This tool runs on Windows only.");
    return 2;
}

return Run(args);

[SupportedOSPlatform("windows")]
static int Run(string[] args)
{
    var cmd = args.FirstOrDefault()?.ToLowerInvariant();
    try
    {
        switch (cmd)
        {
            case "list":
                ListCmd();
                return 0;
            case "swap":
                return SwapCmd(args);
            default:
                Usage();
                return 1;
        }
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine("Error: " + ex.Message);
        return 1;
    }
}

[SupportedOSPlatform("windows")]
static void ListCmd()
{
    Console.WriteLine("Local Windows profiles (HKLM ProfileList):");
    Console.WriteLine();
    foreach (var p in ProfileManager.ListProfiles().OrderBy(p => p.IsSystem))
    {
        var tag = p.IsSystem ? "  [system]" : "";
        Console.WriteLine($"  {p.Account ?? "(unresolved)",-40} {p.Sid}");
        Console.WriteLine($"      {p.ProfilePath}{tag}");
    }
}

[SupportedOSPlatform("windows")]
static int SwapCmd(string[] args)
{
    var old = GetOpt(args, "--old");
    var @new = GetOpt(args, "--new");
    if (old is null || @new is null)
    {
        Console.Error.WriteLine("swap requires --old <account|SID> and --new <account|SID>.");
        return 1;
    }

    var plan = ProfileManager.Plan(old, @new);
    Console.WriteLine(plan.Describe());
    Console.WriteLine();

    var execute = args.Contains("--execute", StringComparer.OrdinalIgnoreCase);
    if (!execute)
    {
        Console.WriteLine("DRY RUN — nothing changed. Re-run with --execute --yes to apply.");
        return 0;
    }
    if (!args.Contains("--yes", StringComparer.OrdinalIgnoreCase))
    {
        Console.Error.WriteLine("Refusing to modify the system without --yes (this is destructive).");
        return 1;
    }
    if (!ProfileManager.IsElevated())
    {
        Console.Error.WriteLine("Run this from an elevated (Administrator) prompt.");
        return 1;
    }

    var backupDir = GetOpt(args, "--backup")
        ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "m365-migrate", "profileswap");
    var result = ProfileManager.Execute(plan, backupDir);
    Console.WriteLine($"Done. Registry backup: {result.BackupPath}");
    Console.WriteLine($"ACL grant: {result.AclSummary}");
    Console.WriteLine("Have the user sign out and back in to the new account to load the migrated profile.");
    return 0;
}

static string? GetOpt(string[] args, string name)
{
    var i = Array.FindIndex(args, a => string.Equals(a, name, StringComparison.OrdinalIgnoreCase));
    return i >= 0 && i + 1 < args.Length ? args[i + 1] : null;
}

static void Usage()
{
    Console.WriteLine("m365-migrate profile swapper (Windows, run as Administrator)");
    Console.WriteLine();
    Console.WriteLine("  list");
    Console.WriteLine("      Show local Windows profiles (SID, account, path).");
    Console.WriteLine();
    Console.WriteLine("  swap --old <account|SID> --new <account|SID> [--execute --yes] [--backup <dir>]");
    Console.WriteLine("      Re-point the new tenant account to the existing local profile.");
    Console.WriteLine("      Without --execute it's a dry run. The new account must have signed");
    Console.WriteLine("      into this PC once so its SID exists locally.");
}
