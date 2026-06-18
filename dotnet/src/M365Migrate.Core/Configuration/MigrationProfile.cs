using System.Text.Json;

namespace M365Migrate.Core.Configuration;

/// <summary>
/// A saved set of connection + option inputs so an admin doesn't re-enter them each launch
/// (e.g. one profile per source tenant). Client secrets are stored encrypted — see
/// <see cref="ProfileStore"/>, which takes a protect/unprotect function (the app supplies a
/// Windows DPAPI one, scoped to the current user).
/// </summary>
public sealed record MigrationProfile
{
    public string Name { get; init; } = "";

    public string SourceTenantId { get; init; } = "";
    public string SourceClientId { get; init; } = "";
    public string SourceClientSecret { get; init; } = "";
    public string SourceDomain { get; init; } = "";

    public string TargetTenantId { get; init; } = "";
    public string TargetClientId { get; init; } = "";
    public string TargetClientSecret { get; init; } = "";
    public string TargetDomain { get; init; } = "";

    public bool RewriteUpn { get; init; } = true;
    public bool SkipGuests { get; init; } = true;
    public bool AssignLicenses { get; init; }
    public string DefaultUsageLocation { get; init; } = "";
    public string NamePrefix { get; init; } = "";
    public string NameSuffix { get; init; } = "";
    public string DisplayNameSuffix { get; init; } = "";
}

/// <summary>
/// Reads/writes <see cref="MigrationProfile"/>s as JSON files in a directory. The two client
/// secrets are passed through a caller-supplied <c>protect</c>/<c>unprotect</c> function so the
/// engine stays cross-platform while the app encrypts at rest with Windows DPAPI. Tests inject
/// an identity function.
/// </summary>
public static class ProfileStore
{
    private static readonly JsonSerializerOptions Options = new() { WriteIndented = true };

    public static string FileFor(string dir, string name) =>
        Path.Combine(dir, SafeName(name) + ".json");

    public static void Save(string dir, MigrationProfile profile, Func<string, string> protect)
    {
        if (string.IsNullOrWhiteSpace(profile.Name))
            throw new ArgumentException("A profile needs a name.");
        Directory.CreateDirectory(dir);
        var onDisk = profile with
        {
            SourceClientSecret = profile.SourceClientSecret.Length > 0 ? protect(profile.SourceClientSecret) : "",
            TargetClientSecret = profile.TargetClientSecret.Length > 0 ? protect(profile.TargetClientSecret) : "",
        };
        File.WriteAllText(FileFor(dir, profile.Name), JsonSerializer.Serialize(onDisk, Options));
    }

    public static MigrationProfile Load(string dir, string name, Func<string, string> unprotect)
    {
        var profile = JsonSerializer.Deserialize<MigrationProfile>(File.ReadAllText(FileFor(dir, name)))
            ?? throw new InvalidOperationException("Profile file is empty or invalid.");
        return profile with
        {
            SourceClientSecret = profile.SourceClientSecret.Length > 0 ? unprotect(profile.SourceClientSecret) : "",
            TargetClientSecret = profile.TargetClientSecret.Length > 0 ? unprotect(profile.TargetClientSecret) : "",
        };
    }

    public static List<string> List(string dir) =>
        Directory.Exists(dir)
            ? Directory.GetFiles(dir, "*.json").Select(Path.GetFileNameWithoutExtension).Where(n => n is not null).Cast<string>().OrderBy(n => n).ToList()
            : new List<string>();

    private static string SafeName(string name)
    {
        var safe = name.Trim();
        foreach (var c in Path.GetInvalidFileNameChars())
            safe = safe.Replace(c, '_');
        return safe.Length == 0 ? "profile" : safe;
    }
}
