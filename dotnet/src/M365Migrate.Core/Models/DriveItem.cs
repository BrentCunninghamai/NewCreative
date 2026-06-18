using System.Text.Json;
using M365Migrate.Core.Graph;

namespace M365Migrate.Core.Models;

/// <summary>A direct user permission grant on a drive item (file or folder).</summary>
public sealed class DriveGrant
{
    public string Upn { get; set; } = "";
    public List<string> Roles { get; set; } = new();
}

/// <summary>
/// A OneDrive / SharePoint file or folder as read from a source drive. Both are
/// exposed by Graph as drives of driveItems, so this serves both.
/// </summary>
public sealed class DriveItem
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    /// <summary>Parent folder path relative to the drive root ("" at the root).</summary>
    public string ParentPath { get; set; } = "";
    public bool IsFolder { get; set; }
    public long Size { get; set; }
    /// <summary>OneDrive/SharePoint content hash (quickXorHash); deterministic for identical
    /// bytes across tenants, so it's a reliable "already copied / unchanged" signal. Null when
    /// the service didn't return a hash (e.g. some SharePoint libraries) — then don't skip.</summary>
    public string? QuickXorHash { get; set; }
    public List<DriveGrant> Grants { get; set; } = new();

    /// <summary>Full path of this item relative to the drive root, including its name.</summary>
    public string RelativePath => string.IsNullOrEmpty(ParentPath) ? Name : $"{ParentPath}/{Name}";

    public static DriveItem FromGraph(JsonElement data)
    {
        var rel = "";
        if (data.TryGetProperty("parentReference", out var parentRef)
            && parentRef.ValueKind == JsonValueKind.Object)
        {
            var rawPath = parentRef.GetStringOrNull("path") ?? "";
            var marker = rawPath.IndexOf("root:", StringComparison.Ordinal);
            if (marker >= 0)
                rel = rawPath[(marker + "root:".Length)..].TrimStart('/');
        }

        var grants = new List<DriveGrant>();
        foreach (var perm in data.GetArrayOrEmpty("permissions"))
        {
            JsonElement user = default;
            var hasUser = false;
            if (perm.TryGetProperty("grantedToV2", out var g2) && g2.ValueKind == JsonValueKind.Object
                && g2.TryGetProperty("user", out user))
                hasUser = true;
            else if (perm.TryGetProperty("grantedTo", out var g1) && g1.ValueKind == JsonValueKind.Object
                && g1.TryGetProperty("user", out user))
                hasUser = true;

            if (!hasUser)
                continue;
            var upn = user.GetStringOrNull("userPrincipalName") ?? user.GetStringOrNull("email");
            if (upn is null)
                continue;

            var roles = new List<string>();
            foreach (var r in perm.GetArrayOrEmpty("roles"))
                if (r.ValueKind == JsonValueKind.String)
                    roles.Add(r.GetString()!);
            grants.Add(new DriveGrant { Upn = upn, Roles = roles });
        }

        string? quickXor = null;
        if (data.TryGetProperty("file", out var file) && file.ValueKind == JsonValueKind.Object
            && file.TryGetProperty("hashes", out var hashes) && hashes.ValueKind == JsonValueKind.Object)
            quickXor = hashes.GetStringOrNull("quickXorHash");

        return new DriveItem
        {
            Id = data.GetStringOrNull("id") ?? "",
            Name = data.GetStringOrNull("name") ?? "",
            ParentPath = rel,
            IsFolder = data.TryGetProperty("folder", out _),
            Size = data.TryGetProperty("size", out var s) && s.ValueKind == JsonValueKind.Number ? s.GetInt64() : 0,
            QuickXorHash = quickXor,
            Grants = grants,
        };
    }
}

/// <summary>
/// A drive item copy planned for the target. <see cref="Action"/> is <c>copy</c>
/// (recreate folder / upload file) or <c>skip</c>.
/// </summary>
public sealed class PlannedDriveItem
{
    public string SourceId { get; set; } = "";
    public string Name { get; set; } = "";
    public string RelativePath { get; set; } = "";
    public bool IsFolder { get; set; }
    public long Size { get; set; }
    /// <summary>Source content hash (quickXorHash), used for delta skip; null if unavailable.</summary>
    public string? ContentHash { get; set; }
    public string Action { get; set; } = "copy";
    public string? Reason { get; set; }
    public List<DriveGrant> TargetGrants { get; set; } = new();
}
