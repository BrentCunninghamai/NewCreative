using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;

namespace M365Migrate.Core.Workloads;

/// <summary>A SharePoint/Teams document library (a Graph drive on a site).</summary>
public sealed record SiteDrive(string Id, string Name);

/// <summary>
/// SharePoint site content: resolve a site by URL, list its document libraries, and copy each
/// library's files to the matching library on a target site. The actual copy reuses
/// <see cref="FilesWorkload"/> (delta-aware, large-file sessions, direct-grant reapply) by
/// addressing each library as a Graph drive (<c>/drives/{id}</c>). A Microsoft Team's files
/// live in its SharePoint site, so this also moves Teams files.
/// </summary>
public sealed class SharePointWorkload
{
    private readonly MigrationConfig _config;

    public SharePointWorkload(MigrationConfig config) => _config = config;

    /// <summary>Split a SharePoint site URL into (hostname, server-relative path).</summary>
    public static (string Host, string Path) ParseSiteUrl(string url)
    {
        var u = new Uri(url);
        return (u.Host, u.AbsolutePath.TrimEnd('/'));
    }

    /// <summary>
    /// Resolve a site URL (e.g. <c>https://contoso.sharepoint.com/sites/Marketing</c>) to its
    /// Graph site id. If the value isn't a URL it's assumed to already be a site id.
    /// </summary>
    public async Task<string> ResolveSiteIdAsync(GraphClient client, string siteUrlOrId, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(siteUrlOrId))
            throw new ArgumentException("Provide a SharePoint site URL (or site id).");
        var s = siteUrlOrId.Trim();
        if (!s.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            return s;

        var (host, path) = ParseSiteUrl(s);
        var url = string.IsNullOrEmpty(path) ? $"/sites/{host}" : $"/sites/{host}:{path}";
        var el = await client.GetAsync(url, ct);
        return el.GetStringOrNull("id")
            ?? throw new InvalidOperationException($"Could not resolve SharePoint site '{s}'.");
    }

    /// <summary>List the document libraries (drives) on a site.</summary>
    public async Task<List<SiteDrive>> ListDrivesAsync(GraphClient client, string siteId, CancellationToken ct = default)
    {
        var raw = await client.GetAllAsync($"/sites/{siteId}/drives?$select=id,name", ct);
        var drives = new List<SiteDrive>();
        foreach (var d in raw)
        {
            var id = d.GetStringOrNull("id");
            if (id is not null)
                drives.Add(new SiteDrive(id, d.GetStringOrNull("name") ?? id));
        }
        return drives;
    }

    /// <summary>
    /// Match source libraries to target libraries by name. Returns pairs to copy and the names
    /// of source libraries with no target counterpart (reported, not copied).
    /// </summary>
    public static (List<(SiteDrive Source, SiteDrive Target)> Pairs, List<string> Unmatched) MatchDrives(
        IReadOnlyList<SiteDrive> source, IReadOnlyList<SiteDrive> target)
    {
        var byName = new Dictionary<string, SiteDrive>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in target)
            byName[t.Name] = t;

        var pairs = new List<(SiteDrive, SiteDrive)>();
        var unmatched = new List<string>();
        foreach (var s in source)
        {
            if (byName.TryGetValue(s.Name, out var t))
                pairs.Add((s, t));
            else
                unmatched.Add(s.Name);
        }
        return (pairs, unmatched);
    }

    /// <summary>The Graph drive-root ref for a library id, for use with <see cref="FilesWorkload"/>.</summary>
    public static string DriveRoot(string driveId) => $"/drives/{driveId}";

    /// <summary>
    /// Derive a tenant's OneDrive ("-my") host from its SPO root web URL, e.g.
    /// <c>https://contoso.sharepoint.com</c> → <c>contoso-my.sharepoint.com</c>. Works for
    /// any cloud (sharepoint.com / .us / .de) by inserting "-my" after the tenant label.
    /// </summary>
    public static string MyHostFromRoot(string rootWebUrl)
    {
        var host = new Uri(rootWebUrl).Host;            // contoso.sharepoint.com
        var dot = host.IndexOf('.');
        return dot < 0 ? host + "-my" : host[..dot] + "-my" + host[dot..];
    }

    /// <summary>Read the tenant's SPO root and return its OneDrive host (cached by the caller).</summary>
    public static async Task<string> GetMyHostAsync(GraphClient client, CancellationToken ct = default)
    {
        var root = await client.GetAsync("/sites/root?$select=webUrl", ct);
        var url = root.GetStringOrNull("webUrl")
            ?? throw new InvalidOperationException("Could not read the tenant's SharePoint root URL.");
        return MyHostFromRoot(url);
    }

    /// <summary>
    /// Build a user's OneDrive personal-site URL from the tenant OneDrive host and their UPN
    /// (the personal-site segment is the UPN with each non-alphanumeric char replaced by '_').
    /// Resolving this URL reaches the drive even in multi-geo, where /users/{id}/drive fails.
    /// </summary>
    public static string OneDriveUrl(string myHost, string upn)
    {
        var seg = new string(upn.Select(c => char.IsLetterOrDigit(c) ? c : '_').ToArray());
        return $"https://{myHost}/personal/{seg}";
    }
}
