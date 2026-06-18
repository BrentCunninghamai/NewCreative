using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Mapping;
using M365Migrate.Core.Models;

namespace M365Migrate.Core.Workloads;

/// <summary>
/// Files workload: discover, plan, and copy OneDrive / SharePoint drive content.
/// Both are Graph drives of driveItems, so one set of primitives serves both.
/// Small files upload in a single PUT; larger files use a resumable upload
/// session. Only direct user grants that resolve in the target are reapplied.
/// </summary>
public sealed class FilesWorkload
{
    public const long SimpleUploadLimit = 4 * 1024 * 1024; // 4 MiB

    private readonly MigrationConfig _config;

    public FilesWorkload(MigrationConfig config) => _config = config;

    private string Rewrite(string upn) => TargetNaming.TargetUpn(upn, _config);

    /// <summary>
    /// True only when the user genuinely has no OneDrive: Graph returns <c>404</c> when
    /// the drive doesn't exist and (under app-only auth) can't be auto-provisioned. Other
    /// failures — including <c>400 notSupported</c> — are NOT "no drive" (a user with a
    /// provisioned OneDrive can still hit those), so they must surface, not be skipped.
    /// </summary>
    public static bool IsDriveNotProvisioned(GraphException ex) => ex.StatusCode == 404;

    /// <summary>
    /// A human-readable hint for why reading a user's OneDrive failed, so the operator can
    /// act instead of seeing a raw Graph error. Distinguishes a missing drive (404) from
    /// the common misconfigurations that return <c>400 notSupported</c> / <c>403</c>.
    /// </summary>
    public static string DriveErrorHint(GraphException ex)
    {
        if (ex.StatusCode == 404)
            return "No OneDrive provisioned for this user — nothing to copy.";
        if (ex.StatusCode == 400 && ex.Message.Contains("notSupported", StringComparison.OrdinalIgnoreCase))
            return "Graph returned 'notSupported' reading OneDrive. The user HAS a drive, so this is usually a config issue: "
                 + "grant the SOURCE app Files.ReadWrite.All + Sites.ReadWrite.All (Application) and Grant admin consent; "
                 + "if the tenant is multi-geo, the drive may live in another geo.";
        if (ex.StatusCode is 401 or 403)
            return "Access denied reading OneDrive — grant Files.ReadWrite.All + Sites.ReadWrite.All on the source app and admin-consent.";
        return ex.Message;
    }

    /// <summary>
    /// Return the (source, target) drive root paths for a user or a site.
    /// <paramref name="targetUserOverride"/> overrides the domain-rewrite for the
    /// target OneDrive owner — needed when the target UPN isn't a clean rewrite of
    /// the source (e.g. a user already partly migrated by Microsoft's cross-tenant
    /// orchestrator / cross-tenant sync).
    /// </summary>
    public (string Source, string Target) ResolveDriveRoots(string? user = null, string? site = null, string? targetSite = null, string? targetUserOverride = null)
    {
        if (!string.IsNullOrWhiteSpace(user))
        {
            var targetUser = string.IsNullOrWhiteSpace(targetUserOverride) ? Rewrite(user) : targetUserOverride.Trim();
            return ($"/users/{user}/drive", $"/users/{targetUser}/drive");
        }
        if (!string.IsNullOrWhiteSpace(site))
        {
            if (string.IsNullOrWhiteSpace(targetSite))
                throw new ArgumentException("A site migration requires a target site id.");
            return ($"/sites/{site}/drive", $"/sites/{targetSite}/drive");
        }
        throw new ArgumentException("Specify either a user (OneDrive) or a site (SharePoint).");
    }

    /// <summary>Walk a drive breadth-first, returning folders before their children.</summary>
    public async Task<List<DriveItem>> DiscoverDriveItemsAsync(GraphClient client, string driveRoot, CancellationToken ct = default)
    {
        var items = new List<DriveItem>();
        var queue = new Queue<string>();
        queue.Enqueue("root/children");
        while (queue.Count > 0)
        {
            var segment = queue.Dequeue();
            var raw = await client.GetAllAsync($"{driveRoot}/{segment}?$expand=permissions&$top=200", ct);
            foreach (var element in raw)
            {
                var item = DriveItem.FromGraph(element);
                items.Add(item);
                if (item.IsFolder)
                    queue.Enqueue($"items/{item.Id}/children");
            }
        }
        return items;
    }

    /// <summary>Build a copy plan; the migrate phase picks the upload method by size.</summary>
    public List<PlannedDriveItem> Plan(IEnumerable<DriveItem> items)
    {
        var planned = new List<PlannedDriveItem>();
        foreach (var item in items)
        {
            planned.Add(new PlannedDriveItem
            {
                SourceId = item.Id,
                Name = item.Name,
                RelativePath = item.RelativePath,
                IsFolder = item.IsFolder,
                Size = item.Size,
                ContentHash = item.QuickXorHash,
                Action = "copy",
                TargetGrants = item.Grants
                    .Select(g => new DriveGrant { Upn = Rewrite(g.Upn), Roles = g.Roles })
                    .ToList(),
            });
        }
        return planned;
    }

    /// <summary>
    /// Index a target drive's items for delta comparison: file paths→content hash
    /// (quickXorHash), plus the set of folder paths. Used to skip items already present
    /// (pre-sync / repeatable sync), so each pass only moves what's new or changed.
    /// </summary>
    public static (Dictionary<string, string?> Files, HashSet<string> Folders) IndexTarget(IEnumerable<DriveItem> targetItems)
    {
        var files = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        var folders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in targetItems)
        {
            if (t.IsFolder) folders.Add(t.RelativePath);
            else files[t.RelativePath] = t.QuickXorHash;
        }
        return (files, folders);
    }

    /// <summary>
    /// Mark planned items already present in the target as skip, in place: a folder that
    /// exists ("exists"), or a file at the same path whose **content hash matches** the source
    /// ("unchanged"). Hash (quickXorHash) is the content signal — a file edited to the same
    /// byte length is NOT skipped. When either side lacks a hash, the file is copied (never
    /// skipped on uncertain equality). Returns how many files were marked unchanged.
    /// </summary>
    public static int MarkUnchanged(
        IEnumerable<PlannedDriveItem> planned,
        IReadOnlyDictionary<string, string?> targetFiles,
        IReadOnlySet<string> targetFolders)
    {
        var unchanged = 0;
        foreach (var p in planned)
        {
            if (p.Action != "copy") continue;
            if (p.IsFolder)
            {
                if (targetFolders.Contains(p.RelativePath)) { p.Action = "skip"; p.Reason = "exists"; }
            }
            else if (p.ContentHash is not null
                && targetFiles.TryGetValue(p.RelativePath, out var targetHash)
                && targetHash is not null
                && string.Equals(targetHash, p.ContentHash, StringComparison.OrdinalIgnoreCase))
            {
                p.Action = "skip"; p.Reason = "unchanged";
                unchanged++;
            }
        }
        return unchanged;
    }

    private static string ParentSegment(string relativePath)
    {
        var slash = relativePath.LastIndexOf('/');
        return slash < 0 ? "root/children" : $"root:/{relativePath[..slash]}:/children";
    }

    private static string GrantRole(IEnumerable<string> roles) =>
        roles.Any(r => r is "write" or "owner") ? "write" : "read";

    /// <summary>Recreate folders, upload file content, and reapply resolvable grants.</summary>
    public async Task<List<WorkloadResult>> MigrateDriveItemsAsync(
        GraphClient source,
        GraphClient target,
        IEnumerable<PlannedDriveItem> planned,
        string sourceRoot,
        string targetRoot,
        bool dryRun = true,
        CancellationToken ct = default)
    {
        var targetUsers = await UsersWorkload.GetTargetUserIdsAsync(target, ct);
        var results = new List<WorkloadResult>();

        foreach (var p in planned)
        {
            var record = new WorkloadResult { Name = p.RelativePath, Detail = { ["type"] = p.IsFolder ? "folder" : "file" } };

            if (p.Action != "copy")
            {
                record.Status = $"skipped:{p.Action}";
                record.Reason = p.Reason;
                results.Add(record);
                continue;
            }

            try
            {
                if (p.IsFolder)
                {
                    if (dryRun)
                    {
                        record.Status = "would-create-folder";
                    }
                    else
                    {
                        await target.PostJsonAsync(
                            $"{targetRoot}/{ParentSegment(p.RelativePath)}",
                            new Dictionary<string, object>
                            {
                                ["name"] = p.Name,
                                ["folder"] = new Dictionary<string, object>(),
                                ["@microsoft.graph.conflictBehavior"] = "replace",
                            },
                            ct);
                        record.Status = "folder-created";
                    }
                }
                else
                {
                    var large = p.Size > SimpleUploadLimit;
                    if (dryRun)
                    {
                        record.Status = large ? "would-upload-session" : "would-upload";
                    }
                    else
                    {
                        var data = await source.GetBytesAsync($"{sourceRoot}/items/{p.SourceId}/content", ct);
                        if (large)
                        {
                            await target.UploadLargeFileAsync(
                                $"{targetRoot}/root:/{p.RelativePath}:/createUploadSession", data, ct: ct);
                            record.Status = "uploaded-session";
                        }
                        else
                        {
                            await target.PutBytesAsync($"{targetRoot}/root:/{p.RelativePath}:/content", data, ct);
                            record.Status = "uploaded";
                        }
                    }
                }
            }
            catch (GraphException ex)
            {
                record.Status = "error";
                record.Reason = ex.StatusCode.ToString();
                results.Add(record);
                continue;
            }

            // --- reapply direct user grants that resolve to a target account ---
            var resolvable = p.TargetGrants.Where(g => targetUsers.ContainsKey(g.Upn)).ToList();
            if (resolvable.Count > 0)
            {
                if (dryRun)
                {
                    record.Detail["grants"] = $"would-grant:{resolvable.Count}";
                }
                else
                {
                    var granted = 0;
                    foreach (var g in resolvable)
                    {
                        try
                        {
                            await target.PostJsonAsync(
                                $"{targetRoot}/root:/{p.RelativePath}:/invite",
                                new Dictionary<string, object>
                                {
                                    ["recipients"] = new[] { new Dictionary<string, object> { ["email"] = g.Upn } },
                                    ["roles"] = new[] { GrantRole(g.Roles) },
                                    ["requireSignIn"] = true,
                                    ["sendInvitation"] = false,
                                },
                                ct);
                            granted++;
                        }
                        catch (GraphException)
                        {
                            // best-effort grant reapply
                        }
                    }
                    record.Detail["grants"] = $"granted:{granted}";
                }
            }

            results.Add(record);
        }

        return results;
    }
}
