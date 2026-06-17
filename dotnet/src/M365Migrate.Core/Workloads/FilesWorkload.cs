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

    /// <summary>Return the (source, target) drive root paths for a user or a site.</summary>
    public (string Source, string Target) ResolveDriveRoots(string? user = null, string? site = null, string? targetSite = null)
    {
        if (!string.IsNullOrWhiteSpace(user))
            return ($"/users/{user}/drive", $"/users/{Rewrite(user)}/drive");
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
                Action = "copy",
                TargetGrants = item.Grants
                    .Select(g => new DriveGrant { Upn = Rewrite(g.Upn), Roles = g.Roles })
                    .ToList(),
            });
        }
        return planned;
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
