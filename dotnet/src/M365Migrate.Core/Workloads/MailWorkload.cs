using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Mapping;
using M365Migrate.Core.Models;

namespace M365Migrate.Core.Workloads;

/// <summary>
/// Mail content workload: copy a user's mail folders and messages source -> target,
/// item by item, via Microsoft Graph. Messages are copied in full-fidelity MIME
/// (headers, body, attachments preserved). Re-runs are idempotent: a message whose
/// internetMessageId already exists in the target folder is skipped.
///
/// This is per-user (a source UPN), like the files workload. The target mailbox is
/// the rewritten/affixed UPN. Requires Mail.ReadWrite application permission with
/// admin consent in both tenants. Large mailboxes take time — Graph throttles, and
/// the client retries with backoff.
/// </summary>
public sealed class MailWorkload
{
    private readonly MigrationConfig _config;

    public MailWorkload(MigrationConfig config) => _config = config;

    /// <summary>Return the (source, target) user resource refs for a source UPN.</summary>
    public (string Source, string Target) ResolveUserRefs(string sourceUpn)
    {
        if (string.IsNullOrWhiteSpace(sourceUpn))
            throw new ArgumentException("A mail migration needs a source user UPN.");
        return ($"/users/{sourceUpn}", $"/users/{TargetNaming.TargetUpn(sourceUpn, _config)}");
    }

    private const string FolderSelect = "id,displayName,parentFolderId,totalItemCount,childFolderCount";

    /// <summary>Walk a mailbox's folder tree breadth-first (parents before children).</summary>
    public async Task<List<MailFolderInfo>> DiscoverFoldersAsync(GraphClient client, string userRef, CancellationToken ct = default)
    {
        var folders = new List<MailFolderInfo>();
        var queue = new Queue<(string Segment, string ParentPath)>();
        queue.Enqueue(("mailFolders", ""));
        while (queue.Count > 0)
        {
            var (segment, parentPath) = queue.Dequeue();
            var raw = await client.GetAllAsync($"{userRef}/{segment}?$select={FolderSelect}&$top=100", ct);
            foreach (var element in raw)
            {
                var info = MailFolderInfo.FromGraph(element, parentPath);
                folders.Add(info);
                if (info.ChildFolderCount > 0)
                    queue.Enqueue(($"mailFolders/{info.Id}/childFolders", info.Path));
            }
        }
        return folders;
    }

    /// <summary>Build a plan (read-only): one row per source folder with its item count.</summary>
    public List<PlannedMailFolder> Plan(IEnumerable<MailFolderInfo> folders) =>
        folders.Select(f => new PlannedMailFolder
        {
            SourceId = f.Id,
            DisplayName = f.DisplayName,
            Path = f.Path,
            ItemCount = f.TotalItemCount,
            Action = "copy",
        }).ToList();

    /// <summary>Copy mail folders and their messages to the target mailbox.</summary>
    public async Task<List<WorkloadResult>> MigrateAsync(
        GraphClient source,
        GraphClient target,
        string sourceUserRef,
        string targetUserRef,
        IEnumerable<PlannedMailFolder> planned,
        bool dryRun = true,
        CancellationToken ct = default)
    {
        var results = new List<WorkloadResult>();

        // Map of existing target folder path -> id, seeded from the target mailbox so
        // well-known folders (Inbox, Sent Items, ...) are reused, not duplicated.
        var targetFolders = await DiscoverFoldersAsync(target, targetUserRef, ct);
        var targetByPath = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var f in targetFolders)
            targetByPath[f.Path] = f.Id;

        foreach (var folder in planned)
        {
            ct.ThrowIfCancellationRequested();
            var record = new WorkloadResult { Name = folder.Path };

            if (dryRun)
            {
                record.Status = "would-copy";
                record.Detail["messages"] = $"~{folder.ItemCount}";
                results.Add(record);
                continue;
            }

            // --- find or create the target folder ---
            string targetFolderId;
            try
            {
                if (!targetByPath.TryGetValue(folder.Path, out targetFolderId!))
                {
                    var created = targetByPath.TryGetValue(folder.ParentPath, out var parentId) && folder.ParentPath.Length > 0
                        ? await target.PostJsonAsync($"{targetUserRef}/mailFolders/{parentId}/childFolders",
                            new Dictionary<string, object> { ["displayName"] = folder.DisplayName }, ct)
                        : await target.PostJsonAsync($"{targetUserRef}/mailFolders",
                            new Dictionary<string, object> { ["displayName"] = folder.DisplayName }, ct);
                    targetFolderId = created.GetStringOrNull("id") ?? "";
                    targetByPath[folder.Path] = targetFolderId;
                }
            }
            catch (GraphException ex)
            {
                record.Status = "error";
                record.Reason = $"folder:{ex.StatusCode}";
                results.Add(record);
                continue;
            }

            // --- internetMessageIds already in the target folder (for idempotency) ---
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var m in await target.GetAllAsync($"{targetUserRef}/mailFolders/{targetFolderId}/messages?$select=internetMessageId&$top=100", ct))
            {
                var mid = m.GetStringOrNull("internetMessageId");
                if (mid is not null)
                    existing.Add(mid);
            }

            // --- copy each source message not already present ---
            var copied = 0;
            var skipped = 0;
            var errors = 0;
            var sourceMessages = await source.GetAllAsync($"{sourceUserRef}/mailFolders/{folder.SourceId}/messages?$select=id,internetMessageId&$top=100", ct);
            foreach (var msg in sourceMessages)
            {
                ct.ThrowIfCancellationRequested();
                var imId = msg.GetStringOrNull("internetMessageId");
                if (imId is not null && existing.Contains(imId))
                {
                    skipped++;
                    continue;
                }
                var msgId = msg.GetStringOrNull("id");
                if (msgId is null)
                    continue;
                try
                {
                    var mime = await source.GetBytesAsync($"{sourceUserRef}/messages/{msgId}/$value", ct);
                    await target.PostMimeMessageAsync($"{targetUserRef}/mailFolders/{targetFolderId}/messages", mime, ct);
                    copied++;
                }
                catch (GraphException)
                {
                    errors++;
                }
            }

            record.Status = "ok";
            record.Detail["copied"] = copied.ToString();
            if (skipped > 0)
                record.Detail["skipped"] = skipped.ToString();
            if (errors > 0)
                record.Detail["errors"] = errors.ToString();
            results.Add(record);
        }

        return results;
    }
}
