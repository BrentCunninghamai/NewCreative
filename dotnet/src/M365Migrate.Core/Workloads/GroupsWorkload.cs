using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Mapping;
using M365Migrate.Core.Models;

namespace M365Migrate.Core.Workloads;

/// <summary>
/// Groups workload: discover, plan, and sync groups + their membership and
/// ownership. Only security and Microsoft 365 groups can be provisioned via Graph;
/// dynamic groups are recreated with their membership rule.
/// </summary>
public sealed class GroupsWorkload
{
    private static readonly HashSet<string> CreatableKinds = new() { "security", "microsoft365" };

    private readonly MigrationConfig _config;

    public GroupsWorkload(MigrationConfig config) => _config = config;

    /// <summary>Read all groups from the source tenant, including members and owners.</summary>
    public async Task<List<SourceGroup>> DiscoverAsync(GraphClient source, CancellationToken ct = default)
    {
        var select = string.Join(",", SourceGroup.SelectFields);
        var raw = await source.GetAllAsync(
            $"/groups?$select={select}&$expand={SourceGroup.Expand}&$top=999", ct);
        return raw.Select(SourceGroup.FromGraph).ToList();
    }

    /// <summary>Return a map of lowercased target mailNickname -> group id.</summary>
    public static async Task<Dictionary<string, string>> GetTargetGroupIdsAsync(GraphClient target, CancellationToken ct = default)
    {
        var raw = await target.GetAllAsync("/groups?$select=id,mailNickname&$top=999", ct);
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in raw)
        {
            var nick = item.GetStringOrNull("mailNickname");
            var id = item.GetStringOrNull("id");
            if (nick is not null && id is not null)
                map[nick] = id;
        }
        return map;
    }

    private static async Task<HashSet<string>> GetDirectoryIdsAsync(GraphClient target, string relationship, string groupId, CancellationToken ct)
    {
        var raw = await target.GetAllAsync($"/groups/{groupId}/{relationship}?$select=id&$top=999", ct);
        var set = new HashSet<string>();
        foreach (var item in raw)
        {
            var id = item.GetStringOrNull("id");
            if (id is not null)
                set.Add(id);
        }
        return set;
    }

    private string Rewrite(string upn) => TargetNaming.TargetUpn(upn, _config);

    /// <summary>Build a group reconciliation plan without writing anything.</summary>
    public List<PlannedGroup> Plan(IEnumerable<SourceGroup> groups, IDictionary<string, string>? existingTargetGroups = null)
    {
        var existing = existingTargetGroups is null
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, string>(existingTargetGroups, StringComparer.OrdinalIgnoreCase);

        var planned = new List<PlannedGroup>();
        foreach (var group in groups)
        {
            var kind = group.Kind;
            var sourceNickname = group.MailNickname;
            // The target nickname carries the optional prefix/suffix so groups from
            // different source tenants don't collide in a shared target.
            var targetNickname = string.IsNullOrEmpty(sourceNickname)
                ? sourceNickname
                : TargetNaming.TargetMailNickname(sourceNickname, _config);
            string action;
            string? reason;

            if (!CreatableKinds.Contains(kind))
            {
                action = "skip";
                reason = $"{kind} group not provisionable via Graph";
            }
            else if (string.IsNullOrEmpty(sourceNickname))
            {
                action = "skip";
                reason = "group has no mailNickname";
            }
            else if (existing.ContainsKey(targetNickname!))
            {
                action = "exists";
                reason = null;
            }
            else
            {
                action = "create";
                reason = null;
            }

            planned.Add(new PlannedGroup
            {
                SourceId = group.Id,
                MailNickname = targetNickname,
                DisplayName = group.DisplayName,
                Kind = kind,
                Action = action,
                Reason = reason,
                Description = group.Description,
                TargetMemberUpns = group.MemberUpns.Select(Rewrite).ToList(),
                TargetOwnerUpns = group.OwnerUpns.Select(Rewrite).ToList(),
                IsDynamic = group.IsDynamic,
                MembershipRule = group.MembershipRule,
            });
        }
        return planned;
    }

    private string DirectoryObjectRef(GraphClient client, string objectId) =>
        $"{client.BaseUrl}/directoryObjects/{objectId}";

    /// <summary>Provision missing groups and reconcile membership + ownership in the target.</summary>
    public async Task<List<WorkloadResult>> SyncAsync(
        GraphClient target,
        IEnumerable<PlannedGroup> planned,
        bool dryRun = true,
        CancellationToken ct = default)
    {
        var targetGroups = await GetTargetGroupIdsAsync(target, ct);
        var targetUsers = await UsersWorkload.GetTargetUserIdsAsync(target, ct);
        var results = new List<WorkloadResult>();

        foreach (var p in planned)
        {
            var record = new WorkloadResult { Name = p.MailNickname ?? "", Detail = { ["kind"] = p.Kind } };

            if (p.Action == "skip")
            {
                record.Status = $"skipped:{p.Reason}";
                results.Add(record);
                continue;
            }

            // --- resolve or create the target group ---
            targetGroups.TryGetValue(p.MailNickname ?? "", out var groupId);
            var existingMembers = new HashSet<string>();

            if (p.Action == "create" && groupId is null)
            {
                if (dryRun)
                {
                    record.Detail["group"] = "would-create";
                }
                else
                {
                    try
                    {
                        var created = await target.PostJsonAsync("/groups", p.ToGraphBody(), ct);
                        groupId = created.GetStringOrNull("id");
                        record.Detail["group"] = "created";
                    }
                    catch (GraphException ex)
                    {
                        record.Detail["group"] = $"error:{ex.StatusCode}";
                        record.Status = "error";
                        results.Add(record);
                        continue;
                    }
                }
            }
            else
            {
                record.Detail["group"] = "exists";
                if (groupId is not null && !p.IsDynamic)
                    existingMembers = await GetDirectoryIdsAsync(target, "members", groupId, ct);
            }

            // --- reconcile membership (dynamic groups are rule-driven) ---
            if (p.IsDynamic)
            {
                record.Detail["members"] = "dynamic-rule";
            }
            else
            {
                await ReconcileAsync(
                    target, p.TargetMemberUpns, targetUsers, existingMembers, groupId,
                    relationship: "members", dryRun, record, "members", ct);
            }

            // --- reconcile owners (needed before a group can be Teams-enabled) ---
            if (p.TargetOwnerUpns.Count > 0)
            {
                var existingOwners = groupId is not null
                    ? await GetDirectoryIdsAsync(target, "owners", groupId, ct)
                    : new HashSet<string>();
                await ReconcileAsync(
                    target, p.TargetOwnerUpns, targetUsers, existingOwners, groupId,
                    relationship: "owners", dryRun, record, "owners", ct);
            }

            record.Status = "ok";
            results.Add(record);
        }

        return results;
    }

    private async Task ReconcileAsync(
        GraphClient target,
        IReadOnlyList<string> targetUpns,
        IReadOnlyDictionary<string, string> targetUsers,
        ISet<string> existing,
        string? groupId,
        string relationship,
        bool dryRun,
        WorkloadResult record,
        string detailKey,
        CancellationToken ct)
    {
        var resolved = new List<string>();
        var unresolved = 0;
        foreach (var upn in targetUpns)
        {
            if (targetUsers.TryGetValue(upn, out var uid))
                resolved.Add(uid);
            else
                unresolved++;
        }

        var toAdd = resolved.Where(uid => !existing.Contains(uid)).ToList();
        if (unresolved > 0)
            record.Detail[$"{detailKey}_unresolved"] = unresolved.ToString();

        if (toAdd.Count == 0)
        {
            record.Detail[detailKey] = "none";
        }
        else if (dryRun || groupId is null)
        {
            record.Detail[detailKey] = $"would-add:{toAdd.Count}";
        }
        else
        {
            var added = 0;
            var errors = 0;
            foreach (var uid in toAdd)
            {
                try
                {
                    await target.PostJsonAsync(
                        $"/groups/{groupId}/{relationship}/$ref",
                        new Dictionary<string, object> { ["@odata.id"] = DirectoryObjectRef(target, uid) },
                        ct);
                    added++;
                }
                catch (GraphException)
                {
                    errors++;
                }
            }
            record.Detail[detailKey] = $"added:{added}";
            if (errors > 0)
                record.Detail[$"{detailKey}_errors"] = errors.ToString();
        }
    }
}
