using System.Text.Json;
using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Mapping;
using M365Migrate.Core.Models;

namespace M365Migrate.Core.Workloads;

/// <summary>
/// Teams workload: discover, plan, and provision Microsoft Teams + channels. A
/// team is a Teams-enabled M365 group, so this composes with the others: team
/// membership rides on the group (groups workload) and channel files on SharePoint
/// (files workload). This owns enabling Teams and recreating standard channels.
/// Run after groups sync so the backing group exists to be Teams-enabled.
/// </summary>
public sealed class TeamsWorkload
{
    private readonly MigrationConfig _config;

    public TeamsWorkload(MigrationConfig config) => _config = config;

    /// <summary>True if a Graph group payload represents a Teams-enabled M365 group.</summary>
    public static bool GroupIsTeam(JsonElement group)
    {
        foreach (var option in group.GetArrayOrEmpty("resourceProvisioningOptions"))
            if (option.ValueKind == JsonValueKind.String
                && option.GetString()!.Equals("Team", StringComparison.OrdinalIgnoreCase))
                return true;
        return false;
    }

    /// <summary>Read all teams from the source tenant, including their channels.</summary>
    public async Task<List<SourceTeam>> DiscoverAsync(GraphClient source, CancellationToken ct = default)
    {
        var select = string.Join(",", SourceTeam.GroupSelectFields);
        var groups = await source.GetAllAsync($"/groups?$select={select}&$top=999", ct);
        var teams = new List<SourceTeam>();
        foreach (var group in groups)
        {
            if (!GroupIsTeam(group))
                continue;
            var id = group.GetStringOrNull("id");
            if (id is null)
                continue;
            var channelsRaw = await source.GetAllAsync($"/teams/{id}/channels?$top=999", ct);
            var channels = channelsRaw.Select(Channel.FromGraph).ToList();
            teams.Add(SourceTeam.FromGraph(group, channels));
        }
        return teams;
    }

    /// <summary>Return lowercased mailNicknames of target groups that are already teams.</summary>
    public static async Task<HashSet<string>> GetTargetTeamNicknamesAsync(GraphClient target, CancellationToken ct = default)
    {
        var raw = await target.GetAllAsync("/groups?$select=mailNickname,resourceProvisioningOptions&$top=999", ct);
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var group in raw)
        {
            var nick = group.GetStringOrNull("mailNickname");
            if (nick is not null && GroupIsTeam(group))
                set.Add(nick);
        }
        return set;
    }

    private static PlannedChannel PlanChannel(Channel channel)
    {
        string action;
        string? reason;
        if (channel.IsDefault)
        {
            action = "skip";
            reason = "default channel created automatically with the team";
        }
        else if (channel.MembershipType != "standard")
        {
            action = "skip";
            reason = $"{channel.MembershipType} channel not yet supported";
        }
        else
        {
            action = "create";
            reason = null;
        }
        return new PlannedChannel
        {
            DisplayName = channel.DisplayName,
            Description = channel.Description,
            MembershipType = channel.MembershipType,
            Action = action,
            Reason = reason,
        };
    }

    /// <summary>Build a team provisioning plan without writing anything.</summary>
    public List<PlannedTeam> Plan(IEnumerable<SourceTeam> teams, IDictionary<string, string>? existingTargetGroups = null)
    {
        var existing = existingTargetGroups is null
            ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, string>(existingTargetGroups, StringComparer.OrdinalIgnoreCase);

        var planned = new List<PlannedTeam>();
        foreach (var team in teams)
        {
            var sourceNickname = team.MailNickname;
            // Match the backing group by its target nickname (prefix/suffix applied),
            // the same name the groups workload provisions.
            var targetNickname = string.IsNullOrEmpty(sourceNickname)
                ? sourceNickname
                : TargetNaming.TargetMailNickname(sourceNickname, _config);
            string action;
            string? reason;
            if (string.IsNullOrEmpty(sourceNickname))
            {
                action = "skip";
                reason = "team has no mailNickname";
            }
            else if (!existing.ContainsKey(targetNickname!))
            {
                action = "skip";
                reason = "target M365 group missing — run groups sync first";
            }
            else
            {
                action = "provision";
                reason = null;
            }

            planned.Add(new PlannedTeam
            {
                SourceId = team.Id,
                MailNickname = targetNickname,
                DisplayName = team.DisplayName,
                Action = action,
                Reason = reason,
                Channels = team.Channels.Select(PlanChannel).ToList(),
            });
        }
        return planned;
    }

    /// <summary>Enable Teams on matching target groups and recreate standard channels.</summary>
    public async Task<List<WorkloadResult>> MigrateAsync(
        GraphClient target,
        IEnumerable<PlannedTeam> planned,
        bool dryRun = true,
        CancellationToken ct = default)
    {
        var targetGroups = await GroupsWorkload.GetTargetGroupIdsAsync(target, ct);
        var targetTeams = await GetTargetTeamNicknamesAsync(target, ct);
        var results = new List<WorkloadResult>();

        foreach (var p in planned)
        {
            var record = new WorkloadResult { Name = p.MailNickname ?? "" };

            if (p.Action == "skip")
            {
                record.Status = $"skipped:{p.Reason}";
                results.Add(record);
                continue;
            }

            var nickname = p.MailNickname ?? "";
            if (!targetGroups.TryGetValue(nickname, out var teamId))
            {
                record.Status = "skipped:target group missing";
                results.Add(record);
                continue;
            }

            var alreadyTeam = targetTeams.Contains(nickname);

            // --- enable Teams on the group (team id == group id) ---
            if (alreadyTeam)
            {
                record.Detail["team"] = "exists";
            }
            else if (dryRun)
            {
                record.Detail["team"] = "would-enable";
            }
            else
            {
                try
                {
                    await target.PutJsonAsync($"/groups/{teamId}/team", new Dictionary<string, object>(), ct);
                    record.Detail["team"] = "enabled";
                    alreadyTeam = true;
                }
                catch (GraphException ex)
                {
                    record.Detail["team"] = $"error:{ex.StatusCode}";
                    record.Status = "error";
                    results.Add(record);
                    continue;
                }
            }

            // --- recreate standard channels ---
            var toCreate = p.Channels.Where(c => c.Action == "create").ToList();
            if (toCreate.Count == 0)
            {
                record.Detail["channels"] = "none";
                record.Status = "ok";
                results.Add(record);
                continue;
            }
            if (dryRun)
            {
                record.Detail["channels"] = $"would-create:{toCreate.Count}";
                record.Status = "ok";
                results.Add(record);
                continue;
            }

            var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in await target.GetAllAsync($"/teams/{teamId}/channels?$top=999", ct))
            {
                var name = c.GetStringOrNull("displayName");
                if (name is not null)
                    existingNames.Add(name);
            }

            var created = 0;
            var errors = 0;
            foreach (var channel in toCreate)
            {
                if (existingNames.Contains(channel.DisplayName))
                    continue;
                try
                {
                    await target.PostJsonAsync(
                        $"/teams/{teamId}/channels",
                        new Dictionary<string, object>
                        {
                            ["displayName"] = channel.DisplayName,
                            ["description"] = channel.Description ?? "",
                            ["membershipType"] = "standard",
                        },
                        ct);
                    created++;
                }
                catch (GraphException)
                {
                    errors++;
                }
            }
            record.Detail["channels"] = $"created:{created}";
            if (errors > 0)
                record.Detail["channels_errors"] = errors.ToString();
            record.Status = "ok";
            results.Add(record);
        }

        return results;
    }
}
