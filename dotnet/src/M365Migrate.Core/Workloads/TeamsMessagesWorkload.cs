using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Mapping;
using M365Migrate.Core.Models;

namespace M365Migrate.Core.Workloads;

/// <summary>
/// Migrates Teams channel <b>message history</b> using Microsoft Graph "migration
/// mode": create a fresh team with teamCreationMode=migration, create channels in
/// migration mode, import each message with its original author + timestamp, then
/// completeMigration to unlock the team.
///
/// This is a separate path from the structure-only Teams workload (which enables
/// Teams on the already-migrated M365 group). A migration-mode team is created
/// fresh, so it has its own backing group; add members after migration completes.
///
/// After completeMigration it also adds the source team's owners and members to the
/// new team (mapped to target accounts), so it's usable immediately.
///
/// Not idempotent: completeMigration is one-shot, so re-running creates a duplicate
/// team — run once per team. Requires Teamwork.Migrate.All + TeamMember.ReadWrite.All
/// (and user read) with admin consent. Imports top-level messages and their threaded
/// replies (each with original author + timestamp).
/// </summary>
public sealed class TeamsMessagesWorkload
{
    private const string DefaultCreated = "2020-01-01T00:00:00.000Z";

    private readonly MigrationConfig _config;

    public TeamsMessagesWorkload(MigrationConfig config) => _config = config;

    /// <summary>Discover source teams (Teams-enabled M365 groups) and their channels.</summary>
    public async Task<List<MigratableTeam>> DiscoverAsync(GraphClient source, CancellationToken ct = default)
    {
        var groups = await source.GetAllAsync(
            "/groups?$select=id,displayName,mailNickname,description,createdDateTime,resourceProvisioningOptions&$top=999", ct);
        var teams = new List<MigratableTeam>();
        foreach (var group in groups)
        {
            if (!TeamsWorkload.GroupIsTeam(group))
                continue;
            var id = group.GetStringOrNull("id");
            if (id is null)
                continue;
            var channelsRaw = await source.GetAllAsync(
                $"/teams/{id}/channels?$select=id,displayName,description,membershipType,createdDateTime&$top=999", ct);
            var channels = channelsRaw.Select(MigratableChannel.FromGraph).ToList();
            teams.Add(MigratableTeam.FromGraph(group, channels));
        }
        return teams;
    }

    private static string? ParseTeamId(string? location)
    {
        if (location is null)
            return null;
        var open = location.IndexOf("('", StringComparison.Ordinal);
        var close = location.IndexOf("')", StringComparison.Ordinal);
        if (open >= 0 && close > open)
            return location.Substring(open + 2, close - (open + 2));
        return null;
    }

    private async Task<Dictionary<string, string>> SourceUserUpnMapAsync(GraphClient source, CancellationToken ct)
    {
        var raw = await source.GetAllAsync("/users?$select=id,userPrincipalName&$top=999", ct);
        var map = new Dictionary<string, string>();
        foreach (var u in raw)
        {
            var id = u.GetStringOrNull("id");
            var upn = u.GetStringOrNull("userPrincipalName");
            if (id is not null && upn is not null)
                map[id] = upn;
        }
        return map;
    }

    private string? ResolveAuthor(string? sourceUserId, IReadOnlyDictionary<string, string> sourceIdToUpn, IReadOnlyDictionary<string, string> targetUpnToId)
    {
        if (sourceUserId is null || !sourceIdToUpn.TryGetValue(sourceUserId, out var upn))
            return null;
        var targetUpn = TargetNaming.TargetUpn(upn, _config);
        return targetUpnToId.TryGetValue(targetUpn, out var id) ? id : null;
    }

    /// <summary>Build the migration-mode message body (author + timestamp + content) for a
    /// top-level message or a reply.</summary>
    internal static Dictionary<string, object> MigrationMessageBody(string authorId, ChannelMessage m, string fallbackCreated) =>
        new()
        {
            ["createdDateTime"] = m.CreatedDateTime ?? fallbackCreated,
            ["from"] = new Dictionary<string, object>
            {
                ["user"] = new Dictionary<string, object>
                {
                    ["id"] = authorId,
                    ["displayName"] = m.FromDisplayName ?? "",
                    ["userIdentityType"] = "aadUser",
                },
            },
            ["body"] = new Dictionary<string, object>
            {
                ["contentType"] = m.BodyContentType,
                ["content"] = m.BodyContent,
            },
        };

    private static async Task<List<string>> GroupUpnsAsync(GraphClient source, string url, CancellationToken ct)
    {
        var raw = await source.GetAllAsync(url, ct);
        var upns = new List<string>();
        foreach (var u in raw)
        {
            var upn = u.GetStringOrNull("userPrincipalName");
            if (upn is not null)
                upns.Add(upn);
        }
        return upns;
    }

    private bool TryTargetId(string sourceUpn, IReadOnlyDictionary<string, string> map, out string id)
    {
        id = "";
        if (map.TryGetValue(TargetNaming.TargetUpn(sourceUpn, _config), out var v))
        {
            id = v;
            return true;
        }
        return false;
    }

    private static async Task<bool> AddOneMemberAsync(GraphClient target, string teamId, string userId, bool owner, CancellationToken ct)
    {
        try
        {
            await target.PostJsonAsync($"/teams/{teamId}/members", new Dictionary<string, object>
            {
                ["@odata.type"] = "#microsoft.graph.aadUserConversationMember",
                ["roles"] = owner ? new[] { "owner" } : Array.Empty<string>(),
                ["user@odata.bind"] = $"https://graph.microsoft.com/v1.0/users('{userId}')",
            }, ct);
            return true;
        }
        catch (GraphException)
        {
            return false; // best-effort per member; the batch continues
        }
    }

    /// <summary>
    /// Add the source team's owners and members to the (now unlocked) target team, mapping each
    /// source UPN to its target account. Owners are added first (a team needs an owner); a user
    /// who is both owner and member is added once as owner. Returns (owners, members, errors).
    /// </summary>
    private async Task<(int Owners, int Members, int Errors)> AddTeamMembersAsync(
        GraphClient source, GraphClient target, string sourceGroupId, string teamId,
        IReadOnlyDictionary<string, string> targetUpnToId, CancellationToken ct)
    {
        var owners = await GroupUpnsAsync(source, $"/groups/{sourceGroupId}/owners?$select=id,userPrincipalName&$top=999", ct);
        var members = await GroupUpnsAsync(source, $"/groups/{sourceGroupId}/members?$select=id,userPrincipalName&$top=999", ct);
        var ownerUpns = new HashSet<string>(owners, StringComparer.OrdinalIgnoreCase);
        var addedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int addedOwners = 0, addedMembers = 0, errors = 0;

        foreach (var upn in owners)
        {
            ct.ThrowIfCancellationRequested();
            if (!TryTargetId(upn, targetUpnToId, out var id) || !addedIds.Add(id)) continue;
            if (await AddOneMemberAsync(target, teamId, id, owner: true, ct)) addedOwners++; else errors++;
        }
        foreach (var upn in members)
        {
            ct.ThrowIfCancellationRequested();
            if (ownerUpns.Contains(upn)) continue;
            if (!TryTargetId(upn, targetUpnToId, out var id) || !addedIds.Add(id)) continue;
            if (await AddOneMemberAsync(target, teamId, id, owner: false, ct)) addedMembers++; else errors++;
        }
        return (addedOwners, addedMembers, errors);
    }

    /// <summary>Create migration-mode teams, import message history, and complete migration.</summary>
    public async Task<List<WorkloadResult>> MigrateAsync(
        GraphClient source,
        GraphClient target,
        IEnumerable<MigratableTeam> teams,
        bool dryRun = true,
        CancellationToken ct = default)
    {
        var results = new List<WorkloadResult>();
        Dictionary<string, string>? sourceIdToUpn = null;
        Dictionary<string, string>? targetUpnToId = null;

        foreach (var team in teams)
        {
            ct.ThrowIfCancellationRequested();
            var record = new WorkloadResult { Name = team.MailNickname ?? team.DisplayName ?? team.GroupId };

            if (dryRun)
            {
                record.Status = "would-migrate";
                record.Detail["channels"] = team.Channels.Count.ToString();
                results.Add(record);
                continue;
            }

            sourceIdToUpn ??= await SourceUserUpnMapAsync(source, ct);
            targetUpnToId ??= await UsersWorkload.GetTargetUserIdsAsync(target, ct);

            string? newTeamId;
            try
            {
                var created = team.CreatedDateTime ?? DefaultCreated;
                var location = await target.PostJsonGetLocationAsync("/teams", new Dictionary<string, object>
                {
                    ["@microsoft.graph.teamCreationMode"] = "migration",
                    ["template@odata.bind"] = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
                    ["displayName"] = team.DisplayName ?? team.MailNickname ?? "Migrated team",
                    ["description"] = team.Description ?? "",
                    ["createdDateTime"] = created,
                }, ct);
                newTeamId = ParseTeamId(location);
            }
            catch (GraphException ex)
            {
                record.Status = "error";
                record.Reason = $"team:{ex.StatusCode}";
                results.Add(record);
                continue;
            }

            if (newTeamId is null)
            {
                record.Status = "error";
                record.Reason = "could not determine new team id";
                results.Add(record);
                continue;
            }

            // Map source channels -> target channel ids (General exists already).
            var targetChannels = await target.GetAllAsync($"/teams/{newTeamId}/channels?$select=id,displayName&$top=999", ct);
            var targetByName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in targetChannels)
            {
                var name = c.GetStringOrNull("displayName");
                var id = c.GetStringOrNull("id");
                if (name is not null && id is not null)
                    targetByName[name] = id;
            }

            var imported = 0;
            var replies = 0;
            var skipped = 0;
            var channelErrors = 0;

            foreach (var channel in team.Channels)
            {
                ct.ThrowIfCancellationRequested();
                string? targetChannelId;
                if (channel.IsDefault)
                {
                    targetByName.TryGetValue("General", out targetChannelId);
                }
                else if (!targetByName.TryGetValue(channel.DisplayName, out targetChannelId))
                {
                    try
                    {
                        var createdChannel = await target.PostJsonAsync($"/teams/{newTeamId}/channels", new Dictionary<string, object>
                        {
                            ["@microsoft.graph.channelCreationMode"] = "migration",
                            ["displayName"] = channel.DisplayName,
                            ["description"] = channel.Description ?? "",
                            ["membershipType"] = "standard",
                            ["createdDateTime"] = channel.CreatedDateTime ?? team.CreatedDateTime ?? DefaultCreated,
                        }, ct);
                        targetChannelId = createdChannel.GetStringOrNull("id");
                        if (targetChannelId is not null)
                            targetByName[channel.DisplayName] = targetChannelId;
                    }
                    catch (GraphException)
                    {
                        channelErrors++;
                        continue;
                    }
                }

                if (targetChannelId is null)
                    continue;

                // Import each top-level message.
                var messages = await source.GetAllAsync(
                    $"/teams/{team.GroupId}/channels/{channel.Id}/messages?$top=50", ct);
                foreach (var raw in messages)
                {
                    ct.ThrowIfCancellationRequested();
                    var m = ChannelMessage.FromGraph(raw);
                    if (m.MessageType != "message")
                        continue; // skip system messages
                    var authorId = ResolveAuthor(m.FromUserId, sourceIdToUpn, targetUpnToId);
                    if (authorId is null)
                    {
                        skipped++;
                        continue; // migration import requires a valid target author
                    }
                    var fallbackCreated = team.CreatedDateTime ?? DefaultCreated;
                    string? targetMsgId;
                    try
                    {
                        var createdMsg = await target.PostJsonAsync(
                            $"/teams/{newTeamId}/channels/{targetChannelId}/messages",
                            MigrationMessageBody(authorId, m, fallbackCreated), ct);
                        targetMsgId = createdMsg.GetStringOrNull("id");
                        imported++;
                    }
                    catch (GraphException)
                    {
                        skipped++;
                        continue; // can't thread replies under a parent that failed
                    }

                    // Import this message's threaded replies under the new parent message.
                    var sourceMsgId = raw.GetStringOrNull("id");
                    if (targetMsgId is null || sourceMsgId is null)
                        continue;
                    var sourceReplies = await source.GetAllAsync(
                        $"/teams/{team.GroupId}/channels/{channel.Id}/messages/{sourceMsgId}/replies?$top=50", ct);
                    foreach (var rawReply in sourceReplies)
                    {
                        ct.ThrowIfCancellationRequested();
                        var reply = ChannelMessage.FromGraph(rawReply);
                        if (reply.MessageType != "message")
                            continue;
                        var replyAuthor = ResolveAuthor(reply.FromUserId, sourceIdToUpn, targetUpnToId);
                        if (replyAuthor is null)
                        {
                            skipped++;
                            continue;
                        }
                        try
                        {
                            await target.PostJsonAsync(
                                $"/teams/{newTeamId}/channels/{targetChannelId}/messages/{targetMsgId}/replies",
                                MigrationMessageBody(replyAuthor, reply, m.CreatedDateTime ?? fallbackCreated), ct);
                            replies++;
                        }
                        catch (GraphException)
                        {
                            skipped++;
                        }
                    }
                }
            }

            // Unlock the team so it becomes a normal, usable team.
            try
            {
                await target.PostNoBodyAsync($"/teams/{newTeamId}/completeMigration", ct);
                record.Status = "ok";
            }
            catch (GraphException ex)
            {
                record.Status = "error";
                record.Reason = $"completeMigration:{ex.StatusCode}";
            }

            // After completeMigration, add the source team's owners + members (a migration-mode
            // team is created empty). Members can only be added once the team is unlocked.
            if (record.Status == "ok")
            {
                try
                {
                    var (owners, members, memberErrors) =
                        await AddTeamMembersAsync(source, target, team.GroupId, newTeamId, targetUpnToId, ct);
                    record.Detail["owners"] = owners.ToString();
                    record.Detail["members"] = members.ToString();
                    if (memberErrors > 0)
                        record.Detail["member_errors"] = memberErrors.ToString();
                }
                catch (GraphException ex)
                {
                    record.Detail["members"] = $"error:{ex.StatusCode}";
                }
            }

            record.Detail["messages"] = $"imported:{imported}";
            if (replies > 0)
                record.Detail["replies"] = replies.ToString();
            if (skipped > 0)
                record.Detail["skipped"] = skipped.ToString();
            if (channelErrors > 0)
                record.Detail["channel_errors"] = channelErrors.ToString();
            results.Add(record);
        }

        return results;
    }
}
