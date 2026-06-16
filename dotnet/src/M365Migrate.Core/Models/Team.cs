using System.Text.Json;
using M365Migrate.Core.Graph;

namespace M365Migrate.Core.Models;

/// <summary>A channel within a team.</summary>
public sealed class Channel
{
    public string? Id { get; set; }
    public string DisplayName { get; set; } = "";
    public string? Description { get; set; }
    public string MembershipType { get; set; } = "standard"; // standard | private | shared

    /// <summary>True for the auto-created primary channel (named "General").</summary>
    public bool IsDefault => DisplayName.Trim().Equals("General", StringComparison.OrdinalIgnoreCase);

    public static Channel FromGraph(JsonElement data) => new()
    {
        Id = data.GetStringOrNull("id"),
        DisplayName = data.GetStringOrNull("displayName") ?? "",
        Description = data.GetStringOrNull("description"),
        MembershipType = data.GetStringOrNull("membershipType") ?? "standard",
    };
}

/// <summary>
/// A team (Teams-enabled M365 group) as read from the source tenant. Team
/// membership is the backing group's membership (groups workload) and channel
/// files live in the team's SharePoint library (files workload); this captures
/// the Teams-specific layer.
/// </summary>
public sealed class SourceTeam
{
    /// <summary>resourceProvisioningOptions contains "Team" iff the group is Teams-enabled.</summary>
    public static readonly string[] GroupSelectFields =
    {
        "id", "displayName", "mailNickname", "description", "resourceProvisioningOptions",
    };

    public string Id { get; set; } = "";
    public string? DisplayName { get; set; }
    public string? MailNickname { get; set; }
    public string? Description { get; set; }
    public List<Channel> Channels { get; set; } = new();

    public static SourceTeam FromGraph(JsonElement group, List<Channel> channels) => new()
    {
        Id = group.GetStringOrNull("id") ?? "",
        DisplayName = group.GetStringOrNull("displayName"),
        MailNickname = group.GetStringOrNull("mailNickname"),
        Description = group.GetStringOrNull("description"),
        Channels = channels,
    };
}

/// <summary>
/// A channel recreation planned for the target team. <see cref="Action"/> is
/// <c>create</c> (recreate a standard channel) or <c>skip</c> (default General, or
/// a private/shared channel not yet supported).
/// </summary>
public sealed class PlannedChannel
{
    public string DisplayName { get; set; } = "";
    public string? Description { get; set; }
    public string MembershipType { get; set; } = "standard";
    public string Action { get; set; } = "create";
    public string? Reason { get; set; }
}

/// <summary>
/// A team provisioning planned for the target tenant. <see cref="Action"/> is
/// <c>provision</c> (enable Teams on the matching group + recreate channels) or
/// <c>skip</c> (no matching target group yet).
/// </summary>
public sealed class PlannedTeam
{
    public string SourceId { get; set; } = "";
    public string? MailNickname { get; set; }
    public string? DisplayName { get; set; }
    public string Action { get; set; } = "provision";
    public string? Reason { get; set; }
    public List<PlannedChannel> Channels { get; set; } = new();
}
