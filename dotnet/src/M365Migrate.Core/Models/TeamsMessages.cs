using System.Text.Json;
using M365Migrate.Core.Graph;

namespace M365Migrate.Core.Models;

/// <summary>A source channel to recreate in a migration-mode team.</summary>
public sealed class MigratableChannel
{
    public string Id { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string? Description { get; set; }
    public string MembershipType { get; set; } = "standard";
    public string? CreatedDateTime { get; set; }

    public bool IsDefault => DisplayName.Trim().Equals("General", StringComparison.OrdinalIgnoreCase);

    public static MigratableChannel FromGraph(JsonElement data) => new()
    {
        Id = data.GetStringOrNull("id") ?? "",
        DisplayName = data.GetStringOrNull("displayName") ?? "",
        Description = data.GetStringOrNull("description"),
        MembershipType = data.GetStringOrNull("membershipType") ?? "standard",
        CreatedDateTime = data.GetStringOrNull("createdDateTime"),
    };
}

/// <summary>A source team whose channels + message history will be migrated.</summary>
public sealed class MigratableTeam
{
    public string GroupId { get; set; } = "";
    public string? DisplayName { get; set; }
    public string? MailNickname { get; set; }
    public string? Description { get; set; }
    public string? CreatedDateTime { get; set; }
    public List<MigratableChannel> Channels { get; set; } = new();

    public static MigratableTeam FromGraph(JsonElement group, List<MigratableChannel> channels) => new()
    {
        GroupId = group.GetStringOrNull("id") ?? "",
        DisplayName = group.GetStringOrNull("displayName"),
        MailNickname = group.GetStringOrNull("mailNickname"),
        Description = group.GetStringOrNull("description"),
        CreatedDateTime = group.GetStringOrNull("createdDateTime"),
        Channels = channels,
    };
}

/// <summary>A single channel message to import.</summary>
public sealed class ChannelMessage
{
    public string? CreatedDateTime { get; set; }
    public string? FromUserId { get; set; }
    public string? FromDisplayName { get; set; }
    public string BodyContentType { get; set; } = "html";
    public string BodyContent { get; set; } = "";
    public string MessageType { get; set; } = "message";

    public static ChannelMessage FromGraph(JsonElement data)
    {
        var msg = new ChannelMessage
        {
            CreatedDateTime = data.GetStringOrNull("createdDateTime"),
            MessageType = data.GetStringOrNull("messageType") ?? "message",
        };
        if (data.TryGetProperty("from", out var from) && from.ValueKind == JsonValueKind.Object
            && from.TryGetProperty("user", out var user) && user.ValueKind == JsonValueKind.Object)
        {
            msg.FromUserId = user.GetStringOrNull("id");
            msg.FromDisplayName = user.GetStringOrNull("displayName");
        }
        if (data.TryGetProperty("body", out var body) && body.ValueKind == JsonValueKind.Object)
        {
            msg.BodyContentType = body.GetStringOrNull("contentType") ?? "html";
            msg.BodyContent = body.GetStringOrNull("content") ?? "";
        }
        return msg;
    }
}
