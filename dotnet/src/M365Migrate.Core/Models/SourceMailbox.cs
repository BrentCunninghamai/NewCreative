using System.Text.Json;

namespace M365Migrate.Core.Models;

/// <summary>
/// A user's mailbox configuration as read from the source tenant. Only the
/// writable <c>mailboxSettings</c> subset is captured; mailbox *content* (mail,
/// calendar, contacts) is out of scope for the Graph layer.
/// </summary>
public sealed class SourceMailbox
{
    /// <summary>
    /// The writable subset of Graph mailboxSettings. Read-only fields (e.g.
    /// userPurpose) are intentionally excluded so they are never PATCHed back.
    /// </summary>
    public static readonly string[] SettableFields =
    {
        "automaticRepliesSetting", "timeZone", "language", "workingHours",
        "dateFormat", "timeFormat", "delegateMeetingMessageDeliveryOptions",
    };

    public string UserPrincipalName { get; set; } = "";
    public Dictionary<string, JsonElement> Settings { get; set; } = new();

    public static SourceMailbox FromGraph(string upn, JsonElement data)
    {
        var settings = new Dictionary<string, JsonElement>();
        if (data.ValueKind == JsonValueKind.Object)
            foreach (var field in SettableFields)
                if (data.TryGetProperty(field, out var value) && value.ValueKind != JsonValueKind.Null)
                    settings[field] = value.Clone();
        return new SourceMailbox { UserPrincipalName = upn, Settings = settings };
    }
}
