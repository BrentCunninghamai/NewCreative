using System.Text.Json;
using M365Migrate.Core.Graph;

namespace M365Migrate.Core.Models;

/// <summary>A group as read from the source tenant, with its user members and owners.</summary>
public sealed class SourceGroup
{
    public static readonly string[] SelectFields =
    {
        "id", "displayName", "mailNickname", "description", "groupTypes",
        "securityEnabled", "mailEnabled", "visibility",
        "membershipRule", "membershipRuleProcessingState",
    };

    public const string Expand =
        "members($select=id,userPrincipalName),owners($select=id,userPrincipalName)";

    public string Id { get; set; } = "";
    public string? DisplayName { get; set; }
    public string? MailNickname { get; set; }
    public string? Description { get; set; }
    public List<string> GroupTypes { get; set; } = new();
    public bool SecurityEnabled { get; set; }
    public bool MailEnabled { get; set; }
    public string? Visibility { get; set; }
    public List<string> MemberUpns { get; set; } = new();
    public List<string> OwnerUpns { get; set; } = new();
    public string? MembershipRule { get; set; }

    public bool IsUnified =>
        GroupTypes.Any(t => t.Equals("Unified", StringComparison.OrdinalIgnoreCase));

    public bool IsDynamic =>
        GroupTypes.Any(t => t.Equals("DynamicMembership", StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Classify the group: <c>microsoft365</c>, <c>security</c>,
    /// <c>mail-enabled-security</c>, <c>distribution</c>, or <c>unknown</c>. Only the
    /// first two can be provisioned through Graph.
    /// </summary>
    public string Kind
    {
        get
        {
            if (IsUnified) return "microsoft365";
            if (SecurityEnabled && MailEnabled) return "mail-enabled-security";
            if (SecurityEnabled) return "security";
            if (MailEnabled) return "distribution";
            return "unknown";
        }
    }

    public static SourceGroup FromGraph(JsonElement data)
    {
        var group = new SourceGroup
        {
            Id = data.GetStringOrNull("id") ?? "",
            DisplayName = data.GetStringOrNull("displayName"),
            MailNickname = data.GetStringOrNull("mailNickname"),
            Description = data.GetStringOrNull("description"),
            SecurityEnabled = data.GetBoolOrDefault("securityEnabled"),
            MailEnabled = data.GetBoolOrDefault("mailEnabled"),
            Visibility = data.GetStringOrNull("visibility"),
            MembershipRule = data.GetStringOrNull("membershipRule"),
        };

        foreach (var t in data.GetArrayOrEmpty("groupTypes"))
            if (t.ValueKind == JsonValueKind.String)
                group.GroupTypes.Add(t.GetString()!);

        foreach (var m in data.GetArrayOrEmpty("members"))
        {
            var upn = m.GetStringOrNull("userPrincipalName");
            if (upn is not null)
                group.MemberUpns.Add(upn);
        }

        foreach (var o in data.GetArrayOrEmpty("owners"))
        {
            var upn = o.GetStringOrNull("userPrincipalName");
            if (upn is not null)
                group.OwnerUpns.Add(upn);
        }

        return group;
    }
}
