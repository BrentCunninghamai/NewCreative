namespace M365Migrate.Core.Models;

/// <summary>
/// A group reconciliation planned for the target tenant. <see cref="Action"/> is
/// one of <c>create</c>, <c>exists</c>, or <c>skip</c>.
/// </summary>
public sealed class PlannedGroup
{
    public string SourceId { get; set; } = "";
    public string? MailNickname { get; set; }
    public string? DisplayName { get; set; }
    public string Kind { get; set; } = "";
    public string Action { get; set; } = "create";
    public string? Reason { get; set; }
    public string? Description { get; set; }
    public List<string> TargetMemberUpns { get; set; } = new();
    public List<string> TargetOwnerUpns { get; set; } = new();
    public bool IsDynamic { get; set; }
    public string? MembershipRule { get; set; }

    /// <summary>Build the Graph <c>POST /groups</c> request body for this planned group.</summary>
    public Dictionary<string, object> ToGraphBody()
    {
        var body = new Dictionary<string, object>
        {
            ["displayName"] = DisplayName ?? MailNickname ?? "",
            ["mailNickname"] = MailNickname ?? "",
        };
        if (Description is not null)
            body["description"] = Description;

        var groupTypes = new List<string>();
        if (Kind == "microsoft365")
        {
            groupTypes.Add("Unified");
            body["mailEnabled"] = true;
            body["securityEnabled"] = false;
        }
        else // security
        {
            body["mailEnabled"] = false;
            body["securityEnabled"] = true;
        }

        if (IsDynamic)
        {
            groupTypes.Add("DynamicMembership");
            body["membershipRule"] = MembershipRule ?? "";
            body["membershipRuleProcessingState"] = "On";
        }

        body["groupTypes"] = groupTypes;
        return body;
    }
}
