namespace M365Migrate.Core.Models;

/// <summary>
/// A user creation planned for the target tenant. <see cref="Action"/> is one of
/// <c>create</c>, <c>skip</c> (filtered out), or <c>conflict</c> (target UPN exists).
/// </summary>
public sealed class PlannedUser
{
    public string SourceId { get; set; } = "";
    public string SourceUpn { get; set; } = "";
    public string TargetUpn { get; set; } = "";
    public string? DisplayName { get; set; }
    public string UserType { get; set; } = "Member";
    /// <summary>True if the source user is on-prem synced (hybrid) — mailbox may live on-premises.</summary>
    public bool OnPremisesSynced { get; set; }
    /// <summary>The source user's usageLocation (required before a license can be assigned).</summary>
    public string? UsageLocation { get; set; }
    /// <summary>The source tenant license SKU ids assigned to this user (mapped to target by part number).</summary>
    public List<string> SourceSkuIds { get; set; } = new();
    public string Action { get; set; } = "create";
    public string? Reason { get; set; }

    /// <summary>Build the Graph <c>POST /users</c> request body for this planned user.</summary>
    public Dictionary<string, object> ToGraphBody(string password, string? defaultUsageLocation = null)
    {
        var mailNickname = TargetUpn.Split('@', 2)[0];
        var body = new Dictionary<string, object>
        {
            ["accountEnabled"] = true,
            ["displayName"] = DisplayName ?? mailNickname,
            ["mailNickname"] = mailNickname,
            ["userPrincipalName"] = TargetUpn,
            ["passwordProfile"] = new Dictionary<string, object>
            {
                ["forceChangePasswordNextSignIn"] = true,
                ["password"] = password,
            },
        };
        // usageLocation is required before a license can be assigned; prefer the source value.
        var usageLocation = !string.IsNullOrEmpty(UsageLocation) ? UsageLocation : defaultUsageLocation;
        if (!string.IsNullOrEmpty(usageLocation))
            body["usageLocation"] = usageLocation;
        return body;
    }
}
