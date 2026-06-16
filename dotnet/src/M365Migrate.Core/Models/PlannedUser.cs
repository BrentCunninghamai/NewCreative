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
        if (!string.IsNullOrEmpty(defaultUsageLocation))
            body["usageLocation"] = defaultUsageLocation;
        return body;
    }
}
