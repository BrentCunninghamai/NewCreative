using System.Text.Json;
using M365Migrate.Core.Graph;

namespace M365Migrate.Core.Models;

/// <summary>A user as read from the source tenant.</summary>
public sealed class SourceUser
{
    /// <summary>The Graph user properties we read from the source tenant.</summary>
    public static readonly string[] SelectFields =
    {
        "id", "userPrincipalName", "displayName", "givenName", "surname", "mail",
        "jobTitle", "department", "officeLocation", "mobilePhone", "accountEnabled",
        "userType", "usageLocation", "assignedLicenses",
    };

    /// <summary>Expanded alongside the user so we learn each user's manager in one request.</summary>
    public const string Expand = "manager($select=id,userPrincipalName)";

    public string Id { get; set; } = "";
    public string UserPrincipalName { get; set; } = "";
    public string? DisplayName { get; set; }
    public string? GivenName { get; set; }
    public string? Surname { get; set; }
    public string? Mail { get; set; }
    public string? JobTitle { get; set; }
    public string? Department { get; set; }
    public string? OfficeLocation { get; set; }
    public string? MobilePhone { get; set; }
    public bool AccountEnabled { get; set; } = true;
    public string UserType { get; set; } = "Member";
    public string? UsageLocation { get; set; }
    public string? ManagerUpn { get; set; }
    public List<string> AssignedSkuIds { get; set; } = new();

    public static SourceUser FromGraph(JsonElement data)
    {
        var user = new SourceUser
        {
            Id = data.GetStringOrNull("id") ?? "",
            UserPrincipalName = data.GetStringOrNull("userPrincipalName") ?? "",
            DisplayName = data.GetStringOrNull("displayName"),
            GivenName = data.GetStringOrNull("givenName"),
            Surname = data.GetStringOrNull("surname"),
            Mail = data.GetStringOrNull("mail"),
            JobTitle = data.GetStringOrNull("jobTitle"),
            Department = data.GetStringOrNull("department"),
            OfficeLocation = data.GetStringOrNull("officeLocation"),
            MobilePhone = data.GetStringOrNull("mobilePhone"),
            AccountEnabled = data.GetBoolOrDefault("accountEnabled", true),
            UserType = data.GetStringOrNull("userType") ?? "Member",
            UsageLocation = data.GetStringOrNull("usageLocation"),
        };

        if (data.TryGetProperty("manager", out var manager) && manager.ValueKind == JsonValueKind.Object)
            user.ManagerUpn = manager.GetStringOrNull("userPrincipalName");

        foreach (var license in data.GetArrayOrEmpty("assignedLicenses"))
        {
            var sku = license.GetStringOrNull("skuId");
            if (sku is not null)
                user.AssignedSkuIds.Add(sku);
        }

        return user;
    }
}
