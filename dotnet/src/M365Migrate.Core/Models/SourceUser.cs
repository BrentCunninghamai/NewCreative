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
        "onPremisesSyncEnabled", "onPremisesImmutableId", "proxyAddresses",
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

    /// <summary>True if this user is mastered in on-prem AD and synced via Entra Connect
    /// (a hybrid tenant). Such users' mailboxes may live on Exchange on-premises.</summary>
    public bool OnPremisesSyncEnabled { get; set; }
    /// <summary>The on-prem anchor (immutableId / ms-DS-ConsistencyGuid), when hybrid-synced.</summary>
    public string? OnPremisesImmutableId { get; set; }
    /// <summary>All SMTP addresses (primary + aliases), used to match a target by any shared address.</summary>
    public List<string> ProxyAddresses { get; set; } = new();

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
            OnPremisesSyncEnabled = data.GetBoolOrDefault("onPremisesSyncEnabled", false),
            OnPremisesImmutableId = data.GetStringOrNull("onPremisesImmutableId"),
        };

        if (data.TryGetProperty("manager", out var manager) && manager.ValueKind == JsonValueKind.Object)
            user.ManagerUpn = manager.GetStringOrNull("userPrincipalName");

        foreach (var license in data.GetArrayOrEmpty("assignedLicenses"))
        {
            var sku = license.GetStringOrNull("skuId");
            if (sku is not null)
                user.AssignedSkuIds.Add(sku);
        }

        foreach (var proxy in data.GetArrayOrEmpty("proxyAddresses"))
        {
            var addr = SmtpProxyAddress(proxy.GetString());
            if (addr is not null)
                user.ProxyAddresses.Add(addr);
        }

        return user;
    }

    /// <summary>
    /// Return the email address from an SMTP proxyAddresses entry ("SMTP:primary@x" /
    /// "smtp:alias@x"), or null for non-SMTP schemes (SIP:, X500:, …) which must not be
    /// treated as email addresses for matching.
    /// </summary>
    public static string? SmtpProxyAddress(string? proxy)
    {
        if (string.IsNullOrEmpty(proxy)) return null;
        var colon = proxy.IndexOf(':');
        if (colon < 0) return proxy; // unprefixed (e.g. the mail attribute) — treat as an address
        return proxy[..colon].Equals("smtp", StringComparison.OrdinalIgnoreCase) ? proxy[(colon + 1)..] : null;
    }
}
