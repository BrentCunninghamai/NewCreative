using System.Text;
using System.Text.Json;

namespace M365Migrate.Core.Configuration;

/// <summary>A Microsoft Graph application permission the tool needs, with its app-role id.</summary>
public sealed record GraphPermission(string Name, string Id, string UsedFor);

/// <summary>
/// One-click admin setup: the exact set of Microsoft Graph **application** permissions this
/// tool uses, plus generators for (a) the app-registration **manifest** block a Global Admin
/// pastes once, and (b) the **admin-consent URL** the GA clicks to approve them all at once —
/// instead of adding each permission and granting consent by hand. Permission role ids are
/// the well-known Microsoft Graph values (verified against the Graph permissions reference).
/// </summary>
public static class GraphSetup
{
    /// <summary>The Microsoft Graph resource (service principal) app id.</summary>
    public const string GraphResourceAppId = "00000003-0000-0000-c000-000000000000";

    /// <summary>Every application permission the tool's workloads can use.</summary>
    public static readonly IReadOnlyList<GraphPermission> Permissions = new[]
    {
        new GraphPermission("User.ReadWrite.All",       "741f803b-c850-494e-b5df-cde7c675a1ca", "Users: read source, create in target"),
        new GraphPermission("Directory.ReadWrite.All",  "19dbc75e-c2e2-444c-a770-ec69d8559fc7", "Users enrich: licenses, manager"),
        new GraphPermission("Group.ReadWrite.All",      "62a82d76-70ea-41e2-9197-370581804d09", "Groups + Teams provisioning"),
        new GraphPermission("User.Read.All",            "df021288-bdef-4463-88db-98f22de89214", "Read users for mapping/membership"),
        new GraphPermission("Organization.Read.All",    "498476ce-e0fe-48b0-b801-37ba7e2685c6", "Test connections (org name)"),
        new GraphPermission("MailboxSettings.ReadWrite", "6931bccd-447a-43d1-b442-00a195474933", "Mailbox settings"),
        new GraphPermission("Mail.ReadWrite",           "e2a3a72e-5f79-4c64-b1b1-878b674786c9", "Mail content (folders + messages)"),
        new GraphPermission("Calendars.ReadWrite",      "ef54d2bf-783f-4e0f-bca1-3210c0444d99", "Calendar content"),
        new GraphPermission("Contacts.ReadWrite",       "6918b873-d17a-4dc1-b314-35f528134491", "Contacts content"),
        new GraphPermission("Files.ReadWrite.All",      "75359482-378d-4052-8f01-80520e7db3cd", "OneDrive/SharePoint files"),
        new GraphPermission("Sites.ReadWrite.All",      "9492366f-7969-46a4-8d15-ed1a20078fff", "OneDrive/SharePoint sites"),
        new GraphPermission("Team.Create",              "23fc2474-f741-46ce-8465-674744c5c361", "Teams: create/enable"),
        new GraphPermission("Channel.ReadBasic.All",    "59a6b24b-4225-4393-8165-ebaec5f55d7a", "Teams: read channels"),
        new GraphPermission("Teamwork.Migrate.All",     "dfb0dd15-61de-45b2-be36-d6a69fba3c79", "Teams: import message history"),
        new GraphPermission("TeamMember.ReadWrite.All", "0121dc95-1b9f-4aed-8bac-58c5ac466691", "Teams: add members/owners"),
    };

    /// <summary>
    /// The <c>requiredResourceAccess</c> manifest block. Paste into the app registration's
    /// **Manifest** (replacing the existing <c>requiredResourceAccess</c> array) so every
    /// permission is added at once; then use the admin-consent URL to approve them.
    /// </summary>
    public static string ManifestJson()
    {
        var block = new[]
        {
            new
            {
                resourceAppId = GraphResourceAppId,
                resourceAccess = Permissions.Select(p => new { id = p.Id, type = "Role" }).ToArray(),
            },
        };
        return JsonSerializer.Serialize(block, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    /// The admin-consent URL a Global Admin opens to grant the app all configured permissions
    /// for <paramref name="tenantId"/> in one click. <paramref name="tenantId"/> may be a tenant
    /// id or domain; <paramref name="clientId"/> is the app registration's Application (client) id.
    /// </summary>
    public static string AdminConsentUrl(string tenantId, string clientId)
    {
        var tenant = string.IsNullOrWhiteSpace(tenantId) ? "common" : tenantId.Trim();
        return $"https://login.microsoftonline.com/{Uri.EscapeDataString(tenant)}" +
               $"/adminconsent?client_id={Uri.EscapeDataString(clientId.Trim())}";
    }

    /// <summary>A copy-paste setup guide: permission list, manifest block, and consent links.</summary>
    public static string SetupGuide(
        string sourceTenantId, string sourceClientId,
        string targetTenantId, string targetClientId)
    {
        var sb = new StringBuilder();
        sb.AppendLine("m365-migrate — app setup (run once per tenant, as Global Administrator)");
        sb.AppendLine("======================================================================");
        sb.AppendLine();
        sb.AppendLine("STEP 1 — Add all permissions at once");
        sb.AppendLine("  In each tenant's app registration: Manage > Manifest, replace the");
        sb.AppendLine("  \"requiredResourceAccess\": [...] array with the block below, then Save.");
        sb.AppendLine();
        sb.AppendLine(ManifestJson());
        sb.AppendLine();
        sb.AppendLine("STEP 2 — Approve everything in one click (Global Admin)");
        sb.AppendLine("  Open each link, sign in as a Global Admin, and click Accept.");
        sb.AppendLine();
        if (!string.IsNullOrWhiteSpace(sourceClientId))
            sb.AppendLine("  SOURCE: " + AdminConsentUrl(sourceTenantId, sourceClientId));
        if (!string.IsNullOrWhiteSpace(targetClientId))
            sb.AppendLine("  TARGET: " + AdminConsentUrl(targetTenantId, targetClientId));
        sb.AppendLine();
        sb.AppendLine("Permissions granted (Microsoft Graph, Application):");
        foreach (var p in Permissions)
            sb.AppendLine($"  - {p.Name,-26} {p.UsedFor}");
        sb.AppendLine();
        sb.AppendLine("Optional: MultiTenantOrganization.Read.All (target only) lets Test");
        sb.AppendLine("connections confirm an MTO link — add it by name if you want that.");
        return sb.ToString();
    }
}
