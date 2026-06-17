using M365Migrate.Core.Graph;

namespace M365Migrate.Core.Workloads;

/// <summary>A member tenant of a Multi-Tenant Organization (MTO).</summary>
public sealed record MtoTenant(string TenantId, string? DisplayName, string? State);

/// <summary>
/// The result of inspecting a tenant's Multi-Tenant Organization configuration.
/// <see cref="PermissionGranted"/> is false when the app lacks
/// <c>MultiTenantOrganization.Read.All</c> (Graph returns 403); <see cref="Configured"/>
/// is false when the tenant simply isn't part of an MTO (404 / no MTO).
/// </summary>
public sealed record MultiTenantOrgInfo(
    bool PermissionGranted,
    bool Configured,
    string? DisplayName,
    string? State,
    IReadOnlyList<MtoTenant> Tenants)
{
    public static MultiTenantOrgInfo NoPermission { get; } =
        new(false, false, null, null, Array.Empty<MtoTenant>());

    public static MultiTenantOrgInfo NotConfigured { get; } =
        new(true, false, null, null, Array.Empty<MtoTenant>());

    /// <summary>True if a tenant with the given id is a member of this MTO.</summary>
    public bool ContainsTenant(string tenantId) =>
        !string.IsNullOrWhiteSpace(tenantId)
        && Tenants.Any(t => string.Equals(t.TenantId, tenantId, StringComparison.OrdinalIgnoreCase));
}

/// <summary>
/// Reads a tenant's Multi-Tenant Organization (MTO) relationship directly from Graph
/// (<c>/tenantRelationships/multiTenantOrganization</c> and its <c>/tenants</c>), so the
/// tool can authoritatively confirm that source and target are MTO-linked rather than
/// only inferring it from <c>#EXT#</c> rows. Requires <c>MultiTenantOrganization.Read.All</c>;
/// degrades gracefully (returns <see cref="MultiTenantOrgInfo.NoPermission"/>) if absent.
/// </summary>
public static class MultiTenantOrgInspector
{
    public static async Task<MultiTenantOrgInfo> InspectAsync(GraphClient client, CancellationToken ct = default)
    {
        System.Text.Json.JsonElement mto;
        try
        {
            mto = await client.GetAsync("/tenantRelationships/multiTenantOrganization", ct);
        }
        catch (GraphException ex) when (ex.StatusCode is 403 or 401)
        {
            return MultiTenantOrgInfo.NoPermission;
        }
        catch (GraphException ex) when (ex.StatusCode == 404)
        {
            return MultiTenantOrgInfo.NotConfigured;
        }

        var state = mto.GetStringOrNull("state");
        var displayName = mto.GetStringOrNull("displayName");
        // A tenant with no MTO reports state "inactive" (or omits it). Treat that as
        // not configured so the UI doesn't claim an MTO that doesn't exist.
        if (state is not null && state.Equals("inactive", StringComparison.OrdinalIgnoreCase))
            return MultiTenantOrgInfo.NotConfigured;

        var tenants = new List<MtoTenant>();
        try
        {
            foreach (var t in await client.GetAllAsync("/tenantRelationships/multiTenantOrganization/tenants", ct))
            {
                var id = t.GetStringOrNull("tenantId");
                if (id is not null)
                    tenants.Add(new MtoTenant(id, t.GetStringOrNull("displayName"), t.GetStringOrNull("state")));
            }
        }
        catch (GraphException ex) when (ex.StatusCode is 403 or 401)
        {
            // The MTO itself was readable but the tenants collection wasn't.
            return new MultiTenantOrgInfo(true, true, displayName, state, Array.Empty<MtoTenant>());
        }

        return new MultiTenantOrgInfo(true, true, displayName, state, tenants);
    }

    /// <summary>
    /// A one-line, human-readable summary for the connection test, given the target's
    /// MTO view and the source tenant id. Returns null when there's nothing useful to say.
    /// </summary>
    public static string? Summarize(MultiTenantOrgInfo target, string sourceTenantId)
    {
        if (!target.PermissionGranted)
            return null; // permission not consented — stay silent rather than guess
        if (!target.Configured)
            return "No Multi-Tenant Organization on the target.";

        var linked = target.ContainsTenant(sourceTenantId);
        var name = string.IsNullOrWhiteSpace(target.DisplayName) ? "Multi-Tenant Organization" : target.DisplayName;
        return linked
            ? $"MTO \"{name}\" links these tenants ({target.Tenants.Count} members) — some source users are expected to already exist in the target (cross-tenant sync); the plan flags them as conflict."
            : $"Target is in MTO \"{name}\" ({target.Tenants.Count} members) but the source tenant isn't listed.";
    }
}
