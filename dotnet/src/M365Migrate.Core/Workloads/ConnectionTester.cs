using M365Migrate.Core.Graph;

namespace M365Migrate.Core.Workloads;

/// <summary>
/// Verifies a tenant connection before any workload runs: that the app
/// registration authenticates and Graph is reachable. Returns the organization's
/// display name so the operator can confirm they connected to the right tenant.
/// </summary>
public static class ConnectionTester
{
    public static async Task<string> CheckAsync(GraphClient client, CancellationToken ct = default)
    {
        var org = await client.GetAsync("/organization", ct);
        if (org.TryGetProperty("value", out var value)
            && value.ValueKind == System.Text.Json.JsonValueKind.Array
            && value.GetArrayLength() > 0)
        {
            return value[0].GetStringOrNull("displayName") ?? "(unknown organization)";
        }
        return "(unknown organization)";
    }
}
