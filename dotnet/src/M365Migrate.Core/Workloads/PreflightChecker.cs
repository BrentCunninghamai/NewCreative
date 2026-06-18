using System.Text;
using System.Text.Json;

namespace M365Migrate.Core.Workloads;

/// <summary>The result of a per-tenant readiness check.</summary>
public sealed record PreflightResult(IReadOnlyList<string> Granted, IReadOnlyList<string> Missing)
{
    public bool AllGranted => Missing.Count == 0;
}

/// <summary>
/// Pre-flight readiness: confirm the app's admin-consented Graph **application** permissions
/// before a run, so missing consent surfaces up front instead of as a mid-run 403. A
/// client-credentials access token carries a <c>roles</c> claim listing exactly the app
/// permissions that are effective, so we read that directly — no extra permission needed.
/// </summary>
public static class PreflightChecker
{
    /// <summary>Extract the application-permission role values from a JWT access token.</summary>
    public static HashSet<string> RolesFromToken(string accessToken)
    {
        var roles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(accessToken))
            return roles;
        var parts = accessToken.Split('.');
        if (parts.Length < 2)
            return roles;
        try
        {
            using var doc = JsonDocument.Parse(DecodeSegment(parts[1]));
            if (doc.RootElement.TryGetProperty("roles", out var rolesEl) && rolesEl.ValueKind == JsonValueKind.Array)
                foreach (var r in rolesEl.EnumerateArray())
                {
                    var v = r.GetString();
                    if (v is not null)
                        roles.Add(v);
                }
        }
        catch (Exception)
        {
            // Malformed token — treat as no roles; the caller reports everything missing.
        }
        return roles;
    }

    /// <summary>Compare a token's granted roles against the required permission names.</summary>
    public static PreflightResult Check(string accessToken, IEnumerable<string> required)
    {
        var granted = RolesFromToken(accessToken);
        var req = required.ToList();
        var have = req.Where(granted.Contains).ToList();
        var missing = req.Where(r => !granted.Contains(r)).ToList();
        return new PreflightResult(have, missing);
    }

    private static string DecodeSegment(string segment)
    {
        var s = segment.Replace('-', '+').Replace('_', '/');
        switch (s.Length % 4)
        {
            case 2: s += "=="; break;
            case 3: s += "="; break;
        }
        return Encoding.UTF8.GetString(Convert.FromBase64String(s));
    }
}
