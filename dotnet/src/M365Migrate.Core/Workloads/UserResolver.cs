using M365Migrate.Core.Graph;

namespace M365Migrate.Core.Workloads;

/// <summary>A resolved directory user, used to confirm targeting before any write.</summary>
public sealed record ResolvedUser(string Id, string? DisplayName, string? UserPrincipalName, string? Mail)
{
    /// <summary>Human-readable "Display Name (upn-or-mail-or-id)".</summary>
    public string Describe() =>
        $"{DisplayName ?? "(no name)"} ({UserPrincipalName ?? Mail ?? Id})";
}

/// <summary>
/// Resolves a user object from Graph by UPN or object id, so the per-user content
/// workloads can show the operator exactly which source and target identity a run will
/// touch — important when the target isn't a clean rewrite (cross-tenant orchestrator /
/// hybrid sync). Returns null if the user can't be found (e.g. a wrong/typo'd key).
/// </summary>
public static class UserResolver
{
    /// <summary>
    /// Resolve a user by key. <paramref name="userRefOrKey"/> may be a bare UPN/object id
    /// or a resource ref like <c>/users/{key}</c> or <c>/users/{key}/drive</c>.
    /// </summary>
    public static async Task<ResolvedUser?> TryResolveAsync(GraphClient client, string userRefOrKey, CancellationToken ct = default)
    {
        var key = NormalizeKey(userRefOrKey);
        if (string.IsNullOrWhiteSpace(key))
            return null;
        try
        {
            var el = await client.GetAsync($"/users/{key}?$select=id,displayName,userPrincipalName,mail", ct);
            var id = el.GetStringOrNull("id");
            if (id is null)
                return null;
            return new ResolvedUser(
                id,
                el.GetStringOrNull("displayName"),
                el.GetStringOrNull("userPrincipalName"),
                el.GetStringOrNull("mail"));
        }
        catch (GraphException ex) when (ex.StatusCode is 404 or 400)
        {
            // 404: no such user. 400: malformed key (e.g. not a valid UPN/GUID).
            return null;
        }
    }

    public static string NormalizeKey(string userRefOrKey)
    {
        var key = userRefOrKey?.Trim() ?? "";
        const string prefix = "/users/";
        if (key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            key = key[prefix.Length..];
            var slash = key.IndexOf('/');
            if (slash >= 0)
                key = key[..slash];
        }
        return key;
    }
}
