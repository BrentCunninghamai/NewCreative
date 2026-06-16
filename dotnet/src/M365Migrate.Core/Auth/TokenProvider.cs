using Azure.Core;
using Azure.Identity;
using M365Migrate.Core.Configuration;

namespace M365Migrate.Core.Auth;

/// <summary>
/// A token provider yields a valid Graph bearer token. Injecting it (rather than
/// hard-coding credential acquisition) keeps the Graph client unit-testable.
/// </summary>
public delegate Task<string> TokenProvider(CancellationToken ct);

/// <summary>Builds <see cref="TokenProvider"/>s for the OAuth2 client-credentials flow.</summary>
public static class TokenProviders
{
    private const string GraphDefaultScope = "https://graph.microsoft.com/.default";

    /// <summary>
    /// Returns a provider that yields a Graph access token for <paramref name="tenant"/>
    /// using its app registration. Token acquisition/refresh/caching is handled by
    /// Azure.Identity.
    /// </summary>
    public static TokenProvider ForTenant(TenantConfig tenant)
    {
        var credential = new ClientSecretCredential(
            tenant.TenantId, tenant.ClientId, tenant.ClientSecret);
        var context = new TokenRequestContext(new[] { GraphDefaultScope });

        return async ct =>
        {
            AccessToken token = await credential.GetTokenAsync(context, ct);
            return token.Token;
        };
    }

    /// <summary>A provider that always returns the same fixed token (for tests/offline use).</summary>
    public static TokenProvider Static(string token) => _ => Task.FromResult(token);
}
