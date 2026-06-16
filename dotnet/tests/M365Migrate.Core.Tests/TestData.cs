using M365Migrate.Core.Configuration;

namespace M365Migrate.Core.Tests;

internal static class TestData
{
    public static MigrationConfig Config() => new()
    {
        Source = new TenantConfig
        {
            TenantId = "src-tenant",
            ClientId = "src-client",
            ClientSecret = "src-secret",
            PrimaryDomain = "contoso.onmicrosoft.com",
        },
        Target = new TenantConfig
        {
            TenantId = "tgt-tenant",
            ClientId = "tgt-client",
            ClientSecret = "tgt-secret",
            PrimaryDomain = "fabrikam.onmicrosoft.com",
        },
        Options = new MigrationOptions { RewriteUpnDomain = true, SkipGuests = true },
    };
}
