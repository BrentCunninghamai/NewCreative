namespace M365Migrate.Core.Configuration;

/// <summary>Credentials and identity for a single Microsoft 365 tenant.</summary>
public sealed class TenantConfig
{
    public string TenantId { get; set; } = "";
    public string ClientId { get; set; } = "";
    public string ClientSecret { get; set; } = "";
    public string PrimaryDomain { get; set; } = "";
}
