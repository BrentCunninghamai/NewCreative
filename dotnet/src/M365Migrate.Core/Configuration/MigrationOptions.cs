namespace M365Migrate.Core.Configuration;

/// <summary>Behavioral options for migration runs.</summary>
public sealed class MigrationOptions
{
    public bool RewriteUpnDomain { get; set; } = true;
    public bool SkipGuests { get; set; } = true;
}

/// <summary>Top-level configuration: a source tenant, a target tenant, and options.</summary>
public sealed class MigrationConfig
{
    public TenantConfig Source { get; set; } = new();
    public TenantConfig Target { get; set; } = new();
    public MigrationOptions Options { get; set; } = new();
}
