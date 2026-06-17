namespace M365Migrate.Core.Configuration;

/// <summary>Behavioral options for migration runs.</summary>
public sealed class MigrationOptions
{
    public bool RewriteUpnDomain { get; set; } = true;
    public bool SkipGuests { get; set; } = true;

    /// <summary>
    /// Optional text prepended/appended to each target identity name — the user
    /// UPN local-part and the group mailNickname — to avoid collisions when
    /// merging multiple source tenants into one target. Empty = no change.
    /// </summary>
    public string NamePrefix { get; set; } = "";
    public string NameSuffix { get; set; } = "";

    /// <summary>
    /// Optional text appended to a migrated user's display name (e.g. "(Contoso)")
    /// so merged users are distinguishable in the target GAL. Empty = no change.
    /// </summary>
    public string DisplayNameSuffix { get; set; } = "";
}

/// <summary>Top-level configuration: a source tenant, a target tenant, and options.</summary>
public sealed class MigrationConfig
{
    public TenantConfig Source { get; set; } = new();
    public TenantConfig Target { get; set; } = new();
    public MigrationOptions Options { get; set; } = new();
}
