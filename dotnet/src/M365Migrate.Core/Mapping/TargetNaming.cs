using M365Migrate.Core.Configuration;

namespace M365Migrate.Core.Mapping;

/// <summary>
/// Computes target-tenant names from source names, applying the configured UPN
/// domain rewrite and the optional name prefix/suffix used to keep identities
/// distinct when several source tenants are merged into one target.
/// </summary>
public static class TargetNaming
{
    /// <summary>
    /// The target UPN for a source UPN: domain rewritten (when enabled), then the
    /// name prefix/suffix applied to the local part.
    /// </summary>
    public static string TargetUpn(string sourceUpn, MigrationConfig config)
    {
        var upn = config.Options.RewriteUpnDomain
            ? UpnMapper.Rewrite(sourceUpn, config.Source.PrimaryDomain, config.Target.PrimaryDomain)
            : sourceUpn;

        if (config.Options.NamePrefix.Length == 0 && config.Options.NameSuffix.Length == 0)
            return upn;

        var at = upn.IndexOf('@');
        if (at < 0)
            return config.Options.NamePrefix + upn + config.Options.NameSuffix;
        return $"{config.Options.NamePrefix}{upn[..at]}{config.Options.NameSuffix}@{upn[(at + 1)..]}";
    }

    /// <summary>The target group mailNickname: source nickname with prefix/suffix applied.</summary>
    public static string TargetMailNickname(string? sourceNickname, MigrationConfig config)
        => $"{config.Options.NamePrefix}{sourceNickname}{config.Options.NameSuffix}";
}
