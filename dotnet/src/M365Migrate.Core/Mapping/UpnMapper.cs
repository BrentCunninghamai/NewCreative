namespace M365Migrate.Core.Mapping;

/// <summary>UPN domain rewriting between the source and target tenants.</summary>
public static class UpnMapper
{
    /// <summary>
    /// Rewrite the domain portion of a UPN from source to target, but only when the
    /// UPN's domain matches <paramref name="sourceDomain"/> (case-insensitive). UPNs
    /// on other (e.g. already-vanity) domains are left intact.
    /// </summary>
    public static string Rewrite(string upn, string sourceDomain, string targetDomain)
    {
        var at = upn.IndexOf('@');
        if (at < 0)
            return upn;
        var local = upn[..at];
        var domain = upn[(at + 1)..];
        return domain.Equals(sourceDomain, StringComparison.OrdinalIgnoreCase)
            ? $"{local}@{targetDomain}"
            : upn;
    }
}
