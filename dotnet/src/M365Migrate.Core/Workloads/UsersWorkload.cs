using System.Security.Cryptography;
using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Mapping;
using M365Migrate.Core.Models;

namespace M365Migrate.Core.Workloads;

/// <summary>
/// Users / Identities workload: discover, plan, migrate, and enrich user objects.
/// Staged so a plan can be reviewed before anything is written to the target.
/// </summary>
public sealed class UsersWorkload
{
    private readonly MigrationConfig _config;

    public UsersWorkload(MigrationConfig config) => _config = config;

    /// <summary>Read all users from the source tenant, including manager and licenses.</summary>
    public async Task<List<SourceUser>> DiscoverAsync(GraphClient source, CancellationToken ct = default)
    {
        var select = string.Join(",", SourceUser.SelectFields);
        var raw = await source.GetAllAsync(
            $"/users?$select={select}&$expand={SourceUser.Expand}&$top=999", ct);
        return raw.Select(SourceUser.FromGraph).ToList();
    }

    /// <summary>Return a map of lowercased target UPN -> target user id.</summary>
    public static async Task<Dictionary<string, string>> GetTargetUserIdsAsync(GraphClient target, CancellationToken ct = default)
    {
        var raw = await target.GetAllAsync("/users?$select=id,userPrincipalName&$top=999", ct);
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in raw)
        {
            var upn = item.GetStringOrNull("userPrincipalName");
            var id = item.GetStringOrNull("id");
            if (upn is not null && id is not null)
                map[upn] = id;
        }
        return map;
    }

    /// <summary>Read existing userPrincipalNames from the target tenant for conflict checks.</summary>
    public static async Task<HashSet<string>> DiscoverTargetUpnsAsync(GraphClient target, CancellationToken ct = default)
    {
        var raw = await target.GetAllAsync("/users?$select=userPrincipalName&$top=999", ct);
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in raw)
        {
            var upn = item.GetStringOrNull("userPrincipalName");
            if (upn is not null)
                set.Add(upn);
        }
        return set;
    }

    private string TargetUpn(string sourceUpn) => TargetNaming.TargetUpn(sourceUpn, _config);

    /// <summary>
    /// Decode a B2B/external (#EXT#) UPN back to the original source UPN. e.g.
    /// "alice_contoso.onmicrosoft.com#EXT#@target..." -> "alice@contoso.onmicrosoft.com".
    /// Returns null if the value isn't an #EXT# UPN.
    /// </summary>
    internal static string? DecodeExtUpn(string upn)
    {
        var marker = upn.IndexOf("#EXT#", StringComparison.OrdinalIgnoreCase);
        if (marker < 0)
            return null;
        var prefix = upn[..marker];
        var lastUnderscore = prefix.LastIndexOf('_');
        if (lastUnderscore < 0)
            return null;
        return prefix[..lastUnderscore] + "@" + prefix[(lastUnderscore + 1)..];
    }

    /// <summary>
    /// Identities already present in the target via cross-tenant sync / B2B (e.g. a
    /// Multi-Tenant Organization): decoded #EXT# source UPNs plus target mail
    /// addresses. Used to avoid creating native duplicates of users who are already
    /// represented in the target.
    /// </summary>
    public static async Task<HashSet<string>> DiscoverCrossTenantIdentitiesAsync(GraphClient target, CancellationToken ct = default)
    {
        var raw = await target.GetAllAsync("/users?$select=userPrincipalName,mail&$top=999", ct);
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in raw)
        {
            var upn = item.GetStringOrNull("userPrincipalName");
            if (upn is not null)
            {
                var decoded = DecodeExtUpn(upn);
                if (decoded is not null)
                    set.Add(decoded);
            }
            var mail = item.GetStringOrNull("mail");
            if (mail is not null)
                set.Add(mail);
        }
        return set;
    }

    /// <summary>Build a migration plan without writing anything.</summary>
    public List<PlannedUser> Plan(
        IEnumerable<SourceUser> users,
        ISet<string>? existingTargetUpns = null,
        ISet<string>? crossTenantIdentities = null)
    {
        var existing = existingTargetUpns is null
            ? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            : new HashSet<string>(existingTargetUpns, StringComparer.OrdinalIgnoreCase);
        var crossTenant = crossTenantIdentities is null
            ? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            : new HashSet<string>(crossTenantIdentities, StringComparer.OrdinalIgnoreCase);

        var planned = new List<PlannedUser>();
        foreach (var user in users)
        {
            var targetUpn = TargetUpn(user.UserPrincipalName);
            string action;
            string? reason;

            // #EXT# UPNs are B2B/external (guest) identities homed in another tenant.
            // They can't be recreated as normal users, so treat them as guests.
            var isExternal = user.UserPrincipalName.Contains("#EXT#", StringComparison.OrdinalIgnoreCase);
            var isGuest = isExternal || user.UserType.Equals("guest", StringComparison.OrdinalIgnoreCase);

            // Already present in the target via cross-tenant sync / B2B (MTO): the
            // source user appears as a decoded #EXT# identity or shares a target mail.
            var presentCrossTenant =
                crossTenant.Contains(user.UserPrincipalName)
                || (user.Mail is not null && crossTenant.Contains(user.Mail));

            if (_config.Options.SkipGuests && isGuest)
            {
                action = "skip";
                reason = isExternal ? "external/guest (#EXT#)" : "guest user";
            }
            else if (presentCrossTenant)
            {
                action = "conflict";
                reason = "already in target (cross-tenant/B2B sync)";
            }
            else if (existing.Contains(targetUpn))
            {
                action = "conflict";
                reason = "target UPN already exists";
            }
            else
            {
                action = "create";
                reason = null;
            }

            planned.Add(new PlannedUser
            {
                SourceId = user.Id,
                SourceUpn = user.UserPrincipalName,
                TargetUpn = targetUpn,
                DisplayName = TargetNaming.TargetDisplayName(user.DisplayName, _config),
                UserType = user.UserType,
                OnPremisesSynced = user.OnPremisesSyncEnabled,
                UsageLocation = user.UsageLocation,
                SourceSkuIds = new List<string>(user.AssignedSkuIds),
                Action = action,
                Reason = reason,
            });
        }
        return planned;
    }

    /// <summary>Map skuId → skuPartNumber for a tenant (subscribedSkus). The part number is the
    /// product identifier that's consistent across tenants; skuId is per-tenant.</summary>
    public static async Task<Dictionary<string, string>> SkuIdToPartAsync(GraphClient client, CancellationToken ct = default)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var s in await client.GetAllAsync("/subscribedSkus?$select=skuId,skuPartNumber", ct))
        {
            var id = s.GetStringOrNull("skuId");
            var part = s.GetStringOrNull("skuPartNumber");
            if (id is not null && part is not null)
                map[id] = part;
        }
        return map;
    }

    /// <summary>Map skuPartNumber → skuId for a tenant (to translate a source license to the target).</summary>
    public static async Task<Dictionary<string, string>> PartToSkuIdAsync(GraphClient client, CancellationToken ct = default)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var s in await client.GetAllAsync("/subscribedSkus?$select=skuId,skuPartNumber", ct))
        {
            var id = s.GetStringOrNull("skuId");
            var part = s.GetStringOrNull("skuPartNumber");
            if (id is not null && part is not null)
                map[part] = id;
        }
        return map;
    }

    /// <summary>Translate a planned user's source SKU ids to target SKU ids via part number.</summary>
    internal static List<string> TargetSkuIds(
        IEnumerable<string> sourceSkuIds,
        IReadOnlyDictionary<string, string> sourceIdToPart,
        IReadOnlyDictionary<string, string> targetPartToId)
    {
        var ids = new List<string>();
        foreach (var src in sourceSkuIds)
            if (sourceIdToPart.TryGetValue(src, out var part) && targetPartToId.TryGetValue(part, out var tgt))
                ids.Add(tgt);
        return ids;
    }

    /// <summary>Create the planned users in the target tenant (dry run unless executed). When
    /// license maps are supplied, also assign the target equivalents of each user's source
    /// licenses (matched by SKU part number) after creation.</summary>
    public async Task<List<WorkloadResult>> MigrateAsync(
        GraphClient target,
        IEnumerable<PlannedUser> planned,
        bool dryRun = true,
        IReadOnlyDictionary<string, string>? sourceSkuIdToPart = null,
        IReadOnlyDictionary<string, string>? targetPartToSkuId = null,
        string? defaultUsageLocation = null,
        CancellationToken ct = default)
    {
        var assignLicenses = sourceSkuIdToPart is not null && targetPartToSkuId is not null;
        var results = new List<WorkloadResult>();
        foreach (var p in planned)
        {
            if (p.Action != "create")
            {
                results.Add(new WorkloadResult(p.TargetUpn, $"skipped:{p.Action}", p.Reason));
                continue;
            }
            if (dryRun)
            {
                var would = new WorkloadResult(p.TargetUpn, "would-create");
                if (assignLicenses)
                {
                    var skus = TargetSkuIds(p.SourceSkuIds, sourceSkuIdToPart!, targetPartToSkuId!);
                    if (skus.Count > 0) would.Detail["licenses"] = $"would-assign:{skus.Count}";
                }
                results.Add(would);
                continue;
            }

            var body = p.ToGraphBody(GeneratePassword(), defaultUsageLocation);
            WorkloadResult result;
            string? id;
            try
            {
                var created = await target.PostJsonAsync("/users", body, ct);
                result = new WorkloadResult(p.TargetUpn, "created");
                id = created.GetStringOrNull("id");
                if (id is not null)
                    result.Detail["target_id"] = id;
            }
            catch (GraphException ex)
            {
                results.Add(new WorkloadResult(p.TargetUpn, "error", ex.Message));
                continue;
            }

            if (assignLicenses && id is not null)
            {
                var skuIds = TargetSkuIds(p.SourceSkuIds, sourceSkuIdToPart!, targetPartToSkuId!);
                if (skuIds.Count == 0)
                {
                    if (p.SourceSkuIds.Count > 0)
                        result.Detail["licenses"] = "none-matched";
                }
                else if (string.IsNullOrEmpty(p.UsageLocation) && string.IsNullOrEmpty(defaultUsageLocation))
                {
                    result.Detail["licenses"] = "skipped:no-usageLocation";
                }
                else
                {
                    try
                    {
                        await target.PostJsonAsync($"/users/{id}/assignLicense", new Dictionary<string, object>
                        {
                            ["addLicenses"] = skuIds.Select(s => new Dictionary<string, object>
                            {
                                ["skuId"] = s,
                                ["disabledPlans"] = Array.Empty<string>(),
                            }).ToArray(),
                            ["removeLicenses"] = Array.Empty<string>(),
                        }, ct);
                        result.Detail["licenses"] = $"assigned:{skuIds.Count}";
                    }
                    catch (GraphException ex)
                    {
                        result.Detail["licenses"] = $"error:{ex.StatusCode}";
                    }
                }
            }
            results.Add(result);
        }
        return results;
    }

    /// <summary>Generate a strong random initial password (user resets on first sign-in).</summary>
    public static string GeneratePassword(int length = 20)
    {
        const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*-_";
        var chars = new char[length];
        for (var i = 0; i < length; i++)
            chars[i] = alphabet[RandomNumberGenerator.GetInt32(alphabet.Length)];
        return new string(chars);
    }
}
