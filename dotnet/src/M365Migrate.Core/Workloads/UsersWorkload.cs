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

    /// <summary>Build a migration plan without writing anything.</summary>
    public List<PlannedUser> Plan(IEnumerable<SourceUser> users, ISet<string>? existingTargetUpns = null)
    {
        var existing = existingTargetUpns is null
            ? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            : new HashSet<string>(existingTargetUpns, StringComparer.OrdinalIgnoreCase);

        var planned = new List<PlannedUser>();
        foreach (var user in users)
        {
            var targetUpn = TargetUpn(user.UserPrincipalName);
            string action;
            string? reason;

            if (_config.Options.SkipGuests && user.UserType.Equals("guest", StringComparison.OrdinalIgnoreCase))
            {
                action = "skip";
                reason = "guest user";
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
                DisplayName = user.DisplayName,
                Action = action,
                Reason = reason,
            });
        }
        return planned;
    }

    /// <summary>Create the planned users in the target tenant (dry run unless executed).</summary>
    public async Task<List<WorkloadResult>> MigrateAsync(
        GraphClient target,
        IEnumerable<PlannedUser> planned,
        bool dryRun = true,
        CancellationToken ct = default)
    {
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
                results.Add(new WorkloadResult(p.TargetUpn, "would-create"));
                continue;
            }

            var body = p.ToGraphBody(GeneratePassword());
            try
            {
                var created = await target.PostJsonAsync("/users", body, ct);
                var result = new WorkloadResult(p.TargetUpn, "created");
                var id = created.GetStringOrNull("id");
                if (id is not null)
                    result.Detail["target_id"] = id;
                results.Add(result);
            }
            catch (GraphException ex)
            {
                results.Add(new WorkloadResult(p.TargetUpn, "error", ex.Message));
            }
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
