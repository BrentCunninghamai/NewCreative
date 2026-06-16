using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Mapping;
using M365Migrate.Core.Models;

namespace M365Migrate.Core.Workloads;

/// <summary>
/// Mailbox workload: discover, plan, and migrate Exchange Online mailbox
/// *settings* (time zone, language, working hours, automatic replies, date/time
/// formats, delegate options). Mailbox *content* requires a native cross-tenant
/// move and is out of scope for the Graph layer.
/// </summary>
public sealed class MailboxesWorkload
{
    private readonly MigrationConfig _config;

    public MailboxesWorkload(MigrationConfig config) => _config = config;

    private string Rewrite(string upn) =>
        _config.Options.RewriteUpnDomain
            ? UpnMapper.Rewrite(upn, _config.Source.PrimaryDomain, _config.Target.PrimaryDomain)
            : upn;

    /// <summary>Read mailbox settings for every mailbox-enabled source user.</summary>
    public async Task<List<SourceMailbox>> DiscoverAsync(GraphClient source, CancellationToken ct = default)
    {
        var users = await source.GetAllAsync("/users?$select=id,userPrincipalName,mail&$top=999", ct);
        var mailboxes = new List<SourceMailbox>();
        foreach (var user in users)
        {
            var upn = user.GetStringOrNull("userPrincipalName");
            var mail = user.GetStringOrNull("mail");
            var id = user.GetStringOrNull("id");
            // No SMTP address almost always means no Exchange mailbox; skip the probe.
            if (upn is null || mail is null || id is null)
                continue;
            try
            {
                var data = await source.GetAsync($"/users/{id}/mailboxSettings", ct);
                mailboxes.Add(SourceMailbox.FromGraph(upn, data));
            }
            catch (GraphException ex) when (ex.StatusCode == 404)
            {
                // mailbox not provisioned for this user
            }
        }
        return mailboxes;
    }

    /// <summary>Build a mailbox-settings migration plan without writing anything.</summary>
    public List<PlannedMailbox> Plan(IEnumerable<SourceMailbox> mailboxes, ISet<string>? targetUpns = null)
    {
        var existing = targetUpns is null
            ? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            : new HashSet<string>(targetUpns, StringComparer.OrdinalIgnoreCase);

        var planned = new List<PlannedMailbox>();
        foreach (var mb in mailboxes)
        {
            var target = Rewrite(mb.UserPrincipalName);
            string action;
            string? reason;

            if (!existing.Contains(target))
            {
                action = "skip";
                reason = "target account not found";
            }
            else if (mb.Settings.Count == 0)
            {
                action = "skip";
                reason = "no settings to migrate";
            }
            else
            {
                action = "settings";
                reason = null;
            }

            planned.Add(new PlannedMailbox
            {
                SourceUpn = mb.UserPrincipalName,
                TargetUpn = target,
                Action = action,
                Reason = reason,
                Settings = mb.Settings,
            });
        }
        return planned;
    }

    /// <summary>Apply planned mailbox settings to the target tenant (dry run unless executed).</summary>
    public async Task<List<WorkloadResult>> MigrateAsync(
        GraphClient target,
        IEnumerable<PlannedMailbox> planned,
        bool dryRun = true,
        CancellationToken ct = default)
    {
        var targetIds = await UsersWorkload.GetTargetUserIdsAsync(target, ct);
        var results = new List<WorkloadResult>();

        foreach (var p in planned)
        {
            var record = new WorkloadResult { Name = p.TargetUpn };

            if (p.Action != "settings")
            {
                record.Status = $"skipped:{p.Action}";
                record.Reason = p.Reason;
                results.Add(record);
                continue;
            }
            if (!targetIds.TryGetValue(p.TargetUpn, out var id))
            {
                record.Status = "skipped:not-in-target";
                results.Add(record);
                continue;
            }
            if (dryRun)
            {
                record.Status = "would-update";
                results.Add(record);
                continue;
            }

            try
            {
                await target.PatchJsonAsync($"/users/{id}/mailboxSettings", p.Settings, ct);
                record.Status = "updated";
            }
            catch (GraphException ex)
            {
                record.Status = "error";
                record.Reason = ex.StatusCode.ToString();
            }
            results.Add(record);
        }
        return results;
    }
}
