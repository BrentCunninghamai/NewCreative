using System.Collections.ObjectModel;
using System.Net.Http;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Models;
using M365Migrate.Core.Workloads;

namespace M365Migrate.App.ViewModels;

/// <summary>
/// Drives the single-window UI: holds connection inputs, runs discover/plan and
/// migrate against the engine, and exposes results for the grid.
/// </summary>
public sealed class MainViewModel : ViewModelBase
{
    // --- connection inputs (populated from the view before each run) ---
    public string SourceTenantId { get; set; } = "";
    public string SourceClientId { get; set; } = "";
    public string SourceClientSecret { get; set; } = "";
    public string SourceDomain { get; set; } = "";
    public string TargetTenantId { get; set; } = "";
    public string TargetClientId { get; set; } = "";
    public string TargetClientSecret { get; set; } = "";
    public string TargetDomain { get; set; } = "";

    public string Workload { get; set; } = "Users";
    // For the Files workload: the source user's UPN (their OneDrive).
    public string Scope { get; set; } = "";
    public bool RewriteUpn { get; set; } = true;
    public bool SkipGuests { get; set; } = true;
    public bool Execute { get; set; }

    private bool _isBusy;
    public bool IsBusy
    {
        get => _isBusy;
        private set => SetField(ref _isBusy, value);
    }

    private string _status = "Enter both tenants' app-registration details, then Discover & Plan.";
    public string Status
    {
        get => _status;
        private set => SetField(ref _status, value);
    }

    public ObservableCollection<PlanRow> Rows { get; } = new();

    private List<PlannedUser>? _plannedUsers;
    private List<PlannedGroup>? _plannedGroups;
    private List<PlannedMailbox>? _plannedMailboxes;
    private List<PlannedDriveItem>? _plannedFiles;
    private string? _filesSourceRoot;
    private string? _filesTargetRoot;
    private List<PlannedTeam>? _plannedTeams;

    private MigrationConfig BuildConfig() => new()
    {
        Source = new TenantConfig
        {
            TenantId = SourceTenantId.Trim(),
            ClientId = SourceClientId.Trim(),
            ClientSecret = SourceClientSecret,
            PrimaryDomain = SourceDomain.Trim(),
        },
        Target = new TenantConfig
        {
            TenantId = TargetTenantId.Trim(),
            ClientId = TargetClientId.Trim(),
            ClientSecret = TargetClientSecret,
            PrimaryDomain = TargetDomain.Trim(),
        },
        Options = new MigrationOptions { RewriteUpnDomain = RewriteUpn, SkipGuests = SkipGuests },
    };

    private string? Validate()
    {
        if (string.IsNullOrWhiteSpace(SourceTenantId) || string.IsNullOrWhiteSpace(SourceClientId)
            || string.IsNullOrWhiteSpace(SourceClientSecret) || string.IsNullOrWhiteSpace(SourceDomain))
            return "Fill in all source tenant fields.";
        if (string.IsNullOrWhiteSpace(TargetTenantId) || string.IsNullOrWhiteSpace(TargetClientId)
            || string.IsNullOrWhiteSpace(TargetClientSecret) || string.IsNullOrWhiteSpace(TargetDomain))
            return "Fill in all target tenant fields.";
        return null;
    }

    /// <summary>Verify both tenants authenticate and Graph is reachable.</summary>
    public async Task TestConnectionsAsync()
    {
        if (IsBusy) return;
        var error = Validate();
        if (error is not null) { Status = error; return; }

        IsBusy = true;
        try
        {
            var config = BuildConfig();
            Status = "Testing connections to both tenants...";

            using var sourceHttp = new HttpClient();
            using var targetHttp = new HttpClient();
            var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
            var target = new GraphClient(targetHttp, TokenProviders.ForTenant(config.Target));

            var sourceOrg = await ConnectionTester.CheckAsync(source);
            var targetOrg = await ConnectionTester.CheckAsync(target);
            Status = $"Connected OK.  Source: \"{sourceOrg}\"   Target: \"{targetOrg}\".  Ready to Discover & Plan.";
        }
        catch (Exception ex)
        {
            Status = "Connection failed: " + ex.Message;
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Connect to both tenants, discover the source, and build a plan.</summary>
    public async Task PlanAsync()
    {
        if (IsBusy) return;
        var error = Validate();
        if (error is not null) { Status = error; return; }

        IsBusy = true;
        Rows.Clear();
        _plannedUsers = null;
        _plannedGroups = null;
        _plannedMailboxes = null;
        _plannedFiles = null;
        _plannedTeams = null;
        try
        {
            var config = BuildConfig();
            Status = $"Connecting and planning {Workload}...";

            using var sourceHttp = new HttpClient();
            using var targetHttp = new HttpClient();
            var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
            var target = new GraphClient(targetHttp, TokenProviders.ForTenant(config.Target));

            if (Workload == "Groups")
            {
                var workload = new GroupsWorkload(config);
                var groups = await workload.DiscoverAsync(source);
                var existing = await GroupsWorkload.GetTargetGroupIdsAsync(target);
                _plannedGroups = workload.Plan(groups, existing);
                foreach (var p in _plannedGroups)
                    Rows.Add(new PlanRow
                    {
                        Name = p.MailNickname ?? "",
                        Action = p.Action,
                        Detail = p.IsDynamic ? "dynamic" : $"members:{p.TargetMemberUpns.Count} owners:{p.TargetOwnerUpns.Count}",
                        Reason = p.Reason ?? "",
                    });
                Status = $"Planned {_plannedGroups.Count} groups. Review, then Migrate.";
            }
            else if (Workload == "Mailboxes")
            {
                var workload = new MailboxesWorkload(config);
                var mailboxes = await workload.DiscoverAsync(source);
                var existing = await UsersWorkload.DiscoverTargetUpnsAsync(target);
                _plannedMailboxes = workload.Plan(mailboxes, existing);
                foreach (var p in _plannedMailboxes)
                    Rows.Add(new PlanRow
                    {
                        Name = p.TargetUpn,
                        Action = p.Action,
                        Detail = $"settings:{p.Settings.Count}",
                        Reason = p.Reason ?? "",
                    });
                Status = $"Planned {_plannedMailboxes.Count} mailboxes. Review, then Migrate.";
            }
            else if (Workload == "Files (OneDrive)")
            {
                var workload = new FilesWorkload(config);
                (_filesSourceRoot, _filesTargetRoot) = workload.ResolveDriveRoots(user: Scope);
                var driveItems = await workload.DiscoverDriveItemsAsync(source, _filesSourceRoot);
                _plannedFiles = workload.Plan(driveItems);
                foreach (var p in _plannedFiles)
                    Rows.Add(new PlanRow
                    {
                        Name = p.RelativePath,
                        Action = p.Action,
                        Detail = p.IsFolder ? "folder" : $"file ({p.Size} bytes)",
                        Reason = p.Reason ?? "",
                    });
                Status = $"Planned {_plannedFiles.Count} drive items for {Scope}. Review, then Migrate.";
            }
            else if (Workload == "Teams")
            {
                var workload = new TeamsWorkload(config);
                var teams = await workload.DiscoverAsync(source);
                var existing = await GroupsWorkload.GetTargetGroupIdsAsync(target);
                _plannedTeams = workload.Plan(teams, existing);
                foreach (var p in _plannedTeams)
                {
                    var creatable = p.Channels.Count(c => c.Action == "create");
                    Rows.Add(new PlanRow
                    {
                        Name = p.MailNickname ?? "",
                        Action = p.Action,
                        Detail = $"channels {creatable}/{p.Channels.Count}",
                        Reason = p.Reason ?? "",
                    });
                }
                Status = $"Planned {_plannedTeams.Count} teams. Review, then Migrate.";
            }
            else
            {
                var workload = new UsersWorkload(config);
                var users = await workload.DiscoverAsync(source);
                var existing = await UsersWorkload.DiscoverTargetUpnsAsync(target);
                _plannedUsers = workload.Plan(users, existing);
                foreach (var p in _plannedUsers)
                    Rows.Add(new PlanRow
                    {
                        Name = p.TargetUpn,
                        Action = p.Action,
                        Detail = p.SourceUpn,
                        Reason = p.Reason ?? "",
                    });
                Status = $"Planned {_plannedUsers.Count} users. Review, then Migrate.";
            }
        }
        catch (Exception ex)
        {
            Status = "Error: " + ex.Message;
        }
        finally
        {
            IsBusy = false;
        }
    }

    /// <summary>Run the migration for the current plan (dry run unless Execute is set).</summary>
    public async Task MigrateAsync()
    {
        if (IsBusy) return;
        if (_plannedUsers is null && _plannedGroups is null && _plannedMailboxes is null
            && _plannedFiles is null && _plannedTeams is null)
        {
            Status = "Nothing planned yet — run Discover & Plan first.";
            return;
        }

        IsBusy = true;
        try
        {
            var config = BuildConfig();
            var mode = Execute ? "EXECUTE" : "DRY RUN";
            Status = $"{mode}: migrating {Workload}...";

            using var targetHttp = new HttpClient();
            var target = new GraphClient(targetHttp, TokenProviders.ForTenant(config.Target));

            List<WorkloadResult> results;
            if (Workload == "Groups" && _plannedGroups is not null)
            {
                results = await new GroupsWorkload(config).SyncAsync(target, _plannedGroups, dryRun: !Execute);
            }
            else if (Workload == "Mailboxes" && _plannedMailboxes is not null)
            {
                results = await new MailboxesWorkload(config).MigrateAsync(target, _plannedMailboxes, dryRun: !Execute);
            }
            else if (Workload == "Files (OneDrive)" && _plannedFiles is not null)
            {
                using var sourceHttp = new HttpClient();
                var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
                results = await new FilesWorkload(config).MigrateDriveItemsAsync(
                    source, target, _plannedFiles, _filesSourceRoot!, _filesTargetRoot!, dryRun: !Execute);
            }
            else if (Workload == "Teams" && _plannedTeams is not null)
            {
                results = await new TeamsWorkload(config).MigrateAsync(target, _plannedTeams, dryRun: !Execute);
            }
            else if (_plannedUsers is not null)
            {
                results = await new UsersWorkload(config).MigrateAsync(target, _plannedUsers, dryRun: !Execute);
            }
            else
            {
                Status = "Plan and workload do not match — re-run Discover & Plan.";
                return;
            }

            Rows.Clear();
            foreach (var r in results)
                Rows.Add(new PlanRow
                {
                    Name = r.Name,
                    Action = r.Status,
                    Detail = string.Join("  ", r.Detail.Select(kv => $"{kv.Key}={kv.Value}")),
                    Reason = r.Reason ?? "",
                });
            Status = $"{mode} complete — {results.Count} items processed." +
                     (Execute ? "" : " No changes were made; tick Execute to apply.");
        }
        catch (Exception ex)
        {
            Status = "Error: " + ex.Message;
        }
        finally
        {
            IsBusy = false;
        }
    }
}
