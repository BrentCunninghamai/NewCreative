using System.Collections.ObjectModel;
using System.IO;
using System.Net.Http;
using M365Migrate.App.Logging;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Models;
using M365Migrate.Core.Reporting;
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
    // For the per-user content workloads (Files/Mail/Calendar): the SOURCE user's UPN.
    public string Scope { get; set; } = "";
    // Optional explicit TARGET user UPN. Use when the target identity isn't a clean
    // domain-rewrite of the source — e.g. a user already partly migrated by Microsoft's
    // cross-tenant orchestrator / cross-tenant sync. Blank = derive by rewrite.
    public string ScopeTargetUpn { get; set; } = "";
    public bool RewriteUpn { get; set; } = true;
    public bool SkipGuests { get; set; } = true;
    public bool Execute { get; set; }

    // Optional name prefix/suffix to keep identities distinct when merging
    // multiple source tenants into one target (e.g. prefix "contoso-").
    public string NamePrefix { get; set; } = "";
    public string NameSuffix { get; set; } = "";
    // Appended to migrated users' display names, e.g. "(Contoso)".
    public string DisplayNameSuffix { get; set; } = "";

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

    private CancellationTokenSource? _cts;

    /// <summary>Request cancellation of the in-flight operation.</summary>
    public void Cancel()
    {
        if (_cts is not null)
        {
            _cts.Cancel();
            Status = "Canceling...";
        }
    }

    public ObservableCollection<PlanRow> Rows { get; } = new();

    /// <summary>Where plan/result CSV reports are written.</summary>
    public string ReportsDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "m365-migrate", "reports");

    /// <summary>Write the current grid rows to a timestamped CSV. Never throws.</summary>
    /// <summary>Write the editable source→target mapping CSV (re-importable for a bulk run).</summary>
    private void WriteMappingCsv(IEnumerable<UserMatch> matches)
    {
        try
        {
            Directory.CreateDirectory(ReportsDirectory);
            var file = Path.Combine(ReportsDirectory, $"mapping-{DateTime.Now:yyyyMMdd-HHmmss}.csv");
            File.WriteAllText(file, UserMatchingWorkload.ToCsv(matches));
            AppLog.Write($"mapping CSV written: {file}");
        }
        catch (Exception ex)
        {
            AppLog.Write($"failed to write mapping CSV: {ex.Message}");
        }
    }

    private void WriteReport(string kind)
    {
        try
        {
            Directory.CreateDirectory(ReportsDirectory);
            var safeWorkload = Workload.Replace(" ", "_").Replace("(", "").Replace(")", "");
            var file = Path.Combine(ReportsDirectory, $"{kind}-{safeWorkload}-{DateTime.Now:yyyyMMdd-HHmmss}.csv");
            var csv = CsvReport.ToCsv(
                new[] { "Name", "Action", "Detail", "Reason" },
                Rows.Select(r => (IReadOnlyList<string>)new[] { r.Name, r.Action, r.Detail, r.Reason }));
            File.WriteAllText(file, csv);
            AppLog.Write($"{kind} report written: {file} ({Rows.Count} rows)");
        }
        catch (Exception ex)
        {
            AppLog.Write($"failed to write {kind} report: {ex.Message}");
        }
    }

    private List<PlannedUser>? _plannedUsers;
    private List<PlannedGroup>? _plannedGroups;
    private List<PlannedMailbox>? _plannedMailboxes;
    private List<PlannedDriveItem>? _plannedFiles;
    private string? _filesSourceRoot;
    private string? _filesTargetRoot;
    private List<PlannedTeam>? _plannedTeams;
    private List<PlannedMailFolder>? _plannedMailFolders;
    private string? _mailSourceRef;
    private string? _mailTargetRef;
    private bool _calContactsPlanned;
    private string? _ccSourceRef;
    private string? _ccTargetRef;
    private List<MigratableTeam>? _messageTeams;
    private List<UserMatch>? _userMatches;

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
        Options = new MigrationOptions
        {
            RewriteUpnDomain = RewriteUpn,
            SkipGuests = SkipGuests,
            NamePrefix = NamePrefix.Trim(),
            NameSuffix = NameSuffix.Trim(),
            DisplayNameSuffix = DisplayNameSuffix.Trim(),
        },
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
        _cts = new CancellationTokenSource();
        var ct = _cts.Token;
        try
        {
            var config = BuildConfig();
            Status = "Testing connections to both tenants...";

            using var sourceHttp = new HttpClient();
            using var targetHttp = new HttpClient();
            var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
            var target = new GraphClient(targetHttp, TokenProviders.ForTenant(config.Target));

            var sourceOrg = await ConnectionTester.CheckAsync(source, ct);
            var targetOrg = await ConnectionTester.CheckAsync(target, ct);
            var status = $"Connected OK.  Source: \"{sourceOrg}\"   Target: \"{targetOrg}\".  Ready to Discover & Plan.";

            // Authoritatively confirm a Multi-Tenant Organization link (if the target
            // app has MultiTenantOrganization.Read.All); silent if not consented.
            try
            {
                var mto = await MultiTenantOrgInspector.InspectAsync(target, ct);
                var mtoLine = MultiTenantOrgInspector.Summarize(mto, config.Source.TenantId);
                if (mtoLine is not null)
                    status += "  " + mtoLine;
            }
            catch (Exception)
            {
                // MTO is advisory only — never fail the connection test over it.
            }

            Status = status;
        }
        catch (OperationCanceledException)
        {
            Status = "Connection test canceled.";
        }
        catch (Exception ex)
        {
            Status = "Connection failed: " + ex.Message;
        }
        finally
        {
            _cts.Dispose();
            _cts = null;
            IsBusy = false;
        }
    }

    /// <summary>
    /// Resolve and display the source and target identities for a per-user content run,
    /// so the operator can confirm targeting before any write. Adds an "Identity" row and
    /// returns a short status suffix; warns clearly if either side can't be resolved.
    /// </summary>
    private async Task<string> VerifyIdentitiesAsync(
        GraphClient source, GraphClient target, string sourceRef, string targetRef, CancellationToken ct)
    {
        var src = await UserResolver.TryResolveAsync(source, sourceRef, ct);
        var tgt = await UserResolver.TryResolveAsync(target, targetRef, ct);
        var srcText = src?.Describe() ?? $"NOT FOUND ({UserResolver.NormalizeKey(sourceRef)})";
        var tgtText = tgt?.Describe() ?? $"NOT FOUND ({UserResolver.NormalizeKey(targetRef)})";
        Rows.Add(new PlanRow
        {
            Name = "Identity",
            Action = src is not null && tgt is not null ? "resolved" : "check",
            Detail = $"{srcText}  →  {tgtText}",
            Reason = tgt is null ? "Target not found — set 'Target user UPN' (UPN or object ID)."
                   : src is null ? "Source not found — check the Scope UPN."
                   : "",
        });
        return src is null || tgt is null
            ? "  ⚠ Identity check failed — see the Identity row before migrating."
            : $"  {srcText}  →  {tgtText}.";
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
        _plannedMailFolders = null;
        _calContactsPlanned = false;
        _messageTeams = null;
        _userMatches = null;
        _cts = new CancellationTokenSource();
        var ct = _cts.Token;
        try
        {
            var config = BuildConfig();
            Status = $"Connecting and planning {Workload}...";

            using var sourceHttp = new HttpClient();
            using var targetHttp = new HttpClient();
            var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
            var target = new GraphClient(targetHttp, TokenProviders.ForTenant(config.Target));

            if (Workload == "User mapping (preview)")
            {
                var workload = new UserMatchingWorkload(config);
                _userMatches = await workload.BuildAsync(source, target, ct: ct);
                WriteMappingCsv(_userMatches);
                foreach (var m in _userMatches.OrderBy(m => m.Matched ? 1 : 0))
                    Rows.Add(new PlanRow
                    {
                        Name = m.SourceUpn,
                        Action = m.Matched ? m.Method : "unmatched",
                        Detail = m.TargetUpn ?? "—",
                        Reason = m.Note ?? "",
                    });
                var matched = _userMatches.Count(m => m.Matched);
                Status = $"Mapped {matched}/{_userMatches.Count} source users to the target " +
                         $"({_userMatches.Count - matched} unmatched). Exported mapping CSV to {ReportsDirectory}. " +
                         "Review; edit the CSV to override matches for a bulk run.";
            }
            else if (Workload == "Groups")
            {
                var workload = new GroupsWorkload(config);
                var groups = await workload.DiscoverAsync(source, ct);
                var existing = await GroupsWorkload.GetTargetGroupIdsAsync(target, ct);
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
                var mailboxes = await workload.DiscoverAsync(source, ct);
                var existing = await UsersWorkload.DiscoverTargetUpnsAsync(target, ct);
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
                (_filesSourceRoot, _filesTargetRoot) = workload.ResolveDriveRoots(user: Scope, targetUserOverride: ScopeTargetUpn);
                var idSuffix = await VerifyIdentitiesAsync(source, target, _filesSourceRoot, _filesTargetRoot, ct);
                List<DriveItem> driveItems;
                try
                {
                    driveItems = await workload.DiscoverDriveItemsAsync(source, _filesSourceRoot, ct);
                }
                catch (GraphException ex)
                {
                    // Reading the source OneDrive failed. A 404 means no drive: record an
                    // empty, valid plan so Migrate is a clean no-op. Anything else
                    // (notSupported / 403 / ...) is a fixable failure — leave the plan NULL
                    // so Migrate refuses to run and can't mask it as "0 items processed".
                    var notProvisioned = FilesWorkload.IsDriveNotProvisioned(ex);
                    _plannedFiles = notProvisioned ? new List<PlannedDriveItem>() : null;
                    var hint = FilesWorkload.DriveErrorHint(ex);
                    Rows.Add(new PlanRow
                    {
                        Name = "OneDrive",
                        Action = notProvisioned ? "unavailable" : "error",
                        Detail = "couldn't read source OneDrive",
                        Reason = hint,
                    });
                    Status = $"OneDrive for {Scope}: {hint}" + idSuffix;
                    WriteReport("plan");
                    return;
                }
                _plannedFiles = workload.Plan(driveItems);
                foreach (var p in _plannedFiles)
                    Rows.Add(new PlanRow
                    {
                        Name = p.RelativePath,
                        Action = p.Action,
                        Detail = p.IsFolder ? "folder" : $"file ({p.Size} bytes)",
                        Reason = p.Reason ?? "",
                    });
                Status = $"Planned {_plannedFiles.Count} drive items for {Scope}. Review, then Migrate." + idSuffix;
            }
            else if (Workload == "Teams")
            {
                var workload = new TeamsWorkload(config);
                var teams = await workload.DiscoverAsync(source, ct);
                var existing = await GroupsWorkload.GetTargetGroupIdsAsync(target, ct);
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
            else if (Workload == "Mail (content)")
            {
                var workload = new MailWorkload(config);
                (_mailSourceRef, _mailTargetRef) = workload.ResolveUserRefs(Scope, ScopeTargetUpn);
                var idSuffix = await VerifyIdentitiesAsync(source, target, _mailSourceRef, _mailTargetRef, ct);
                var folders = await workload.DiscoverFoldersAsync(source, _mailSourceRef, ct);
                _plannedMailFolders = workload.Plan(folders);
                foreach (var p in _plannedMailFolders)
                    Rows.Add(new PlanRow
                    {
                        Name = p.Path,
                        Action = p.Action,
                        Detail = $"{p.ItemCount} items",
                        Reason = "",
                    });
                Status = $"Planned {_plannedMailFolders.Count} mail folders " +
                         $"({_plannedMailFolders.Sum(f => f.ItemCount)} items) for {Scope}. Review, then Migrate." + idSuffix;
            }
            else if (Workload == "Calendar & Contacts (content)")
            {
                var workload = new CalendarContactsWorkload(config);
                (_ccSourceRef, _ccTargetRef) = workload.ResolveUserRefs(Scope, ScopeTargetUpn);
                var idSuffix = await VerifyIdentitiesAsync(source, target, _ccSourceRef, _ccTargetRef, ct);
                var (events, contacts) = await workload.CountAsync(source, _ccSourceRef, ct);
                _calContactsPlanned = true;
                Rows.Add(new PlanRow { Name = "Calendar", Action = "copy", Detail = $"{events} events", Reason = "" });
                Rows.Add(new PlanRow { Name = "Contacts", Action = "copy", Detail = $"{contacts} contacts", Reason = "" });
                Status = $"Planned {events} events and {contacts} contacts for {Scope}. Review, then Migrate." + idSuffix;
            }
            else if (Workload == "Teams (messages)")
            {
                var workload = new TeamsMessagesWorkload(config);
                _messageTeams = await workload.DiscoverAsync(source, ct);
                foreach (var t in _messageTeams)
                    Rows.Add(new PlanRow
                    {
                        Name = t.MailNickname ?? t.DisplayName ?? "",
                        Action = "migrate",
                        Detail = $"channels {t.Channels.Count}",
                        Reason = "",
                    });
                Status = $"Planned {_messageTeams.Count} teams for message-history migration. " +
                         "Each becomes a NEW migration-mode team; run once. Review, then Migrate.";
            }
            else
            {
                var workload = new UsersWorkload(config);
                var users = await workload.DiscoverAsync(source, ct);
                var existing = await UsersWorkload.DiscoverTargetUpnsAsync(target, ct);
                var crossTenant = await UsersWorkload.DiscoverCrossTenantIdentitiesAsync(target, ct);
                _plannedUsers = workload.Plan(users, existing, crossTenant);
                foreach (var p in _plannedUsers)
                    Rows.Add(new PlanRow
                    {
                        Name = p.TargetUpn,
                        Action = p.Action,
                        Detail = $"{p.SourceUpn}  [{p.UserType}]",
                        Reason = p.Reason ?? "",
                    });
                Status = $"Planned {_plannedUsers.Count} users. Review, then Migrate.";
            }

            WriteReport("plan");
        }
        catch (OperationCanceledException)
        {
            AppLog.Write($"plan {Workload} canceled");
            Status = "Plan canceled.";
        }
        catch (Exception ex)
        {
            AppLog.Write($"plan {Workload} failed: {ex}");
            Status = "Error: " + ex.Message;
        }
        finally
        {
            _cts.Dispose();
            _cts = null;
            IsBusy = false;
        }
    }

    /// <summary>Run the migration for the current plan (dry run unless Execute is set).</summary>
    public async Task MigrateAsync()
    {
        if (IsBusy) return;
        if (Workload == "User mapping (preview)")
        {
            Status = "User mapping is a preview — no changes to apply. Use it to review/export the " +
                     "source→target map, then run the actual workloads (Mail, Files, ...).";
            return;
        }
        if (_plannedUsers is null && _plannedGroups is null && _plannedMailboxes is null
            && _plannedFiles is null && _plannedTeams is null && _plannedMailFolders is null
            && !_calContactsPlanned && _messageTeams is null)
        {
            Status = "Nothing planned yet — run Discover & Plan first.";
            return;
        }

        IsBusy = true;
        _cts = new CancellationTokenSource();
        var ct = _cts.Token;
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
                results = await new GroupsWorkload(config).SyncAsync(target, _plannedGroups, dryRun: !Execute, ct: ct);
            }
            else if (Workload == "Mailboxes" && _plannedMailboxes is not null)
            {
                results = await new MailboxesWorkload(config).MigrateAsync(target, _plannedMailboxes, dryRun: !Execute, ct: ct);
            }
            else if (Workload == "Files (OneDrive)" && _plannedFiles is not null)
            {
                using var sourceHttp = new HttpClient();
                var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
                results = await new FilesWorkload(config).MigrateDriveItemsAsync(
                    source, target, _plannedFiles, _filesSourceRoot!, _filesTargetRoot!, dryRun: !Execute, ct: ct);
            }
            else if (Workload == "Teams" && _plannedTeams is not null)
            {
                results = await new TeamsWorkload(config).MigrateAsync(target, _plannedTeams, dryRun: !Execute, ct: ct);
            }
            else if (Workload == "Mail (content)" && _plannedMailFolders is not null)
            {
                using var sourceHttp = new HttpClient();
                var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
                results = await new MailWorkload(config).MigrateAsync(
                    source, target, _mailSourceRef!, _mailTargetRef!, _plannedMailFolders, dryRun: !Execute, ct: ct);
            }
            else if (Workload == "Calendar & Contacts (content)" && _calContactsPlanned)
            {
                using var sourceHttp = new HttpClient();
                var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
                results = await new CalendarContactsWorkload(config).MigrateAsync(
                    source, target, _ccSourceRef!, _ccTargetRef!, dryRun: !Execute, ct: ct);
            }
            else if (Workload == "Teams (messages)" && _messageTeams is not null)
            {
                using var sourceHttp = new HttpClient();
                var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
                results = await new TeamsMessagesWorkload(config).MigrateAsync(
                    source, target, _messageTeams, dryRun: !Execute, ct: ct);
            }
            else if (_plannedUsers is not null)
            {
                results = await new UsersWorkload(config).MigrateAsync(target, _plannedUsers, dryRun: !Execute, ct: ct);
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
            WriteReport("results");
            Status = $"{mode} complete — {results.Count} items processed." +
                     (Execute ? "" : " No changes were made; tick Execute to apply.") +
                     $"  Report saved to {ReportsDirectory}.";
            AppLog.Write($"{mode} {Workload}: {results.Count} items processed");
        }
        catch (OperationCanceledException)
        {
            AppLog.Write($"migrate {Workload} canceled");
            Status = "Migration canceled. Items already processed were applied; re-run to continue.";
        }
        catch (Exception ex)
        {
            AppLog.Write($"migrate {Workload} failed: {ex}");
            Status = "Error: " + ex.Message;
        }
        finally
        {
            _cts.Dispose();
            _cts = null;
            IsBusy = false;
        }
    }
}
