using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Net.Http;
using System.Windows.Data;
using System.Windows.Media;
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
    // Users workload: assign target-equivalent licenses (matched by SKU part number) on create.
    public bool AssignLicenses { get; set; }
    public string DefaultUsageLocation { get; set; } = "";

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
        set
        {
            if (SetField(ref _status, value))
                OnPropertyChanged(nameof(StatusBrush));
        }
    }

    /// <summary>Colour the status line by severity (red error / amber warning / green done).</summary>
    public Brush StatusBrush
    {
        get
        {
            var s = _status.ToLowerInvariant();
            if (s.Contains("error") || s.Contains("failed") || s.Contains("missing") || s.Contains("⚠"))
                return Brushes.Firebrick;
            if (s.Contains("canceled") || s.Contains("skipped") || s.Contains("issue") || s.Contains("conflict"))
                return Brushes.DarkGoldenrod;
            if (s.Contains("complete") || s.Contains("passed") || s.Contains("done") || s.Contains("ready") || s.Contains("saved") || s.Contains("loaded"))
                return Brushes.ForestGreen;
            return Brushes.Black;
        }
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

    /// <summary>Filtered/searchable view of <see cref="Rows"/> bound by the grid.</summary>
    public ICollectionView RowsView { get; }

    private string _filterText = "";
    /// <summary>Free-text filter over the results grid (name / action / detail / reason).</summary>
    public string FilterText
    {
        get => _filterText;
        set
        {
            if (SetField(ref _filterText, value))
                RowsView.Refresh();
        }
    }

    public MainViewModel()
    {
        RowsView = CollectionViewSource.GetDefaultView(Rows);
        RowsView.Filter = RowMatchesFilter;
    }

    private bool RowMatchesFilter(object item)
    {
        if (string.IsNullOrWhiteSpace(_filterText))
            return true;
        if (item is not PlanRow r)
            return true;
        var q = _filterText.Trim();
        return (r.Name?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
            || (r.Action?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
            || (r.Detail?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
            || (r.Reason?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false);
    }

    /// <summary>Where plan/result CSV reports are written.</summary>
    public string ReportsDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "m365-migrate", "reports");

    /// <summary>Write the current grid rows to a timestamped CSV. Never throws.</summary>
    /// <summary>
    /// Write the one-time app-setup guide (permission manifest + 1-click admin-consent links)
    /// to the reports folder and return its path, so a Global Admin can grant everything fast.
    /// </summary>
    public string? WriteSetupGuide()
    {
        try
        {
            Directory.CreateDirectory(ReportsDirectory);
            var guide = GraphSetup.SetupGuide(
                SourceTenantId.Trim(), SourceClientId.Trim(),
                TargetTenantId.Trim(), TargetClientId.Trim());
            var file = Path.Combine(ReportsDirectory, $"app-setup-{DateTime.Now:yyyyMMdd-HHmmss}.txt");
            File.WriteAllText(file, guide);
            Status = "App setup guide written. Paste the manifest into each app registration's " +
                     "Manifest, then open the admin-consent links (in the file) as Global Admin.";
            AppLog.Write($"setup guide written: {file}");
            return file;
        }
        catch (Exception ex)
        {
            Status = "Couldn't write setup guide: " + ex.Message;
            return null;
        }
    }

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
    private List<(string Library, string SourceRoot, string TargetRoot, List<PlannedDriveItem> Items)>? _spJobs;

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

    /// <summary>
    /// Pre-flight readiness: for each tenant, acquire a token and verify which required Graph
    /// application permissions are admin-consented (from the token's roles claim), so missing
    /// consent is caught before a run instead of as a mid-run 403.
    /// </summary>
    public async Task PreflightAsync()
    {
        if (IsBusy) return;
        var error = Validate();
        if (error is not null) { Status = error; return; }

        IsBusy = true;
        Rows.Clear();
        _cts = new CancellationTokenSource();
        var ct = _cts.Token;
        try
        {
            var config = BuildConfig();
            // Check only what the selected workload needs (least-privilege friendly).
            var required = GraphSetup.RequiredFor(Workload);
            var totalMissing = 0;
            foreach (var (label, tenant) in new[] { ("Source", config.Source), ("Target", config.Target) })
            {
                Status = $"Pre-flight: acquiring {label} token...";
                try
                {
                    var token = await TokenProviders.ForTenant(tenant)(ct);
                    var result = PreflightChecker.Check(token, required);
                    totalMissing += result.Missing.Count;
                    Rows.Add(new PlanRow
                    {
                        Name = $"{label} — {tenant.TenantId}",
                        Action = result.AllGranted ? "ready" : "missing",
                        Detail = $"{result.Granted.Count}/{required.Count} consented",
                        Reason = result.Missing.Count == 0 ? "" : "missing: " + string.Join(", ", result.Missing),
                    });
                }
                catch (Exception ex)
                {
                    totalMissing++;
                    Rows.Add(new PlanRow { Name = $"{label} — {tenant.TenantId}", Action = "error", Detail = "token/auth failed", Reason = ex.Message });
                }
            }
            WriteReport("preflight");
            Status = totalMissing == 0
                ? $"Pre-flight passed for “{Workload}” — both tenants authenticate and the needed permissions are consented. Ready to migrate."
                : $"Pre-flight for “{Workload}” found issues ({totalMissing}). Fix consent with “App setup”, then re-check. " +
                  "Note: this verifies admin-consented permissions; some workloads still need the right license on each user.";
        }
        catch (OperationCanceledException) { Status = "Pre-flight canceled."; }
        catch (Exception ex) { Status = "Pre-flight error: " + ex.Message; }
        finally { _cts.Dispose(); _cts = null; IsBusy = false; }
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
        _spJobs = null;
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
                foreach (var m in _userMatches.OrderBy(m => m.ContentReady ? 2 : m.Matched ? 1 : 0))
                    Rows.Add(new PlanRow
                    {
                        Name = m.SourceUpn,
                        Action = !m.Matched ? "unmatched" : m.TargetIsGuest ? "guest" : m.Method,
                        Detail = m.TargetUpn ?? "—",
                        Reason = m.Note ?? "",
                    });
                var contentReady = _userMatches.Count(m => m.ContentReady);
                var guestOnly = _userMatches.Count(m => m.Matched && m.TargetIsGuest);
                var unmatched = _userMatches.Count(m => !m.Matched);
                Status = $"{_userMatches.Count} source users: {contentReady} content-ready (native target), " +
                         $"{guestOnly} matched only to a guest (#EXT#), {unmatched} unmatched. " +
                         $"Mapping CSV exported to {ReportsDirectory}. Filter “guest”/“unmatched” to review.";
            }
            else if (Workload == "Bulk Mail (mapped users)" || Workload == "Bulk OneDrive (mapped users)")
            {
                var content = Workload.Contains("Mail") ? "Mail" : "OneDrive";
                _userMatches = await new UserMatchingWorkload(config).BuildAsync(source, target, ct: ct);
                WriteMappingCsv(_userMatches);
                var ready = _userMatches.Where(m => m.ContentReady).ToList();
                foreach (var m in ready)
                    Rows.Add(new PlanRow
                    {
                        Name = m.SourceUpn,
                        Action = "queued",
                        Detail = m.TargetUpn ?? "",
                        Reason = m.Method,
                    });
                var guestOnly = _userMatches.Count(m => m.Matched && m.TargetIsGuest);
                var unmatched = _userMatches.Count(m => !m.Matched);
                Status = $"{ready.Count} content-ready users queued for bulk {content} " +
                         $"({guestOnly} matched only to a guest, {unmatched} unmatched — both skipped; create native " +
                         $"target accounts for them first). Migrate runs the queued set — dry run unless Execute. " +
                         $"Mapping CSV exported to {ReportsDirectory}.";
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
                    // For a non-404 failure (notSupported/403), probe the SOURCE user's SharePoint
                    // service-plan so we can tell licensing (cut-over user) from multi-geo.
                    if (!notProvisioned)
                    {
                        try { hint += "  |  source " + await UsersWorkload.SharePointPlanStatusAsync(source, UserResolver.NormalizeKey(_filesSourceRoot!), ct); }
                        catch (GraphException) { /* diagnostic is best-effort */ }
                    }
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
                // Delta: mark items already present in the target so the plan shows only changes.
                var fileTotal = _plannedFiles.Count(x => !x.IsFolder);
                var alreadyThere = 0;
                try
                {
                    var (tf, tfo) = FilesWorkload.IndexTarget(await workload.DiscoverDriveItemsAsync(target, _filesTargetRoot!, ct));
                    alreadyThere = FilesWorkload.MarkUnchanged(_plannedFiles, tf, tfo);
                }
                catch (GraphException ex) when (FilesWorkload.IsDriveNotProvisioned(ex)) { /* empty target → copy all */ }
                foreach (var p in _plannedFiles)
                    Rows.Add(new PlanRow
                    {
                        Name = p.RelativePath,
                        Action = p.Action,
                        Detail = p.IsFolder ? "folder" : $"file ({p.Size} bytes)",
                        Reason = p.Reason ?? "",
                    });
                var pct = fileTotal == 0 ? 100 : 100 * alreadyThere / fileTotal;
                Status = $"Planned {_plannedFiles.Count} drive items for {Scope} — {alreadyThere}/{fileTotal} files " +
                         $"already in target ({pct}% synced); only changes will copy. Review, then Migrate." + idSuffix;
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
                List<MailFolderInfo> folders;
                try
                {
                    folders = await workload.DiscoverFoldersAsync(source, _mailSourceRef, ct);
                }
                catch (GraphException ex) when (MailWorkload.IsMailboxUnavailable(ex))
                {
                    // Hybrid / on-prem mailbox (or none) — leave the plan null so Migrate won't run.
                    _plannedMailFolders = null;
                    var hint = MailWorkload.MailboxErrorHint(ex);
                    Rows.Add(new PlanRow { Name = "Mailbox", Action = "unavailable", Detail = "couldn't read source mailbox", Reason = hint });
                    Status = $"Mail for {Scope}: {hint}" + idSuffix;
                    WriteReport("plan");
                    return;
                }
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
            else if (Workload == "SharePoint (site)")
            {
                var sp = new SharePointWorkload(config);
                var files = new FilesWorkload(config);
                var srcSite = await sp.ResolveSiteIdAsync(source, Scope, ct);
                var tgtSite = await sp.ResolveSiteIdAsync(target, ScopeTargetUpn, ct);
                var (pairs, unmatched) = SharePointWorkload.MatchDrives(
                    await sp.ListDrivesAsync(source, srcSite, ct),
                    await sp.ListDrivesAsync(target, tgtSite, ct));

                _spJobs = new();
                var totalFiles = 0;
                var totalSynced = 0;
                foreach (var (sd, td) in pairs)
                {
                    var srcRoot = SharePointWorkload.DriveRoot(sd.Id);
                    var tgtRoot = SharePointWorkload.DriveRoot(td.Id);
                    var planned = files.Plan(await files.DiscoverDriveItemsAsync(source, srcRoot, ct));
                    var fileCount = planned.Count(x => !x.IsFolder);
                    var (tf, tfo) = FilesWorkload.IndexTarget(await files.DiscoverDriveItemsAsync(target, tgtRoot, ct));
                    var unchanged = FilesWorkload.MarkUnchanged(planned, tf, tfo);
                    _spJobs.Add((sd.Name, srcRoot, tgtRoot, planned));
                    totalFiles += fileCount;
                    totalSynced += unchanged;
                    Rows.Add(new PlanRow
                    {
                        Name = sd.Name,
                        Action = "queued",
                        Detail = $"{fileCount} files, {unchanged} already in target",
                        Reason = fileCount == 0 ? "" : $"{100 * unchanged / fileCount}% synced",
                    });
                }
                foreach (var name in unmatched)
                    Rows.Add(new PlanRow { Name = name, Action = "unmatched", Detail = "no target library of this name", Reason = "create it on the target site" });

                var pct = totalFiles == 0 ? 100 : 100 * totalSynced / totalFiles;
                Status = $"Planned {pairs.Count} document libraries ({totalFiles} files, {pct}% already synced; " +
                         $"{unmatched.Count} unmatched). Review, then Migrate (dry run unless Execute).";
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
                        Detail = $"{p.SourceUpn}  [{p.UserType}{(p.OnPremisesSynced ? ", hybrid" : "")}]",
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
        var isBulk = Workload == "Bulk Mail (mapped users)" || Workload == "Bulk OneDrive (mapped users)";
        if (_plannedUsers is null && _plannedGroups is null && _plannedMailboxes is null
            && _plannedFiles is null && _plannedTeams is null && _plannedMailFolders is null
            && !_calContactsPlanned && _messageTeams is null && _spJobs is null
            && !(isBulk && _userMatches is not null))
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
            else if ((Workload == "Bulk Mail (mapped users)" || Workload == "Bulk OneDrive (mapped users)") && _userMatches is not null)
            {
                using var sourceHttp = new HttpClient();
                var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
                results = Workload.Contains("Mail")
                    ? await RunBulkMailAsync(config, source, target, mode, ct)
                    : await RunBulkOneDriveAsync(config, source, target, mode, ct);
            }
            else if (Workload == "SharePoint (site)" && _spJobs is not null)
            {
                using var sourceHttp = new HttpClient();
                var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
                var files = new FilesWorkload(config);
                results = new List<WorkloadResult>();
                var n = 0;
                foreach (var job in _spJobs)
                {
                    ct.ThrowIfCancellationRequested();
                    n++;
                    Status = $"{mode}: SharePoint library {n}/{_spJobs.Count} — {job.Library}";
                    var per = await files.MigrateDriveItemsAsync(source, target, job.Items, job.SourceRoot, job.TargetRoot, dryRun: !Execute, ct);
                    var r = new WorkloadResult { Name = job.Library, Status = Execute ? "ok" : "would-copy" };
                    r.Detail["items"] = job.Items.Count.ToString();
                    var errs = per.Count(x => x.Status == "error");
                    if (errs > 0) r.Detail["errors"] = errs.ToString();
                    results.Add(r);
                }
            }
            else if (_plannedUsers is not null)
            {
                if (AssignLicenses)
                {
                    using var sourceHttp = new HttpClient();
                    var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
                    var sourceSkuMap = await UsersWorkload.SkuIdToPartAsync(source, ct);
                    var targetSkuMap = await UsersWorkload.PartToSkuIdAsync(target, ct);
                    var ul = DefaultUsageLocation.Trim();
                    results = await new UsersWorkload(config).MigrateAsync(
                        target, _plannedUsers, dryRun: !Execute,
                        sourceSkuMap, targetSkuMap, ul.Length > 0 ? ul : null, ct);
                }
                else
                {
                    results = await new UsersWorkload(config).MigrateAsync(target, _plannedUsers, dryRun: !Execute, ct: ct);
                }
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

    private static int SumDetail(IEnumerable<WorkloadResult> rs, string key) =>
        rs.Sum(x => x.Detail.TryGetValue(key, out var v) && int.TryParse(v, out var n) ? n : 0);

    private static int AverageSynced(IReadOnlyList<WorkloadResult> rs)
    {
        var vals = rs.Where(r => r.Detail.TryGetValue("synced%", out _))
            .Select(r => int.TryParse(r.Detail["synced%"], out var n) ? n : 0).ToList();
        return vals.Count == 0 ? 0 : (int)Math.Round(vals.Average());
    }

    /// <summary>
    /// Repeatedly run the selected Bulk workload on an interval — pre-seed to ~100% while
    /// users keep working on the source, so cutover is a tiny final pass. Each round rebuilds
    /// the mapping (picks up new users) and copies only the delta. Stops on Cancel.
    /// </summary>
    public async Task StartAutoSyncAsync(int intervalMinutes)
    {
        if (IsBusy) return;
        var isBulk = Workload == "Bulk Mail (mapped users)" || Workload == "Bulk OneDrive (mapped users)";
        if (!isBulk) { Status = "Auto-sync needs a Bulk workload — pick Bulk Mail or Bulk OneDrive."; return; }
        if (!Execute) { Status = "Tick Execute to auto-sync — each round copies the delta."; return; }
        var error = Validate();
        if (error is not null) { Status = error; return; }
        if (intervalMinutes < 1) intervalMinutes = 30;

        IsBusy = true;
        _cts = new CancellationTokenSource();
        var ct = _cts.Token;
        try
        {
            var config = BuildConfig();
            using var sourceHttp = new HttpClient();
            using var targetHttp = new HttpClient();
            var source = new GraphClient(sourceHttp, TokenProviders.ForTenant(config.Source));
            var target = new GraphClient(targetHttp, TokenProviders.ForTenant(config.Target));
            var round = 0;
            while (!ct.IsCancellationRequested)
            {
                round++;
                _userMatches = await new UserMatchingWorkload(config).BuildAsync(source, target, ct: ct);
                var results = Workload.Contains("Mail")
                    ? await RunBulkMailAsync(config, source, target, $"AUTO-SYNC #{round}", ct)
                    : await RunBulkOneDriveAsync(config, source, target, $"AUTO-SYNC #{round}", ct);

                Rows.Clear();
                foreach (var r in results)
                    Rows.Add(new PlanRow
                    {
                        Name = r.Name,
                        Action = r.Status,
                        Detail = string.Join("  ", r.Detail.Select(kv => $"{kv.Key}={kv.Value}")),
                        Reason = r.Reason ?? "",
                    });
                WriteReport("autosync");
                var avg = AverageSynced(results);
                Status = $"Auto-sync round {round} done — {results.Count} users, avg {avg}% synced. " +
                         $"Next pass in {intervalMinutes} min. Cancel to stop (then do a final pass at cutover).";
                AppLog.Write($"auto-sync round {round}: {results.Count} users, avg {avg}% synced");

                try { await Task.Delay(TimeSpan.FromMinutes(intervalMinutes), ct); }
                catch (OperationCanceledException) { break; }
            }
            Status = $"Auto-sync stopped after {round} round(s).";
        }
        catch (OperationCanceledException) { Status = "Auto-sync canceled."; }
        catch (Exception ex) { AppLog.Write($"auto-sync failed: {ex}"); Status = "Auto-sync error: " + ex.Message; }
        finally { _cts.Dispose(); _cts = null; IsBusy = false; }
    }

    /// <summary>Run the Mail workload for every mapped user, one per row. Resumable (dedup).</summary>
    private async Task<List<WorkloadResult>> RunBulkMailAsync(
        MigrationConfig config, GraphClient source, GraphClient target, string mode, CancellationToken ct)
    {
        var mailWl = new MailWorkload(config);
        var matched = _userMatches!.Where(m => m.ContentReady).ToList();
        var results = new List<WorkloadResult>();
        var i = 0;
        foreach (var m in matched)
        {
            ct.ThrowIfCancellationRequested();
            i++;
            Status = $"{mode}: Mail {i}/{matched.Count} — {m.SourceUpn}";
            var srcRef = $"/users/{m.SourceUpn}";
            var tgtRef = $"/users/{m.TargetId ?? m.TargetUpn}";
            var r = new WorkloadResult { Name = m.SourceUpn };

            // Read the SOURCE mailbox first; only this path is "mailbox unavailable"
            // (hybrid/on-prem source). A failure copying to the TARGET is a real error,
            // handled separately below, so a missing target mailbox isn't masked as a skip.
            List<MailFolderInfo> folders;
            try
            {
                folders = await mailWl.DiscoverFoldersAsync(source, srcRef, ct);
            }
            catch (GraphException ex) when (MailWorkload.IsMailboxUnavailable(ex))
            {
                r.Status = "skipped";
                r.Reason = MailWorkload.MailboxErrorHint(ex);
                results.Add(r);
                continue;
            }
            catch (GraphException ex)
            {
                // Other source-read failure (e.g. 403 missing permission, transient 5xx) —
                // record this user and continue the batch rather than aborting the whole run.
                r.Status = "error";
                r.Reason = MailWorkload.MailboxErrorHint(ex);
                results.Add(r);
                continue;
            }

            try
            {
                var planned = mailWl.Plan(folders);
                var perFolder = await mailWl.MigrateAsync(source, target, srcRef, tgtRef, planned, dryRun: !Execute, ct);
                r.Detail["folders"] = planned.Count.ToString();
                if (Execute)
                {
                    r.Status = "ok";
                    var copied = SumDetail(perFolder, "copied");
                    var skipped = SumDetail(perFolder, "skipped");
                    r.Detail["copied"] = copied.ToString();
                    r.Detail["skipped"] = skipped.ToString();
                    r.Detail["synced%"] = (copied + skipped == 0 ? 100 : 100 * skipped / (copied + skipped)).ToString();
                    var errs = SumDetail(perFolder, "errors");
                    if (errs > 0) r.Detail["errors"] = errs.ToString();
                }
                else
                {
                    r.Status = "would-copy";
                    r.Detail["~items"] = planned.Sum(f => f.ItemCount).ToString();
                }
            }
            catch (GraphException ex)
            {
                // Target-side or other failure — surface it (hint distinguishes a missing
                // target mailbox / licensing from a transient error).
                r.Status = "error";
                r.Reason = MailWorkload.MailboxErrorHint(ex);
            }
            results.Add(r);
        }
        return results;
    }

    /// <summary>Run the OneDrive workload for every mapped user, one per row. Resumable.</summary>
    private async Task<List<WorkloadResult>> RunBulkOneDriveAsync(
        MigrationConfig config, GraphClient source, GraphClient target, string mode, CancellationToken ct)
    {
        var filesWl = new FilesWorkload(config);
        var matched = _userMatches!.Where(m => m.ContentReady).ToList();
        var results = new List<WorkloadResult>();
        var i = 0;
        foreach (var m in matched)
        {
            ct.ThrowIfCancellationRequested();
            i++;
            Status = $"{mode}: OneDrive {i}/{matched.Count} — {m.SourceUpn}";
            var (sRoot, tRoot) = filesWl.ResolveDriveRoots(user: m.SourceUpn, targetUserOverride: m.TargetId ?? m.TargetUpn);
            var r = new WorkloadResult { Name = m.SourceUpn };
            try
            {
                var items = await filesWl.DiscoverDriveItemsAsync(source, sRoot, ct);
                var planned = filesWl.Plan(items);
                // Delta: skip items already in the target (pre-sync / repeatable passes).
                var fileCount = planned.Count(x => !x.IsFolder);
                var unchanged = 0;
                try
                {
                    var (tf, tfo) = FilesWorkload.IndexTarget(await filesWl.DiscoverDriveItemsAsync(target, tRoot, ct));
                    unchanged = FilesWorkload.MarkUnchanged(planned, tf, tfo);
                }
                catch (GraphException ex) when (FilesWorkload.IsDriveNotProvisioned(ex)) { /* empty target → copy all */ }
                var per = await filesWl.MigrateDriveItemsAsync(source, target, planned, sRoot, tRoot, dryRun: !Execute, ct);
                r.Status = Execute ? "ok" : "would-copy";
                r.Detail["files"] = fileCount.ToString();
                r.Detail["unchanged"] = unchanged.ToString();
                r.Detail["synced%"] = (fileCount == 0 ? 100 : 100 * unchanged / fileCount).ToString();
                var errs = per.Count(x => x.Status == "error");
                if (errs > 0) r.Detail["errors"] = errs.ToString();
            }
            catch (GraphException ex) when (FilesWorkload.IsDriveNotProvisioned(ex))
            {
                r.Status = "skipped"; r.Reason = "no OneDrive provisioned";
            }
            catch (GraphException ex)
            {
                r.Status = "error"; r.Reason = FilesWorkload.DriveErrorHint(ex);
            }
            results.Add(r);
        }
        return results;
    }
}
