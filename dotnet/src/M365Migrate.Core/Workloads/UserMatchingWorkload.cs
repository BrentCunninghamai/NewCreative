using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Mapping;
using M365Migrate.Core.Models;

namespace M365Migrate.Core.Workloads;

/// <summary>A target tenant user, indexed for matching against source users.</summary>
public sealed record TargetUserRec(string Id, string Upn, string? Mail, IReadOnlyList<string> SmtpAddresses, string? UserType = null)
{
    /// <summary>True if this target object is a guest — by userType or an #EXT# UPN. A guest can't
    /// receive migrated mailbox/OneDrive content.</summary>
    public bool IsGuest =>
        string.Equals(UserType, "Guest", StringComparison.OrdinalIgnoreCase)
        || Upn.Contains("#EXT#", StringComparison.OrdinalIgnoreCase);

    public static TargetUserRec FromGraph(System.Text.Json.JsonElement el)
    {
        var smtp = new List<string>();
        foreach (var p in el.GetArrayOrEmpty("proxyAddresses"))
        {
            var v = p.GetString();
            if (v is null) continue;
            // proxyAddresses look like "SMTP:primary@x" / "smtp:alias@x"; keep the address.
            var colon = v.IndexOf(':');
            smtp.Add(colon >= 0 ? v[(colon + 1)..] : v);
        }
        return new TargetUserRec(
            el.GetStringOrNull("id") ?? "",
            el.GetStringOrNull("userPrincipalName") ?? "",
            el.GetStringOrNull("mail"),
            smtp,
            el.GetStringOrNull("userType"));
    }
}

/// <summary>
/// One source user mapped to a target user (or not). <see cref="Method"/> records how the
/// match was made so the operator can trust/override it, ShareGate-style.
/// </summary>
public sealed record UserMatch(
    string SourceId,
    string SourceUpn,
    string? SourceMail,
    string? TargetId,
    string? TargetUpn,
    string Method,
    string? Note = null,
    bool TargetIsGuest = false)
{
    public bool Matched => TargetUpn is not null;

    /// <summary>True only if matched to a real (native, non-guest) target account that can
    /// receive migrated content. A match to a <c>#EXT#</c> guest rep is present but not usable.</summary>
    public bool ContentReady => Matched && !TargetIsGuest;
}

/// <summary>
/// Builds a source→target user mapping for the whole tenant — the backbone of a
/// ShareGate-style bulk migration. For each source user it picks the best target match
/// by precedence: explicit CSV override → same-UPN native account → any shared SMTP
/// address (primary or alias, native preferred) → cross-tenant/B2B (#EXT#) identity →
/// UPN domain rewrite. A match to a guest (#EXT# or userType Guest)
/// guest rep is flagged (TargetIsGuest) — present, but not a content destination. Unmatched
/// users are reported so they can be created or mapped by hand.
/// </summary>
public sealed class UserMatchingWorkload
{
    private readonly MigrationConfig _config;

    public UserMatchingWorkload(MigrationConfig config) => _config = config;

    /// <summary>Load target users with the fields needed for matching.</summary>
    public static async Task<List<TargetUserRec>> LoadTargetUsersAsync(GraphClient target, CancellationToken ct = default)
    {
        var raw = await target.GetAllAsync(
            "/users?$select=id,userPrincipalName,mail,proxyAddresses,userType&$top=999", ct);
        return raw.Select(TargetUserRec.FromGraph).ToList();
    }

    /// <summary>Pure matcher (no I/O) so the precedence logic is unit-testable.</summary>
    public List<UserMatch> Match(
        IEnumerable<SourceUser> sources,
        IReadOnlyCollection<TargetUserRec> targets,
        IReadOnlyDictionary<string, string>? csvOverrides = null)
    {
        var byUpn = new Dictionary<string, TargetUserRec>(StringComparer.OrdinalIgnoreCase);
        var byMail = new Dictionary<string, TargetUserRec>(StringComparer.OrdinalIgnoreCase);
        var byDecodedExt = new Dictionary<string, TargetUserRec>(StringComparer.OrdinalIgnoreCase);

        // When two targets share an address, prefer a native account over a guest rep.
        void AddMail(string key, TargetUserRec t)
        {
            if (!byMail.TryGetValue(key, out var existing) || (existing.IsGuest && !t.IsGuest))
                byMail[key] = t;
        }

        foreach (var t in targets)
        {
            if (!string.IsNullOrEmpty(t.Upn))
            {
                byUpn[t.Upn] = t;
                var decoded = UsersWorkload.DecodeExtUpn(t.Upn);
                if (decoded is not null)
                    byDecodedExt[decoded] = t;
            }
            if (t.Mail is not null)
                AddMail(t.Mail, t);
            foreach (var smtp in t.SmtpAddresses)
                AddMail(smtp, t);
        }

        var result = new List<UserMatch>();
        foreach (var s in sources)
        {
            TargetUserRec? hit;
            string method;
            string? note = null;

            if (csvOverrides is not null && csvOverrides.TryGetValue(s.UserPrincipalName, out var mappedUpn))
            {
                method = "csv";
                if (byUpn.TryGetValue(mappedUpn, out var rec))
                    hit = rec;
                else
                {
                    // Honor the operator's explicit target even if we couldn't pre-verify it.
                    result.Add(new UserMatch(s.Id, s.UserPrincipalName, s.Mail, null, mappedUpn, "csv", "mapped UPN not found in target (will be tried as-is)"));
                    continue;
                }
            }
            else if (byUpn.TryGetValue(s.UserPrincipalName, out var sameUpn))
            {
                // The target has a native account with the SAME UPN — e.g. a hybrid/orchestrator
                // account that preserved the UPN. This is the real mailbox, so prefer it over a
                // cross-tenant #EXT# guest rep.
                hit = sameUpn; method = "upn";
            }
            else if (TryMatchByAddress(s, byMail, out var byAddr))
            {
                // Any source address (primary mail OR an alias) equals a target address — catches
                // a native account reachable by a shared proxy even when UPN/primary differ.
                hit = byAddr!; method = "mail";
            }
            else if (byDecodedExt.TryGetValue(s.UserPrincipalName, out var ext))
            {
                hit = ext; method = "cross-tenant";
            }
            else if (byUpn.TryGetValue(TargetNaming.TargetUpn(s.UserPrincipalName, _config), out var rewritten))
            {
                hit = rewritten; method = "rewrite";
            }
            else
            {
                result.Add(new UserMatch(s.Id, s.UserPrincipalName, s.Mail, null, null, "none", "no target match — create the user or map manually"));
                continue;
            }

            // A guest target (by userType or #EXT# UPN) means the user exists only as a guest —
            // fine for "don't duplicate", but NOT a destination for mailbox/OneDrive content.
            var targetIsGuest = hit.IsGuest;
            if (targetIsGuest && note is null)
                note = "target is a guest — create/convert a native account before migrating content";

            result.Add(new UserMatch(s.Id, s.UserPrincipalName, s.Mail, hit.Id, hit.Upn, method, note, targetIsGuest));
        }
        return result;
    }

    /// <summary>Match a source user by any of its SMTP addresses (primary mail or an alias)
    /// against the target address index. Prefers a native (non-guest) target if both exist.</summary>
    private static bool TryMatchByAddress(
        SourceUser s, IReadOnlyDictionary<string, TargetUserRec> byMail, out TargetUserRec? hit)
    {
        hit = null;
        TargetUserRec? guestFallback = null;
        var addresses = new List<string>();
        if (s.Mail is not null) addresses.Add(s.Mail);
        addresses.AddRange(s.ProxyAddresses);
        foreach (var addr in addresses)
        {
            if (!byMail.TryGetValue(addr, out var rec)) continue;
            if (!rec.IsGuest) { hit = rec; return true; } // native wins immediately
            guestFallback ??= rec;
        }
        if (guestFallback is not null) { hit = guestFallback; return true; }
        return false;
    }

    /// <summary>Discover both tenants and build the full mapping.</summary>
    public async Task<List<UserMatch>> BuildAsync(
        GraphClient source,
        GraphClient target,
        IReadOnlyDictionary<string, string>? csvOverrides = null,
        CancellationToken ct = default)
    {
        var sources = await new UsersWorkload(_config).DiscoverAsync(source, ct);
        var targets = await LoadTargetUsersAsync(target, ct);
        return Match(sources, targets, csvOverrides);
    }

    /// <summary>Serialize a mapping to CSV.</summary>
    public static string ToCsv(IEnumerable<UserMatch> matches)
    {
        var headers = new[] { "source_upn", "target_upn", "method", "target_is_guest", "target_id", "note" };
        var rows = matches.Select(m => (IReadOnlyList<string>)new[]
        {
            m.SourceUpn, m.TargetUpn ?? "", m.Method, m.TargetIsGuest ? "yes" : "", m.TargetId ?? "", m.Note ?? "",
        });
        return Reporting.CsvReport.ToCsv(headers, rows);
    }

    /// <summary>
    /// Parse an override CSV (a header plus source_upn,target_upn rows) into a
    /// source→target map. Blank target cells are ignored. Extra columns are ignored.
    /// </summary>
    public static Dictionary<string, string> ParseOverrides(string csv)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(csv)) return map;
        var lines = csv.Replace("\r\n", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries);
        var start = lines.Length > 0 && lines[0].StartsWith("source_upn", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
        for (var i = start; i < lines.Length; i++)
        {
            var cells = SplitCsvLine(lines[i]);
            if (cells.Count < 2) continue;
            var src = cells[0].Trim();
            var tgt = cells[1].Trim();
            if (src.Length > 0 && tgt.Length > 0)
                map[src] = tgt;
        }
        return map;
    }

    private static List<string> SplitCsvLine(string line)
    {
        var cells = new List<string>();
        var sb = new System.Text.StringBuilder();
        var inQuotes = false;
        for (var i = 0; i < line.Length; i++)
        {
            var c = line[i];
            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < line.Length && line[i + 1] == '"') { sb.Append('"'); i++; }
                    else inQuotes = false;
                }
                else sb.Append(c);
            }
            else if (c == '"') inQuotes = true;
            else if (c == ',') { cells.Add(sb.ToString()); sb.Clear(); }
            else sb.Append(c);
        }
        cells.Add(sb.ToString());
        return cells;
    }
}
