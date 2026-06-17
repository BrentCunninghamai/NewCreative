using System.Text.Json;
using M365Migrate.Core.Configuration;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Mapping;

namespace M365Migrate.Core.Workloads;

/// <summary>
/// Copies a user's calendar events and contacts source -> target via Graph,
/// item by item. Writable fields are copied (read-only ids/etags are dropped, so
/// the target owns the new items). Re-runs are best-effort idempotent: an event
/// matching an existing target event by (subject, start, end), or a contact by
/// (display name, primary email), is skipped.
///
/// Per-user (a source UPN). Requires Calendars.ReadWrite and Contacts.ReadWrite
/// application permissions with admin consent in both tenants.
/// </summary>
public sealed class CalendarContactsWorkload
{
    private readonly MigrationConfig _config;

    public CalendarContactsWorkload(MigrationConfig config) => _config = config;

    private static readonly string[] EventFields =
    {
        "subject", "body", "start", "end", "location", "locations", "attendees",
        "categories", "importance", "sensitivity", "showAs", "isAllDay",
        "isReminderOn", "reminderMinutesBeforeStart", "recurrence",
        "responseRequested", "allowNewTimeProposals", "hideAttendees",
    };

    private static readonly string[] ContactFields =
    {
        "givenName", "surname", "middleName", "displayName", "nickName",
        "emailAddresses", "businessPhones", "homePhones", "mobilePhone",
        "companyName", "department", "jobTitle", "officeLocation", "profession",
        "businessHomePage", "personalNotes", "spouseName", "title", "imAddresses",
        "businessAddress", "homeAddress", "otherAddress", "birthday", "fileAs",
    };

    /// <summary>
    /// (source, target) user refs. <paramref name="targetUpnOverride"/> overrides the
    /// domain-rewrite for users whose target identity isn't a clean rewrite (e.g.
    /// already partly migrated by Microsoft's cross-tenant orchestrator).
    /// </summary>
    public (string Source, string Target) ResolveUserRefs(string sourceUpn, string? targetUpnOverride = null)
    {
        if (string.IsNullOrWhiteSpace(sourceUpn))
            throw new ArgumentException("A calendar/contacts migration needs a source user UPN.");
        var target = string.IsNullOrWhiteSpace(targetUpnOverride)
            ? TargetNaming.TargetUpn(sourceUpn, _config)
            : targetUpnOverride.Trim();
        return ($"/users/{sourceUpn}", $"/users/{target}");
    }

    /// <summary>Count source events and contacts (for the plan).</summary>
    public async Task<(int Events, int Contacts)> CountAsync(GraphClient source, string sourceRef, CancellationToken ct = default)
    {
        var events = await source.GetAllAsync($"{sourceRef}/events?$select=id&$top=100", ct);
        var contacts = await source.GetAllAsync($"{sourceRef}/contacts?$select=id&$top=100", ct);
        return (events.Count, contacts.Count);
    }

    /// <summary>Copy calendar events and contacts to the target mailbox.</summary>
    public async Task<List<WorkloadResult>> MigrateAsync(
        GraphClient source,
        GraphClient target,
        string sourceRef,
        string targetRef,
        bool dryRun = true,
        CancellationToken ct = default)
    {
        return new List<WorkloadResult>
        {
            await CopyAsync(source, target, sourceRef, targetRef, "Calendar", "events",
                EventFields, "subject,start,end", EventKey, dryRun, ct),
            await CopyAsync(source, target, sourceRef, targetRef, "Contacts", "contacts",
                ContactFields, "displayName,emailAddresses", ContactKey, dryRun, ct),
        };
    }

    private static async Task<WorkloadResult> CopyAsync(
        GraphClient source,
        GraphClient target,
        string sourceRef,
        string targetRef,
        string label,
        string collection,
        string[] copyFields,
        string keySelect,
        Func<JsonElement, string?> keyFn,
        bool dryRun,
        CancellationToken ct)
    {
        var record = new WorkloadResult { Name = label };

        var sourceItems = await source.GetAllAsync(
            $"{sourceRef}/{collection}?$select={string.Join(",", copyFields)}&$top=100", ct);

        if (dryRun)
        {
            record.Status = "would-copy";
            record.Detail["items"] = $"~{sourceItems.Count}";
            return record;
        }

        var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in await target.GetAllAsync($"{targetRef}/{collection}?$select={keySelect}&$top=100", ct))
        {
            var k = keyFn(t);
            if (k is not null)
                existing.Add(k);
        }

        int copied = 0, skipped = 0, errors = 0;
        foreach (var item in sourceItems)
        {
            ct.ThrowIfCancellationRequested();
            var key = keyFn(item);
            if (key is not null && existing.Contains(key))
            {
                skipped++;
                continue;
            }
            var body = BuildBody(item, copyFields);
            if (body.Count == 0)
                continue;
            try
            {
                await target.PostJsonAsync($"{targetRef}/{collection}", body, ct);
                copied++;
                if (key is not null)
                    existing.Add(key);
            }
            catch (GraphException)
            {
                errors++;
            }
        }

        record.Status = "ok";
        record.Detail["copied"] = copied.ToString();
        if (skipped > 0)
            record.Detail["skipped"] = skipped.ToString();
        if (errors > 0)
            record.Detail["errors"] = errors.ToString();
        return record;
    }

    private static Dictionary<string, object> BuildBody(JsonElement element, string[] fields)
    {
        var body = new Dictionary<string, object>();
        foreach (var field in fields)
            if (element.TryGetProperty(field, out var value) && value.ValueKind != JsonValueKind.Null)
                body[field] = value;
        return body;
    }

    private static string EventKey(JsonElement e)
    {
        var subject = e.GetStringOrNull("subject") ?? "";
        var start = e.TryGetProperty("start", out var s) ? s.GetStringOrNull("dateTime") ?? "" : "";
        var end = e.TryGetProperty("end", out var en) ? en.GetStringOrNull("dateTime") ?? "" : "";
        return $"{subject}|{start}|{end}";
    }

    private static string ContactKey(JsonElement c)
    {
        var name = c.GetStringOrNull("displayName") ?? "";
        var email = "";
        if (c.TryGetProperty("emailAddresses", out var arr) && arr.ValueKind == JsonValueKind.Array && arr.GetArrayLength() > 0)
            email = arr[0].GetStringOrNull("address") ?? "";
        return $"{name}|{email}";
    }
}
