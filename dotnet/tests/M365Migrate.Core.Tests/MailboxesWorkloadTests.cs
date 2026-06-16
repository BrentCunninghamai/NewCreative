using System.Net;
using System.Text.Json;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Models;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class MailboxesWorkloadTests
{
    private static JsonElement Value(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    [Fact]
    public void FromGraph_KeepsOnlyWritableSettings()
    {
        var data = Value("{\"timeZone\":\"UTC\",\"language\":{\"locale\":\"en-US\"},\"userPurpose\":\"user\"}");
        var mb = SourceMailbox.FromGraph("jane@contoso.onmicrosoft.com", data);

        Assert.True(mb.Settings.ContainsKey("timeZone"));
        Assert.True(mb.Settings.ContainsKey("language"));
        Assert.False(mb.Settings.ContainsKey("userPurpose")); // read-only, dropped
    }

    [Fact]
    public void Plan_SkipsMissingTargetAndEmptySettings()
    {
        var workload = new MailboxesWorkload(TestData.Config());
        var mailboxes = new[]
        {
            new SourceMailbox { UserPrincipalName = "jane@contoso.onmicrosoft.com", Settings = { ["timeZone"] = Value("\"UTC\"") } },
            new SourceMailbox { UserPrincipalName = "bob@contoso.onmicrosoft.com", Settings = { ["timeZone"] = Value("\"UTC\"") } },
            new SourceMailbox { UserPrincipalName = "empty@contoso.onmicrosoft.com" },
        };
        var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "jane@fabrikam.onmicrosoft.com",
            "empty@fabrikam.onmicrosoft.com",
        };

        var planned = workload.Plan(mailboxes, existing).ToDictionary(p => p.SourceUpn);

        Assert.Equal("settings", planned["jane@contoso.onmicrosoft.com"].Action);
        Assert.Equal("skip", planned["bob@contoso.onmicrosoft.com"].Action); // no target account
        Assert.Equal("skip", planned["empty@contoso.onmicrosoft.com"].Action); // no settings
    }

    [Fact]
    public async Task Migrate_Execute_PatchesSettings()
    {
        var patched = false;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            if (req.Method == HttpMethod.Get && req.RequestUri!.ToString().Contains("/users"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"tj\",\"userPrincipalName\":\"jane@fabrikam.onmicrosoft.com\"}]}");
            if (req.Method == HttpMethod.Patch)
            {
                patched = true;
                return new HttpResponseMessage(HttpStatusCode.NoContent);
            }
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));
        var workload = new MailboxesWorkload(TestData.Config());

        var planned = new[]
        {
            new PlannedMailbox
            {
                SourceUpn = "jane@contoso.onmicrosoft.com",
                TargetUpn = "jane@fabrikam.onmicrosoft.com",
                Action = "settings",
                Settings = { ["timeZone"] = Value("\"UTC\"") },
            },
        };
        var results = await workload.MigrateAsync(client, planned, dryRun: false);

        Assert.True(patched);
        Assert.Equal("updated", results[0].Status);
        Assert.Contains("timeZone", handler.Bodies.First(b => b.Contains("timeZone")));
    }
}
