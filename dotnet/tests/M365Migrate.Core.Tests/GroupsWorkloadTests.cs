using System.Net;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Models;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class GroupsWorkloadTests
{
    [Fact]
    public void Plan_ClassifiesKinds()
    {
        var workload = new GroupsWorkload(TestData.Config());
        var groups = new[]
        {
            new SourceGroup { Id = "g1", MailNickname = "marketing", GroupTypes = { "Unified" }, MailEnabled = true },
            new SourceGroup { Id = "g2", MailNickname = "sec", SecurityEnabled = true },
            new SourceGroup { Id = "g3", MailNickname = "dl", MailEnabled = true }, // distribution
        };
        var existing = new Dictionary<string, string> { ["sec"] = "tsec" };

        var planned = workload.Plan(groups, existing).ToDictionary(p => p.SourceId);

        Assert.Equal("create", planned["g1"].Action);
        Assert.Equal("exists", planned["g2"].Action);
        Assert.Equal("skip", planned["g3"].Action);
        Assert.Contains("not provisionable", planned["g3"].Reason);
    }

    [Fact]
    public void DynamicGroup_Body_IncludesRule()
    {
        var p = new PlannedGroup
        {
            MailNickname = "sales",
            DisplayName = "Sales",
            Kind = "microsoft365",
            IsDynamic = true,
            MembershipRule = "user.department -eq \"Sales\"",
        };
        var body = p.ToGraphBody();

        var groupTypes = Assert.IsAssignableFrom<List<string>>(body["groupTypes"]);
        Assert.Contains("Unified", groupTypes);
        Assert.Contains("DynamicMembership", groupTypes);
        Assert.Equal("On", body["membershipRuleProcessingState"]);
    }

    [Fact]
    public async Task Sync_DynamicGroup_CreatesWithRuleAndSkipsMembers()
    {
        var memberRefPosted = false;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (req.Method == HttpMethod.Get && uri.Contains("/groups") && uri.Contains("mailNickname"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
            if (req.Method == HttpMethod.Get && uri.Contains("/users"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"tj\",\"userPrincipalName\":\"jane@fabrikam.onmicrosoft.com\"}]}");
            if (req.Method == HttpMethod.Post && uri.Contains("/$ref"))
            {
                memberRefPosted = true;
                return new HttpResponseMessage(HttpStatusCode.NoContent);
            }
            if (req.Method == HttpMethod.Post && uri.EndsWith("/groups"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"dyn\"}");
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));
        var workload = new GroupsWorkload(TestData.Config());

        var planned = new[]
        {
            new PlannedGroup
            {
                SourceId = "g1",
                MailNickname = "sales",
                DisplayName = "Sales",
                Kind = "microsoft365",
                Action = "create",
                IsDynamic = true,
                MembershipRule = "user.department -eq \"Sales\"",
                TargetMemberUpns = { "jane@fabrikam.onmicrosoft.com" },
            },
        };
        var results = await workload.SyncAsync(client, planned, dryRun: false);

        Assert.False(memberRefPosted); // membership is rule-driven
        Assert.Equal("dynamic-rule", results[0].Detail["members"]);
        Assert.Contains("DynamicMembership", handler.Bodies.First(b => b.Contains("displayName")));
    }
}
