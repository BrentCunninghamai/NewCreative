using System.Net;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Mapping;
using M365Migrate.Core.Models;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class UpnMapperTests
{
    [Fact]
    public void Rewrite_OnlyRewritesMatchingDomain()
    {
        Assert.Equal("jane@fabrikam.onmicrosoft.com",
            UpnMapper.Rewrite("jane@contoso.onmicrosoft.com", "contoso.onmicrosoft.com", "fabrikam.onmicrosoft.com"));
        // Vanity / other domains are left intact.
        Assert.Equal("jane@vanity.com",
            UpnMapper.Rewrite("jane@vanity.com", "contoso.onmicrosoft.com", "fabrikam.onmicrosoft.com"));
    }
}

public class UsersWorkloadTests
{
    private static SourceUser User(string upn, string type = "Member") =>
        new() { Id = upn, UserPrincipalName = upn, DisplayName = upn, UserType = type };

    [Fact]
    public void Plan_MarksCreateConflictAndSkipGuest()
    {
        var workload = new UsersWorkload(TestData.Config());
        var users = new[]
        {
            User("jane@contoso.onmicrosoft.com"),
            User("bob@contoso.onmicrosoft.com"),
            User("guest@contoso.onmicrosoft.com", type: "Guest"),
        };
        var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "bob@fabrikam.onmicrosoft.com" };

        var planned = workload.Plan(users, existing).ToDictionary(p => p.SourceUpn);

        Assert.Equal("create", planned["jane@contoso.onmicrosoft.com"].Action);
        Assert.Equal("jane@fabrikam.onmicrosoft.com", planned["jane@contoso.onmicrosoft.com"].TargetUpn);
        Assert.Equal("conflict", planned["bob@contoso.onmicrosoft.com"].Action);
        Assert.Equal("skip", planned["guest@contoso.onmicrosoft.com"].Action);
    }

    [Fact]
    public void Plan_SkipsExternalExtAccountsEvenWhenMemberType()
    {
        var workload = new UsersWorkload(TestData.Config());
        // A B2B account invited as Member but homed elsewhere (#EXT# UPN).
        var users = new[]
        {
            new SourceUser
            {
                Id = "x",
                UserPrincipalName = "vendor_othercorp.com#EXT#@contoso.onmicrosoft.com",
                DisplayName = "Vendor",
                UserType = "Member",
            },
        };
        var planned = workload.Plan(users);
        Assert.Equal("skip", planned[0].Action);
        Assert.Contains("#EXT#", planned[0].Reason);
    }

    [Fact]
    public async Task Migrate_DryRun_WritesNothing()
    {
        var posted = false;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            if (req.Method == HttpMethod.Post) posted = true;
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));
        var workload = new UsersWorkload(TestData.Config());

        var planned = new[] { new PlannedUser { TargetUpn = "jane@fabrikam.onmicrosoft.com", Action = "create" } };
        var results = await workload.MigrateAsync(client, planned, dryRun: true);

        Assert.False(posted);
        Assert.Equal("would-create", results[0].Status);
    }

    [Fact]
    public async Task Migrate_Execute_CreatesUser()
    {
        var handler = new FakeHttpMessageHandler((req, _) =>
            FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"new-id\"}"));
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));
        var workload = new UsersWorkload(TestData.Config());

        var planned = new[] { new PlannedUser { TargetUpn = "jane@fabrikam.onmicrosoft.com", Action = "create" } };
        var results = await workload.MigrateAsync(client, planned, dryRun: false);

        Assert.Equal("created", results[0].Status);
        Assert.Equal("new-id", results[0].Detail["target_id"]);
        // The create body carries a generated password and the target UPN.
        Assert.Contains("userPrincipalName", handler.Bodies[0]);
        Assert.Contains("passwordProfile", handler.Bodies[0]);
    }
}
