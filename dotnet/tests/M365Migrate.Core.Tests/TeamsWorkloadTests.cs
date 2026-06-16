using System.Net;
using System.Text.Json;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Models;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class TeamsWorkloadTests
{
    private static string GroupJson(string id, string nick, bool team) =>
        $"{{\"id\":\"{id}\",\"displayName\":\"{nick}\",\"mailNickname\":\"{nick}\"," +
        $"\"resourceProvisioningOptions\":[{(team ? "\"Team\"" : "")}]}}";

    private static JsonElement Parse(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    [Fact]
    public void GroupIsTeam_DetectsTeamOption()
    {
        Assert.True(TeamsWorkload.GroupIsTeam(Parse("{\"resourceProvisioningOptions\":[\"Team\"]}")));
        Assert.True(TeamsWorkload.GroupIsTeam(Parse("{\"resourceProvisioningOptions\":[\"team\"]}")));
        Assert.False(TeamsWorkload.GroupIsTeam(Parse("{\"resourceProvisioningOptions\":[]}")));
        Assert.False(TeamsWorkload.GroupIsTeam(Parse("{}")));
    }

    [Fact]
    public void Channel_DefaultDetection()
    {
        Assert.True(new Channel { DisplayName = "General" }.IsDefault);
        Assert.True(new Channel { DisplayName = " general " }.IsDefault);
        Assert.False(new Channel { DisplayName = "Engineering" }.IsDefault);
    }

    [Fact]
    public async Task Discover_OnlyReturnsTeams()
    {
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (uri.Contains("/groups"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[" + GroupJson("g1", "sales", true) + "," + GroupJson("g2", "plain", false) + "]}");
            if (uri.Contains("/teams/g1/channels"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"c0\",\"displayName\":\"General\"},{\"id\":\"c1\",\"displayName\":\"Deals\"}]}");
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[]}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var teams = await new TeamsWorkload(TestData.Config()).DiscoverAsync(client);
        Assert.Single(teams);
        Assert.Equal("sales", teams[0].MailNickname);
        Assert.Equal(new[] { "General", "Deals" }, teams[0].Channels.Select(c => c.DisplayName));
    }

    [Fact]
    public void Plan_ClassifiesTeamAndChannels()
    {
        var team = new SourceTeam
        {
            Id = "g1", MailNickname = "sales", DisplayName = "Sales",
            Channels =
            {
                new Channel { DisplayName = "General" },
                new Channel { DisplayName = "Deals" },
                new Channel { DisplayName = "Execs", MembershipType = "private" },
            },
        };
        var planned = new TeamsWorkload(TestData.Config())
            .Plan(new[] { team }, new Dictionary<string, string> { ["sales"] = "tg1" });

        Assert.Equal("provision", planned[0].Action);
        var byName = planned[0].Channels.ToDictionary(c => c.DisplayName);
        Assert.Equal("skip", byName["General"].Action);
        Assert.Equal("create", byName["Deals"].Action);
        Assert.Equal("skip", byName["Execs"].Action);
    }

    [Fact]
    public void Plan_SkipsWhenNoTargetGroup()
    {
        var team = new SourceTeam { Id = "g1", MailNickname = "sales" };
        var planned = new TeamsWorkload(TestData.Config()).Plan(new[] { team });
        Assert.Equal("skip", planned[0].Action);
        Assert.Contains("groups sync", planned[0].Reason);
    }

    [Fact]
    public async Task Migrate_Execute_EnablesTeamAndCreatesChannel()
    {
        var enabled = false;
        var channelCreated = false;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (req.Method == HttpMethod.Get && uri.Contains("/groups") && uri.Contains("resourceProvisioningOptions"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[" + GroupJson("tg1", "sales", false) + "]}");
            if (req.Method == HttpMethod.Get && uri.Contains("/groups"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[{\"id\":\"tg1\",\"mailNickname\":\"sales\"}]}");
            if (req.Method == HttpMethod.Put && uri.Contains("/groups/tg1/team"))
            {
                enabled = true;
                return FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"tg1\"}");
            }
            if (req.Method == HttpMethod.Get && uri.Contains("/teams/tg1/channels"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[{\"id\":\"c0\",\"displayName\":\"General\"}]}");
            if (req.Method == HttpMethod.Post && uri.Contains("/teams/tg1/channels"))
            {
                channelCreated = true;
                return FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{\"id\":\"c1\"}");
            }
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var planned = new[]
        {
            new PlannedTeam
            {
                SourceId = "g1", MailNickname = "sales", DisplayName = "Sales", Action = "provision",
                Channels = { new PlannedChannel { DisplayName = "Deals", Action = "create" } },
            },
        };
        var results = await new TeamsWorkload(TestData.Config()).MigrateAsync(client, planned, dryRun: false);

        Assert.True(enabled && channelCreated);
        Assert.Equal("enabled", results[0].Detail["team"]);
        Assert.Equal("created:1", results[0].Detail["channels"]);
    }

    [Fact]
    public async Task Migrate_SkipsExistingTeamAndChannel()
    {
        var channelPosted = false;
        var handler = new FakeHttpMessageHandler((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (req.Method == HttpMethod.Get && uri.Contains("/groups") && uri.Contains("resourceProvisioningOptions"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[" + GroupJson("tg1", "sales", true) + "]}");
            if (req.Method == HttpMethod.Get && uri.Contains("/groups"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"value\":[{\"id\":\"tg1\",\"mailNickname\":\"sales\"}]}");
            if (req.Method == HttpMethod.Get && uri.Contains("/teams/tg1/channels"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[{\"id\":\"c0\",\"displayName\":\"General\"},{\"id\":\"c1\",\"displayName\":\"Deals\"}]}");
            if (req.Method == HttpMethod.Post && uri.Contains("/channels"))
            {
                channelPosted = true;
                return FakeHttpMessageHandler.Json(HttpStatusCode.Created, "{}");
            }
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{}");
        });
        var client = new GraphClient(new HttpClient(handler), TokenProviders.Static("t"));

        var planned = new[]
        {
            new PlannedTeam
            {
                SourceId = "g1", MailNickname = "sales", DisplayName = "Sales", Action = "provision",
                Channels = { new PlannedChannel { DisplayName = "Deals", Action = "create" } },
            },
        };
        var results = await new TeamsWorkload(TestData.Config()).MigrateAsync(client, planned, dryRun: false);

        Assert.Equal("exists", results[0].Detail["team"]);
        Assert.Equal("created:0", results[0].Detail["channels"]);
        Assert.False(channelPosted);
    }
}
