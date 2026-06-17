using System.Net;
using M365Migrate.Core.Auth;
using M365Migrate.Core.Graph;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class MultiTenantOrgInspectorTests
{
    private static GraphClient Client(Func<HttpRequestMessage, string, HttpResponseMessage> responder) =>
        new(new HttpClient(new FakeHttpMessageHandler(responder)), TokenProviders.Static("t"));

    [Fact]
    public async Task Inspect_ReturnsNoPermission_On403()
    {
        var client = Client((_, _) => FakeHttpMessageHandler.Json(HttpStatusCode.Forbidden, "{\"error\":{\"code\":\"Authorization_RequestDenied\"}}"));
        var info = await MultiTenantOrgInspector.InspectAsync(client);
        Assert.False(info.PermissionGranted);
        Assert.False(info.Configured);
    }

    [Fact]
    public async Task Inspect_ReturnsNotConfigured_On404()
    {
        var client = Client((_, _) => FakeHttpMessageHandler.Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"ResourceNotFound\"}}"));
        var info = await MultiTenantOrgInspector.InspectAsync(client);
        Assert.True(info.PermissionGranted);
        Assert.False(info.Configured);
    }

    [Fact]
    public async Task Inspect_TreatsInactiveStateAsNotConfigured()
    {
        var client = Client((req, _) =>
            FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"state\":\"inactive\"}"));
        var info = await MultiTenantOrgInspector.InspectAsync(client);
        Assert.True(info.PermissionGranted);
        Assert.False(info.Configured);
    }

    [Fact]
    public async Task Inspect_ReturnsTenants_WhenActive()
    {
        var client = Client((req, _) =>
        {
            var uri = req.RequestUri!.ToString();
            if (uri.Contains("/multiTenantOrganization/tenants"))
                return FakeHttpMessageHandler.Json(HttpStatusCode.OK,
                    "{\"value\":[" +
                    "{\"tenantId\":\"aaaa\",\"displayName\":\"Source Co\",\"state\":\"active\"}," +
                    "{\"tenantId\":\"bbbb\",\"displayName\":\"Target Co\",\"state\":\"active\"}]}");
            return FakeHttpMessageHandler.Json(HttpStatusCode.OK, "{\"state\":\"active\",\"displayName\":\"Contoso Group\"}");
        });

        var info = await MultiTenantOrgInspector.InspectAsync(client);

        Assert.True(info.Configured);
        Assert.Equal("Contoso Group", info.DisplayName);
        Assert.Equal(2, info.Tenants.Count);
        Assert.True(info.ContainsTenant("AAAA")); // case-insensitive
        Assert.False(info.ContainsTenant("cccc"));
    }

    [Fact]
    public void Summarize_FlagsLinkedTenants()
    {
        var info = new MultiTenantOrgInfo(true, true, "Contoso Group", "active",
            new[] { new MtoTenant("aaaa", "Source Co", "active"), new MtoTenant("bbbb", "Target Co", "active") });

        var linked = MultiTenantOrgInspector.Summarize(info, "aaaa");
        Assert.Contains("MTO", linked);
        Assert.Contains("already exist", linked);

        var notListed = MultiTenantOrgInspector.Summarize(info, "zzzz");
        Assert.Contains("isn't listed", notListed);
    }

    [Fact]
    public void Summarize_SilentWhenNoPermission()
    {
        Assert.Null(MultiTenantOrgInspector.Summarize(MultiTenantOrgInfo.NoPermission, "aaaa"));
    }
}
