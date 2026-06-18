using System.Text.Json;
using M365Migrate.Core.Configuration;
using Xunit;

namespace M365Migrate.Core.Tests;

public class GraphSetupTests
{
    [Fact]
    public void AdminConsentUrl_IsWellFormed()
    {
        var url = GraphSetup.AdminConsentUrl("contoso.onmicrosoft.com", "11111111-2222-3333-4444-555555555555");
        Assert.Equal(
            "https://login.microsoftonline.com/contoso.onmicrosoft.com/adminconsent?client_id=11111111-2222-3333-4444-555555555555",
            url);
    }

    [Fact]
    public void AdminConsentUrl_FallsBackToCommonWhenTenantBlank()
    {
        Assert.Contains("/common/adminconsent", GraphSetup.AdminConsentUrl("", "abc"));
    }

    [Fact]
    public void Manifest_IsValidJson_AndCoversEveryPermissionAsRole()
    {
        var json = GraphSetup.ManifestJson();
        using var doc = JsonDocument.Parse(json);
        var block = doc.RootElement[0];

        Assert.Equal(GraphSetup.GraphResourceAppId, block.GetProperty("resourceAppId").GetString());
        var access = block.GetProperty("resourceAccess");
        Assert.Equal(GraphSetup.Permissions.Count, access.GetArrayLength());
        foreach (var entry in access.EnumerateArray())
            Assert.Equal("Role", entry.GetProperty("type").GetString());

        // Every declared permission id appears in the manifest.
        var ids = access.EnumerateArray().Select(e => e.GetProperty("id").GetString()).ToHashSet();
        foreach (var p in GraphSetup.Permissions)
            Assert.Contains(p.Id, ids);
    }

    [Fact]
    public void Permissions_HaveUniqueNamesAndGuidIds()
    {
        Assert.Equal(GraphSetup.Permissions.Count, GraphSetup.Permissions.Select(p => p.Name).Distinct().Count());
        foreach (var p in GraphSetup.Permissions)
            Assert.True(Guid.TryParse(p.Id, out _), $"{p.Name} has a non-GUID id: {p.Id}");
    }

    [Fact]
    public void RequiredFor_ScopesToWorkload()
    {
        var users = GraphSetup.RequiredFor("Users");
        Assert.Contains("User.ReadWrite.All", users);
        Assert.Contains("Organization.Read.All", users); // baseline always included
        Assert.DoesNotContain("Mail.ReadWrite", users);  // unrelated workload not required

        var mail = GraphSetup.RequiredFor("Bulk Mail (mapped users)");
        Assert.Contains("Mail.ReadWrite", mail);
        Assert.DoesNotContain("Files.ReadWrite.All", mail);

        // Unknown workload falls back to the full set.
        var all = GraphSetup.RequiredFor("???");
        foreach (var p in GraphSetup.Permissions)
            Assert.Contains(p.Name, all);
    }

    [Fact]
    public void SetupGuide_IncludesManifestAndBothConsentLinks()
    {
        var guide = GraphSetup.SetupGuide("srcTenant", "srcClient", "tgtTenant", "tgtClient");
        Assert.Contains("requiredResourceAccess".Length > 0 ? "resourceAppId" : "", guide);
        Assert.Contains("login.microsoftonline.com/srcTenant/adminconsent?client_id=srcClient", guide);
        Assert.Contains("login.microsoftonline.com/tgtTenant/adminconsent?client_id=tgtClient", guide);
        Assert.Contains("Mail.ReadWrite", guide);
    }
}
