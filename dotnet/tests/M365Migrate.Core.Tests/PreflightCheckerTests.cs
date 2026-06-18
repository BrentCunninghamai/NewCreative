using System.Text;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class PreflightCheckerTests
{
    // Build a fake JWT (header.payload.signature) whose payload carries the given roles.
    private static string TokenWithRoles(params string[] roles)
    {
        var payload = "{\"roles\":[" + string.Join(",", roles.Select(r => $"\"{r}\"")) + "]}";
        return "x." + Base64Url(payload) + ".y";
    }

    private static string Base64Url(string s) =>
        Convert.ToBase64String(Encoding.UTF8.GetBytes(s)).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    [Fact]
    public void RolesFromToken_ParsesRolesClaim()
    {
        var roles = PreflightChecker.RolesFromToken(TokenWithRoles("User.ReadWrite.All", "Mail.ReadWrite"));
        Assert.Contains("User.ReadWrite.All", roles);
        Assert.Contains("Mail.ReadWrite", roles);
        Assert.DoesNotContain("Sites.ReadWrite.All", roles);
    }

    [Fact]
    public void RolesFromToken_ToleratesGarbage()
    {
        Assert.Empty(PreflightChecker.RolesFromToken(""));
        Assert.Empty(PreflightChecker.RolesFromToken("not-a-jwt"));
        Assert.Empty(PreflightChecker.RolesFromToken("a.b.c")); // payload not valid base64 json
    }

    [Fact]
    public void Check_SplitsGrantedAndMissing()
    {
        var token = TokenWithRoles("User.ReadWrite.All", "Mail.ReadWrite");
        var result = PreflightChecker.Check(token, new[] { "User.ReadWrite.All", "Mail.ReadWrite", "Sites.ReadWrite.All" });

        Assert.False(result.AllGranted);
        Assert.Equal(2, result.Granted.Count);
        Assert.Equal(new[] { "Sites.ReadWrite.All" }, result.Missing);
    }

    [Fact]
    public void Check_AllGranted_WhenTokenHasEverything()
    {
        var token = TokenWithRoles("A", "B");
        var result = PreflightChecker.Check(token, new[] { "A", "B" });
        Assert.True(result.AllGranted);
        Assert.Empty(result.Missing);
    }
}
