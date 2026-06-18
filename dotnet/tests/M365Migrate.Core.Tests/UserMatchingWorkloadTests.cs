using M365Migrate.Core.Models;
using M365Migrate.Core.Workloads;
using Xunit;

namespace M365Migrate.Core.Tests;

public class UserMatchingWorkloadTests
{
    private static SourceUser Src(string upn, string? mail = null) =>
        new() { Id = "s-" + upn, UserPrincipalName = upn, DisplayName = upn, Mail = mail };

    private static TargetUserRec Tgt(string upn, string? mail = null, params string[] smtp) =>
        new("t-" + upn, upn, mail, smtp);

    private readonly UserMatchingWorkload _wl = new(TestData.Config()); // contoso -> fabrikam

    [Fact]
    public void Match_PrefersCsvOverride()
    {
        var sources = new[] { Src("jane@contoso.onmicrosoft.com") };
        var targets = new[]
        {
            Tgt("jane@fabrikam.onmicrosoft.com"),            // would match by rewrite
            Tgt("jane.doe@fabrikam.onmicrosoft.com"),        // the explicit choice
        };
        var csv = new Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
        {
            ["jane@contoso.onmicrosoft.com"] = "jane.doe@fabrikam.onmicrosoft.com",
        };

        var m = _wl.Match(sources, targets, csv).Single();
        Assert.Equal("csv", m.Method);
        Assert.Equal("jane.doe@fabrikam.onmicrosoft.com", m.TargetUpn);
        Assert.True(m.Matched);
    }

    [Fact]
    public void Match_CsvTargetMissing_StillHonoredButNoted()
    {
        var csv = new Dictionary<string, string> { ["jane@contoso.onmicrosoft.com"] = "ghost@fabrikam.onmicrosoft.com" };
        var m = _wl.Match(new[] { Src("jane@contoso.onmicrosoft.com") }, System.Array.Empty<TargetUserRec>(), csv).Single();
        Assert.Equal("csv", m.Method);
        Assert.Equal("ghost@fabrikam.onmicrosoft.com", m.TargetUpn);
        Assert.Contains("not found", m.Note);
    }

    [Fact]
    public void Match_DetectsCrossTenantExtIdentity()
    {
        // Source jane is already in the target as a B2B #EXT# member (decoded == her source UPN).
        var sources = new[] { Src("jane@contoso.onmicrosoft.com") };
        var targets = new[] { Tgt("jane_contoso.onmicrosoft.com#EXT#@fabrikam.onmicrosoft.com") };

        var m = _wl.Match(sources, targets).Single();
        Assert.Equal("cross-tenant", m.Method);
        Assert.Equal("t-jane_contoso.onmicrosoft.com#EXT#@fabrikam.onmicrosoft.com", m.TargetId);
    }

    [Fact]
    public void Match_FallsBackToMailThenRewrite()
    {
        // Mail match wins over rewrite: target UPN differs but a proxy SMTP equals source mail.
        var byMail = _wl.Match(
            new[] { Src("bob@contoso.onmicrosoft.com", mail: "bob@contoso.com") },
            new[] { Tgt("robert@fabrikam.onmicrosoft.com", mail: null, "bob@contoso.com") }).Single();
        Assert.Equal("mail", byMail.Method);
        Assert.Equal("robert@fabrikam.onmicrosoft.com", byMail.TargetUpn);

        // No mail/ext match -> rewrite to fabrikam and find it.
        var byRewrite = _wl.Match(
            new[] { Src("amy@contoso.onmicrosoft.com") },
            new[] { Tgt("amy@fabrikam.onmicrosoft.com") }).Single();
        Assert.Equal("rewrite", byRewrite.Method);
    }

    [Fact]
    public void Match_ReportsUnmatched()
    {
        var m = _wl.Match(new[] { Src("nomatch@contoso.onmicrosoft.com") }, System.Array.Empty<TargetUserRec>()).Single();
        Assert.False(m.Matched);
        Assert.Equal("none", m.Method);
        Assert.Null(m.TargetUpn);
    }

    [Fact]
    public void Csv_RoundTripsThroughParse()
    {
        var matches = new[]
        {
            new UserMatch("s1", "a@contoso.onmicrosoft.com", null, "t1", "a@fabrikam.onmicrosoft.com", "rewrite"),
            new UserMatch("s2", "b@contoso.onmicrosoft.com", null, null, null, "none"),
        };
        var csv = UserMatchingWorkload.ToCsv(matches);
        var parsed = UserMatchingWorkload.ParseOverrides(csv);

        // Only the matched row yields an override (blank target is skipped).
        Assert.Single(parsed);
        Assert.Equal("a@fabrikam.onmicrosoft.com", parsed["a@contoso.onmicrosoft.com"]);
    }

    [Fact]
    public void ParseOverrides_HandlesHeaderlessAndQuoted()
    {
        var csv = "src@x.com,\"tgt@y.com\"\n\"o'brien@x.com\",ob@y.com";
        var map = UserMatchingWorkload.ParseOverrides(csv);
        Assert.Equal("tgt@y.com", map["src@x.com"]);
        Assert.Equal("ob@y.com", map["o'brien@x.com"]);
    }
}
