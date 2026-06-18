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
    public void Match_PrefersSameUpnNativeAccountOverGuestRep()
    {
        // The Paul case: target has a native account with the SAME UPN (preserved by a
        // hybrid/orchestrator move) AND a cross-tenant #EXT# guest rep. Prefer the native one.
        var sources = new[] { Src("paule@net1.com") };
        var targets = new[]
        {
            Tgt("paule@net1.com"),                                  // native account (real mailbox)
            Tgt("paule_net1.com#EXT#@lesaka.onmicrosoft.com"),     // guest rep
        };

        var m = _wl.Match(sources, targets).Single();
        Assert.Equal("upn", m.Method);
        Assert.Equal("t-paule@net1.com", m.TargetId);
        Assert.False(m.TargetIsGuest);
        Assert.True(m.ContentReady);
    }

    [Fact]
    public void Match_FlagsGuestOnlyTargetAsNotContentReady()
    {
        // Only a #EXT# guest rep exists in the target — matched (don't duplicate), but content
        // can't be migrated into a guest, so it must be flagged.
        var sources = new[] { Src("abdul@net1.com") };
        var targets = new[] { Tgt("abdul_net1.com#EXT#@lesaka.onmicrosoft.com") };

        var m = _wl.Match(sources, targets).Single();
        Assert.Equal("cross-tenant", m.Method);
        Assert.True(m.Matched);
        Assert.True(m.TargetIsGuest);
        Assert.False(m.ContentReady);
        Assert.Contains("guest", m.Note);
    }

    [Fact]
    public void Match_MatchesNativeByAliasWhenUpnAndPrimaryDiffer()
    {
        // Paul: source UPN/primary differ from the target's, but a shared alias links them to
        // the native target account. Should match (mail), native, content-ready — not unmatched.
        var src = new SourceUser
        {
            Id = "s-paul",
            UserPrincipalName = "PaulE@net1.com",
            DisplayName = "Paul",
            Mail = "paule@net1.com",
            ProxyAddresses = { "paule@net1.com", "paulenc@net1.com" },
        };
        var targets = new[]
        {
            new TargetUserRec("t-paul", "paul.encarnacao@lesakatechnologies.onmicrosoft.com",
                "paule@lesaka.tech", new[] { "paule@lesaka.tech", "paulenc@net1.com" }, "Member"),
        };

        var m = _wl.Match(new[] { src }, targets).Single();
        Assert.Equal("mail", m.Method);
        Assert.Equal("t-paul", m.TargetId);
        Assert.True(m.ContentReady);
    }

    [Fact]
    public void Match_PrefersNativeOverGuestWhenBothShareAnAlias()
    {
        var src = new SourceUser { Id = "s", UserPrincipalName = "x@net1.com", Mail = "x@net1.com", ProxyAddresses = { "x@net1.com" } };
        var targets = new[]
        {
            new TargetUserRec("t-guest", "x_net1.com#EXT#@lesaka.onmicrosoft.com", null, new[] { "x@net1.com" }, null),
            new TargetUserRec("t-native", "x@lesaka.tech", "x@net1.com", new[] { "x@net1.com" }, "Member"),
        };
        var m = _wl.Match(new[] { src }, targets).Single();
        Assert.Equal("t-native", m.TargetId);
        Assert.False(m.TargetIsGuest);
    }

    [Fact]
    public void Match_FlagsGuestByUserTypeEvenWithNormalUpn()
    {
        // A target Guest object with a normal-looking UPN (no #EXT#) is still not a content target.
        var sources = new[] { Src("vendor@contoso.onmicrosoft.com") };
        var targets = new[] { new TargetUserRec("t-g", "vendor@fabrikam.onmicrosoft.com", null, System.Array.Empty<string>(), "Guest") };

        // Matches by rewrite (contoso -> fabrikam), but the target is a Guest.
        var m = _wl.Match(sources, targets).Single();
        Assert.True(m.Matched);
        Assert.True(m.TargetIsGuest);
        Assert.False(m.ContentReady);
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
