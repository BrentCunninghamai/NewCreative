using M365Migrate.Core.Mapping;
using Xunit;

namespace M365Migrate.Core.Tests;

public class TargetNamingTests
{
    [Fact]
    public void TargetUpn_RewritesDomainOnly_WhenNoAffix()
    {
        var config = TestData.Config();
        Assert.Equal("jane@fabrikam.onmicrosoft.com",
            TargetNaming.TargetUpn("jane@contoso.onmicrosoft.com", config));
    }

    [Fact]
    public void TargetUpn_AppliesPrefixAndSuffixToLocalPart()
    {
        var config = TestData.Config();
        config.Options.NamePrefix = "contoso-";
        config.Options.NameSuffix = "-mig";

        Assert.Equal("contoso-jane-mig@fabrikam.onmicrosoft.com",
            TargetNaming.TargetUpn("jane@contoso.onmicrosoft.com", config));
    }

    [Fact]
    public void TargetMailNickname_AppliesAffix()
    {
        var config = TestData.Config();
        config.Options.NamePrefix = "contoso-";
        Assert.Equal("contoso-sales", TargetNaming.TargetMailNickname("sales", config));

        var bare = TestData.Config();
        Assert.Equal("sales", TargetNaming.TargetMailNickname("sales", bare));
    }

    [Fact]
    public void Prefix_DisambiguatesCollidingSources()
    {
        // Two sources both have "john"; per-source prefixes keep them distinct in target.
        var a = TestData.Config();
        a.Options.NamePrefix = "contoso-";
        var b = TestData.Config();
        b.Options.NamePrefix = "northwind-";

        var fromA = TargetNaming.TargetUpn("john@contoso.onmicrosoft.com", a);
        var fromB = TargetNaming.TargetUpn("john@contoso.onmicrosoft.com", b);

        Assert.NotEqual(fromA, fromB);
        Assert.Equal("contoso-john@fabrikam.onmicrosoft.com", fromA);
        Assert.Equal("northwind-john@fabrikam.onmicrosoft.com", fromB);
    }
}
