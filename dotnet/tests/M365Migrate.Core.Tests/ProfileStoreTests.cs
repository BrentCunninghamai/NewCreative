using M365Migrate.Core.Configuration;
using Xunit;

namespace M365Migrate.Core.Tests;

public class ProfileStoreTests
{
    // A reversible "protector" stand-in for DPAPI so the round-trip is testable off-Windows.
    private static string Wrap(string s) => "ENC(" + s + ")";
    private static string Unwrap(string s) => s.StartsWith("ENC(") && s.EndsWith(")") ? s[4..^1] : s;

    [Fact]
    public void SaveLoad_RoundTrips_AndProtectsSecretsAtRest()
    {
        var dir = Path.Combine(Path.GetTempPath(), "m365-profiles-" + Guid.NewGuid().ToString("N"));
        try
        {
            var profile = new MigrationProfile
            {
                Name = "Contoso source",
                SourceTenantId = "t-src",
                SourceClientId = "c-src",
                SourceClientSecret = "super-secret",
                SourceDomain = "contoso.onmicrosoft.com",
                TargetTenantId = "t-tgt",
                TargetClientSecret = "tgt-secret",
                AssignLicenses = true,
                DefaultUsageLocation = "ZA",
                NamePrefix = "contoso-",
            };

            ProfileStore.Save(dir, profile, Wrap);

            // Secret is encrypted on disk, not plaintext.
            var raw = File.ReadAllText(ProfileStore.FileFor(dir, profile.Name));
            Assert.DoesNotContain("super-secret", raw);
            Assert.Contains("ENC(super-secret)", raw);

            var loaded = ProfileStore.Load(dir, "Contoso source", Unwrap);
            Assert.Equal("super-secret", loaded.SourceClientSecret);
            Assert.Equal("tgt-secret", loaded.TargetClientSecret);
            Assert.True(loaded.AssignLicenses);
            Assert.Equal("ZA", loaded.DefaultUsageLocation);
            Assert.Equal("contoso-", loaded.NamePrefix);

            Assert.Contains("Contoso source", ProfileStore.List(dir));
        }
        finally
        {
            if (Directory.Exists(dir)) Directory.Delete(dir, true);
        }
    }

    [Fact]
    public void Save_RequiresName()
    {
        Assert.Throws<ArgumentException>(() => ProfileStore.Save(Path.GetTempPath(), new MigrationProfile(), Wrap));
    }

    [Fact]
    public void List_EmptyWhenNoDir() =>
        Assert.Empty(ProfileStore.List(Path.Combine(Path.GetTempPath(), "nope-" + Guid.NewGuid().ToString("N"))));
}
