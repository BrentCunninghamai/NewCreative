using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using M365Migrate.App.ViewModels;
using M365Migrate.Core.Configuration;

namespace M365Migrate.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _vm = new();

    private static readonly string ProfilesDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "m365-migrate", "profiles");

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _vm;
        RefreshProfiles();
    }

    // --- Saved profiles (secrets encrypted at rest with Windows DPAPI, current user) ---

    private static string Protect(string s) =>
        Convert.ToBase64String(ProtectedData.Protect(Encoding.UTF8.GetBytes(s), null, DataProtectionScope.CurrentUser));

    private static string Unprotect(string s) =>
        Encoding.UTF8.GetString(ProtectedData.Unprotect(Convert.FromBase64String(s), null, DataProtectionScope.CurrentUser));

    private void RefreshProfiles()
    {
        var selected = ProfileCombo.Text;
        ProfileCombo.ItemsSource = ProfileStore.List(ProfilesDir);
        ProfileCombo.Text = selected;
    }

    private void OnSaveProfileClick(object sender, RoutedEventArgs e)
    {
        var name = ProfileCombo.Text?.Trim();
        if (string.IsNullOrEmpty(name))
        {
            _vm.Status = "Enter a profile name first.";
            return;
        }
        try
        {
            var profile = new MigrationProfile
            {
                Name = name,
                SourceTenantId = SourceTenantId.Text,
                SourceClientId = SourceClientId.Text,
                SourceClientSecret = SourceClientSecret.Password,
                SourceDomain = SourceDomain.Text,
                TargetTenantId = TargetTenantId.Text,
                TargetClientId = TargetClientId.Text,
                TargetClientSecret = TargetClientSecret.Password,
                TargetDomain = TargetDomain.Text,
                RewriteUpn = RewriteUpnCheck.IsChecked == true,
                SkipGuests = SkipGuestsCheck.IsChecked == true,
                AssignLicenses = AssignLicensesCheck.IsChecked == true,
                DefaultUsageLocation = UsageLocationBox.Text,
                NamePrefix = NamePrefixBox.Text,
                NameSuffix = NameSuffixBox.Text,
                DisplayNameSuffix = DisplayNameSuffixBox.Text,
            };
            ProfileStore.Save(ProfilesDir, profile, Protect);
            RefreshProfiles();
            ProfileCombo.Text = name;
            _vm.Status = $"Saved profile “{name}”. Secrets are encrypted (DPAPI, current Windows user).";
        }
        catch (Exception ex)
        {
            _vm.Status = "Couldn't save profile: " + ex.Message;
        }
    }

    private void OnLoadProfileClick(object sender, RoutedEventArgs e)
    {
        var name = ProfileCombo.Text?.Trim();
        if (string.IsNullOrEmpty(name))
        {
            _vm.Status = "Pick a profile to load.";
            return;
        }
        try
        {
            var p = ProfileStore.Load(ProfilesDir, name, Unprotect);
            SourceTenantId.Text = p.SourceTenantId;
            SourceClientId.Text = p.SourceClientId;
            SourceClientSecret.Password = p.SourceClientSecret;
            SourceDomain.Text = p.SourceDomain;
            TargetTenantId.Text = p.TargetTenantId;
            TargetClientId.Text = p.TargetClientId;
            TargetClientSecret.Password = p.TargetClientSecret;
            TargetDomain.Text = p.TargetDomain;
            RewriteUpnCheck.IsChecked = p.RewriteUpn;
            SkipGuestsCheck.IsChecked = p.SkipGuests;
            AssignLicensesCheck.IsChecked = p.AssignLicenses;
            UsageLocationBox.Text = p.DefaultUsageLocation;
            NamePrefixBox.Text = p.NamePrefix;
            NameSuffixBox.Text = p.NameSuffix;
            DisplayNameSuffixBox.Text = p.DisplayNameSuffix;
            _vm.Status = $"Loaded profile “{name}”.";
        }
        catch (Exception ex)
        {
            _vm.Status = "Couldn't load profile: " + ex.Message;
        }
    }

    /// <summary>Copy the form fields (including PasswordBoxes) into the view model.</summary>
    private void ApplyInputs()
    {
        _vm.SourceTenantId = SourceTenantId.Text;
        _vm.SourceClientId = SourceClientId.Text;
        _vm.SourceClientSecret = SourceClientSecret.Password;
        _vm.SourceDomain = SourceDomain.Text;
        _vm.TargetTenantId = TargetTenantId.Text;
        _vm.TargetClientId = TargetClientId.Text;
        _vm.TargetClientSecret = TargetClientSecret.Password;
        _vm.TargetDomain = TargetDomain.Text;

        _vm.Workload = (WorkloadCombo.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "Users";
        _vm.Scope = ScopeBox.Text;
        _vm.ScopeTargetUpn = ScopeTargetBox.Text;
        _vm.NamePrefix = NamePrefixBox.Text;
        _vm.NameSuffix = NameSuffixBox.Text;
        _vm.DisplayNameSuffix = DisplayNameSuffixBox.Text;
        _vm.RewriteUpn = RewriteUpnCheck.IsChecked == true;
        _vm.SkipGuests = SkipGuestsCheck.IsChecked == true;
        _vm.AssignLicenses = AssignLicensesCheck.IsChecked == true;
        _vm.DefaultUsageLocation = UsageLocationBox.Text;
        _vm.Execute = ExecuteCheck.IsChecked == true;
    }

    private async void OnTestClick(object sender, RoutedEventArgs e)
    {
        ApplyInputs();
        await _vm.TestConnectionsAsync();
    }

    private async void OnPreflightClick(object sender, RoutedEventArgs e)
    {
        ApplyInputs();
        await _vm.PreflightAsync();
    }

    private void OnSetupClick(object sender, RoutedEventArgs e)
    {
        ApplyInputs();
        var file = _vm.WriteSetupGuide();
        if (file is not null)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = file,
                    UseShellExecute = true,
                });
            }
            catch
            {
                // opening the file is best-effort; the path is in the status line.
            }
        }
    }

    private async void OnPlanClick(object sender, RoutedEventArgs e)
    {
        ApplyInputs();
        await _vm.PlanAsync();
    }

    private async void OnMigrateClick(object sender, RoutedEventArgs e)
    {
        ApplyInputs();
        await _vm.MigrateAsync();
    }

    private async void OnAutoSyncClick(object sender, RoutedEventArgs e)
    {
        ApplyInputs();
        if (!int.TryParse(AutoSyncBox.Text, out var minutes) || minutes < 1)
            minutes = 60;
        await _vm.StartAutoSyncAsync(minutes);
    }

    private void OnCancelClick(object sender, RoutedEventArgs e) => _vm.Cancel();

    private void OnOpenFolderClick(object sender, RoutedEventArgs e)
    {
        try
        {
            System.IO.Directory.CreateDirectory(_vm.ReportsDirectory);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = _vm.ReportsDirectory,
                UseShellExecute = true,
            });
        }
        catch
        {
            // opening the folder is best-effort
        }
    }
}
