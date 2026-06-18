using System.Windows;
using System.Windows.Controls;
using M365Migrate.App.ViewModels;

namespace M365Migrate.App;

public partial class MainWindow : Window
{
    private readonly MainViewModel _vm = new();

    public MainWindow()
    {
        InitializeComponent();
        DataContext = _vm;
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
        _vm.Execute = ExecuteCheck.IsChecked == true;
    }

    private async void OnTestClick(object sender, RoutedEventArgs e)
    {
        ApplyInputs();
        await _vm.TestConnectionsAsync();
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
