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
        _vm.RewriteUpn = RewriteUpnCheck.IsChecked == true;
        _vm.SkipGuests = SkipGuestsCheck.IsChecked == true;
        _vm.Execute = ExecuteCheck.IsChecked == true;
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
}
