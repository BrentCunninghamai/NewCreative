using System.Windows;
using System.Windows.Threading;
using M365Migrate.App.Logging;

namespace M365Migrate.App;

public partial class App : Application
{
    public App()
    {
        // Surface any unhandled error as a copyable dialog instead of a silent
        // crash, so problems can be reported and diagnosed.
        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnDomainUnhandledException;
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        AppLog.Write($"unhandled UI exception: {e.Exception}");
        MessageBox.Show(
            e.Exception.ToString(),
            "m365-migrate — unexpected error",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        e.Handled = true; // keep the app alive after a handled UI-thread error
    }

    private void OnDomainUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        var message = (e.ExceptionObject as Exception)?.ToString() ?? "Unknown fatal error.";
        AppLog.Write($"fatal exception: {message}");
        MessageBox.Show(
            message,
            "m365-migrate — fatal error",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
    }
}
