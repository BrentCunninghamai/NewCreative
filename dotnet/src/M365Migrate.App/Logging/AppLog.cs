using System.IO;

namespace M365Migrate.App.Logging;

/// <summary>
/// Lightweight append-only file logger under %LOCALAPPDATA%\m365-migrate\logs.
/// Logging must never disrupt the app, so all failures are swallowed.
/// </summary>
public static class AppLog
{
    public static string Directory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "m365-migrate", "logs");

    public static void Write(string message)
    {
        try
        {
            System.IO.Directory.CreateDirectory(Directory);
            var file = Path.Combine(Directory, $"app-{DateTime.Now:yyyyMMdd}.log");
            File.AppendAllText(file, $"{DateTime.Now:o}  {message}{Environment.NewLine}");
        }
        catch
        {
            // never let logging break the app
        }
    }
}
