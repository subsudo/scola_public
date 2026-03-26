using System.IO;
using System.Text;

namespace VerlaufsakteApp.Services;

public static class AppLogger
{
    private static readonly object Sync = new();
    private static bool _debugEnabled;

    public static string LogDirectoryPath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AkteX",
            "logs");

    public static string CurrentLogFilePath =>
        Path.Combine(LogDirectoryPath, $"app-{DateTime.Now:yyyy-MM-dd}.log");

    public static void Info(string message)
    {
        Write("INFO", message);
    }

    public static void Debug(string message)
    {
        if (!_debugEnabled)
        {
            return;
        }

        Write("DEBUG", message);
    }

    public static void Warn(string message)
    {
        Write("WARN", message);
    }

    public static void Error(string message, Exception? exception = null)
    {
        Write("ERROR", message, exception);
    }

    public static void SetDebugEnabled(bool enabled)
    {
        _debugEnabled = enabled;
        Write("INFO", $"Debug-Logging {(enabled ? "aktiviert" : "deaktiviert")}.");
    }

    private static void Write(string level, string message, Exception? exception = null)
    {
        try
        {
            lock (Sync)
            {
                Directory.CreateDirectory(LogDirectoryPath);
                using var stream = new FileStream(CurrentLogFilePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

                writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}\t{level}\t{message}");
                if (exception is not null)
                {
                    writer.WriteLine(exception.ToString());
                }
            }
        }
        catch
        {
            // Logging darf nie die App-Funktionalitaet blockieren.
        }
    }
}
