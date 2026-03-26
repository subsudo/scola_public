using System.Diagnostics;
using System.Text.Json;
using System.Windows.Forms;

namespace ScolaUpdater;

internal static class Program
{
    private const int ReplaceRetryCount = 5;
    private const int ReplaceRetryDelayMs = 500;

    [STAThread]
    private static int Main(string[] args)
    {
        ApplicationConfiguration.Initialize();

        try
        {
            var launchArguments = ParseArguments(args);
            Log($"Updater gestartet. Ziel='{launchArguments.TargetExecutablePath}', Quelle='{launchArguments.DownloadedExecutablePath}', PID={launchArguments.SourceProcessId}, Version={launchArguments.VersionString}");

            WaitForSourceProcessExit(launchArguments.SourceProcessId);

            var appDataDirectory = ResolveAppDataDirectory();
            var updatesRoot = Path.Combine(appDataDirectory, "updates");
            var backupDirectory = Path.Combine(updatesRoot, "backup");
            Directory.CreateDirectory(backupDirectory);

            var backupPath = Path.Combine(
                backupDirectory,
                $"{Path.GetFileNameWithoutExtension(launchArguments.TargetExecutablePath)}-{MakeSafeFileSegment(launchArguments.VersionString)}-{DateTime.Now:yyyyMMddHHmmss}.bak.exe");

            ReplaceExecutableWithRetry(launchArguments.TargetExecutablePath, launchArguments.DownloadedExecutablePath, backupPath);
            WriteCleanupMarker(updatesRoot, launchArguments.VersionString);
            StartUpdatedApplication(launchArguments.TargetExecutablePath);

            Log("Updater erfolgreich abgeschlossen.");
            return 0;
        }
        catch (Exception ex)
        {
            Log($"Updater-Fehler: {ex}");
            MessageBox.Show(
                $"Das Update konnte nicht abgeschlossen werden.{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                "Scola Update fehlgeschlagen",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return 1;
        }
    }

    private static UpdaterLaunchArguments ParseArguments(string[] args)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < args.Length - 1; index += 2)
        {
            values[args[index]] = args[index + 1];
        }

        if (!values.TryGetValue("--target-exe", out var targetExecutablePath) ||
            !values.TryGetValue("--downloaded-exe", out var downloadedExecutablePath) ||
            !values.TryGetValue("--source-pid", out var sourcePidText) ||
            !values.TryGetValue("--version", out var versionString) ||
            !int.TryParse(sourcePidText, out var sourceProcessId))
        {
            throw new InvalidOperationException("Updater wurde mit unvollstaendigen Parametern gestartet.");
        }

        return new UpdaterLaunchArguments
        {
            TargetExecutablePath = targetExecutablePath,
            DownloadedExecutablePath = downloadedExecutablePath,
            SourceProcessId = sourceProcessId,
            VersionString = versionString
        };
    }

    private static void WaitForSourceProcessExit(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            if (!process.HasExited)
            {
                Log($"Warte auf Ende von Prozess {processId}.");
                process.WaitForExit(30000);
            }
        }
        catch (ArgumentException)
        {
            Log($"Quellprozess {processId} ist bereits beendet.");
        }

        Thread.Sleep(500);
    }

    private static void ReplaceExecutableWithRetry(string targetExecutablePath, string downloadedExecutablePath, string backupPath)
    {
        Exception? lastException = null;
        for (var attempt = 1; attempt <= ReplaceRetryCount; attempt++)
        {
            try
            {
                ReplaceExecutable(targetExecutablePath, downloadedExecutablePath, backupPath);
                return;
            }
            catch (Exception ex) when (attempt < ReplaceRetryCount)
            {
                lastException = ex;
                Log($"Ersetzungsversuch {attempt}/{ReplaceRetryCount} fehlgeschlagen: {ex.Message}");
                Thread.Sleep(ReplaceRetryDelayMs);
            }
            catch (Exception ex)
            {
                lastException = ex;
                break;
            }
        }

        throw new IOException("Die bestehende EXE konnte nach mehreren Versuchen nicht ersetzt werden.", lastException);
    }

    private static void ReplaceExecutable(string targetExecutablePath, string downloadedExecutablePath, string backupPath)
    {
        if (!File.Exists(downloadedExecutablePath))
        {
            throw new FileNotFoundException("Die heruntergeladene Update-Datei wurde nicht gefunden.", downloadedExecutablePath);
        }

        var targetDirectory = Path.GetDirectoryName(targetExecutablePath);
        if (string.IsNullOrWhiteSpace(targetDirectory))
        {
            throw new InvalidOperationException("Zielordner der EXE konnte nicht bestimmt werden.");
        }

        Directory.CreateDirectory(targetDirectory);

        if (File.Exists(targetExecutablePath))
        {
            File.Copy(targetExecutablePath, backupPath, overwrite: true);
            File.Delete(targetExecutablePath);
        }

        try
        {
            File.Copy(downloadedExecutablePath, targetExecutablePath, overwrite: true);
        }
        catch
        {
            RestoreBackupIfPossible(targetExecutablePath, backupPath);
            throw;
        }
    }

    private static void StartUpdatedApplication(string targetExecutablePath)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = targetExecutablePath,
            UseShellExecute = true,
            WorkingDirectory = Path.GetDirectoryName(targetExecutablePath) ?? Environment.CurrentDirectory
        };

        Process.Start(startInfo);
    }

    private static void WriteCleanupMarker(string updatesRoot, string versionString)
    {
        Directory.CreateDirectory(updatesRoot);
        var cleanupMarker = new UpdateCleanupMarker
        {
            TargetVersion = versionString,
            CreatedUtc = DateTimeOffset.UtcNow
        };

        var cleanupMarkerPath = Path.Combine(updatesRoot, "update-cleanup.json");
        File.WriteAllText(cleanupMarkerPath, JsonSerializer.Serialize(cleanupMarker, new JsonSerializerOptions
        {
            WriteIndented = true
        }));
    }

    private static void RestoreBackupIfPossible(string targetExecutablePath, string backupPath)
    {
        try
        {
            if (File.Exists(backupPath) && !File.Exists(targetExecutablePath))
            {
                File.Copy(backupPath, targetExecutablePath, overwrite: true);
            }
        }
        catch (Exception ex)
        {
            Log($"Backup-Wiederherstellung fehlgeschlagen: {ex.Message}");
        }
    }

    private static string ResolveAppDataDirectory()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AkteX");
    }

    private static string MakeSafeFileSegment(string value)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        return new string(value.Select(character => invalidCharacters.Contains(character) ? '_' : character).ToArray());
    }

    private static void Log(string message)
    {
        try
        {
            var logDirectoryPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AkteX",
                "logs");
            Directory.CreateDirectory(logDirectoryPath);
            var logPath = Path.Combine(logDirectoryPath, $"updater-{DateTime.Now:yyyy-MM-dd}.log");
            File.AppendAllText(logPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}\t{message}{Environment.NewLine}");
        }
        catch
        {
        }
    }

    private sealed class UpdaterLaunchArguments
    {
        public required string TargetExecutablePath { get; init; }
        public required string DownloadedExecutablePath { get; init; }
        public required int SourceProcessId { get; init; }
        public required string VersionString { get; init; }
    }

    private sealed class UpdateCleanupMarker
    {
        public string? TargetVersion { get; set; }
        public DateTimeOffset CreatedUtc { get; set; }
    }
}
