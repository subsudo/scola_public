using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace VerlaufsakteApp.Services;

public static class JsonStorage
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    private const int SaveRetryCount = 4;
    private const int SaveRetryDelayMs = 60;

    public static T Load<T>(string path, string? backupPath, Func<T> createDefault)
    {
        try
        {
            if (!File.Exists(path))
            {
                return createDefault();
            }

            return DeserializeFromFile<T>(path) ?? createDefault();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"JSON-Laden fehlgeschlagen '{path}': {ex.Message}");
            if (!string.IsNullOrWhiteSpace(backupPath) && File.Exists(backupPath))
            {
                try
                {
                    AppLogger.Warn($"Fallback auf Backup-Datei '{backupPath}'.");
                    return DeserializeFromFile<T>(backupPath) ?? createDefault();
                }
                catch (Exception backupEx)
                {
                    AppLogger.Warn($"Backup-Laden fehlgeschlagen '{backupPath}': {backupEx.Message}");
                }
            }

            return createDefault();
        }
    }

    public static void SaveAtomic<T>(string path, string? backupPath, T value)
    {
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (!string.IsNullOrWhiteSpace(backupPath))
        {
            var backupDirectory = Path.GetDirectoryName(backupPath);
            if (!string.IsNullOrWhiteSpace(backupDirectory))
            {
                Directory.CreateDirectory(backupDirectory);
            }
        }

        var json = JsonSerializer.Serialize(value, JsonOptions);
        string? tempPath = null;

        try
        {
            ExecuteWithRetries(() => tempPath = WriteTempFile(path, json));
            ExecuteWithRetries(() => ReplaceTargetFile(tempPath!, path, backupPath));
        }
        catch
        {
            if (!string.IsNullOrWhiteSpace(tempPath))
            {
                TryDeleteTempFile(tempPath);
            }

            throw;
        }
    }

    private static void ExecuteWithRetries(Action action)
    {
        Exception? lastException = null;

        for (var attempt = 0; attempt < SaveRetryCount; attempt++)
        {
            try
            {
                action();
                return;
            }
            catch (Exception ex) when (IsRetryableIo(ex) && attempt < SaveRetryCount - 1)
            {
                lastException = ex;
                Thread.Sleep(SaveRetryDelayMs * (attempt + 1));
            }
            catch (Exception ex)
            {
                lastException = ex;
                break;
            }
        }

        throw lastException ?? new IOException("Unbekannter Fehler beim Speichern.");
    }

    private static string WriteTempFile(string targetPath, string json)
    {
        foreach (var tempPath in GetTempPathCandidates(targetPath))
        {
            try
            {
                var tempDirectory = Path.GetDirectoryName(tempPath);
                if (!string.IsNullOrWhiteSpace(tempDirectory))
                {
                    Directory.CreateDirectory(tempDirectory);
                }

                File.WriteAllText(tempPath, json, new UTF8Encoding(false));
                return tempPath;
            }
            catch (Exception ex) when (IsRetryableIo(ex))
            {
                AppLogger.Warn($"Temp-Datei konnte nicht geschrieben werden '{tempPath}': {ex.Message}");
            }
        }

        throw new UnauthorizedAccessException($"Temp-Datei fuer '{targetPath}' konnte nicht geschrieben werden.");
    }

    private static IEnumerable<string> GetTempPathCandidates(string targetPath)
    {
        var targetDirectory = Path.GetDirectoryName(targetPath) ?? Path.GetTempPath();
        var targetFileName = Path.GetFileName(targetPath);
        var guid = Guid.NewGuid().ToString("N");

        yield return Path.Combine(targetDirectory, $"{targetFileName}.{guid}.tmp");
        yield return Path.Combine(Path.GetTempPath(), $"AkteX.{targetFileName}.{guid}.tmp");
    }

    private static void ReplaceTargetFile(string tempPath, string path, string? backupPath)
    {
        if (File.Exists(path))
        {
            File.Replace(tempPath, path, backupPath, ignoreMetadataErrors: true);
            return;
        }

        File.Move(tempPath, path);
    }

    private static bool IsRetryableIo(Exception ex)
    {
        return ex is IOException or UnauthorizedAccessException;
    }

    private static T? DeserializeFromFile<T>(string path)
    {
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<T>(json, JsonOptions);
    }

    private static void TryDeleteTempFile(string tempPath)
    {
        try
        {
            if (File.Exists(tempPath))
            {
                File.Delete(tempPath);
            }
        }
        catch
        {
        }
    }
}
