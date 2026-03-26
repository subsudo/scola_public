using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp.Services;

public sealed class AppUpdateService
{
    private const string Owner = "subsudo";
    private const string Repository = "scola_public";
    private const string ReleaseAssetName = "Scola.exe";
    private const string EmbeddedUpdaterResourceName = "VerlaufsakteApp.Updater.ScolaUpdater.exe";
    private static readonly TimeSpan UpdateCheckInterval = TimeSpan.FromHours(24);
    private static readonly TimeSpan SnoozeDuration = TimeSpan.FromDays(7);
    private static readonly HttpClient HttpClient = CreateHttpClient();

    private readonly string _statePath;
    private readonly string _stateBackupPath;
    private readonly string _updatesRootPath;
    private readonly string _pendingDirectoryPath;
    private readonly string _backupDirectoryPath;
    private readonly string _updaterDirectoryPath;
    private readonly string _cleanupMarkerPath;

    public AppUpdateService()
    {
        _statePath = Path.Combine(App.AppDataDirectoryPath, "update-state.json");
        _stateBackupPath = Path.Combine(App.AppDataDirectoryPath, "update-state.bak");
        _updatesRootPath = Path.Combine(App.AppDataDirectoryPath, "updates");
        _pendingDirectoryPath = Path.Combine(_updatesRootPath, "pending");
        _backupDirectoryPath = Path.Combine(_updatesRootPath, "backup");
        _updaterDirectoryPath = Path.Combine(App.AppDataDirectoryPath, "Updater");
        _cleanupMarkerPath = Path.Combine(_updatesRootPath, "update-cleanup.json");
        CurrentVersion = ResolveCurrentVersion();
        CurrentVersionDisplay = CurrentVersion.ToString(3);
    }

    public Version CurrentVersion { get; }
    public string CurrentVersionDisplay { get; }

    public async Task<GitHubReleaseInfo?> GetAvailableUpdateAsync(CancellationToken cancellationToken)
    {
        var state = LoadState();
        if (IsCachedCheckStillFresh(state))
        {
            var cachedRelease = TryBuildCachedRelease(state);
            if (cachedRelease is null)
            {
                AppLogger.Info("Updater: Kein Online-Check noetig, letzter Check ist noch frisch und es gibt kein gecachtes Update.");
                return null;
            }

            if (IsSnoozed(state, cachedRelease.VersionString))
            {
                AppLogger.Info($"Updater: Version {cachedRelease.VersionString} bleibt bis {state.SkipUntilUtc:yyyy-MM-dd HH:mm:ss} unterdrueckt.");
                return null;
            }

            state.LastOfferedVersion = cachedRelease.VersionString;
            SaveState(state);
            AppLogger.Info($"Updater: Nutze gecachte Release-Information fuer Version {cachedRelease.VersionString}.");
            return cachedRelease;
        }

        AppLogger.Info("Updater: Starte GitHub-Release-Check.");

        try
        {
            var latestRelease = await FetchLatestReleaseAsync(cancellationToken).ConfigureAwait(false);
            state.LastCheckedUtc = DateTimeOffset.UtcNow;
            ApplyCachedRelease(state, latestRelease);
            SaveState(state);

            if (latestRelease is null)
            {
                AppLogger.Info("Updater: Keine verwertbare Release von GitHub gefunden.");
                return null;
            }

            if (latestRelease.Version <= CurrentVersion)
            {
                AppLogger.Info($"Updater: Keine neuere Version verfuegbar. Lokal={CurrentVersionDisplay}, Remote={latestRelease.VersionString}.");
                return null;
            }

            if (IsSnoozed(state, latestRelease.VersionString))
            {
                AppLogger.Info($"Updater: Neuere Version {latestRelease.VersionString} ist vorhanden, aber aktuell auf Spaeter gesetzt bis {state.SkipUntilUtc:yyyy-MM-dd HH:mm:ss}.");
                return null;
            }

            state.LastOfferedVersion = latestRelease.VersionString;
            SaveState(state);
            AppLogger.Info($"Updater: Neue Version verfuegbar: {latestRelease.VersionString}.");
            return latestRelease;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Updater: GitHub-Check fehlgeschlagen: {ex.Message}");

            var cachedRelease = TryBuildCachedRelease(state);
            if (cachedRelease is not null && !IsSnoozed(state, cachedRelease.VersionString))
            {
                AppLogger.Info($"Updater: Nutze nach Fehler die letzte gecachte neuere Version {cachedRelease.VersionString}.");
                return cachedRelease;
            }

            return null;
        }
    }

    public void SnoozeRelease(GitHubReleaseInfo release)
    {
        var state = LoadState();
        state.SkippedVersion = release.VersionString;
        state.SkipUntilUtc = DateTimeOffset.UtcNow.Add(SnoozeDuration);
        SaveState(state);
        AppLogger.Info($"Updater: Version {release.VersionString} bis {state.SkipUntilUtc:yyyy-MM-dd HH:mm:ss} auf Spaeter gesetzt.");
    }

    public async Task<DownloadedUpdateInfo> DownloadUpdateAsync(
        GitHubReleaseInfo release,
        IProgress<UpdateDownloadProgress>? progress,
        CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(_pendingDirectoryPath);
        var safeVersion = MakeSafeFileSegment(release.VersionString);
        var destinationPath = Path.Combine(_pendingDirectoryPath, $"Scola-{safeVersion}.exe");
        var tempPath = $"{destinationPath}.tmp";

        TryDeleteFile(tempPath);

        AppLogger.Info($"Updater: Download startet fuer Version {release.VersionString} von '{release.DownloadUrl}'.");

        using var request = new HttpRequestMessage(HttpMethod.Get, release.DownloadUrl);
        request.Headers.UserAgent.ParseAdd("Scola-Updater/1.0");
        request.Headers.Accept.ParseAdd("application/octet-stream");

        using var response = await HttpClient.SendAsync(
            request,
            HttpCompletionOption.ResponseHeadersRead,
            cancellationToken).ConfigureAwait(false);

        response.EnsureSuccessStatusCode();

        var totalBytes = response.Content.Headers.ContentLength;

        await using (var responseStream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false))
        await using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            var buffer = new byte[81920];
            long totalRead = 0;
            int read;

            while ((read = await responseStream.ReadAsync(buffer, cancellationToken).ConfigureAwait(false)) > 0)
            {
                await fileStream.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
                totalRead += read;

                int? percentage = null;
                if (totalBytes is long knownTotal && knownTotal > 0)
                {
                    percentage = (int)Math.Round((double)totalRead / knownTotal * 100, MidpointRounding.AwayFromZero);
                }

                progress?.Report(new UpdateDownloadProgress
                {
                    BytesReceived = totalRead,
                    TotalBytes = totalBytes,
                    Percentage = percentage
                });
            }
        }

        TryDeleteFile(destinationPath);
        File.Move(tempPath, destinationPath);

        AppLogger.Info($"Updater: Download abgeschlossen. Datei='{destinationPath}'.");
        return new DownloadedUpdateInfo
        {
            Release = release,
            LocalFilePath = destinationPath
        };
    }

    public void LaunchUpdater(DownloadedUpdateInfo downloadedUpdate)
    {
        var updaterPath = EnsureUpdaterExecutable();
        var targetExecutablePath = ResolveCurrentExecutablePath();
        var args = new UpdaterLaunchArguments
        {
            TargetExecutablePath = targetExecutablePath,
            DownloadedExecutablePath = downloadedUpdate.LocalFilePath,
            SourceProcessId = Environment.ProcessId,
            VersionString = downloadedUpdate.Release.VersionString
        };

        var startInfo = new ProcessStartInfo
        {
            FileName = updaterPath,
            UseShellExecute = false,
            WorkingDirectory = Path.GetDirectoryName(updaterPath) ?? App.AppDataDirectoryPath
        };

        startInfo.ArgumentList.Add("--target-exe");
        startInfo.ArgumentList.Add(args.TargetExecutablePath);
        startInfo.ArgumentList.Add("--downloaded-exe");
        startInfo.ArgumentList.Add(args.DownloadedExecutablePath);
        startInfo.ArgumentList.Add("--source-pid");
        startInfo.ArgumentList.Add(args.SourceProcessId.ToString());
        startInfo.ArgumentList.Add("--version");
        startInfo.ArgumentList.Add(args.VersionString);

        AppLogger.Info($"Updater: Starte Updater-Prozess '{updaterPath}' fuer Version {args.VersionString}.");

        var process = Process.Start(startInfo);
        if (process is null)
        {
            throw new InvalidOperationException("Updater konnte nicht gestartet werden.");
        }
    }

    public void TryCleanupSuccessfulUpdateArtifactsOnStartup()
    {
        var marker = JsonStorage.Load(_cleanupMarkerPath, null, static () => new UpdateCleanupMarker());
        if (string.IsNullOrWhiteSpace(marker.TargetVersion))
        {
            return;
        }

        if (!TryParseVersion(marker.TargetVersion, out var targetVersion))
        {
            AppLogger.Warn($"Updater: Cleanup-Marker enthaelt ungueltige Version '{marker.TargetVersion}'.");
            TryDeleteFile(_cleanupMarkerPath);
            return;
        }

        if (CurrentVersion < targetVersion)
        {
            AppLogger.Info($"Updater: Cleanup wird noch nicht ausgefuehrt. Aktuell={CurrentVersionDisplay}, Marker={marker.TargetVersion}.");
            return;
        }

        AppLogger.Info($"Updater: Fuehre Startup-Cleanup fuer Version {marker.TargetVersion} aus.");
        TryDeleteDirectoryContents(_pendingDirectoryPath);
        TryDeleteDirectoryContents(_backupDirectoryPath);
        TryDeleteFile(_cleanupMarkerPath);
    }

    private static HttpClient CreateHttpClient()
    {
        var client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(20)
        };
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Scola-Updater/1.0");
        client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return client;
    }

    private UpdateState LoadState()
    {
        return JsonStorage.Load(_statePath, _stateBackupPath, static () => new UpdateState());
    }

    private void SaveState(UpdateState state)
    {
        JsonStorage.SaveAtomic(_statePath, _stateBackupPath, state);
    }

    private bool IsCachedCheckStillFresh(UpdateState state)
    {
        return state.LastCheckedUtc != default &&
               DateTimeOffset.UtcNow - state.LastCheckedUtc < UpdateCheckInterval;
    }

    private bool IsSnoozed(UpdateState state, string versionString)
    {
        return string.Equals(state.SkippedVersion, versionString, StringComparison.OrdinalIgnoreCase) &&
               state.SkipUntilUtc is DateTimeOffset skipUntil &&
               skipUntil > DateTimeOffset.UtcNow;
    }

    private GitHubReleaseInfo? TryBuildCachedRelease(UpdateState state)
    {
        if (string.IsNullOrWhiteSpace(state.CachedLatestVersion) ||
            string.IsNullOrWhiteSpace(state.CachedTagName) ||
            string.IsNullOrWhiteSpace(state.CachedDownloadUrl))
        {
            return null;
        }

        if (!TryParseVersion(state.CachedLatestVersion, out var version))
        {
            return null;
        }

        if (version <= CurrentVersion)
        {
            return null;
        }

        return new GitHubReleaseInfo
        {
            TagName = state.CachedTagName,
            VersionString = state.CachedLatestVersion,
            Version = version,
            ReleaseTitle = state.CachedReleaseTitle ?? state.CachedLatestVersion,
            ReleaseNotes = state.CachedReleaseNotes ?? string.Empty,
            DownloadUrl = state.CachedDownloadUrl,
            PublishedAtUtc = state.CachedPublishedAtUtc
        };
    }

    private void ApplyCachedRelease(UpdateState state, GitHubReleaseInfo? release)
    {
        state.CachedLatestVersion = release?.VersionString;
        state.CachedTagName = release?.TagName;
        state.CachedReleaseTitle = release?.ReleaseTitle;
        state.CachedReleaseNotes = release?.ReleaseNotes;
        state.CachedDownloadUrl = release?.DownloadUrl;
        state.CachedPublishedAtUtc = release?.PublishedAtUtc;
    }

    private async Task<GitHubReleaseInfo?> FetchLatestReleaseAsync(CancellationToken cancellationToken)
    {
        using var response = await HttpClient.GetAsync(
            $"https://api.github.com/repos/{Owner}/{Repository}/releases/latest",
            cancellationToken).ConfigureAwait(false);

        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken).ConfigureAwait(false);
        using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken).ConfigureAwait(false);
        var root = document.RootElement;

        if (root.TryGetProperty("draft", out var draftElement) && draftElement.GetBoolean())
        {
            return null;
        }

        if (root.TryGetProperty("prerelease", out var prereleaseElement) && prereleaseElement.GetBoolean())
        {
            return null;
        }

        var tagName = root.TryGetProperty("tag_name", out var tagNameElement)
            ? tagNameElement.GetString()
            : null;

        if (string.IsNullOrWhiteSpace(tagName) || !TryParseVersion(tagName, out var version))
        {
            AppLogger.Warn($"Updater: Release-Tag '{tagName ?? "<leer>"}' ist ungueltig.");
            return null;
        }

        string? downloadUrl = null;
        if (root.TryGetProperty("assets", out var assetsElement) && assetsElement.ValueKind == JsonValueKind.Array)
        {
            foreach (var asset in assetsElement.EnumerateArray())
            {
                var assetName = asset.TryGetProperty("name", out var assetNameElement)
                    ? assetNameElement.GetString()
                    : null;

                if (!string.Equals(assetName, ReleaseAssetName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                downloadUrl = asset.TryGetProperty("browser_download_url", out var urlElement)
                    ? urlElement.GetString()
                    : null;
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(downloadUrl))
        {
            AppLogger.Warn("Updater: Release enthaelt kein Asset 'Scola.exe'.");
            return null;
        }

        var title = root.TryGetProperty("name", out var nameElement)
            ? nameElement.GetString()
            : null;
        var body = root.TryGetProperty("body", out var bodyElement)
            ? bodyElement.GetString()
            : null;

        DateTimeOffset? publishedAtUtc = null;
        if (root.TryGetProperty("published_at", out var publishedElement) &&
            publishedElement.ValueKind == JsonValueKind.String &&
            DateTimeOffset.TryParse(publishedElement.GetString(), out var parsedPublished))
        {
            publishedAtUtc = parsedPublished;
        }

        return new GitHubReleaseInfo
        {
            TagName = tagName,
            VersionString = NormalizeVersionString(tagName),
            Version = version,
            ReleaseTitle = string.IsNullOrWhiteSpace(title) ? NormalizeVersionString(tagName) : title.Trim(),
            ReleaseNotes = BuildReleaseNotesExcerpt(body),
            DownloadUrl = downloadUrl,
            PublishedAtUtc = publishedAtUtc
        };
    }

    private string EnsureUpdaterExecutable()
    {
        Directory.CreateDirectory(_updaterDirectoryPath);
        var updaterPath = Path.Combine(_updaterDirectoryPath, "ScolaUpdater.exe");
        var tempPath = $"{updaterPath}.tmp";

        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(EmbeddedUpdaterResourceName);
        if (stream is null)
        {
            throw new InvalidOperationException("Eingebetteter Updater wurde nicht gefunden.");
        }

        using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            stream.CopyTo(fileStream);
        }

        File.Copy(tempPath, updaterPath, overwrite: true);
        TryDeleteFile(tempPath);
        return updaterPath;
    }

    private static Version ResolveCurrentVersion()
    {
        var version = Assembly.GetExecutingAssembly().GetName().Version;
        if (version is null)
        {
            return new Version(0, 0, 0);
        }

        return new Version(version.Major, Math.Max(0, version.Minor), Math.Max(0, version.Build));
    }

    private static string ResolveCurrentExecutablePath()
    {
        var path = Environment.ProcessPath;
        if (!string.IsNullOrWhiteSpace(path))
        {
            return path;
        }

        using var process = Process.GetCurrentProcess();
        var mainModulePath = process.MainModule?.FileName;
        if (string.IsNullOrWhiteSpace(mainModulePath))
        {
            throw new InvalidOperationException("Pfad der aktuellen EXE konnte nicht bestimmt werden.");
        }

        return mainModulePath;
    }

    private static bool TryParseVersion(string value, out Version version)
    {
        var normalized = NormalizeVersionString(value);
        if (Version.TryParse(normalized, out var parsedVersion))
        {
            version = parsedVersion;
            return true;
        }

        version = new Version(0, 0, 0);
        return false;
    }

    private static string NormalizeVersionString(string value)
    {
        var trimmed = (value ?? string.Empty).Trim();
        if (trimmed.StartsWith("v", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed[1..];
        }

        var dashIndex = trimmed.IndexOf('-');
        if (dashIndex >= 0)
        {
            trimmed = trimmed[..dashIndex];
        }

        var segments = trimmed
            .Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Take(3)
            .ToList();

        while (segments.Count < 3)
        {
            segments.Add("0");
        }

        return string.Join('.', segments);
    }

    private static string BuildReleaseNotesExcerpt(string? body)
    {
        var text = (body ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return "Es ist eine neue Version von Scola verfuegbar.";
        }

        const int maxLength = 420;
        if (text.Length <= maxLength)
        {
            return text;
        }

        return $"{text[..maxLength].Trim()}...";
    }

    private static string MakeSafeFileSegment(string value)
    {
        var invalidCharacters = Path.GetInvalidFileNameChars();
        return new string(value.Select(character => invalidCharacters.Contains(character) ? '_' : character).ToArray());
    }

    private static void TryDeleteDirectoryContents(string directoryPath)
    {
        if (!Directory.Exists(directoryPath))
        {
            return;
        }

        try
        {
            foreach (var filePath in Directory.GetFiles(directoryPath))
            {
                TryDeleteFile(filePath);
            }

            foreach (var childDirectoryPath in Directory.GetDirectories(directoryPath))
            {
                try
                {
                    Directory.Delete(childDirectoryPath, recursive: true);
                }
                catch (Exception ex)
                {
                    AppLogger.Warn($"Updater: Unterordner '{childDirectoryPath}' konnte nicht geloescht werden: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Updater: Inhalt von '{directoryPath}' konnte nicht bereinigt werden: {ex.Message}");
        }
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }
}
