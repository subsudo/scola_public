using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Globalization;
using System.IO;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp.Services;

public sealed class ParticipantHintsService
{
    public const string DefaultStorePath = @"K:\FuturX\34_Bildung\02_Grundlagen\20_Arbeitsinstrumente\300_AppData_Scola_Acta\ParticipantHints\participant-hints.json";
    private const int SchemaVersion = 1;
    private const string MissingFileHash = "<missing>";
    private const int UniversalNameInfoLevel = 1;
    private const int ErrorMoreData = 234;
    private const int NoError = 0;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly string _storePath;
    private readonly Dictionary<string, string?> _uncRootCache = new(StringComparer.OrdinalIgnoreCase);

    public ParticipantHintsService(string? configuredStorePath)
    {
        _storePath = string.IsNullOrWhiteSpace(configuredStorePath)
            ? DefaultStorePath
            : configuredStorePath.Trim();
    }

    public string StorePath => _storePath;

    public IReadOnlyList<ParticipantHintDisplay> LoadActiveDisplays(string documentPath)
    {
        var session = LoadEditorSession(documentPath);
        return !session.IsAvailable
            ? Array.Empty<ParticipantHintDisplay>()
            : BuildDisplayHints(session.Record.Hints);
    }

    public ParticipantHintEditSession LoadEditorSession(string documentPath)
    {
        if (!TryCanonicalizeDocumentPath(documentPath, out var canonicalPath, out var error))
        {
            return new ParticipantHintEditSession
            {
                IsAvailable = false,
                ErrorMessage = error,
                DocumentPath = documentPath
            };
        }

        try
        {
            var document = LoadDocument();
            var record = FindRecord(document, canonicalPath)
                         ?? CreateEmptyRecord(canonicalPath, documentPath);

            return new ParticipantHintEditSession
            {
                IsAvailable = true,
                DocumentPath = documentPath,
                ExpectedHash = ComputeStoreHash(),
                Record = CloneRecord(record)
            };
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Hinweise konnten nicht geladen werden. Path='{_storePath}', Error='{ex.Message}'");
            return new ParticipantHintEditSession
            {
                IsAvailable = false,
                ErrorMessage = $"Hinweisdatei konnte nicht gelesen werden: {ex.Message}",
                DocumentPath = documentPath
            };
        }
    }

    public ParticipantHintSaveResult SaveEditorSession(ParticipantHintEditSession session, IReadOnlyList<ParticipantHintEditorItem> items)
    {
        if (!session.IsAvailable)
        {
            return new ParticipantHintSaveResult { Success = false, ErrorMessage = session.ErrorMessage };
        }

        if (!TryCanonicalizeDocumentPath(session.DocumentPath, out var canonicalPath, out var error))
        {
            return new ParticipantHintSaveResult { Success = false, ErrorMessage = error };
        }

        try
        {
            var currentHash = ComputeStoreHash();
            if (!string.Equals(currentHash, session.ExpectedHash, StringComparison.Ordinal))
            {
                return new ParticipantHintSaveResult
                {
                    Success = false,
                    Conflict = true,
                    ErrorMessage = "Die Hinweise wurden inzwischen von einer anderen Stelle geändert. Bitte den Editor erneut öffnen."
                };
            }

            var document = LoadDocument();
            var record = FindRecord(document, canonicalPath);
            if (record is null)
            {
                record = CreateEmptyRecord(canonicalPath, session.DocumentPath);
                document.Participants.Add(record);
            }

            var updatedBy = Environment.UserName ?? string.Empty;
            record.Key = canonicalPath;
            record.CanonicalDocumentPath = canonicalPath;
            record.OriginalDocumentPath = session.DocumentPath;
            record.FolderName = ResolveFolderName(session.DocumentPath);
            record.UpdatedAtUtc = DateTime.UtcNow;
            record.Hints = items
                .Select(item => item.ToEntry(updatedBy))
                .Where(entry => !IsEmpty(entry))
                .ToList();

            document.SchemaVersion = SchemaVersion;
            document.UpdatedAtUtc = DateTime.UtcNow;
            JsonStorage.SaveAtomic(_storePath, BuildBackupPath(), document);

            return new ParticipantHintSaveResult { Success = true };
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Hinweise konnten nicht gespeichert werden. Path='{_storePath}', Error='{ex.Message}'");
            return new ParticipantHintSaveResult
            {
                Success = false,
                ErrorMessage = $"Hinweise konnten nicht gespeichert werden: {ex.Message}"
            };
        }
    }

    public IReadOnlyList<ParticipantHintDisplay> BuildDisplayHints(IReadOnlyList<ParticipantHintEntry> entries)
    {
        return entries
            .Where(entry => string.Equals(entry.Status, ParticipantHintStatuses.Active, StringComparison.OrdinalIgnoreCase))
            .Select(CreateDisplay)
            .Where(display => !string.IsNullOrWhiteSpace(display.Text))
            .OrderBy(display => display.SortDate ?? DateTime.MaxValue)
            .ThenBy(display => GetTypeOrder(display.Type))
            .ToList();
    }

    private ParticipantHintDocument LoadDocument()
    {
        if (!File.Exists(_storePath))
        {
            return new ParticipantHintDocument
            {
                SchemaVersion = SchemaVersion,
                UpdatedAtUtc = DateTime.UtcNow
            };
        }

        var json = File.ReadAllText(_storePath, new UTF8Encoding(false));
        var document = JsonSerializer.Deserialize<ParticipantHintDocument>(json, JsonOptions) ?? new ParticipantHintDocument();
        document.Participants ??= new List<ParticipantHintParticipantRecord>();
        return document;
    }

    private string ComputeStoreHash()
    {
        if (!File.Exists(_storePath))
        {
            return MissingFileHash;
        }

        using var stream = File.Open(_storePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var hash = SHA256.HashData(stream);
        return Convert.ToHexString(hash);
    }

    private string BuildBackupPath()
    {
        return Path.Combine(Path.GetDirectoryName(_storePath) ?? string.Empty, "participant-hints.bak");
    }

    private static ParticipantHintParticipantRecord? FindRecord(ParticipantHintDocument document, string canonicalPath)
    {
        return document.Participants.FirstOrDefault(record =>
            string.Equals(record.Key, canonicalPath, StringComparison.OrdinalIgnoreCase)
            || string.Equals(record.CanonicalDocumentPath, canonicalPath, StringComparison.OrdinalIgnoreCase));
    }

    private static ParticipantHintParticipantRecord CreateEmptyRecord(string canonicalPath, string originalPath)
    {
        return new ParticipantHintParticipantRecord
        {
            Key = canonicalPath,
            CanonicalDocumentPath = canonicalPath,
            OriginalDocumentPath = originalPath,
            FolderName = ResolveFolderName(originalPath),
            UpdatedAtUtc = DateTime.UtcNow
        };
    }

    private static ParticipantHintParticipantRecord CloneRecord(ParticipantHintParticipantRecord record)
    {
        return new ParticipantHintParticipantRecord
        {
            Key = record.Key,
            CanonicalDocumentPath = record.CanonicalDocumentPath,
            OriginalDocumentPath = record.OriginalDocumentPath,
            FolderName = record.FolderName,
            UpdatedAtUtc = record.UpdatedAtUtc,
            Hints = record.Hints.Select(CloneEntry).ToList()
        };
    }

    private static ParticipantHintEntry CloneEntry(ParticipantHintEntry entry)
    {
        return new ParticipantHintEntry
        {
            Id = entry.Id,
            Type = entry.Type,
            Status = entry.Status,
            Details = new ParticipantHintDetails
            {
                Date = entry.Details.Date,
                Month = entry.Details.Month,
                Subject = entry.Details.Subject,
                Note = entry.Details.Note,
                Text = entry.Details.Text
            },
            UpdatedAtUtc = entry.UpdatedAtUtc,
            UpdatedBy = entry.UpdatedBy
        };
    }

    private static bool IsEmpty(ParticipantHintEntry entry)
    {
        var details = entry.Details;
        return entry.Type switch
        {
            ParticipantHintTypes.Free => string.IsNullOrWhiteSpace(details.Text),
            ParticipantHintTypes.AmReport => string.IsNullOrWhiteSpace(details.Date) && string.IsNullOrWhiteSpace(details.Month),
            ParticipantHintTypes.Exit => string.IsNullOrWhiteSpace(details.Date),
            ParticipantHintTypes.StellwerkTest => string.IsNullOrWhiteSpace(details.Date),
            _ => false
        };
    }

    private static ParticipantHintDisplay CreateDisplay(ParticipantHintEntry entry)
    {
        var details = entry.Details;
        var code = ResolveDisplayCode(entry.Type);
        var value = entry.Type switch
        {
            ParticipantHintTypes.Exit => FormatDate(details.Date),
            ParticipantHintTypes.AmReport => FormatDateOrMonth(details.Date, details.Month),
            ParticipantHintTypes.StellwerkTest => FormatDate(details.Date),
            ParticipantHintTypes.Free => LimitNoteText(details.Text),
            _ => string.Empty
        };
        var text = BuildLabel(code, value, string.Empty);

        return new ParticipantHintDisplay
        {
            Type = entry.Type,
            Text = text,
            Code = entry.Type == ParticipantHintTypes.Free ? string.Empty : code,
            Value = value,
            MarkerColor = ResolveMarkerColor(entry.Type),
            PillBackground = ResolvePillBackground(entry.Type),
            PillForeground = ResolvePillForeground(entry.Type),
            IsOverdue = IsOverdue(entry),
            SortDate = GetEntrySortDate(entry)
        };
    }

    private static string BuildLabel(string primary, string secondary, string fallback)
    {
        var parts = new[] { primary, secondary, fallback }
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .Select(part => part.Trim())
            .ToArray();

        return string.Join(" ", parts);
    }

    private static string ResolveDisplayCode(string type)
    {
        return type switch
        {
            ParticipantHintTypes.Exit => "AT",
            ParticipantHintTypes.AmReport => "AM",
            ParticipantHintTypes.StellwerkTest => "STW",
            _ => "N"
        };
    }

    private static string ResolveMarkerColor(string type)
    {
        return type switch
        {
            ParticipantHintTypes.Exit => "#D1493F",
            ParticipantHintTypes.AmReport => "#2E76D0",
            ParticipantHintTypes.StellwerkTest => "#4D8D3E",
            _ => "#C7A126"
        };
    }

    private static string ResolvePillBackground(string type)
    {
        return type switch
        {
            ParticipantHintTypes.Exit => "#F3B3AD",
            ParticipantHintTypes.AmReport => "#B7D4F6",
            ParticipantHintTypes.StellwerkTest => "#BFE2B9",
            _ => "#F1D875"
        };
    }

    private static string ResolvePillForeground(string type)
    {
        return type switch
        {
            ParticipantHintTypes.Exit => "#6E1F1A",
            ParticipantHintTypes.AmReport => "#173F73",
            ParticipantHintTypes.StellwerkTest => "#24501F",
            _ => "#4D3B00"
        };
    }

    private static bool IsOverdue(ParticipantHintEntry entry)
    {
        if (entry.Type is not (ParticipantHintTypes.Exit or ParticipantHintTypes.AmReport or ParticipantHintTypes.StellwerkTest))
        {
            return false;
        }

        return TryParseDate(entry.Details.Date, out var date) && date.Date < DateTime.Today;
    }

    private static DateTime? GetEntrySortDate(ParticipantHintEntry entry)
    {
        if (entry.Type is ParticipantHintTypes.Exit or ParticipantHintTypes.AmReport or ParticipantHintTypes.StellwerkTest
            && TryParseDate(entry.Details.Date, out var date))
        {
            return date.Date;
        }

        if (entry.Type == ParticipantHintTypes.AmReport
            && DateTime.TryParseExact(entry.Details.Month, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var month))
        {
            return month.Date;
        }

        return null;
    }

    private static int GetTypeOrder(string type)
    {
        return type switch
        {
            ParticipantHintTypes.Exit => 0,
            ParticipantHintTypes.AmReport => 1,
            ParticipantHintTypes.StellwerkTest => 2,
            _ => 3
        };
    }

    private static string FormatDate(string value)
    {
        return TryParseDate(value, out var date)
            ? date.ToString("dd.MM.yy")
            : value.Trim();
    }

    private static string FormatMonth(string value)
    {
        if (DateTime.TryParseExact(value, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            return date.ToString("MM.yy");
        }

        return value.Trim();
    }

    private static string FormatDateOrMonth(string dateValue, string monthValue)
    {
        if (!string.IsNullOrWhiteSpace(dateValue))
        {
            return FormatDate(dateValue);
        }

        return FormatMonth(monthValue);
    }

    private static string LimitNoteText(string value)
    {
        var trimmed = value.Trim();
        return trimmed.Length <= 60 ? trimmed : trimmed[..60];
    }

    private static bool TryParseDate(string value, out DateTime date)
    {
        return DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date)
               || DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out date);
    }

    private bool TryCanonicalizeDocumentPath(string documentPath, out string canonicalPath, out string error)
    {
        canonicalPath = string.Empty;
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(documentPath))
        {
            error = "Für diesen Teilnehmer ist noch keine Akte zugeordnet.";
            return false;
        }

        var fullPath = Path.GetFullPath(documentPath.Trim());
        if (fullPath.StartsWith(@"\\", StringComparison.Ordinal))
        {
            canonicalPath = NormalizePath(fullPath);
            return true;
        }

        var root = Path.GetPathRoot(fullPath) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(root) || !root.Contains(':', StringComparison.Ordinal))
        {
            error = $"Aktenpfad kann nicht kanonisiert werden: {documentPath}";
            return false;
        }

        var uncRoot = ResolveUncRoot(root);
        if (string.IsNullOrWhiteSpace(uncRoot))
        {
            error = $"Laufwerk '{root}' konnte nicht nach UNC aufgelöst werden.";
            return false;
        }

        canonicalPath = NormalizePath(Path.Combine(uncRoot, fullPath[root.Length..]));
        return true;
    }

    private string? ResolveUncRoot(string root)
    {
        if (_uncRootCache.TryGetValue(root, out var cached))
        {
            return cached;
        }

        var driveName = root.TrimEnd('\\');
        var resolved = TryResolveMappedDriveRoot(driveName)
                       ?? TryResolveUniversalName(root.TrimEnd('\\') + "\\");
        _uncRootCache[root] = resolved;
        return resolved;
    }

    private static string? TryResolveMappedDriveRoot(string driveName)
    {
        if (string.IsNullOrWhiteSpace(driveName))
        {
            return null;
        }

        var buffer = new StringBuilder(512);
        var bufferSize = buffer.Capacity;
        var result = NativeMethods.WNetGetConnection(driveName, buffer, ref bufferSize);
        if (result == NoError)
        {
            return buffer.ToString().TrimEnd('\\');
        }

        if (result != ErrorMoreData || bufferSize <= buffer.Capacity)
        {
            AppLogger.Debug($"Hinweise: WNetGetConnection konnte Laufwerk '{driveName}' nicht aufloesen. Result={result}.");
            return null;
        }

        buffer = new StringBuilder(bufferSize);
        result = NativeMethods.WNetGetConnection(driveName, buffer, ref bufferSize);
        if (result == NoError)
        {
            return buffer.ToString().TrimEnd('\\');
        }

        AppLogger.Debug($"Hinweise: WNetGetConnection konnte Laufwerk '{driveName}' auch mit erweitertem Buffer nicht aufloesen. Result={result}.");
        return null;
    }

    private static string? TryResolveUniversalName(string path)
    {
        var bufferSize = 0;
        var result = NativeMethods.WNetGetUniversalName(path, UniversalNameInfoLevel, IntPtr.Zero, ref bufferSize);
        if (result != ErrorMoreData || bufferSize <= 0)
        {
            return null;
        }

        var buffer = Marshal.AllocHGlobal(bufferSize);
        try
        {
            result = NativeMethods.WNetGetUniversalName(path, UniversalNameInfoLevel, buffer, ref bufferSize);
            if (result != NoError)
            {
                return null;
            }

            var info = Marshal.PtrToStructure<UniversalNameInfo>(buffer);
            return Marshal.PtrToStringUni(info.UniversalName)?.TrimEnd('\\');
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    private static string NormalizePath(string path)
    {
        return path
            .Replace('/', '\\')
            .Normalize(NormalizationForm.FormC)
            .TrimEnd('\\');
    }

    private static string ResolveFolderName(string documentPath)
    {
        var directory = Path.GetDirectoryName(documentPath);
        return string.IsNullOrWhiteSpace(directory) ? string.Empty : Path.GetFileName(directory);
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private readonly struct UniversalNameInfo
    {
        public readonly IntPtr UniversalName;
    }

    private static class NativeMethods
    {
        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        public static extern int WNetGetUniversalName(string lpLocalPath, int dwInfoLevel, IntPtr lpBuffer, ref int lpBufferSize);

        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        public static extern int WNetGetConnection(string lpLocalName, StringBuilder lpRemoteName, ref int lpnLength);
    }
}
