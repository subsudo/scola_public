using VerlaufsakteApp.Models;
using System.Text;
using System.Text.RegularExpressions;

namespace VerlaufsakteApp.Services;

public class ParticipantParser
{
    private const string LateValueNormalized = "verspaetet";
    private const string AbsentKeywordNormalized = "abwesend";
    private const string PresentKeywordNormalized = "anwesend";
    private const string PresentStatusDisplay = "Anwesend";
    private const string LateStatusDisplay = "Verspätet";
    private const string AbsentExcusedStatusDisplay = "Abwesend (entschuldigt)";
    private const string AbsentUnexcusedStatusDisplay = "Abwesend (unentschuldigt)";
    private static readonly string[] NoteMarkerPrefixes =
    {
        "kommt ",
        "kommt um",
        "komme ",
        "spaeter",
        "später",
        "ab ",
        "heute ",
        "krank",
        "arzt",
        "ferien",
        "urlaub",
        "meldung",
        "notiz"
    };
    private readonly HashSet<string> _absenceValues;
    private readonly HashSet<string> _presenceValues;
    private readonly List<string> _knownStatusSuffixes;
    private readonly string _defaultPresenceLabel;
    private static readonly Regex InlineStatusMarkerRegex = new(
        @"^\s*(?<name>.+?)\s+(?<status>Abwesend\s*\(unentschuldigt\)|Abwesend\s*\(entschuldigt\)|Versp(?:ä|ae|a)tet|Anwesend)(?=\s|$).*$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex TrailingTimeDurationRegex = new(
        @"\s+((?:[01]?\d|2[0-3])[:.][0-5]\d(\s+uhr)?|\d+\s*(min(uten)?|stunden?|std|h)\b|\d{1,2}\s+uhr)\s*$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex WhitespaceRegex = new(@"\s+", RegexOptions.Compiled);

    public ParticipantParser(IEnumerable<string> absenceValues, IEnumerable<string> presenceValues)
    {
        var absence = absenceValues?.ToList() ?? new List<string>();
        var presence = presenceValues?.ToList() ?? new List<string>();

        _absenceValues = new HashSet<string>(absence.Select(NormalizeStatus));
        _presenceValues = new HashSet<string>(presence.Select(NormalizeStatus));
        _knownStatusSuffixes = _absenceValues
            .Concat(_presenceValues)
            .Append(LateValueNormalized)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.Ordinal)
            .OrderByDescending(v => v.Length)
            .ToList();
        _defaultPresenceLabel = presence.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v)) ?? "Anwesend";
    }

    public List<ParsedEntry> ParseEntries(string input, Func<string, string?>? nameFallbackResolver = null)
    {
        var result = new List<ParsedEntry>();
        if (string.IsNullOrWhiteSpace(input))
        {
            return result;
        }

        var lineNumber = 0;
        foreach (var line in EnumerateRelevantLines(input))
        {
            lineNumber++;
            if (!TryParseLine(
                    line,
                    nameFallbackResolver,
                    lineNumber,
                    out var name,
                    out var status,
                    out var absenceReason,
                    out var remark))
            {
                continue;
            }

            result.Add(new ParsedEntry
            {
                Name = name,
                Status = status,
                AbsenceReason = absenceReason,
                Remark = remark
            });
        }

        return result;
    }

    public List<Participant> Parse(string input, Func<string, string?>? nameFallbackResolver = null)
    {
        return ParseEntries(input, nameFallbackResolver)
            .Select(entry =>
            {
                var normalized = NormalizeStatus(entry.Status);
                var isPresent = DetermineIsPresent(normalized);
                var absenceDetail = isPresent
                    ? string.Empty
                    : (string.IsNullOrWhiteSpace(entry.AbsenceReason) ? entry.Status : entry.AbsenceReason);

                return new Participant
                {
                    FullName = entry.Name,
                    Status = entry.Status,
                    AbsenceDetail = absenceDetail,
                    IsPresent = isPresent,
                    MatchStatus = MatchStatus.NotFound
                };
            })
            .ToList();
    }

    private IEnumerable<string> EnumerateRelevantLines(string input)
    {
        var lines = input.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            yield return line;
        }
    }

    private bool TryParseLine(
        string line,
        Func<string, string?>? nameFallbackResolver,
        int lineNumber,
        out string name,
        out string status,
        out string absenceReason,
        out string remark)
    {
        name = string.Empty;
        status = string.Empty;
        absenceReason = string.Empty;
        remark = string.Empty;

        var cols = line.Split('\t');
        if (cols.Length == 0)
        {
            return false;
        }

        name = SanitizeCell(cols[0]);
        if (name.Equals("Teilnehmername", StringComparison.OrdinalIgnoreCase) ||
            string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        var statusRaw = cols.Length > 1 ? SanitizeCell(cols[1]) : string.Empty;
        var absenceReasonRaw = cols.Length > 2 ? SanitizeCell(cols[2]) : string.Empty;
        var remarkRaw = cols.Length > 3 ? SanitizeCell(cols[3]) : string.Empty;

        AppLogger.Debug($"Parser L{lineNumber}: Rohzeile='{ShortForLog(line)}', Cols={cols.Length}, InitialName='{name}', StatusCol='{ShortForLog(statusRaw)}', AbsenceCol='{ShortForLog(absenceReasonRaw)}', RemarkCol='{ShortForLog(remarkRaw)}'.");

        var recognizedStatus = IsRecognizedStatus(statusRaw);
        if (!recognizedStatus)
        {
            if (TryExtractStatusFromLine(name, out var extractedName, out var extractedStatus) ||
                TryExtractStatusFromLine(line, out extractedName, out extractedStatus))
            {
                name = extractedName;
                statusRaw = extractedStatus;
                recognizedStatus = true;
                AppLogger.Info($"Parser L{lineNumber}: Status aus Rohzeile extrahiert. Name='{name}', Status='{statusRaw}'.");
            }
            else
            {
                // Statusspalte enthaelt freie Bemerkung: ignorieren.
                if (!string.IsNullOrWhiteSpace(statusRaw))
                {
                    AppLogger.Info($"Parser L{lineNumber}: Unbekannte Statusspalte ignoriert ('{ShortForLog(statusRaw)}').");
                }
                statusRaw = string.Empty;
            }
        }

        if (TryExtractTrailingNote(name, out var strippedNameAfterStatus, out var extractedRemarkAfterStatus))
        {
            name = strippedNameAfterStatus;
            if (string.IsNullOrWhiteSpace(remarkRaw))
            {
                remarkRaw = extractedRemarkAfterStatus;
            }

            AppLogger.Info($"Parser L{lineNumber}: Freitext-Anhang vom Namen getrennt. Name='{name}', Remark='{ShortForLog(extractedRemarkAfterStatus)}'.");
        }

        if (string.IsNullOrWhiteSpace(statusRaw))
        {
            var extraction = TryExtractInlineStatus(name);
            if (!string.IsNullOrWhiteSpace(extraction.Status))
            {
                name = extraction.Name;
                statusRaw = extraction.Status;
                recognizedStatus = true;
                AppLogger.Info($"Parser L{lineNumber}: Status aus Namenssuffix extrahiert. Name='{name}', Status='{statusRaw}'.");
            }
            else if (nameFallbackResolver is not null)
            {
                var resolvedName = nameFallbackResolver(line);
                if (!string.IsNullOrWhiteSpace(resolvedName))
                {
                    name = SanitizeCell(resolvedName);
                    AppLogger.Info($"Parser L{lineNumber}: Name via Ordner-Fallback aufgeloest zu '{name}'. Rohzeile='{ShortForLog(line)}'.");
                }
                else
                {
                    AppLogger.Warn($"Parser L{lineNumber}: Kein eindeutiger Name via Ordner-Fallback. Verwende erste Spalte='{name}'. Rohzeile='{ShortForLog(line)}'.");
                }
            }
        }

        var canonicalStatus = NormalizeDisplayStatus(statusRaw);
        var normalized = NormalizeStatus(canonicalStatus);
        var isPresent = DetermineIsPresent(normalized);

        if (string.IsNullOrWhiteSpace(canonicalStatus))
        {
            status = _defaultPresenceLabel;
            AppLogger.Info($"Parser L{lineNumber}: Kein Statusmarker gefunden, setze Default-Status '{status}'. Name='{name}'.");
        }
        else
        {
            status = canonicalStatus;
        }

        if (isPresent)
        {
            absenceReason = string.Empty;
            remark = string.Empty;
        }
        else
        {
            absenceReason = string.IsNullOrWhiteSpace(absenceReasonRaw) ? status : absenceReasonRaw;
            remark = remarkRaw;
        }
        AppLogger.Debug($"Parser L{lineNumber}: Final Name='{name}', Status='{status}', IsPresent={isPresent}, AbsenceReason='{ShortForLog(absenceReason)}', Remark='{ShortForLog(remark)}'.");
        return true;
    }

    private static string ShortForLog(string value, int maxLength = 160)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = NormalizeWhitespace(value).Replace('\t', ' ');
        return normalized.Length <= maxLength
            ? normalized
            : $"{normalized[..maxLength]}...";
    }

    private bool DetermineIsPresent(string normalizedStatus)
    {
        if (string.IsNullOrWhiteSpace(normalizedStatus))
        {
        return true;
    }

        if (normalizedStatus.Contains(AbsentKeywordNormalized, StringComparison.Ordinal))
        {
            return false;
        }

        if (normalizedStatus == LateValueNormalized)
        {
        return true;
    }

        if (normalizedStatus.Contains("verspaet", StringComparison.Ordinal))
        {
        return true;
    }

        if (normalizedStatus.Contains(PresentKeywordNormalized, StringComparison.Ordinal))
        {
        return true;
    }

        if (_absenceValues.Contains(normalizedStatus))
        {
            return false;
        }

        if (_presenceValues.Contains(normalizedStatus))
        {
        return true;
    }
        return true;
    }

    private bool IsRecognizedStatus(string statusRaw)
    {
        if (string.IsNullOrWhiteSpace(statusRaw))
        {
            return false;
        }

        var normalized = NormalizeStatus(statusRaw);
        if (normalized.Contains(AbsentKeywordNormalized, StringComparison.Ordinal) ||
            normalized.Contains(PresentKeywordNormalized, StringComparison.Ordinal) ||
            normalized.Contains("verspaet", StringComparison.Ordinal))
        {
        return true;
    }

        return _absenceValues.Contains(normalized) || _presenceValues.Contains(normalized);
    }

    private bool TryExtractStatusFromLine(string line, out string name, out string status)
    {
        name = string.Empty;
        status = string.Empty;

        var prepared = NormalizeWhitespace(SanitizeCell(line));
        var match = InlineStatusMarkerRegex.Match(prepared);
        if (!match.Success)
        {
            return false;
        }

        name = SanitizeCell(match.Groups["name"].Value);
        status = NormalizeDisplayStatus(match.Groups["status"].Value);
        return !string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(status);
    }

    private static string NormalizeDisplayStatus(string rawStatus)
    {
        if (string.IsNullOrWhiteSpace(rawStatus))
        {
            return string.Empty;
        }

        var normalized = NormalizeStatus(rawStatus);
        if (normalized.Contains("unentsch", StringComparison.Ordinal))
        {
            return AbsentUnexcusedStatusDisplay;
        }

        if (normalized.Contains(AbsentKeywordNormalized, StringComparison.Ordinal))
        {
            if (normalized.Contains("entsch", StringComparison.Ordinal))
            {
                return AbsentExcusedStatusDisplay;
            }

            return "Abwesend";
        }

        if (normalized.Contains("verspaet", StringComparison.Ordinal))
        {
            return LateStatusDisplay;
        }

        if (normalized.Contains(PresentKeywordNormalized, StringComparison.Ordinal))
        {
            return PresentStatusDisplay;
        }

        return SanitizeCell(rawStatus);
    }

    private static string NormalizeStatus(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = SanitizeCell(value).ToLowerInvariant();
        normalized = normalized
            .Replace("ä", "ae", StringComparison.Ordinal)
            .Replace("ö", "oe", StringComparison.Ordinal)
            .Replace("ü", "ue", StringComparison.Ordinal)
            .Replace("ß", "ss", StringComparison.Ordinal);

        var parts = normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        return string.Join(" ", parts);
    }

    private static string SanitizeCell(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        return value
            .Normalize(NormalizationForm.FormC)
            .Replace('\u00A0', ' ')
            .Replace('\u2007', ' ')
            .Replace('\u202F', ' ')
            .Replace('\t', ' ')
            .Trim();
    }

    private static string NormalizeWhitespace(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return WhitespaceRegex.Replace(value, " ").Trim();
    }

    private static bool TryExtractTrailingNote(string rawName, out string strippedName, out string remark)
    {
        strippedName = rawName;
        remark = string.Empty;

        var normalized = NormalizeWhitespace(rawName);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return false;
        }

        var tokens = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length < 3)
        {
            return TryStripTrailingTimeDuration(rawName, out strippedName, out remark);
        }

        for (var splitIndex = 2; splitIndex < tokens.Length; splitIndex++)
        {
            var candidateRemark = string.Join(" ", tokens.Skip(splitIndex));
            var candidateRemarkNormalized = NormalizeStatus(candidateRemark);
            if (!StartsWithNoteMarker(candidateRemarkNormalized))
            {
                continue;
            }

            var candidateName = string.Join(" ", tokens.Take(splitIndex)).Trim();
            if (candidateName.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length < 2)
            {
                continue;
            }

            strippedName = candidateName;
            remark = candidateRemark;
        return true;
    }

        return TryStripTrailingTimeDuration(rawName, out strippedName, out remark);
    }

    private static bool TryStripTrailingTimeDuration(string rawName, out string strippedName, out string remark)
    {
        strippedName = rawName;
        remark = string.Empty;

        var match = TrailingTimeDurationRegex.Match(rawName);
        if (!match.Success)
        {
            return false;
        }

        var candidate = rawName[..match.Index].Trim();
        if (candidate.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length < 2)
        {
            return false;
        }

        strippedName = candidate;
        remark = match.Value.Trim();
        return true;
    }

    private static bool StartsWithNoteMarker(string normalizedRemark)
    {
        if (string.IsNullOrWhiteSpace(normalizedRemark))
        {
            return false;
        }

        foreach (var prefix in NoteMarkerPrefixes)
        {
            if (normalizedRemark.StartsWith(prefix, StringComparison.Ordinal))
            {
        return true;
    }
        }

        return false;
    }

    private (string Name, string Status) TryExtractInlineStatus(string rawName)
    {
        var tokens = rawName
            .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
            .ToList();

        if (tokens.Count < 2)
        {
            return (rawName, string.Empty);
        }

        foreach (var suffix in _knownStatusSuffixes)
        {
            var suffixTokens = suffix
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (tokens.Count <= suffixTokens.Length)
            {
                continue;
            }

            var tail = tokens.Skip(tokens.Count - suffixTokens.Length).ToArray();
            var tailNormalized = NormalizeStatus(string.Join(" ", tail));
            if (!tailNormalized.Equals(suffix, StringComparison.Ordinal))
            {
                continue;
            }

            var extractedName = string.Join(" ", tokens.Take(tokens.Count - suffixTokens.Length));
            var extractedStatus = string.Join(" ", tail);
            return (extractedName, extractedStatus);
        }

        return (rawName, string.Empty);
    }
}




