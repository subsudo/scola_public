using System.IO;
using System.Text.RegularExpressions;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp.Services;

public class FolderMatcher
{
    private const int MinRobustTokenLength = 3;
    private const int MinRobustTokenCountForAutoMatch = 2;
    private static readonly Regex TokenRegex = new(@"[\p{L}\p{N}]+", RegexOptions.Compiled);
    private static readonly Regex MultiWhitespaceRegex = new(@"\s+", RegexOptions.Compiled);

    private readonly string _primaryServerBasePath;
    private readonly string _secondaryServerBasePath;
    private readonly string _tertiaryServerBasePath;
    private readonly bool _useSecondaryServerBasePath;
    private readonly bool _useTertiaryServerBasePath;
    private readonly string _verlaufsakteKeyword;
    private readonly InitialsResolver _initialsResolver;
    private readonly object _sync = new();
    private volatile bool _isIndexed;
    private volatile List<FolderEntry> _folders = new();

    public FolderMatcher(
        string serverBasePath,
        bool useSecondaryServerBasePath = false,
        string? secondaryServerBasePath = null,
        bool useTertiaryServerBasePath = false,
        string? tertiaryServerBasePath = null,
        string? verlaufsakteKeyword = null,
        InitialsResolver? initialsResolver = null)
    {
        _primaryServerBasePath = serverBasePath ?? string.Empty;
        _useSecondaryServerBasePath = useSecondaryServerBasePath;
        _secondaryServerBasePath = secondaryServerBasePath ?? string.Empty;
        _useTertiaryServerBasePath = useTertiaryServerBasePath;
        _tertiaryServerBasePath = tertiaryServerBasePath ?? string.Empty;
        _verlaufsakteKeyword = string.IsNullOrWhiteSpace(verlaufsakteKeyword) ? "Verlaufsakte" : verlaufsakteKeyword;
        _initialsResolver = initialsResolver ?? new InitialsResolver();
    }

    public void BuildIndex()
    {
        if (_isIndexed)
        {
            return;
        }

        lock (_sync)
        {
            if (_isIndexed)
            {
                return;
            }

            var folderPaths = new List<string>();

            if (Directory.Exists(_primaryServerBasePath))
            {
                try
                {
                    folderPaths.AddRange(Directory.GetDirectories(_primaryServerBasePath, "*", SearchOption.TopDirectoryOnly));
                }
                catch (Exception ex) when (ex is UnauthorizedAccessException or IOException)
                {
                    AppLogger.Warn($"FolderMatcher: Primaerpfad konnte nicht gelesen werden: '{_primaryServerBasePath}': {ex.Message}");
                }
            }
            else
            {
                AppLogger.Warn($"FolderMatcher: Primaerpfad nicht erreichbar: '{_primaryServerBasePath}'.");
            }

            if (_useSecondaryServerBasePath)
            {
                if (Directory.Exists(_secondaryServerBasePath))
                {
                    try
                    {
                        folderPaths.AddRange(Directory.GetDirectories(_secondaryServerBasePath, "*", SearchOption.TopDirectoryOnly));
                    }
                    catch (Exception ex) when (ex is UnauthorizedAccessException or IOException)
                    {
                        AppLogger.Warn($"FolderMatcher: Sekundaerpfad konnte nicht gelesen werden: '{_secondaryServerBasePath}': {ex.Message}");
                    }
                }
                else if (!string.IsNullOrWhiteSpace(_secondaryServerBasePath))
                {
                    AppLogger.Warn($"FolderMatcher: Sekundaerpfad aktiviert, aber nicht erreichbar: '{_secondaryServerBasePath}'.");
                }
            }

            if (_useTertiaryServerBasePath)
            {
                if (Directory.Exists(_tertiaryServerBasePath))
                {
                    try
                    {
                        folderPaths.AddRange(Directory.GetDirectories(_tertiaryServerBasePath, "*", SearchOption.TopDirectoryOnly));
                    }
                    catch (Exception ex) when (ex is UnauthorizedAccessException or IOException)
                    {
                        AppLogger.Warn($"FolderMatcher: Drittpfad konnte nicht gelesen werden: '{_tertiaryServerBasePath}': {ex.Message}");
                    }
                }
                else if (!string.IsNullOrWhiteSpace(_tertiaryServerBasePath))
                {
                    AppLogger.Warn($"FolderMatcher: Drittpfad aktiviert, aber nicht erreichbar: '{_tertiaryServerBasePath}'.");
                }
            }

            _folders = folderPaths
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(d => new FolderEntry(d, _verlaufsakteKeyword, _initialsResolver))
                .ToList();
            _isIndexed = true;
            AppLogger.Debug($"FolderMatcher: Index aufgebaut. Primaer='{_primaryServerBasePath}', SekundaerAktiv={_useSecondaryServerBasePath}, Sekundaer='{_secondaryServerBasePath}', DrittpfadAktiv={_useTertiaryServerBasePath}, Drittpfad='{_tertiaryServerBasePath}', FolderCount={_folders.Count}.");
        }
    }

    public void MatchParticipant(Participant participant)
    {
        if (!_isIndexed)
        {
            BuildIndex();
        }

        AppLogger.Debug($"FolderMatcher.MatchParticipant start. Name='{participant.FullName}'.");

        var nameTokens = Tokenize(participant.FullName);
        var requiredTokenCounts = BuildRequiredTokenCounts(nameTokens);
        var matches = FindStrictMatches(requiredTokenCounts, useFallback: false);

        if (matches.Count == 0)
        {
            var fallbackTokens = Tokenize(ReplaceUmlauts(participant.FullName));
            requiredTokenCounts = BuildRequiredTokenCounts(fallbackTokens);
            matches = FindStrictMatches(requiredTokenCounts, useFallback: true);
        }

        if (matches.Count == 0)
        {
            AppLogger.Debug($"FolderMatcher.MatchParticipant: kein Treffer fuer '{participant.FullName}'. RequiredTokens={string.Join(",", requiredTokenCounts.Keys)}.");
            participant.MatchStatus = MatchStatus.NotFound;
            participant.MatchedFolderPath = null;
            participant.CandidateFolderPaths = new List<string>();
            participant.SelectedFolderPath = null;
            participant.DocumentPath = string.Empty;
            participant.Initials = string.Empty;
            participant.OdooUrl = string.Empty;
            participant.CounselorInitials = string.Empty;
            participant.IsHeaderMetadataLoaded = false;
            return;
        }

        var candidates = matches
            .Select(m => m.Path)
            .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
            .ToList();
        participant.CandidateFolderPaths = candidates;
        AppLogger.Debug($"FolderMatcher.MatchParticipant: Kandidaten fuer '{participant.FullName}': Count={candidates.Count}, Paths='{ShortForLog(string.Join(" | ", candidates), 400)}'.");

        if (candidates.Count == 1)
        {
            var match = matches[0];
            participant.MatchStatus = MatchStatus.Found;
            participant.MatchedFolderPath = match.Path;
            participant.SelectedFolderPath = match.Path;
            participant.DocumentPath = match.PreferredDocumentPath;
            participant.Initials = match.Initials;
            participant.OdooUrl = string.Empty;
            participant.CounselorInitials = string.Empty;
            participant.IsHeaderMetadataLoaded = false;
            return;
        }

        participant.MatchStatus = MatchStatus.MultipleFound;
        participant.SelectedFolderPath = null;
        participant.MatchedFolderPath = null;
        participant.DocumentPath = string.Empty;
        participant.Initials = string.Empty;
        participant.OdooUrl = string.Empty;
        participant.CounselorInitials = string.Empty;
        participant.IsHeaderMetadataLoaded = false;
    }

    public string? ResolveLikelyNameFromRawLine(string rawLine)
    {
        if (!_isIndexed)
        {
            BuildIndex();
        }

        var orderedTokens = TokenizeOrdered(rawLine);
        AppLogger.Debug($"FolderMatcher.ResolveLikelyNameFromRawLine start. Rohzeile='{ShortForLog(rawLine)}', Tokens='{string.Join(" | ", orderedTokens)}'.");
        if (orderedTokens.Count < MinRobustTokenCountForAutoMatch)
        {
            AppLogger.Warn($"FolderMatcher: Fallback-Nameaufloesung uebersprungen (zu wenige Tokens). Rohzeile='{ShortForLog(rawLine)}'.");
            return null;
        }

        for (var end = orderedTokens.Count; end >= MinRobustTokenCountForAutoMatch; end--)
        {
            var prefix = orderedTokens.Take(end).ToList();
            var requiredTokenCounts = BuildRequiredTokenCounts(prefix);
            if (GetRequiredTokenTotal(requiredTokenCounts) < MinRobustTokenCountForAutoMatch)
            {
                continue;
            }

            var matches = FindStrictMatches(requiredTokenCounts, useFallback: false);
            AppLogger.Debug($"FolderMatcher.ResolveLikelyNameFromRawLine: Prefix='{string.Join(" ", prefix)}', MatchesDirect={matches.Count}.");
            if (matches.Count == 0)
            {
                var fallbackPrefix = TokenizeOrdered(ReplaceUmlauts(string.Join(" ", prefix)));
                requiredTokenCounts = BuildRequiredTokenCounts(fallbackPrefix);
                matches = FindStrictMatches(requiredTokenCounts, useFallback: true);
                AppLogger.Debug($"FolderMatcher.ResolveLikelyNameFromRawLine: PrefixFallback='{string.Join(" ", fallbackPrefix)}', MatchesFallback={matches.Count}.");
            }

            if (matches.Count == 1)
            {
                var resolved = ReconstructCleanName(prefix, matches[0]);
                AppLogger.Info($"FolderMatcher: Fallback-Name eindeutig aufgeloest zu '{resolved}'. Rohzeile='{ShortForLog(rawLine)}'.");
                return resolved;
            }
        }

        AppLogger.Warn($"FolderMatcher: Kein eindeutiger Fallback-Name gefunden. Rohzeile='{ShortForLog(rawLine)}'.");
        return null;
    }

    private List<FolderEntry> FindStrictMatches(Dictionary<string, int> requiredTokenCounts, bool useFallback)
    {
        if (GetRequiredTokenTotal(requiredTokenCounts) < MinRobustTokenCountForAutoMatch)
        {
            return new List<FolderEntry>();
        }

        var matches = new List<FolderEntry>();

        foreach (var folder in _folders)
        {
            var folderTokenCounts = useFallback ? folder.FallbackTokenCounts : folder.OriginalTokenCounts;
            if (HasTokenCountsMatch(requiredTokenCounts, folderTokenCounts))
            {
                matches.Add(folder);
            }
        }

        return matches;
    }

    private static Dictionary<string, int> BuildRequiredTokenCounts(IEnumerable<string> tokens)
    {
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var token in tokens.Where(IsRobustToken))
        {
            counts[token] = counts.TryGetValue(token, out var current) ? current + 1 : 1;
        }

        return counts;
    }

    private static int GetRequiredTokenTotal(Dictionary<string, int> requiredTokenCounts)
    {
        var total = 0;
        foreach (var count in requiredTokenCounts.Values)
        {
            total += count;
        }

        return total;
    }

    private static bool HasTokenCountsMatch(
        Dictionary<string, int> requiredTokenCounts,
        Dictionary<string, int> folderTokenCounts)
    {
        foreach (var required in requiredTokenCounts)
        {
            if (!folderTokenCounts.TryGetValue(required.Key, out var available))
            {
                return false;
            }

            if (available < required.Value)
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsRobustToken(string token)
    {
        if (string.IsNullOrWhiteSpace(token) || token.Length < MinRobustTokenLength)
        {
            return false;
        }

        return token.Any(char.IsLetterOrDigit);
    }

    private static List<string> Tokenize(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return new List<string>();
        }

        return TokenRegex.Matches(name.ToLowerInvariant())
            .Select(m => m.Value.Trim())
            .Where(t => t.Length > 0)
            .ToList();
    }

    private static List<string> TokenizeOrdered(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return new List<string>();
        }

        return TokenRegex.Matches(value)
            .Select(m => m.Value.Trim())
            .Where(t => t.Length > 0)
            .ToList();
    }

    private static string ReplaceUmlauts(string value)
    {
        return value
            .Replace("ä", "ae", StringComparison.OrdinalIgnoreCase)
            .Replace("ö", "oe", StringComparison.OrdinalIgnoreCase)
            .Replace("ü", "ue", StringComparison.OrdinalIgnoreCase)
            .Replace("ß", "ss", StringComparison.OrdinalIgnoreCase);
    }

    private static string ReconstructCleanName(List<string> inputTokens, FolderEntry folder)
    {
        var folderTokenSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in folder.OriginalTokenCounts.Keys)
        {
            folderTokenSet.Add(key);
        }

        foreach (var key in folder.FallbackTokenCounts.Keys)
        {
            folderTokenSet.Add(key);
        }

        var matchedRobustTokenIndexes = inputTokens
            .Select((token, index) => new { token, index })
            .Where(x => IsRobustToken(x.token) &&
                        (folderTokenSet.Contains(x.token) ||
                         folderTokenSet.Contains(ReplaceUmlauts(x.token))))
            .Select(x => x.index)
            .ToList();

        if (matchedRobustTokenIndexes.Count >= MinRobustTokenCountForAutoMatch)
        {
            var startIndex = matchedRobustTokenIndexes.First();
            var endIndex = matchedRobustTokenIndexes.Last();
            var cleanTokens = inputTokens
                .Skip(startIndex)
                .Take(endIndex - startIndex + 1)
                .Where(token => token.Any(char.IsLetter))
                .ToList();

            if (cleanTokens.Count >= MinRobustTokenCountForAutoMatch)
            {
                return string.Join(" ", cleanTokens);
            }
        }

        return folder.DisplayName;
    }

    private static string ShortForLog(string value, int maxLength = 160)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = MultiWhitespaceRegex.Replace(value, " ").Trim();
        return normalized.Length <= maxLength
            ? normalized
            : $"{normalized[..maxLength]}...";
    }

    private sealed class FolderEntry
    {
        public FolderEntry(string path, string keyword, InitialsResolver initialsResolver)
        {
            Path = path;
            DisplayName = System.IO.Path.GetFileName(path);
            OriginalName = DisplayName.ToLowerInvariant();
            FallbackName = ReplaceUmlauts(OriginalName);
            PreferredDocumentPath = TryFindDocument(path, keyword) ?? string.Empty;
            Initials = initialsResolver.TryResolveFromDocumentPath(PreferredDocumentPath);
            OriginalTokenCounts = BuildTokenCounts(OriginalName, Initials, useFallback: false);
            FallbackTokenCounts = BuildTokenCounts(FallbackName, Initials, useFallback: true);
        }

        public string Path { get; }
        public string DisplayName { get; }
        public string OriginalName { get; }
        public string FallbackName { get; }
        public string PreferredDocumentPath { get; }
        public string Initials { get; }
        public Dictionary<string, int> OriginalTokenCounts { get; }
        public Dictionary<string, int> FallbackTokenCounts { get; }

        private static Dictionary<string, int> BuildTokenCounts(string baseName, string initials, bool useFallback)
        {
            var tokens = Tokenize(baseName).ToList();
            if (!string.IsNullOrWhiteSpace(initials))
            {
                tokens.Add(initials.ToLowerInvariant());
                if (useFallback)
                {
                    tokens.Add(ReplaceUmlauts(initials).ToLowerInvariant());
                }
            }

            return BuildRequiredTokenCounts(tokens);
        }

        private static string? TryFindDocument(string folderPath, string keyword)
        {
            try
            {
                return Directory
                    .GetFiles(folderPath, "*.docx", SearchOption.TopDirectoryOnly)
                    .Where(path => System.IO.Path.GetFileName(path).Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(path => path, StringComparer.CurrentCultureIgnoreCase)
                    .FirstOrDefault();
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"FolderMatcher: Dokumente konnten nicht gelesen werden '{folderPath}': {ex.Message}");
                return null;
            }
        }
    }
}
