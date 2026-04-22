using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp.Services;

public sealed class DocxHeaderMetadataService
{
    // Nur gezielt erhöhen, wenn Struktur oder fachliche Bedeutung der gecachten
    // Header-Werte sich ändert und ein alter Cache falsches Verhalten konservieren würde.
    private const int CacheVersion = 5;
    private static readonly TimeSpan MissingDocumentRetention = TimeSpan.FromDays(30);
    private static readonly Regex InlineCounselorRegex = new(@"Beratungsperson\s*:?\s*(?<code>(?:[\p{Lu}\p{N}]\s*){1,8})\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex PlaceholderHeaderRegex = new(@"Nachname|Vorname|Name\s+Vorname|\bXX\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex UpperTokenRegex = new(@"^[\p{Lu}\p{N}]{1,8}$", RegexOptions.Compiled);
    private static readonly Regex CounselorInitialsRegex = new(@"^[\p{Lu}\p{N}]{2}$", RegexOptions.Compiled);

    private readonly object _syncRoot = new();
    private readonly string _cachePath;
    private readonly string _cacheBackupPath;
    private readonly Dictionary<string, CacheEntry> _cache;

    public DocxHeaderMetadataService(string cachePath, string cacheBackupPath)
    {
        _cachePath = cachePath;
        _cacheBackupPath = cacheBackupPath;
        _cache = LoadCache();
        PruneMissingEntries();
    }

    public HeaderMetadata Read(string documentPath)
    {
        var fileInfo = new FileInfo(documentPath);
        if (!fileInfo.Exists)
        {
            return HeaderMetadata.Empty;
        }

        CacheEntry? previousCacheEntry = null;
        lock (_syncRoot)
        {
            if (_cache.TryGetValue(documentPath, out var cached) &&
                cached.LastWriteTimeUtc == fileInfo.LastWriteTimeUtc &&
                cached.Length == fileInfo.Length)
            {
                TouchCacheEntryUnsafe(fileInfo.FullName, cached);
                AppLogger.Debug($"HeaderMetadataCache hit fuer '{documentPath}'.");
                return cached.Metadata;
            }

            _cache.TryGetValue(documentPath, out previousCacheEntry);
        }

        try
        {
            var parsed = ReadFromPackage(documentPath);
            var metadata = MergeMetadataWithCache(parsed, previousCacheEntry);
            var headerSignature = parsed.HeaderSignature;

            if (previousCacheEntry is not null &&
                string.Equals(previousCacheEntry.HeaderSignature, headerSignature, StringComparison.Ordinal))
            {
                metadata = metadata with { CounselorInitials = previousCacheEntry.Metadata.CounselorInitials };
            }

            UpdateCache(fileInfo, headerSignature, metadata);
            return metadata;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Header-Metadaten konnten nicht gelesen werden '{documentPath}': {ex.Message}");
            if (previousCacheEntry is not null)
            {
                AppLogger.Info($"Verwende vorhandenen Header-Metadaten-Cache weiter '{documentPath}'.");
                return previousCacheEntry.Metadata;
            }

            return HeaderMetadata.Empty;
        }
    }

    private HeaderMetadata MergeMetadataWithCache(ParsedHeaderMetadata parsed, CacheEntry? previousCacheEntry)
    {
        if (previousCacheEntry is null)
        {
            return parsed.Metadata;
        }

        var stickyOdooUrl = !string.IsNullOrWhiteSpace(previousCacheEntry.Metadata.OdooUrl)
            ? previousCacheEntry.Metadata.OdooUrl
            : parsed.Metadata.OdooUrl;

        return parsed.Metadata with { OdooUrl = stickyOdooUrl };
    }

    private Dictionary<string, CacheEntry> LoadCache()
    {
        var utcNow = DateTime.UtcNow;
        var document = JsonStorage.Load(_cachePath, _cacheBackupPath, () => new HeaderMetadataCacheDocument());
        if (document.Version != CacheVersion)
        {
            return new Dictionary<string, CacheEntry>(StringComparer.OrdinalIgnoreCase);
        }

        return document.Entries
            .Where(entry => !string.IsNullOrWhiteSpace(entry.DocumentPath))
            .ToDictionary(
                entry => entry.DocumentPath,
                entry => new CacheEntry(
                    entry.LastWriteTimeUtc,
                    entry.Length,
                    entry.LastSeenUtc == default ? utcNow : entry.LastSeenUtc,
                    entry.HeaderSignature ?? string.Empty,
                    new HeaderMetadata(entry.OdooUrl ?? string.Empty, entry.CounselorInitials ?? string.Empty)),
                StringComparer.OrdinalIgnoreCase);
    }

    private void UpdateCache(FileInfo fileInfo, string headerSignature, HeaderMetadata metadata)
    {
        lock (_syncRoot)
        {
            var path = fileInfo.FullName;
            if (_cache.TryGetValue(path, out var existing) &&
                existing.LastWriteTimeUtc == fileInfo.LastWriteTimeUtc &&
                existing.Length == fileInfo.Length &&
                string.Equals(existing.HeaderSignature, headerSignature, StringComparison.Ordinal) &&
                existing.Metadata == metadata)
            {
                TouchCacheEntryUnsafe(path, existing);
                return;
            }

            _cache[path] = new CacheEntry(fileInfo.LastWriteTimeUtc, fileInfo.Length, DateTime.UtcNow, headerSignature, metadata);
            PersistCacheUnsafe();
        }
    }

    private void TouchCacheEntryUnsafe(string documentPath, CacheEntry existing)
    {
        var utcNow = DateTime.UtcNow;
        if (existing.LastSeenUtc.Date == utcNow.Date)
        {
            return;
        }

        _cache[documentPath] = existing with { LastSeenUtc = utcNow };
        PersistCacheUnsafe();
    }

    private void PruneMissingEntries()
    {
        lock (_syncRoot)
        {
            if (_cache.Count == 0)
            {
                return;
            }

            var utcNow = DateTime.UtcNow;
            var threshold = utcNow - MissingDocumentRetention;
            var changed = false;

            foreach (var entry in _cache.Keys.ToList())
            {
                if (File.Exists(entry))
                {
                    continue;
                }

                if (_cache[entry].LastSeenUtc <= threshold)
                {
                    _cache.Remove(entry);
                    changed = true;
                }
            }

            if (changed)
            {
                PersistCacheUnsafe();
            }
        }
    }

    private void PersistCacheUnsafe()
    {
        try
        {
            var document = new HeaderMetadataCacheDocument
            {
                Version = CacheVersion,
                Entries = _cache
                    .OrderBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(entry => new HeaderMetadataCacheEntry
                    {
                        DocumentPath = entry.Key,
                        LastWriteTimeUtc = entry.Value.LastWriteTimeUtc,
                        Length = entry.Value.Length,
                        LastSeenUtc = entry.Value.LastSeenUtc,
                        HeaderSignature = entry.Value.HeaderSignature,
                        OdooUrl = entry.Value.Metadata.OdooUrl,
                        CounselorInitials = entry.Value.Metadata.CounselorInitials
                    })
                    .ToList()
            };

            JsonStorage.SaveAtomic(_cachePath, _cacheBackupPath, document);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Header-Metadaten-Cache konnte nicht gespeichert werden: {ex.Message}");
        }
    }

    private static ParsedHeaderMetadata ReadFromPackage(string documentPath)
    {
        using var package = ZipFile.OpenRead(documentPath);

        var headerEntries = package.Entries
            .Where(entry => entry.FullName.StartsWith("word/header", StringComparison.OrdinalIgnoreCase) &&
                            entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (headerEntries.Count == 0)
        {
            return new ParsedHeaderMetadata(HeaderMetadata.Empty, string.Empty);
        }

        var headers = new List<HeaderCandidate>();
        var headerContents = new List<string>(headerEntries.Count);

        foreach (var headerEntry in headerEntries)
        {
            var headerXml = ReadEntryText(headerEntry);
            headerContents.Add(headerXml);

            var relationshipsEntry = package.GetEntry($"word/_rels/{Path.GetFileName(headerEntry.FullName)}.rels");
            var relationships = relationshipsEntry is null
                ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                : LoadRelationships(relationshipsEntry);

            headers.Add(ReadHeader(headerXml, relationships));
        }

        var bestHeader = headers
            .OrderByDescending(header => header.Score)
            .ThenByDescending(header => !string.IsNullOrWhiteSpace(header.Metadata.OdooUrl))
            .ThenByDescending(header => !string.IsNullOrWhiteSpace(header.Metadata.CounselorInitials))
            .First();

        var odooUrl = !string.IsNullOrWhiteSpace(bestHeader.Metadata.OdooUrl)
            ? bestHeader.Metadata.OdooUrl
            : headers.FirstOrDefault(header => !string.IsNullOrWhiteSpace(header.Metadata.OdooUrl))?.Metadata.OdooUrl ?? string.Empty;

        var counselorInitials = !string.IsNullOrWhiteSpace(bestHeader.Metadata.CounselorInitials)
            ? bestHeader.Metadata.CounselorInitials
            : headers.FirstOrDefault(header => !string.IsNullOrWhiteSpace(header.Metadata.CounselorInitials))?.Metadata.CounselorInitials ?? string.Empty;

        return new ParsedHeaderMetadata(
            new HeaderMetadata(odooUrl, counselorInitials),
            ComputeHeaderSignature(headerContents));
    }

    private static string ReadEntryText(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }

    private static string ComputeHeaderSignature(IEnumerable<string> headerContents)
    {
        var combined = string.Join("\n---header---\n", headerContents);
        var bytes = Encoding.UTF8.GetBytes(combined);
        var hash = SHA256.HashData(bytes);
        return Convert.ToHexString(hash);
    }

    private static Dictionary<string, string> LoadRelationships(ZipArchiveEntry relationshipsEntry)
    {
        using var stream = relationshipsEntry.Open();
        var document = new XmlDocument();
        document.Load(stream);

        var namespaceManager = new XmlNamespaceManager(document.NameTable);
        namespaceManager.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

        var relationships = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var nodes = document.SelectNodes("/rel:Relationships/rel:Relationship", namespaceManager);
        if (nodes is null)
        {
            return relationships;
        }

        foreach (XmlNode node in nodes)
        {
            var id = node.Attributes?["Id"]?.Value;
            var target = node.Attributes?["Target"]?.Value;
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(target))
            {
                continue;
            }

            relationships[id] = target;
        }

        return relationships;
    }

    private static HeaderCandidate ReadHeader(string headerXml, IReadOnlyDictionary<string, string> relationships)
    {
        var document = new XmlDocument();
        document.LoadXml(headerXml);

        var namespaceManager = new XmlNamespaceManager(document.NameTable);
        namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        namespaceManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var textFragments = document.SelectNodes("//w:t", namespaceManager)?
            .Cast<XmlNode>()
            .Select(node => node.InnerText.Trim())
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList() ?? new List<string>();

        var allText = Regex.Replace(string.Join(" ", textFragments), @"\s+", " ").Trim();
        var odooUrl = TryExtractOdooUrl(document, namespaceManager, relationships);
        var counselorInitials = TryExtractCounselorInitials(allText, textFragments);
        var score = ScoreHeader(allText, odooUrl, counselorInitials);

        return new HeaderCandidate(new HeaderMetadata(odooUrl ?? string.Empty, counselorInitials ?? string.Empty), score);
    }

    private static string? TryExtractOdooUrl(XmlDocument document, XmlNamespaceManager namespaceManager, IReadOnlyDictionary<string, string> relationships)
    {
        string? odooUrl = null;
        var hyperlinks = document.SelectNodes("//w:hyperlink", namespaceManager);
        if (hyperlinks is not null)
        {
            foreach (XmlNode hyperlink in hyperlinks)
            {
                var relationshipId = hyperlink.Attributes?["id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"]?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId) || !relationships.TryGetValue(relationshipId, out var target))
                {
                    continue;
                }

                if (target.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    target.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    var anchor = hyperlink.Attributes?["anchor", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"]?.Value;
                    odooUrl ??= CombineTargetAndAnchor(target, anchor);
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(odooUrl))
        {
            return odooUrl;
        }

        var fieldInstructionTexts = document.SelectNodes("//w:instrText", namespaceManager)?
            .Cast<XmlNode>()
            .Select(node => node.InnerText)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList() ?? new List<string>();

        foreach (var instruction in fieldInstructionTexts)
        {
            var normalized = instruction.Trim();
            if (!normalized.Contains("HYPERLINK", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var urlMatch = Regex.Match(normalized, "\"https?://[^\\s\"']+\"", RegexOptions.IgnoreCase);
            if (urlMatch.Success)
            {
                return urlMatch.Value.Trim('"');
            }
        }

        return null;
    }

    private static string? TryExtractCounselorInitials(string allText, IReadOnlyList<string> textFragments)
    {
        foreach (Match match in InlineCounselorRegex.Matches(allText))
        {
            if (!match.Success)
            {
                continue;
            }

            var candidate = NormalizeCounselorToken(match.Groups["code"].Value);
            if (candidate is not null)
            {
                return candidate;
            }
        }

        for (var index = 0; index < textFragments.Count; index++)
        {
            var fragment = textFragments[index];
            if (!fragment.Contains("Beratungsperson", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var combinedCandidate = TryCombineCounselorFragments(textFragments, index + 1);
            if (combinedCandidate is not null)
            {
                return combinedCandidate;
            }

            for (var offset = 0; offset < 5 && index + offset < textFragments.Count; offset++)
            {
                var candidate = NormalizeCounselorToken(textFragments[index + offset]);
                if (candidate is not null)
                {
                    return candidate;
                }
            }
        }

        return null;
    }

    private static string? NormalizeCounselorToken(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var cleaned = value
            .Replace("Beratungsperson", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace(":", string.Empty)
            .Replace(" ", string.Empty)
            .Trim();

        if (string.IsNullOrWhiteSpace(cleaned) || string.Equals(cleaned, "XX", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (!UpperTokenRegex.IsMatch(cleaned) || !CounselorInitialsRegex.IsMatch(cleaned))
        {
            return null;
        }

        return cleaned;
    }

    private static string? TryCombineCounselorFragments(IReadOnlyList<string> textFragments, int startIndex)
    {
        var parts = new List<string>(4);

        for (var index = startIndex; index < textFragments.Count && index < startIndex + 4; index++)
        {
            var normalized = NormalizeCounselorToken(textFragments[index]);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                break;
            }

            if (normalized.Length == 2)
            {
                return normalized;
            }

            parts.Add(normalized);
            if (parts.Count >= 2)
            {
                var combined = string.Concat(parts);
                return CounselorInitialsRegex.IsMatch(combined) ? combined : null;
            }
        }

        return null;
    }

    private static int ScoreHeader(string allText, string? odooUrl, string? counselorInitials)
    {
        var score = 0;

        if (allText.Contains("Verlaufsakte", StringComparison.OrdinalIgnoreCase))
        {
            score += 4;
        }

        if (allText.Contains("Abschlussbericht", StringComparison.OrdinalIgnoreCase))
        {
            score -= 2;
        }

        if (!string.IsNullOrWhiteSpace(odooUrl))
        {
            score += 5;
        }

        if (!string.IsNullOrWhiteSpace(counselorInitials))
        {
            score += 6;
        }

        if (PlaceholderHeaderRegex.IsMatch(allText))
        {
            score -= 8;
        }

        return score;
    }

    private static string CombineTargetAndAnchor(string target, string? anchor)
    {
        if (string.IsNullOrWhiteSpace(anchor))
        {
            return target;
        }

        return target.Contains('#', StringComparison.Ordinal)
            ? $"{target}{anchor}"
            : $"{target}#{anchor}";
    }

    private sealed record CacheEntry(
        DateTime LastWriteTimeUtc,
        long Length,
        DateTime LastSeenUtc,
        string HeaderSignature,
        HeaderMetadata Metadata);

    private sealed record HeaderCandidate(HeaderMetadata Metadata, int Score);
    private sealed record ParsedHeaderMetadata(HeaderMetadata Metadata, string HeaderSignature);
}
