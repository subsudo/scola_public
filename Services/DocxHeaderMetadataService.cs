using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp.Services;

public sealed class DocxHeaderMetadataService
{
    private const int CacheVersion = 1;
    private static readonly TimeSpan MissingDocumentRetention = TimeSpan.FromDays(30);
    private static readonly Regex PlaceholderHeaderRegex = new(@"Nachname|Vorname|Name\s+Vorname|\bXX\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

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
        }

        HeaderMetadata metadata;
        try
        {
            metadata = ReadFromPackage(documentPath);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Header-Metadaten konnten nicht gelesen werden '{documentPath}': {ex.Message}");
            metadata = HeaderMetadata.Empty;
        }

        UpdateCache(fileInfo, metadata);
        return metadata;
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
                    new HeaderMetadata(entry.OdooUrl ?? string.Empty)),
                StringComparer.OrdinalIgnoreCase);
    }

    private void UpdateCache(FileInfo fileInfo, HeaderMetadata metadata)
    {
        lock (_syncRoot)
        {
            var path = fileInfo.FullName;
            if (_cache.TryGetValue(path, out var existing) &&
                existing.LastWriteTimeUtc == fileInfo.LastWriteTimeUtc &&
                existing.Length == fileInfo.Length &&
                existing.Metadata == metadata)
            {
                TouchCacheEntryUnsafe(path, existing);
                return;
            }

            _cache[path] = new CacheEntry(fileInfo.LastWriteTimeUtc, fileInfo.Length, DateTime.UtcNow, metadata);
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
                        OdooUrl = entry.Value.Metadata.OdooUrl
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

    private static HeaderMetadata ReadFromPackage(string documentPath)
    {
        using var package = ZipFile.OpenRead(documentPath);

        var headerEntries = package.Entries
            .Where(entry => entry.FullName.StartsWith("word/header", StringComparison.OrdinalIgnoreCase)
                            && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .OrderBy(entry => entry.FullName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (headerEntries.Count == 0)
        {
            return HeaderMetadata.Empty;
        }

        var headers = new List<HeaderCandidate>();
        foreach (var headerEntry in headerEntries)
        {
            var relationshipsEntry = package.GetEntry($"word/_rels/{Path.GetFileName(headerEntry.FullName)}.rels");
            var relationships = relationshipsEntry is null
                ? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                : LoadRelationships(relationshipsEntry);

            headers.Add(ReadHeader(headerEntry, relationships));
        }

        var bestHeader = headers
            .OrderByDescending(header => header.Score)
            .ThenByDescending(header => !string.IsNullOrWhiteSpace(header.Metadata.OdooUrl))
            .First();

        var odooUrl = !string.IsNullOrWhiteSpace(bestHeader.Metadata.OdooUrl)
            ? bestHeader.Metadata.OdooUrl
            : headers.FirstOrDefault(header => !string.IsNullOrWhiteSpace(header.Metadata.OdooUrl))?.Metadata.OdooUrl ?? string.Empty;

        return new HeaderMetadata(odooUrl);
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

    private static HeaderCandidate ReadHeader(ZipArchiveEntry headerEntry, IReadOnlyDictionary<string, string> relationships)
    {
        using var stream = headerEntry.Open();
        var document = new XmlDocument();
        document.Load(stream);

        var namespaceManager = new XmlNamespaceManager(document.NameTable);
        namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        namespaceManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var textFragments = document.SelectNodes("//w:t", namespaceManager)?
            .Cast<XmlNode>()
            .Select(node => node.InnerText.Trim())
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToList() ?? new List<string>();

        var allText = Regex.Replace(string.Join(" ", textFragments), @"\s+", " ").Trim();

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

        var score = ScoreHeader(allText, odooUrl);
        return new HeaderCandidate(new HeaderMetadata(odooUrl ?? string.Empty), score);
    }

    private static int ScoreHeader(string allText, string? odooUrl)
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

    private sealed record CacheEntry(DateTime LastWriteTimeUtc, long Length, DateTime LastSeenUtc, HeaderMetadata Metadata);
    private sealed record HeaderCandidate(HeaderMetadata Metadata, int Score);
}
