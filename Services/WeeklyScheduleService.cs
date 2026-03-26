using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp.Services;

public sealed class WeeklyScheduleService
{
    private const int CacheVersion = 5;
    private const int MaxDisplayMatchesPerWeek = 10;

    private static readonly Regex TimeMarkerRegex = new(@"\b(08|09|10|11|13|14|15|16|17)[\.:]\s*\d{2}\b", RegexOptions.Compiled);
    private static readonly Regex RoomRegex = new(@"\b([UB])\s*(\d+)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex TokenRegex = new(@"[\p{L}\p{N}]+", RegexOptions.Compiled);
    private static readonly Regex WeekFileRegex = new(@"\bKW[_\s-]*(?<week>\d{1,2})(?:[_\s-]*(?<year>\d{2,4}))?\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private readonly object _syncRoot = new();
    private readonly string _cachePath;
    private readonly string _cacheBackupPath;
    private readonly Dictionary<string, WeeklyScheduleCacheEntryInternal> _cache;

    public WeeklyScheduleService(string cachePath, string cacheBackupPath)
    {
        _cachePath = cachePath;
        _cacheBackupPath = cacheBackupPath;
        _cache = LoadCache();
    }

    public ParticipantMiniScheduleSummary BuildSummary(
        string schedulePath,
        Participant participant,
        IReadOnlyCollection<Participant> participants,
        string? secondaryServerBasePath = null)
    {
        var participantRefs = BuildParticipantRefs(participants, secondaryServerBasePath);
        var participantRef = BuildParticipantRef(participant, secondaryServerBasePath);
        var resolvedSchedulePath = ResolveScheduleDocumentPath(schedulePath);
        if (string.IsNullOrWhiteSpace(resolvedSchedulePath))
        {
            return CreateUnavailableSummary("Kein Stundenplan");
        }

        var document = ReadDocument(resolvedSchedulePath);
        if (document.Slots.Count == 0)
        {
            return CreateUnavailableSummary("Kein Stundenplan");
        }

        var matcher = new ParticipantAliasMatcher(participantRefs);
        return BuildParticipantScheduleSummary(document, participantRef, matcher, out _, out _);
    }


    public void WriteDiagnostics(
        string schedulePath,
        IReadOnlyCollection<Participant> participants,
        string diagnosticsPath,
        string diagnosticsBackupPath,
        string? secondaryServerBasePath = null)
    {
        try
        {
            var participantRefs = BuildParticipantRefs(participants, secondaryServerBasePath);
            var resolvedSchedulePath = ResolveScheduleDocumentPath(schedulePath);
            if (string.IsNullOrWhiteSpace(resolvedSchedulePath))
            {
                JsonStorage.SaveAtomic(diagnosticsPath, diagnosticsBackupPath, new WeeklyScheduleDiagnosticsDocument
                {
                    GeneratedAt = DateTime.Now,
                    RequestedPath = schedulePath ?? string.Empty,
                    ResolvedPath = string.Empty,
                    Status = "Unavailable",
                    Message = "Kein Stundenplan"
                });
                return;
            }

            var document = ReadDocument(resolvedSchedulePath);
            var matcher = new ParticipantAliasMatcher(participantRefs);
            var diagnostics = new WeeklyScheduleDiagnosticsDocument
            {
                GeneratedAt = DateTime.Now,
                RequestedPath = schedulePath ?? string.Empty,
                ResolvedPath = resolvedSchedulePath,
                Status = document.Slots.Count == 0 ? "Unavailable" : "Ready",
                Message = document.Slots.Count == 0 ? "Kein Stundenplan" : string.Empty,
                Slots = document.Slots.Select(slot => new WeeklyScheduleSlotDiagnostics
                {
                    DayKey = slot.DayKey,
                    HalfDay = slot.HalfDay,
                    Blocks = slot.Blocks.Select(block => new WeeklyScheduleBlockDiagnostics
                    {
                        Group = block.Group,
                        Teacher = block.Teacher,
                        Room = block.Room,
                        ParticipantLines = block.ParticipantLines.Select(line => BuildLineDiagnostics(line, block.Group, matcher)).ToList()
                    }).ToList()
                }).ToList(),
                Participants = participantRefs.Select(participant =>
                {
                    var summary = BuildParticipantScheduleSummary(document, participant, matcher, out var ambiguousLines, out var matches);
                    return new WeeklyScheduleParticipantDiagnostics
                    {
                        ParticipantKey = participant.ParticipantKey,
                        DisplayName = participant.DisplayName,
                        ResultState = summary.State.ToString(),
                        Message = summary.Message,
                        Matches = matches,
                        AmbiguousLines = ambiguousLines
                    };
                }).OrderBy(entry => entry.DisplayName, StringComparer.OrdinalIgnoreCase).ToList()
            };

            JsonStorage.SaveAtomic(diagnosticsPath, diagnosticsBackupPath, diagnostics);
            AppLogger.Info($"Stundenplan-Diagnose aktualisiert: {diagnosticsPath}");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Stundenplan-Diagnose konnte nicht geschrieben werden: {ex.Message}");
        }
    }

    private static List<ScheduleParticipantRef> BuildParticipantRefs(
        IReadOnlyCollection<Participant> participants,
        string? secondaryServerBasePath)
    {
        return participants
            .Where(participant => !string.IsNullOrWhiteSpace(participant.FullName))
            .Select(participant => BuildParticipantRef(participant, secondaryServerBasePath))
            .ToList();
    }

    private static ScheduleParticipantRef BuildParticipantRef(Participant participant, string? secondaryServerBasePath)
    {
        return new ScheduleParticipantRef
        {
            ParticipantKey = ResolveParticipantKey(participant),
            DisplayName = participant.FullName,
            StatusTag = ResolveParticipantStatusTag(participant, secondaryServerBasePath)
        };
    }

    private static string ResolveParticipantKey(Participant participant)
    {
        if (!string.IsNullOrWhiteSpace(participant.DocumentPath))
        {
            return participant.DocumentPath;
        }

        if (!string.IsNullOrWhiteSpace(participant.SelectedFolderPath))
        {
            return participant.SelectedFolderPath;
        }

        if (!string.IsNullOrWhiteSpace(participant.MatchedFolderPath))
        {
            return participant.MatchedFolderPath;
        }

        return participant.FullName;
    }

    private static string ResolveParticipantStatusTag(Participant participant, string? secondaryServerBasePath)
    {
        if (string.IsNullOrWhiteSpace(secondaryServerBasePath))
        {
            return string.Empty;
        }

        var effectivePath = participant.SelectedFolderPath
            ?? participant.MatchedFolderPath
            ?? string.Empty;

        if (string.IsNullOrWhiteSpace(effectivePath))
        {
            return string.Empty;
        }

        return effectivePath.StartsWith(secondaryServerBasePath, StringComparison.OrdinalIgnoreCase)
            ? "LB"
            : string.Empty;
    }

    private static ParticipantMiniScheduleSummary CreateUnavailableSummary(string message)
    {
        return new ParticipantMiniScheduleSummary
        {
            State = ParticipantMiniScheduleState.Unavailable,
            Message = message
        };
    }

    private static ParticipantMiniScheduleSummary BuildParticipantScheduleSummary(
        WeeklyScheduleDocument document,
        ScheduleParticipantRef participant,
        ParticipantAliasMatcher matcher,
        out List<string> ambiguousLines,
        out List<WeeklyScheduleParticipantSlotDiagnostics> matches)
    {
        var summary = new ParticipantMiniScheduleSummary
        {
            State = ParticipantMiniScheduleState.Ready,
            Cells = ParticipantMiniScheduleSummary.CreateDefaultCells().ToList()
        };

        ambiguousLines = new List<string>();
        matches = new List<WeeklyScheduleParticipantSlotDiagnostics>();
        var hasEntries = false;
        var isAmbiguous = false;

        foreach (var slot in document.Slots)
        {
            var cell = summary.GetCell(slot.DayKey, ParseHalfDay(slot.HalfDay));
            foreach (var block in slot.Blocks)
            {
                var blockMatchesParticipant = false;
                var blockExternal = false;
                var blockAbsent = false;
                var blockSupplemental = IsDazGroup(block.Group);

                foreach (var line in block.ParticipantLines)
                {
                    if (matcher.TryResolveLine(line, block.Group, out var resolvedMatches))
                    {
                        foreach (var match in resolvedMatches)
                        {
                            if (!string.Equals(match.ParticipantKey, participant.ParticipantKey, StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            blockMatchesParticipant = true;
                            blockAbsent |= match.IsAbsent;
                            blockExternal |= match.IsExternal;
                            // Die Farbmarkierung selbst macht einen Hauptblock nicht zu DAZ.
                            // Zusatzlogik gilt nur fuer echte DAZ-Bloecke.
                            blockSupplemental |= false;
                        }
                    }
                    else if (matcher.LineCouldReferToParticipant(line, block.Group, participant.ParticipantKey))
                    {
                        isAmbiguous = true;
                        ambiguousLines.Add(line.RawText);
                    }
                }

                if (!blockMatchesParticipant)
                {
                    continue;
                }

                hasEntries = true;

                if (blockAbsent)
                {
                    cell.Status = ParticipantMiniScheduleCellStatus.Dispensed;
                    cell.HasSupplementalDaz = false;
                    cell.Entries.Clear();
                    matches.Add(new WeeklyScheduleParticipantSlotDiagnostics
                    {
                        DayKey = slot.DayKey,
                        HalfDay = slot.HalfDay,
                        Group = "disp"
                    });
                    continue;
                }

                if (blockExternal)
                {
                    cell.Status = ParticipantMiniScheduleCellStatus.External;
                    cell.HasSupplementalDaz = false;
                    cell.Entries.Clear();
                    matches.Add(new WeeklyScheduleParticipantSlotDiagnostics
                    {
                        DayKey = slot.DayKey,
                        HalfDay = slot.HalfDay,
                        Group = "ext",
                        IsExternal = true
                    });
                    continue;
                }

                if (cell.Status != ParticipantMiniScheduleCellStatus.None)
                {
                    continue;
                }

                if (blockSupplemental)
                {
                    cell.HasSupplementalDaz = true;
                    continue;
                }

                if (cell.Entries.Any(existing =>
                        string.Equals(existing.Group, block.Group, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(existing.Teacher, block.Teacher, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(existing.Room, block.Room, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                cell.Entries.Add(new ParticipantMiniScheduleEntry
                {
                    Group = block.Group,
                    Teacher = block.Teacher,
                    Room = block.Room,
                    IsExternal = blockExternal
                });
                matches.Add(new WeeklyScheduleParticipantSlotDiagnostics
                {
                    DayKey = slot.DayKey,
                    HalfDay = slot.HalfDay,
                    Group = block.Group,
                    Teacher = block.Teacher,
                    Room = block.Room,
                    IsExternal = blockExternal
                });
            }
        }

        var nonSupplementalMatches = matches
            .Where(match => !IsDazGroup(match.Group))
            .ToList();

        var hasMultipleMatchesPerHalfDay = nonSupplementalMatches
            .GroupBy(match => $"{match.DayKey}|{match.HalfDay}", StringComparer.OrdinalIgnoreCase)
            .Any(group => group.Count() > 1);

        if (hasMultipleMatchesPerHalfDay)
        {
            ambiguousLines = new List<string>();
            matches = new List<WeeklyScheduleParticipantSlotDiagnostics>();
            return new ParticipantMiniScheduleSummary
            {
                State = ParticipantMiniScheduleState.Unavailable,
                Message = "Nicht eindeutig zugeordnet",
                Cells = ParticipantMiniScheduleSummary.CreateDefaultCells().ToList()
            };
        }

        if (nonSupplementalMatches.Count > MaxDisplayMatchesPerWeek)
        {
            ambiguousLines = new List<string>();
            matches = new List<WeeklyScheduleParticipantSlotDiagnostics>();
            return new ParticipantMiniScheduleSummary
            {
                State = ParticipantMiniScheduleState.Unavailable,
                Message = "Nicht eindeutig zugeordnet",
                Cells = ParticipantMiniScheduleSummary.CreateDefaultCells().ToList()
            };
        }

        if (!hasEntries)
        {
            summary.State = ParticipantMiniScheduleState.Unavailable;
            summary.Message = isAmbiguous ? "Nicht eindeutig zugeordnet" : "Kein Stundenplan";
            return summary;
        }

        if (isAmbiguous)
        {
            summary.Message = "Nicht eindeutig zugeordnet";
        }

        return summary;
    }

    private static WeeklyScheduleLineDiagnostics BuildLineDiagnostics(WeeklyScheduleParticipantLine line, string group, ParticipantAliasMatcher matcher)
    {
        var analysis = matcher.AnalyzeLine(line, group);
        return new WeeklyScheduleLineDiagnostics
        {
            RawText = line.RawText,
            Tokens = line.Tokens.Select(token => new WeeklyScheduleTokenDiagnostics
            {
                Text = token.Text,
                IsAbsent = token.IsAbsent,
                IsExternal = token.IsExternal,
                IsSupplemental = token.IsSupplemental
            }).ToList(),
            ResolutionState = analysis.ResolutionState,
            ResolvedParticipants = analysis.ResolvedMatches.Select(match => new WeeklyScheduleResolvedParticipantDiagnostics
            {
                ParticipantKey = match.ParticipantKey,
                DisplayName = matcher.GetDisplayName(match.ParticipantKey),
                IsAbsent = match.IsAbsent,
                IsExternal = match.IsExternal
            }).ToList(),
            CandidateMatches = analysis.Candidates.Select(candidate => new WeeklyScheduleCandidateDiagnostics
            {
                Alias = candidate.Alias,
                StartIndex = candidate.StartIndex,
                Length = candidate.Length,
                Candidates = candidate.ParticipantKeys.Select(key => new WeeklyScheduleResolvedParticipantDiagnostics
                {
                    ParticipantKey = key,
                    DisplayName = matcher.GetDisplayName(key)
                }).ToList()
            }).ToList()
        };
    }
    private static string? ResolveScheduleDocumentPath(string schedulePath)
    {
        if (string.IsNullOrWhiteSpace(schedulePath))
        {
            return null;
        }

        if (File.Exists(schedulePath))
        {
            return schedulePath;
        }

        if (!Directory.Exists(schedulePath))
        {
            return null;
        }

        var candidates = Directory.GetFiles(schedulePath, "*.docx", SearchOption.TopDirectoryOnly)
            .Select(ParseWeekCandidate)
            .Where(candidate => candidate is not null)
            .Cast<WeekFileCandidate>()
            .ToList();

        if (candidates.Count == 0)
        {
            return null;
        }

        var today = DateTime.Today;
        var currentWeek = ISOWeek.GetWeekOfYear(today);
        var currentYear = ISOWeek.GetYear(today);

        return candidates
            .Where(candidate => candidate.Week == currentWeek && candidate.Year == currentYear)
            .OrderBy(candidate => candidate.Path, StringComparer.OrdinalIgnoreCase)
            .Select(candidate => candidate.Path)
            .FirstOrDefault();
    }

    private static WeekFileCandidate? ParseWeekCandidate(string path)
    {
        var match = WeekFileRegex.Match(Path.GetFileNameWithoutExtension(path));
        if (!match.Success || !int.TryParse(match.Groups["week"].Value, out var week))
        {
            return null;
        }

        var year = ISOWeek.GetYear(DateTime.Today);
        if (match.Groups["year"].Success && int.TryParse(match.Groups["year"].Value, out var rawYear))
        {
            year = rawYear < 100 ? 2000 + rawYear : rawYear;
        }

        return new WeekFileCandidate(path, year, week);
    }
    private WeeklyScheduleDocument ReadDocument(string schedulePath)
    {
        var fileInfo = new FileInfo(schedulePath);
        if (!fileInfo.Exists)
        {
            return new WeeklyScheduleDocument();
        }

        lock (_syncRoot)
        {
            if (_cache.TryGetValue(fileInfo.FullName, out var cached) &&
                cached.LastWriteTimeUtc == fileInfo.LastWriteTimeUtc &&
                cached.Length == fileInfo.Length)
            {
                return cached.Document;
            }
        }

        WeeklyScheduleDocument document;
        try
        {
            document = ParseScheduleDocument(fileInfo.FullName);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Stundenplan konnte nicht gelesen werden '{schedulePath}': {ex.Message}");
            document = new WeeklyScheduleDocument();
        }

        UpdateCache(fileInfo, document);
        return document;
    }

    private Dictionary<string, WeeklyScheduleCacheEntryInternal> LoadCache()
    {
        var document = JsonStorage.Load(_cachePath, _cacheBackupPath, () => new WeeklyScheduleCacheDocument());
        if (document.Version != CacheVersion)
        {
            return new Dictionary<string, WeeklyScheduleCacheEntryInternal>(StringComparer.OrdinalIgnoreCase);
        }

        return document.Entries
            .Where(entry => !string.IsNullOrWhiteSpace(entry.DocumentPath))
            .ToDictionary(
                entry => entry.DocumentPath,
                entry => new WeeklyScheduleCacheEntryInternal(entry.LastWriteTimeUtc, entry.Length, entry.Document ?? new WeeklyScheduleDocument()),
                StringComparer.OrdinalIgnoreCase);
    }

    private void UpdateCache(FileInfo fileInfo, WeeklyScheduleDocument document)
    {
        lock (_syncRoot)
        {
            _cache[fileInfo.FullName] = new WeeklyScheduleCacheEntryInternal(fileInfo.LastWriteTimeUtc, fileInfo.Length, document);
            PersistCacheUnsafe();
        }
    }

    private void PersistCacheUnsafe()
    {
        try
        {
            var document = new WeeklyScheduleCacheDocument
            {
                Version = CacheVersion,
                Entries = _cache
                    .OrderBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(entry => new WeeklyScheduleCacheEntry
                    {
                        DocumentPath = entry.Key,
                        LastWriteTimeUtc = entry.Value.LastWriteTimeUtc,
                        Length = entry.Value.Length,
                        Document = entry.Value.Document
                    })
                    .ToList()
            };

            JsonStorage.SaveAtomic(_cachePath, _cacheBackupPath, document);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Stundenplan-Cache konnte nicht gespeichert werden: {ex.Message}");
        }
    }

    private static WeeklyScheduleDocument ParseScheduleDocument(string schedulePath)
    {
        using var archive = ZipFile.OpenRead(schedulePath);
        var documentEntry = archive.GetEntry("word/document.xml") ?? throw new InvalidOperationException("word/document.xml fehlt.");
        var xmlDocument = LoadXml(documentEntry);

        var namespaceManager = new XmlNamespaceManager(xmlDocument.NameTable);
        namespaceManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        var tableNode = xmlDocument.SelectSingleNode("//w:tbl", namespaceManager)
            ?? throw new InvalidOperationException("Keine Tabelle im Stundenplan gefunden.");

        var rows = tableNode.SelectNodes("./w:tr", namespaceManager)?.Cast<XmlNode>()
            .Select(row => ParseRow(row, namespaceManager))
            .ToList() ?? new List<ParsedRow>();

        if (rows.Count == 0)
        {
            return new WeeklyScheduleDocument();
        }

        var dayRanges = ParseDayRanges(rows[0]);
        var slots = new List<WeeklyScheduleSlot>();
        string? currentHalfDay = null;

        foreach (var row in rows.Skip(1))
        {
            var firstCell = row.Cells.FirstOrDefault(cell => cell.StartColumn == 0);
            if (firstCell is not null)
            {
                var rowHalfDay = DetermineHalfDay(firstCell.FullText);
                if (rowHalfDay == "EV")
                {
                    currentHalfDay = null;
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(rowHalfDay))
                {
                    currentHalfDay = rowHalfDay;
                }
            }

            if (string.IsNullOrWhiteSpace(currentHalfDay))
            {
                continue;
            }

            foreach (var cell in row.Cells.Where(cell => cell.StartColumn > 0 && cell.Paragraphs.Count > 0))
            {
                if (string.IsNullOrWhiteSpace(cell.FullText))
                {
                    continue;
                }

                var dayKey = ResolveDayKey(dayRanges, cell.StartColumn, cell.Span);
                if (string.IsNullOrWhiteSpace(dayKey))
                {
                    continue;
                }

                var blocks = ParseBlocks(cell.Paragraphs);
                if (blocks.Count == 0)
                {
                    continue;
                }

                var slot = slots.FirstOrDefault(candidate =>
                    string.Equals(candidate.DayKey, dayKey, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(candidate.HalfDay, currentHalfDay, StringComparison.OrdinalIgnoreCase));
                if (slot is null)
                {
                    slot = new WeeklyScheduleSlot { DayKey = dayKey, HalfDay = currentHalfDay };
                    slots.Add(slot);
                }

                slot.Blocks.AddRange(blocks);
            }
        }

        return new WeeklyScheduleDocument { Slots = slots };
    }

    private static XmlDocument LoadXml(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        var document = new XmlDocument();
        document.Load(stream);
        return document;
    }

    private static ParsedRow ParseRow(XmlNode rowNode, XmlNamespaceManager namespaceManager)
    {
        var row = new ParsedRow();
        var currentColumn = 0;
        foreach (var cellNode in rowNode.SelectNodes("./w:tc", namespaceManager)?.Cast<XmlNode>() ?? Enumerable.Empty<XmlNode>())
        {
            var span = 1;
            var spanNode = cellNode.SelectSingleNode("./w:tcPr/w:gridSpan", namespaceManager);
            if (spanNode?.Attributes is { } spanAttributes)
            {
                span = ParseIntAttribute(spanAttributes, "w:val") ?? ParseIntAttribute(spanAttributes, "val") ?? 1;
            }

            var paragraphs = cellNode.SelectNodes("./w:p", namespaceManager)?.Cast<XmlNode>()
                .Select(paragraphNode => ParseParagraph(paragraphNode, namespaceManager))
                .Where(paragraph => !string.IsNullOrWhiteSpace(paragraph.Text))
                .ToList() ?? new List<ParsedParagraph>();

            row.Cells.Add(new ParsedCell
            {
                StartColumn = currentColumn,
                Span = span,
                Paragraphs = paragraphs,
                FullText = string.Join(" ", paragraphs.Select(paragraph => paragraph.Text)).Trim()
            });
            currentColumn += span;
        }

        return row;
    }

    private static ParsedParagraph ParseParagraph(XmlNode paragraphNode, XmlNamespaceManager namespaceManager)
    {
        var runs = paragraphNode.SelectNodes(".//w:r", namespaceManager)?.Cast<XmlNode>().ToList() ?? new List<XmlNode>();
        var textFragments = new List<string>();
        var tokens = new List<WeeklyScheduleToken>();

        foreach (var run in runs)
        {
            var text = string.Join(" ", run.SelectNodes(".//w:t", namespaceManager)?.Cast<XmlNode>().Select(node => node.InnerText) ?? Array.Empty<string>()).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            textFragments.Add(text);

            var isAbsent = HasFormatValue(run, namespaceManager, "./w:rPr/w:highlight", "red")
                           || HasFormatValue(run, namespaceManager, "./w:rPr/w:shd", "FF0000")
                           || HasFormatValue(run, namespaceManager, "./w:rPr/w:shd", "EE0000");
            var isExternal = HasFormatValue(run, namespaceManager, "./w:rPr/w:highlight", "green");
            var isSupplemental = HasFormatValue(run, namespaceManager, "./w:rPr/w:highlight", "cyan") || HasFormatValue(run, namespaceManager, "./w:rPr/w:color", "00B0F0");

            foreach (Match match in TokenRegex.Matches(text))
            {
                tokens.Add(new WeeklyScheduleToken
                {
                    Text = match.Value,
                    IsAbsent = isAbsent,
                    IsExternal = isExternal,
                    IsSupplemental = isSupplemental
                });
            }
        }

        return new ParsedParagraph
        {
            Text = Regex.Replace(string.Join(" ", textFragments), @"\s+", " ").Trim(),
            Tokens = tokens
        };
    }

    private static bool HasFormatValue(XmlNode run, XmlNamespaceManager namespaceManager, string xpath, string expectedValue)
    {
        var node = run.SelectSingleNode(xpath, namespaceManager);
        if (node?.Attributes is null)
        {
            return false;
        }

        var value = node.Attributes["w:val"]?.Value
            ?? node.Attributes["val"]?.Value
            ?? node.Attributes["w:fill"]?.Value
            ?? node.Attributes["fill"]?.Value
            ?? string.Empty;

        return string.Equals(value, expectedValue, StringComparison.OrdinalIgnoreCase);
    }

    private static List<DayRange> ParseDayRanges(ParsedRow headerRow)
    {
        return headerRow.Cells
            .Where(cell => cell.StartColumn > 0 && !string.IsNullOrWhiteSpace(cell.FullText))
            .Select(cell => new DayRange
            {
                DayKey = NormalizeDayKey(cell.FullText),
                StartColumn = cell.StartColumn,
                EndColumnExclusive = cell.StartColumn + cell.Span
            })
            .Where(range => !string.IsNullOrWhiteSpace(range.DayKey))
            .ToList();
    }

    private static string NormalizeDayKey(string text)
    {
        var normalized = text.Trim().ToLowerInvariant();
        if (normalized.Contains("montag")) return "Mo";
        if (normalized.Contains("dienstag")) return "Di";
        if (normalized.Contains("mittwoch")) return "Mi";
        if (normalized.Contains("donnerstag")) return "Do";
        if (normalized.Contains("freitag")) return "Fr";
        return string.Empty;
    }

    private static string ResolveDayKey(IReadOnlyList<DayRange> dayRanges, int startColumn, int span)
    {
        var midpoint = startColumn + (span / 2.0);
        return dayRanges.FirstOrDefault(range => midpoint >= range.StartColumn && midpoint < range.EndColumnExclusive)?.DayKey ?? string.Empty;
    }

    private static string? DetermineHalfDay(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        var normalized = text.Replace(" ", string.Empty, StringComparison.OrdinalIgnoreCase);
        if (normalized.Contains("17.00") || normalized.Contains("17:00"))
        {
            return "EV";
        }

        if (normalized.Contains("08.45") || normalized.Contains("08:45"))
        {
            return "VM";
        }

        if (normalized.Contains("13.15") || normalized.Contains("13:15"))
        {
            return "NM";
        }

        return TimeMarkerRegex.IsMatch(text) ? "VM" : null;
    }

    private static List<WeeklyScheduleBlock> ParseBlocks(IReadOnlyList<ParsedParagraph> paragraphs)
    {
        var result = new List<WeeklyScheduleBlock>();
        var index = 0;

        while (index < paragraphs.Count)
        {
            var paragraph = paragraphs[index];
            if (!TryNormalizeGroup(paragraph.Text, out var group))
            {
                index++;
                continue;
            }

            index++;
            string teacher = string.Empty;
            string room = string.Empty;

            while (index < paragraphs.Count && !TryNormalizeGroup(paragraphs[index].Text, out _))
            {
                var candidate = paragraphs[index];
                if (IsTimeParagraph(candidate.Text) || IsAdministrativeParagraph(candidate.Text))
                {
                    index++;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(teacher) && IsTeacherParagraph(candidate.Text))
                {
                    teacher = NormalizeTeacher(candidate.Text);
                    index++;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(room) && IsRoomParagraph(candidate.Text))
                {
                    room = NormalizeRoom(candidate.Text);
                    index++;
                    continue;
                }

                break;
            }

            var block = new WeeklyScheduleBlock
            {
                Group = group,
                Teacher = teacher,
                Room = room
            };

            while (index < paragraphs.Count && !TryNormalizeGroup(paragraphs[index].Text, out _))
            {
                var candidate = paragraphs[index];
                if (!IsAdministrativeParagraph(candidate.Text))
                {
                    var line = CreateParticipantLine(candidate);
                    if (line.Tokens.Count > 0)
                    {
                        block.ParticipantLines.Add(line);
                    }
                }

                index++;
            }

            if (!IsAdministrativeBlock(block))
            {
                result.Add(block);
            }
        }

        return result;
    }

    private static bool IsAdministrativeBlock(WeeklyScheduleBlock block)
    {
        return string.IsNullOrWhiteSpace(block.Group)
               || block.Group.Contains("ADMIN", StringComparison.OrdinalIgnoreCase)
               || string.Equals(block.Group, "LB ABEND", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsDazGroup(string group)
    {
        return group.StartsWith("DAZ", StringComparison.OrdinalIgnoreCase);
    }
    private static WeeklyScheduleParticipantLine CreateParticipantLine(ParsedParagraph paragraph)
    {
        return new WeeklyScheduleParticipantLine
        {
            RawText = paragraph.Text,
            Tokens = paragraph.Tokens
                .Select(token => new WeeklyScheduleToken
                {
                    Text = token.Text,
                    IsAbsent = token.IsAbsent,
                    IsExternal = token.IsExternal,
                    IsSupplemental = token.IsSupplemental
                })
                .ToList()
        };
    }

    private static bool TryNormalizeGroup(string text, out string group)
    {
        var normalized = NormalizePhrase(text).ToUpperInvariant();
        group = string.Empty;
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return false;
        }

        if (normalized.StartsWith("DAZ", StringComparison.OrdinalIgnoreCase))
        {
            group = normalized.Replace(".", string.Empty, StringComparison.Ordinal);
            return true;
        }

        if (normalized is "BI" or "BU" or "MO" or "LB" or "PR" or "WIT" or "KONV")
        {
            group = normalized;
            return true;
        }

        if (normalized is "IND" or "IND.")
        {
            group = "IND";
            return true;
        }

        if (normalized.Replace(" ", string.Empty, StringComparison.Ordinal) == "BULV")
        {
            group = "BU LV";
            return true;
        }

        if (normalized.Replace(" ", string.Empty, StringComparison.Ordinal) == "LBABEND")
        {
            group = "LB ABEND";
            return true;
        }

        return false;
    }

    private static bool IsTeacherParagraph(string text)
    {
        if (string.IsNullOrWhiteSpace(text) || IsRoomParagraph(text) || IsTimeParagraph(text))
        {
            return false;
        }

        var normalized = NormalizePhrase(text).ToUpperInvariant();
        if (normalized.Contains("ADMIN", StringComparison.OrdinalIgnoreCase) || normalized.Contains(":", StringComparison.Ordinal))
        {
            return false;
        }

        var tokens = TokenizeLooseNormalized(normalized);
        return tokens.Count is >= 1 and <= 3 && tokens.All(token => token.Length <= 2 && token.All(char.IsLetter));
    }

    private static string NormalizeTeacher(string text)
    {
        return string.Concat(TokenizeLooseNormalized(NormalizePhrase(text))).ToUpperInvariant();
    }

    private static bool IsRoomParagraph(string text)
    {
        return RoomRegex.IsMatch(text);
    }

    private static string NormalizeRoom(string text)
    {
        var match = RoomRegex.Match(text);
        if (!match.Success)
        {
            return NormalizePhrase(text);
        }

        return $"{match.Groups[1].Value.ToUpperInvariant()}{match.Groups[2].Value}";
    }

    private static bool IsTimeParagraph(string text)
    {
        return TimeMarkerRegex.IsMatch(text);
    }

    private static bool IsAdministrativeParagraph(string text)
    {
        var normalized = NormalizePhrase(text).ToUpperInvariant();
        return normalized.Contains("ADMIN", StringComparison.OrdinalIgnoreCase)
               || normalized.Replace(" ", string.Empty, StringComparison.Ordinal) == "VZGL";
    }

    private static string NormalizePhrase(string text)
    {
        return Regex.Replace(text ?? string.Empty, @"\s+", " ").Trim();
    }

    private static List<string> TokenizeNormalized(string value)
    {
        return TokenRegex.Matches(SearchTextUtility.ReplaceUmlauts(value).ToLowerInvariant())
            .Select(match => match.Value)
            .Where(SearchTextUtility.IsRobustToken)
            .ToList();
    }

    private static List<string> TokenizeLooseNormalized(string value)
    {
        return TokenRegex.Matches(SearchTextUtility.ReplaceUmlauts(value).ToLowerInvariant())
            .Select(match => match.Value)
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .ToList();
    }

    private static ParticipantMiniScheduleHalfDay ParseHalfDay(string value)
        => string.Equals(value, "NM", StringComparison.OrdinalIgnoreCase)
            ? ParticipantMiniScheduleHalfDay.Afternoon
            : ParticipantMiniScheduleHalfDay.Morning;

    private static int? ParseIntAttribute(XmlAttributeCollection attributes, string name)
    {
        return int.TryParse(attributes[name]?.Value, out var value) ? value : null;
    }

    private sealed record WeeklyScheduleCacheEntryInternal(DateTime LastWriteTimeUtc, long Length, WeeklyScheduleDocument Document);
    private sealed record WeekFileCandidate(string Path, int Year, int Week);

    private sealed class ParticipantAliasMatcher
    {
        private readonly Dictionary<string, List<string>> _aliasMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, HashSet<string>> _participantAliases = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _displayNames = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _statusTags = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _globallyUniqueSingleAliases;

        public ParticipantAliasMatcher(IReadOnlyList<ScheduleParticipantRef> participants)
        {
            _globallyUniqueSingleAliases = BuildGloballyUniqueSingleAliases(participants);

            foreach (var participant in participants.Where(entry => !string.IsNullOrWhiteSpace(entry.DisplayName)))
            {
                _displayNames[participant.ParticipantKey] = participant.DisplayName;
                _statusTags[participant.ParticipantKey] = participant.StatusTag ?? string.Empty;

                var aliases = BuildAliases(participant.DisplayName, _globallyUniqueSingleAliases);
                if (aliases.Count == 0)
                {
                    continue;
                }

                _participantAliases[participant.ParticipantKey] = aliases;
                foreach (var alias in aliases)
                {
                    if (!_aliasMap.TryGetValue(alias, out var keys))
                    {
                        keys = new List<string>();
                        _aliasMap[alias] = keys;
                    }

                    if (!keys.Contains(participant.ParticipantKey, StringComparer.OrdinalIgnoreCase))
                    {
                        keys.Add(participant.ParticipantKey);
                    }
                }
            }
        }

        public string GetDisplayName(string participantKey)
            => _displayNames.TryGetValue(participantKey, out var displayName) ? displayName : participantKey;

        public bool TryResolveLine(WeeklyScheduleParticipantLine line, string group, out List<ResolvedLineMatch> matches)
        {
            matches = new List<ResolvedLineMatch>();
            if (line.Tokens.Count == 0)
            {
                return false;
            }

            var resolved = Resolve(line.Tokens, group, 0, new Dictionary<int, ResolutionPath?>());
            if (resolved is null || resolved.HasAmbiguity || resolved.Matches.Count == 0)
            {
                return false;
            }

            matches = resolved.Matches
                .GroupBy(match => match.ParticipantKey, StringComparer.OrdinalIgnoreCase)
                .Select(grouped => grouped.Aggregate((current, next) => new ResolvedLineMatch(
                    current.ParticipantKey,
                    current.IsAbsent || next.IsAbsent,
                    current.IsExternal || next.IsExternal,
                    current.IsSupplemental || next.IsSupplemental)))
                .ToList();
            return true;
        }

        public LineAnalysis AnalyzeLine(WeeklyScheduleParticipantLine line, string group)
        {
            var candidates = CollectCandidates(line.Tokens, group);
            var resolved = Resolve(line.Tokens, group, 0, new Dictionary<int, ResolutionPath?>());
            if (resolved is not null && resolved.Matches.Count > 0 && !resolved.HasAmbiguity)
            {
                var distinctMatches = resolved.Matches
                    .GroupBy(match => match.ParticipantKey, StringComparer.OrdinalIgnoreCase)
                    .Select(grouped => grouped.Aggregate((current, next) => new ResolvedLineMatch(
                        current.ParticipantKey,
                        current.IsAbsent || next.IsAbsent,
                        current.IsExternal || next.IsExternal,
                        current.IsSupplemental || next.IsSupplemental)))
                    .ToList();
                return new LineAnalysis("Resolved", distinctMatches, candidates);
            }

            return new LineAnalysis(candidates.Count > 0 ? "Ambiguous" : "Unmatched", new List<ResolvedLineMatch>(), candidates);
        }

        public bool LineCouldReferToParticipant(WeeklyScheduleParticipantLine line, string group, string participantKey)
        {
            if (!_participantAliases.TryGetValue(participantKey, out var aliases) || line.Tokens.Count == 0)
            {
                return false;
            }

            for (var start = 0; start < line.Tokens.Count; start++)
            {
                for (var length = 1; length <= 3 && start + length <= line.Tokens.Count; length++)
                {
                    var alias = string.Join(" ", line.Tokens.Skip(start).Take(length).Select(token => NormalizeToken(token.Text)));
                    if (!aliases.Contains(alias))
                    {
                        continue;
                    }

                    if (IsParticipantEligibleForGroup(participantKey, group))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private List<LineCandidate> CollectCandidates(IReadOnlyList<WeeklyScheduleToken> tokens, string group)
        {
            var result = new List<LineCandidate>();
            for (var start = 0; start < tokens.Count; start++)
            {
                for (var length = 1; length <= 3 && start + length <= tokens.Count; length++)
                {
                    var alias = string.Join(" ", tokens.Skip(start).Take(length).Select(token => NormalizeToken(token.Text)));
                    if (!_aliasMap.TryGetValue(alias, out var keys))
                    {
                        continue;
                    }

                    var filteredKeys = FilterParticipantKeysForGroup(keys, group);
                    if (filteredKeys.Count == 0)
                    {
                        continue;
                    }

                    result.Add(new LineCandidate(alias, start, length, filteredKeys));
                }
            }

            return result
                .OrderBy(candidate => candidate.StartIndex)
                .ThenByDescending(candidate => candidate.Length)
                .ThenBy(candidate => candidate.Alias, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private ResolutionPath? Resolve(
            IReadOnlyList<WeeklyScheduleToken> tokens,
            string group,
            int index,
            Dictionary<int, ResolutionPath?> memo)
        {
            if (memo.TryGetValue(index, out var cached))
            {
                return cached;
            }

            if (index >= tokens.Count)
            {
                return memo[index] = new ResolutionPath(new List<ResolvedLineMatch>(), 0, false);
            }

            var candidatesStartingHere = CollectCandidates(tokens, group)
                .Where(candidate => candidate.StartIndex == index)
                .ToList();
            ResolutionPath? best = null;

            foreach (var candidate in candidatesStartingHere
                         .Where(candidate => candidate.ParticipantKeys.Count == 1)
                         .OrderByDescending(candidate => candidate.Length))
            {
                var tail = Resolve(tokens, group, index + candidate.Length, memo);
                if (tail is null)
                {
                    continue;
                }

                var segmentTokens = tokens.Skip(index).Take(candidate.Length).ToList();
                var match = new ResolvedLineMatch(
                    candidate.ParticipantKeys[0],
                    segmentTokens.Any(token => token.IsAbsent),
                    segmentTokens.Any(token => token.IsExternal),
                    segmentTokens.Any(token => token.IsSupplemental));

                var matches = new List<ResolvedLineMatch> { match };
                matches.AddRange(tail.Matches);
                var path = new ResolutionPath(matches, tail.SkippedTokens, tail.HasAmbiguity);
                best = ChooseBetter(best, path);
            }

            var canSkipCurrentToken = IsIgnorableToken(tokens[index]) || candidatesStartingHere.Count == 0;
            if (canSkipCurrentToken)
            {
                var tail = Resolve(tokens, group, index + 1, memo);
                if (tail is not null)
                {
                    var skipPath = new ResolutionPath(
                        tail.Matches,
                        tail.SkippedTokens + 1,
                        tail.HasAmbiguity || candidatesStartingHere.Any(candidate => candidate.ParticipantKeys.Count > 1));
                    best = ChooseBetter(best, skipPath);
                }
            }
            else if (candidatesStartingHere.Any(candidate => candidate.ParticipantKeys.Count > 1))
            {
                best = ChooseBetter(best, new ResolutionPath(new List<ResolvedLineMatch>(), 0, true));
            }

            return memo[index] = best;
        }

        private static ResolutionPath? ChooseBetter(ResolutionPath? current, ResolutionPath? candidate)
        {
            if (candidate is null)
            {
                return current;
            }

            if (current is null)
            {
                return candidate;
            }

            if (candidate.Matches.Count != current.Matches.Count)
            {
                return candidate.Matches.Count > current.Matches.Count ? candidate : current;
            }

            if (candidate.HasAmbiguity != current.HasAmbiguity)
            {
                return current.HasAmbiguity ? candidate : current;
            }

            if (candidate.SkippedTokens != current.SkippedTokens)
            {
                return candidate.SkippedTokens < current.SkippedTokens ? candidate : current;
            }

            return current;
        }

        private static HashSet<string> BuildAliases(string displayName, HashSet<string> globallyUniqueSingleAliases)
        {
            var tokens = TokenizeLooseNormalized(displayName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var primaryToken = tokens.FirstOrDefault(CanLeadAlias);
            if (!string.IsNullOrWhiteSpace(primaryToken) && globallyUniqueSingleAliases.Contains(primaryToken))
            {
                result.Add(primaryToken);
            }

            // Robust single tokens stay available as candidates so the resolver can
            // still make a safe decision after group-based filtering (e.g. LB vs. LV).
            foreach (var token in tokens.Where(token => token.Length >= 3))
            {
                result.Add(token);
            }

            if (tokens.Count >= 2)
            {
                result.Add(string.Join(" ", tokens));
            }

            for (var i = 0; i < tokens.Count; i++)
            {
                if (!CanLeadAlias(tokens[i]))
                {
                    continue;
                }

                for (var j = 0; j < tokens.Count; j++)
                {
                    if (i == j || string.IsNullOrWhiteSpace(tokens[j]))
                    {
                        continue;
                    }

                    result.Add($"{tokens[i]} {tokens[j][0]}");
                    if (SearchTextUtility.IsRobustToken(tokens[j]))
                    {
                        result.Add($"{tokens[i]} {tokens[j]}");
                    }
                }
            }

            return result;
        }

        private static HashSet<string> BuildGloballyUniqueSingleAliases(IReadOnlyList<ScheduleParticipantRef> participants)
        {
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var participant in participants.Where(entry => !string.IsNullOrWhiteSpace(entry.DisplayName)))
            {
                var tokens = TokenizeLooseNormalized(participant.DisplayName)
                    .Where(CanLeadAlias)
                    .Distinct(StringComparer.OrdinalIgnoreCase);

                foreach (var token in tokens)
                {
                    counts[token] = counts.TryGetValue(token, out var current) ? current + 1 : 1;
                }
            }

            return counts
                .Where(entry => entry.Value == 1)
                .Select(entry => entry.Key)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        private List<string> FilterParticipantKeysForGroup(IEnumerable<string> participantKeys, string group)
        {
            return participantKeys
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(participantKey => IsParticipantEligibleForGroup(participantKey, group))
                .ToList();
        }

        private bool IsParticipantEligibleForGroup(string participantKey, string group)
        {
            var isLbGroup = string.Equals(group, "LB", StringComparison.OrdinalIgnoreCase)
                || string.Equals(group, "LB ABEND", StringComparison.OrdinalIgnoreCase);
            var isLbParticipant = _statusTags.TryGetValue(participantKey, out var statusTag)
                && string.Equals(statusTag, "LB", StringComparison.OrdinalIgnoreCase);

            return !isLbGroup || isLbParticipant;
        }

        private static bool CanLeadAlias(string token)
        {
            return SearchTextUtility.IsRobustToken(token);
        }

        private static bool IsIgnorableToken(WeeklyScheduleToken token)
        {
            var normalized = NormalizeToken(token.Text);
            return normalized.Length <= 2
                   || normalized.All(char.IsDigit)
                   || normalized is "u" or "b";
        }

        private static string NormalizeToken(string value)
        {
            return SearchTextUtility.ReplaceUmlauts(value).ToLowerInvariant();
        }
    }

    private sealed record ResolutionPath(List<ResolvedLineMatch> Matches, int SkippedTokens, bool HasAmbiguity);
    private sealed record ResolvedLineMatch(string ParticipantKey, bool IsAbsent, bool IsExternal, bool IsSupplemental);
    private sealed record LineCandidate(string Alias, int StartIndex, int Length, List<string> ParticipantKeys);
    private sealed record LineAnalysis(string ResolutionState, List<ResolvedLineMatch> ResolvedMatches, List<LineCandidate> Candidates);
}

internal sealed class ScheduleParticipantRef
{
    public string ParticipantKey { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string StatusTag { get; set; } = string.Empty;
}

public sealed class WeeklyScheduleCacheDocument
{
    public int Version { get; set; }
    public List<WeeklyScheduleCacheEntry> Entries { get; set; } = new();
}

public sealed class WeeklyScheduleCacheEntry
{
    public string DocumentPath { get; set; } = string.Empty;
    public DateTime LastWriteTimeUtc { get; set; }
    public long Length { get; set; }
    public WeeklyScheduleDocument Document { get; set; } = new();
}

public sealed class WeeklyScheduleDocument
{
    public List<WeeklyScheduleSlot> Slots { get; set; } = new();
}

public sealed class WeeklyScheduleSlot
{
    public string DayKey { get; set; } = string.Empty;
    public string HalfDay { get; set; } = string.Empty;
    public List<WeeklyScheduleBlock> Blocks { get; set; } = new();
}

public sealed class WeeklyScheduleBlock
{
    public string Group { get; set; } = string.Empty;
    public string Teacher { get; set; } = string.Empty;
    public string Room { get; set; } = string.Empty;
    public List<WeeklyScheduleParticipantLine> ParticipantLines { get; set; } = new();
}

public sealed class WeeklyScheduleParticipantLine
{
    public string RawText { get; set; } = string.Empty;
    public List<WeeklyScheduleToken> Tokens { get; set; } = new();
}

public sealed class WeeklyScheduleToken
{
    public string Text { get; set; } = string.Empty;
    public bool IsAbsent { get; set; }
    public bool IsExternal { get; set; }
    public bool IsSupplemental { get; set; }
}

internal sealed class ParsedRow
{
    public List<ParsedCell> Cells { get; } = new();
}

internal sealed class ParsedCell
{
    public int StartColumn { get; set; }
    public int Span { get; set; }
    public string FullText { get; set; } = string.Empty;
    public List<ParsedParagraph> Paragraphs { get; set; } = new();
}

internal sealed class ParsedParagraph
{
    public string Text { get; set; } = string.Empty;
    public List<WeeklyScheduleToken> Tokens { get; set; } = new();
}

internal sealed class DayRange
{
    public string DayKey { get; set; } = string.Empty;
    public int StartColumn { get; set; }
    public int EndColumnExclusive { get; set; }
}



































