using System.Diagnostics;
using System.IO.Compression;
using System.Xml.Linq;

namespace VerlaufsakteApp.Services;

internal sealed class BiDocxParagraphContent
{
    public string Text { get; init; } = string.Empty;
    public bool IsBullet { get; init; }
}

internal sealed class BiDocxExtractionResult
{
    public string CareerChoice { get; init; } = "-";
    public IReadOnlyList<BiDocxParagraphContent> Paragraphs { get; init; } = Array.Empty<BiDocxParagraphContent>();
}

internal static class BiDocxExtractionService
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static BiDocxExtractionResult Extract(string documentPath, string bookmarkName, string participantName)
    {
        var totalStopwatch = Stopwatch.StartNew();
        var loadStopwatch = Stopwatch.StartNew();
        using var archive = ZipFile.OpenRead(documentPath);
        var documentEntry = archive.GetEntry("word/document.xml");
        if (documentEntry is null)
        {
            throw CreateContentInvalidException(bookmarkName, "Akte konnte nicht gelesen werden. Bitte Vorlage prüfen.");
        }

        using var stream = documentEntry.Open();
        var document = XDocument.Load(stream);
        var body = document.Root?.Element(W + "body");
        if (body is null)
        {
            throw CreateContentInvalidException(bookmarkName, "Akte konnte nicht gelesen werden. Bitte Vorlage prüfen.");
        }
        loadStopwatch.Stop();
        LogParticipantTiming(participantName, "LoadDocxXml", loadStopwatch.ElapsedMilliseconds);

        var careerStopwatch = Stopwatch.StartNew();
        var careerChoice = ExtractCareerChoice(body, bookmarkName);
        careerStopwatch.Stop();
        LogParticipantTiming(participantName, "ReadBiCareerChoiceXml", careerStopwatch.ElapsedMilliseconds);

        var todoStopwatch = Stopwatch.StartNew();
        var paragraphs = ExtractTodoParagraphs(document, bookmarkName);
        todoStopwatch.Stop();
        LogParticipantTiming(participantName, "ReadBiTodoParagraphsXml", todoStopwatch.ElapsedMilliseconds);

        totalStopwatch.Stop();
        LogParticipantTiming(participantName, "ExtractFromDocx", totalStopwatch.ElapsedMilliseconds);

        return new BiDocxExtractionResult
        {
            CareerChoice = careerChoice,
            Paragraphs = paragraphs
        };
    }

    private static string ExtractCareerChoice(XElement body, string bookmarkName)
    {
        var bodyItems = body.Elements().ToList();

        if (TryExtractCareerChoiceBetweenMarkers(bodyItems, "Berufswunsch", "Unterrichtsmaterial", out var buCareerChoice))
        {
            return NormalizeCareerChoiceValue(buCareerChoice);
        }

        if (TryExtractCareerChoiceBetweenMarkers(bodyItems, "Berufswünsche", "Arbeitseinsätze", out var biCareerChoice))
        {
            return NormalizeCareerChoiceValue(biCareerChoice);
        }

        throw CreateContentInvalidException(bookmarkName, "Berufswunsch konnte nicht gelesen werden. Bitte Vorlage prüfen.");
    }

    private static bool TryExtractCareerChoiceBetweenMarkers(
        IReadOnlyList<XElement> bodyItems,
        string startMarker,
        string endMarker,
        out string careerChoice)
    {
        careerChoice = string.Empty;

        var startIndex = FindParagraphIndex(bodyItems, startMarker);
        if (startIndex < 0)
        {
            return false;
        }

        var endIndex = FindParagraphIndex(bodyItems, endMarker, startIndex + 1);
        if (endIndex < 0)
        {
            endIndex = bodyItems.Count;
        }

        for (var index = startIndex + 1; index < endIndex; index++)
        {
            var value = ExtractElementValue(bodyItems[index]);
            if (value is null)
            {
                continue;
            }

            careerChoice = value;
            return true;
        }

        return true;
    }

    private static int FindParagraphIndex(IReadOnlyList<XElement> bodyItems, string marker, int startIndex = 0)
    {
        for (var index = startIndex; index < bodyItems.Count; index++)
        {
            var element = bodyItems[index];
            if (element.Name != W + "p")
            {
                continue;
            }

            var text = NormalizeText(GetElementText(element));
            if (string.Equals(text, marker, StringComparison.OrdinalIgnoreCase))
            {
                return index;
            }
        }

        return -1;
    }

    private static string? ExtractElementValue(XElement element)
    {
        var text = NormalizeText(GetElementText(element));
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        return text;
    }

    private static IReadOnlyList<BiDocxParagraphContent> ExtractTodoParagraphs(XDocument document, string bookmarkName)
    {
        var bookmark = document
            .Descendants(W + "bookmarkStart")
            .FirstOrDefault(x => string.Equals((string?)x.Attribute(W + "name"), bookmarkName, StringComparison.Ordinal));

        if (bookmark is null)
        {
            throw new WordTemplateValidationException(
                WordTemplateValidationErrorKind.BookmarkMissing,
                bookmarkName,
                $"Bookmark '{bookmarkName}' wurde nicht gefunden.");
        }

        var table = bookmark.Ancestors(W + "tbl").FirstOrDefault();
        if (table is null)
        {
            throw CreateTableInvalidException(bookmarkName);
        }

        var rows = table.Elements(W + "tr").ToList();
        if (rows.Count < 1)
        {
            throw CreateTableInvalidException(bookmarkName);
        }

        var cells = rows[^1].Elements(W + "tc").ToList();
        if (cells.Count < 2)
        {
            throw CreateTableInvalidException(bookmarkName);
        }

        var targetCell = cells[^1];
        var result = new List<BiDocxParagraphContent>();
        foreach (var paragraph in targetCell.Elements(W + "p"))
        {
            var text = NormalizeText(GetElementText(paragraph));
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (string.Equals(text, "Nämal:", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "Ampel:", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var isBullet = paragraph.Element(W + "pPr")?.Element(W + "numPr") is not null;
            result.Add(new BiDocxParagraphContent
            {
                Text = text,
                IsBullet = isBullet
            });
        }

        return result;
    }

    private static string NormalizeCareerChoiceValue(string? rawValue)
    {
        var normalized = NormalizeText(rawValue);
        if (string.IsNullOrWhiteSpace(normalized) ||
            string.Equals(normalized, "[Titel]", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalized, "EBA / EFZ, Bereiche", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalized, "Wählen Sie ein Element aus.", StringComparison.OrdinalIgnoreCase))
        {
            return "-";
        }

        return normalized;
    }

    private static string GetElementText(XElement element)
    {
        return string.Concat(element.Descendants(W + "t").Select(x => x.Value));
    }

    private static string NormalizeText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        return WordService.NormalizeWhitespaceForBiDocx(text);
    }

    private static WordTemplateValidationException CreateTableInvalidException(string bookmarkName)
    {
        return new WordTemplateValidationException(
            WordTemplateValidationErrorKind.BiTodoTableInvalid,
            bookmarkName,
            "BI-To-do-Tabelle hat nicht das erwartete Format. Bitte Vorlage prüfen.");
    }

    private static WordTemplateValidationException CreateContentInvalidException(string bookmarkName, string message)
    {
        return new WordTemplateValidationException(
            WordTemplateValidationErrorKind.BiTodoContentInvalid,
            bookmarkName,
            message);
    }

    private static void LogParticipantTiming(string participantName, string step, long elapsedMilliseconds)
    {
        AppLogger.Debug($"Word.CollectBiTodoDocument timing: TN='{participantName}', Step='{step}', Elapsed={elapsedMilliseconds} ms.");
    }
}
