using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;

namespace VerlaufsakteApp.Models;

public static class ParticipantHintTypes
{
    public const string Exit = "austritt";
    public const string AmReport = "am_bericht";
    public const string StellwerkTest = "stellwerk_test";
    public const string Free = "frei";
}

public static class ParticipantHintStatuses
{
    public const string Active = "active";
    public const string Done = "done";
}

public sealed class ParticipantHintDocument
{
    [JsonPropertyName("schema_version")]
    public int SchemaVersion { get; set; } = 1;

    [JsonPropertyName("updated_at_utc")]
    public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;

    [JsonPropertyName("participants")]
    public List<ParticipantHintParticipantRecord> Participants { get; set; } = new();
}

public sealed class ParticipantHintParticipantRecord
{
    [JsonPropertyName("key")]
    public string Key { get; set; } = string.Empty;

    [JsonPropertyName("canonical_document_path")]
    public string CanonicalDocumentPath { get; set; } = string.Empty;

    [JsonPropertyName("original_document_path")]
    public string OriginalDocumentPath { get; set; } = string.Empty;

    [JsonPropertyName("folder_name")]
    public string FolderName { get; set; } = string.Empty;

    [JsonPropertyName("updated_at_utc")]
    public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;

    [JsonPropertyName("hints")]
    public List<ParticipantHintEntry> Hints { get; set; } = new();
}

public sealed class ParticipantHintEntry
{
    [JsonPropertyName("id")]
    public string Id { get; set; } = Guid.NewGuid().ToString("N");

    [JsonPropertyName("type")]
    public string Type { get; set; } = ParticipantHintTypes.Free;

    [JsonPropertyName("status")]
    public string Status { get; set; } = ParticipantHintStatuses.Active;

    [JsonPropertyName("details")]
    public ParticipantHintDetails Details { get; set; } = new();

    [JsonPropertyName("updated_at_utc")]
    public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;

    [JsonPropertyName("updated_by")]
    public string UpdatedBy { get; set; } = string.Empty;
}

public sealed class ParticipantHintDetails
{
    [JsonPropertyName("date")]
    public string Date { get; set; } = string.Empty;

    [JsonPropertyName("month")]
    public string Month { get; set; } = string.Empty;

    [JsonPropertyName("subject")]
    public string Subject { get; set; } = string.Empty;

    [JsonPropertyName("note")]
    public string Note { get; set; } = string.Empty;

    [JsonPropertyName("text")]
    public string Text { get; set; } = string.Empty;
}

public sealed class ParticipantHintDisplay
{
    public string Type { get; init; } = ParticipantHintTypes.Free;
    public string Text { get; init; } = string.Empty;
    public string Code { get; init; } = "N";
    public string Value { get; init; } = string.Empty;
    public string MarkerColor { get; init; } = "#A8A29A";
    public string PillBackground { get; init; } = "#E7E5E4";
    public string PillForeground { get; init; } = "#292524";
    public bool IsOverdue { get; init; }
    public DateTime? SortDate { get; init; }
}

public sealed class ParticipantHintEditSession
{
    public bool IsAvailable { get; init; }
    public string ErrorMessage { get; init; } = string.Empty;
    public string DocumentPath { get; init; } = string.Empty;
    public string ExpectedHash { get; init; } = string.Empty;
    public ParticipantHintParticipantRecord Record { get; init; } = new();
}

public sealed class ParticipantHintSaveResult
{
    public bool Success { get; init; }
    public bool Conflict { get; init; }
    public string ErrorMessage { get; init; } = string.Empty;
}

public sealed class ParticipantHintEditorItem : INotifyPropertyChanged
{
    private const int MaxNoteTextLength = 60;

    private string _status = ParticipantHintStatuses.Active;
    private string _date = string.Empty;
    private string _month = string.Empty;
    private string _subject = string.Empty;
    private string _note = string.Empty;
    private string _text = string.Empty;

    public string Id { get; init; } = Guid.NewGuid().ToString("N");
    public string Type { get; init; } = ParticipantHintTypes.Free;
    public DateTime UpdatedAtUtc { get; init; } = DateTime.UtcNow;
    public string UpdatedBy { get; init; } = string.Empty;

    public string TypeLabel => Type switch
    {
        ParticipantHintTypes.Exit => "Austritt",
        ParticipantHintTypes.AmReport => "AM-Bericht",
        ParticipantHintTypes.StellwerkTest => "Stellwerk-Test",
        _ => "Notiz"
    };

    public string Status
    {
        get => _status;
        set
        {
            if (SetField(ref _status, value))
            {
                OnPropertyChanged(nameof(IsActive));
            }
        }
    }

    public bool IsActive => string.Equals(Status, ParticipantHintStatuses.Active, StringComparison.OrdinalIgnoreCase);
    public bool ShowDate => true;
    public bool ShowMonth => false;
    public bool ShowSubject => false;
    public bool ShowText => Type == ParticipantHintTypes.Free;
    public bool ShowNote => false;

    public string Date
    {
        get => _date;
        set
        {
            if (SetField(ref _date, value))
            {
                OnPropertyChanged(nameof(DateValue));
            }
        }
    }

    public DateTime? DateValue
    {
        get => TryParseDateInput(Date, out var date) ? date : null;
        set
        {
            Date = value.HasValue
                ? value.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : string.Empty;
        }
    }

    public string Month { get => _month; set => SetField(ref _month, value); }
    public string Subject { get => _subject; set => SetField(ref _subject, value); }
    public string Note { get => _note; set => SetField(ref _note, value); }
    public string Text
    {
        get => _text;
        set => SetField(ref _text, LimitText(value));
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public static ParticipantHintEditorItem FromEntry(ParticipantHintEntry entry)
    {
        return new ParticipantHintEditorItem
        {
            Id = string.IsNullOrWhiteSpace(entry.Id) ? Guid.NewGuid().ToString("N") : entry.Id,
            Type = entry.Type,
            Status = string.IsNullOrWhiteSpace(entry.Status) ? ParticipantHintStatuses.Active : entry.Status,
            Date = string.IsNullOrWhiteSpace(entry.Details.Date) && entry.Type == ParticipantHintTypes.AmReport && !string.IsNullOrWhiteSpace(entry.Details.Month)
                ? $"{entry.Details.Month}-01"
                : entry.Details.Date,
            Month = entry.Details.Month,
            Subject = entry.Details.Subject,
            Note = entry.Details.Note,
            Text = entry.Details.Text,
            UpdatedAtUtc = entry.UpdatedAtUtc,
            UpdatedBy = entry.UpdatedBy
        };
    }

    public ParticipantHintEntry ToEntry(string updatedBy)
    {
        return new ParticipantHintEntry
        {
            Id = Id,
            Type = Type,
            Status = IsActive ? ParticipantHintStatuses.Active : ParticipantHintStatuses.Done,
            Details = new ParticipantHintDetails
            {
                Date = NormalizeDateInput(Date),
                Month = string.Empty,
                Subject = string.Empty,
                Note = string.Empty,
                Text = LimitText(Text)
            },
            UpdatedAtUtc = DateTime.UtcNow,
            UpdatedBy = updatedBy
        };
    }

    private static string NormalizeDateInput(string value)
    {
        var trimmed = value.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        if (TryParseDateInput(trimmed, out var isoDate))
        {
            return isoDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        return trimmed;
    }

    private static bool TryParseDateInput(string value, out DateTime date)
    {
        return DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date)
               || DateTime.TryParseExact(value, "dd.MM.yyyy", CultureInfo.GetCultureInfo("de-CH"), DateTimeStyles.None, out date)
               || DateTime.TryParseExact(value, "dd.MM.yy", CultureInfo.GetCultureInfo("de-CH"), DateTimeStyles.None, out date)
               || DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out date);
    }

    private static string LimitText(string value)
    {
        var trimmed = value.Trim();
        return trimmed.Length <= MaxNoteTextLength
            ? trimmed
            : trimmed[..MaxNoteTextLength];
    }

    private static string NormalizeMonthInput(string value)
    {
        var trimmed = value.Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        if (DateTime.TryParseExact(trimmed, "yyyy-MM", CultureInfo.InvariantCulture, DateTimeStyles.None, out var month)
            || DateTime.TryParseExact(trimmed, "MM.yyyy", CultureInfo.GetCultureInfo("de-CH"), DateTimeStyles.None, out month)
            || DateTime.TryParse(trimmed, CultureInfo.CurrentCulture, DateTimeStyles.None, out month))
        {
            return month.ToString("yyyy-MM", CultureInfo.InvariantCulture);
        }

        return trimmed;
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
