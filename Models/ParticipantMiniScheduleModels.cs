namespace VerlaufsakteApp.Models;

public enum ParticipantMiniScheduleState
{
    Hidden,
    Unavailable,
    Ready
}

public enum ParticipantMiniScheduleHalfDay
{
    Morning,
    Afternoon
}

public enum ParticipantMiniScheduleCellStatus
{
    None,
    External,
    Dispensed
}

public sealed class ParticipantMiniScheduleEntry
{
    public string Group { get; set; } = string.Empty;
    public string Teacher { get; set; } = string.Empty;
    public string Room { get; set; } = string.Empty;
    public bool IsExternal { get; set; }
}

public sealed class ParticipantMiniScheduleCell
{
    public string DayKey { get; set; } = string.Empty;
    public ParticipantMiniScheduleHalfDay HalfDay { get; set; }
    public List<ParticipantMiniScheduleEntry> Entries { get; set; } = new();
    public ParticipantMiniScheduleCellStatus Status { get; set; }
    public bool HasSupplementalDaz { get; set; }

    public ParticipantMiniScheduleEntry? PrimaryEntry => Entries.FirstOrDefault();
    public string DisplayGroup => PrimaryEntry?.Group ?? string.Empty;
    public string DisplayTeacher => PrimaryEntry?.Teacher ?? string.Empty;
    public string DisplayRoom => PrimaryEntry?.Room ?? string.Empty;
    public bool HasPrimaryEntry => PrimaryEntry is not null;
    public bool IsExternal => Status == ParticipantMiniScheduleCellStatus.External;
    public bool IsDispensed => Status == ParticipantMiniScheduleCellStatus.Dispensed;
}

public sealed class ParticipantMiniScheduleSummary
{
    private static readonly string[] DayOrder = ["Mo", "Di", "Mi", "Do", "Fr"];

    public ParticipantMiniScheduleState State { get; set; }
    public string Message { get; set; } = string.Empty;
    public List<ParticipantMiniScheduleCell> Cells { get; set; } = CreateDefaultCells();

    public IReadOnlyList<ParticipantMiniScheduleCell> MorningCells =>
        GetCellsForHalfDay(ParticipantMiniScheduleHalfDay.Morning);

    public IReadOnlyList<ParticipantMiniScheduleCell> AfternoonCells =>
        GetCellsForHalfDay(ParticipantMiniScheduleHalfDay.Afternoon);

    public static ParticipantMiniScheduleSummary Create(
        ParticipantMiniScheduleState state,
        string message = "")
    {
        return new ParticipantMiniScheduleSummary
        {
            State = state,
            Message = message,
            Cells = CreateDefaultCells()
        };
    }

    public ParticipantMiniScheduleCell GetCell(string dayKey, ParticipantMiniScheduleHalfDay halfDay)
    {
        return Cells.First(cell =>
            string.Equals(cell.DayKey, dayKey, StringComparison.OrdinalIgnoreCase)
            && cell.HalfDay == halfDay);
    }

    private IReadOnlyList<ParticipantMiniScheduleCell> GetCellsForHalfDay(ParticipantMiniScheduleHalfDay halfDay)
    {
        return DayOrder
            .Select(dayKey => GetCell(dayKey, halfDay))
            .ToList();
    }

    public static List<ParticipantMiniScheduleCell> CreateDefaultCells()
    {
        var cells = new List<ParticipantMiniScheduleCell>();
        foreach (var dayKey in DayOrder)
        {
            cells.Add(new ParticipantMiniScheduleCell
            {
                DayKey = dayKey,
                HalfDay = ParticipantMiniScheduleHalfDay.Morning
            });
            cells.Add(new ParticipantMiniScheduleCell
            {
                DayKey = dayKey,
                HalfDay = ParticipantMiniScheduleHalfDay.Afternoon
            });
        }

        return cells;
    }
}
