namespace VerlaufsakteApp.Models;

public sealed class WeeklyScheduleDiagnosticsDocument
{
    public DateTime GeneratedAt { get; set; }
    public string RequestedPath { get; set; } = string.Empty;
    public string ResolvedPath { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public List<WeeklyScheduleSlotDiagnostics> Slots { get; set; } = new();
    public List<WeeklyScheduleParticipantDiagnostics> Participants { get; set; } = new();
}

public sealed class WeeklyScheduleSlotDiagnostics
{
    public string DayKey { get; set; } = string.Empty;
    public string HalfDay { get; set; } = string.Empty;
    public List<WeeklyScheduleBlockDiagnostics> Blocks { get; set; } = new();
}

public sealed class WeeklyScheduleBlockDiagnostics
{
    public string Group { get; set; } = string.Empty;
    public string Teacher { get; set; } = string.Empty;
    public string Room { get; set; } = string.Empty;
    public List<WeeklyScheduleLineDiagnostics> ParticipantLines { get; set; } = new();
}

public sealed class WeeklyScheduleLineDiagnostics
{
    public string RawText { get; set; } = string.Empty;
    public List<WeeklyScheduleTokenDiagnostics> Tokens { get; set; } = new();
    public string ResolutionState { get; set; } = string.Empty;
    public List<WeeklyScheduleResolvedParticipantDiagnostics> ResolvedParticipants { get; set; } = new();
    public List<WeeklyScheduleCandidateDiagnostics> CandidateMatches { get; set; } = new();
}

public sealed class WeeklyScheduleTokenDiagnostics
{
    public string Text { get; set; } = string.Empty;
    public bool IsAbsent { get; set; }
    public bool IsExternal { get; set; }
    public bool IsSupplemental { get; set; }
}

public sealed class WeeklyScheduleResolvedParticipantDiagnostics
{
    public string ParticipantKey { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public bool IsAbsent { get; set; }
    public bool IsExternal { get; set; }
}

public sealed class WeeklyScheduleCandidateDiagnostics
{
    public string Alias { get; set; } = string.Empty;
    public int StartIndex { get; set; }
    public int Length { get; set; }
    public List<WeeklyScheduleResolvedParticipantDiagnostics> Candidates { get; set; } = new();
}

public sealed class WeeklyScheduleParticipantDiagnostics
{
    public string ParticipantKey { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public string ResultState { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public List<WeeklyScheduleParticipantSlotDiagnostics> Matches { get; set; } = new();
    public List<string> AmbiguousLines { get; set; } = new();
}

public sealed class WeeklyScheduleParticipantSlotDiagnostics
{
    public string DayKey { get; set; } = string.Empty;
    public string HalfDay { get; set; } = string.Empty;
    public string Group { get; set; } = string.Empty;
    public string Teacher { get; set; } = string.Empty;
    public string Room { get; set; } = string.Empty;
    public bool IsExternal { get; set; }
}
