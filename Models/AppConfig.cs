namespace VerlaufsakteApp.Models;

public class AppConfig
{
    public string ServerBasePath { get; set; } = string.Empty;
    public bool UseSecondaryServerBasePath { get; set; }
    public string SecondaryServerBasePath { get; set; } = string.Empty;
    public bool UseTertiaryServerBasePath { get; set; }
    public string TertiaryServerBasePath { get; set; } = string.Empty;
    public string ScheduleRootPath { get; set; } = string.Empty;
    public List<string> AbsenceValues { get; set; } = new();
    public List<string> PresenceValues { get; set; } = new();
    public string VerlaufsakteKeyword { get; set; } = "Verlaufsakte";
    public string WordBookmarkName { get; set; } = "BU_BILDUNG_TABELLE";
    public string WordBuBookmarkName { get; set; } = "_Bildung";
    public string WordBiBookmarkName { get; set; } = "_Berufsintegration";
    public string WordBeBookmarkName { get; set; } = "_Beratung";
    public string WordBiTableBookmarkName { get; set; } = "BI_BERUFSINTEGRATION_TABELLE";
    public string WordBiTodoBookmarkName { get; set; } = "BI_BERUFSINTEGRATION_TODO";
    public string ParticipantHintsStorePath { get; set; } = string.Empty;
}
