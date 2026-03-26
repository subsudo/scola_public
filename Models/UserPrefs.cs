namespace VerlaufsakteApp.Models;

public class UserPrefs
{
    public bool ShowParticipantInitials { get; set; } = true;
    public bool ShowBtnOdoo { get; set; } = false;
    public bool ShowBtnOrdner { get; set; } = true;
    public bool ShowBtnAkte { get; set; } = true;
    public bool ShowBtnBu { get; set; } = true;
    public bool ShowBtnEintrag { get; set; } = true;
    public bool ShowBtnBi { get; set; } = false;
    public bool ShowBtnBe { get; set; } = false;
    public bool ShowBtnEintragBi { get; set; } = false;
    public bool IsDarkTheme { get; set; } = false;
    public string DisplayDensity { get; set; } = DisplayDensityMode.Standard;
    public bool AutoPrefillOnEmptyClipboard { get; set; } = false;
    public string DefaultEntryInitials { get; set; } = string.Empty;
    public bool EnableDebugLogging { get; set; }
    public bool EnableWordLifecycleLogging { get; set; }
    public double? WindowLeft { get; set; }
    public double? WindowTop { get; set; }
    public double? WindowWidth { get; set; }
    public double? WindowHeight { get; set; }
    public string WindowMonitorDeviceName { get; set; } = string.Empty;
    public bool WindowWasMaximized { get; set; }
    public bool OpenWordMaximized { get; set; }
    public string PreferredWordMonitorId { get; set; } = "__PRIMARY__";
}
