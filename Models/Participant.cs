using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace VerlaufsakteApp.Models;

public class Participant : INotifyPropertyChanged
{
    private string _fullName = string.Empty;
    private string _documentPath = string.Empty;
    private string _initials = string.Empty;
    private string _odooUrl = string.Empty;
    private string _counselorInitials = string.Empty;
    private string _status = string.Empty;
    private string _absenceDetail = string.Empty;
    private ParticipantMiniScheduleState _miniScheduleState = ParticipantMiniScheduleState.Hidden;
    private ParticipantMiniScheduleSummary _miniScheduleSummary = ParticipantMiniScheduleSummary.Create(ParticipantMiniScheduleState.Hidden);
    private bool _isMiniScheduleExpanded;
    private bool _isMiniScheduleLoading;
    private bool _isHeaderMetadataLoaded;
    private bool _isPresent;
    private string? _matchedFolderPath;
    private MatchStatus _matchStatus;
    private string? _selectedFolderPath;
    private bool _canOpenOdoo;
    private bool _canOpenFolder;
    private bool _canOpenAkte;
    private bool _canOpenAkteBu;
    private bool _canInsertEntry;
    private bool _canOpenAkteBi;
    private bool _canOpenAkteBe;
    private bool _canInsertEntryBi;
    private IReadOnlyList<ParticipantHintDisplay> _activeHints = Array.Empty<ParticipantHintDisplay>();
    private IReadOnlyList<ParticipantHintDisplay> _hintMarkers = Array.Empty<ParticipantHintDisplay>();
    private string _hintOverflowText = string.Empty;

    public string FullName
    {
        get => _fullName;
        set => SetField(ref _fullName, value);
    }

    public string DocumentPath
    {
        get => _documentPath;
        set => SetField(ref _documentPath, value);
    }

    public string Initials
    {
        get => _initials;
        set
        {
            if (SetField(ref _initials, value))
            {
                OnPropertyChanged(nameof(HasInitials));
            }
        }
    }

    public bool HasInitials => !string.IsNullOrWhiteSpace(Initials);

    public string OdooUrl
    {
        get => _odooUrl;
        set
        {
            if (SetField(ref _odooUrl, value))
            {
                OnPropertyChanged(nameof(HasOdooUrl));
            }
        }
    }

    public bool HasOdooUrl => !string.IsNullOrWhiteSpace(OdooUrl);

    public string CounselorInitials
    {
        get => _counselorInitials;
        set
        {
            if (SetField(ref _counselorInitials, value))
            {
                OnPropertyChanged(nameof(HasCounselorInitials));
            }
        }
    }

    public bool HasCounselorInitials => !string.IsNullOrWhiteSpace(CounselorInitials);

    public string Status
    {
        get => _status;
        set
        {
            if (SetField(ref _status, value))
            {
                OnPropertyChanged(nameof(AbsenceStatusText));
                OnPropertyChanged(nameof(ShowAbsenceDetailSuffix));
            }
        }
    }

    public ParticipantMiniScheduleState MiniScheduleState
    {
        get => _miniScheduleState;
        set => SetField(ref _miniScheduleState, value);
    }

    public ParticipantMiniScheduleSummary MiniScheduleSummary
    {
        get => _miniScheduleSummary;
        set => SetField(ref _miniScheduleSummary, value);
    }

    public bool IsMiniScheduleExpanded
    {
        get => _isMiniScheduleExpanded;
        set => SetField(ref _isMiniScheduleExpanded, value);
    }

    public bool IsMiniScheduleLoading
    {
        get => _isMiniScheduleLoading;
        set => SetField(ref _isMiniScheduleLoading, value);
    }

    public bool IsHeaderMetadataLoaded
    {
        get => _isHeaderMetadataLoaded;
        set => SetField(ref _isHeaderMetadataLoaded, value);
    }

    public string AbsenceDetail
    {
        get => _absenceDetail;
        set
        {
            if (SetField(ref _absenceDetail, value))
            {
                OnPropertyChanged(nameof(HasAbsenceDetail));
                OnPropertyChanged(nameof(ShowAbsenceDetailSuffix));
            }
        }
    }

    public bool HasAbsenceDetail => !string.IsNullOrWhiteSpace(AbsenceDetail);

    public bool IsPresent
    {
        get => _isPresent;
        set
        {
            if (SetField(ref _isPresent, value))
            {
                OnPropertyChanged(nameof(DisplayOpacity));
                OnPropertyChanged(nameof(ShowAbsenceStatus));
                OnPropertyChanged(nameof(PresenceSymbol));
                OnPropertyChanged(nameof(AbsenceStatusText));
                OnPropertyChanged(nameof(ShowAbsenceDetailSuffix));
            }
        }
    }

    public string? MatchedFolderPath
    {
        get => _matchedFolderPath;
        set => SetField(ref _matchedFolderPath, value);
    }

    public MatchStatus MatchStatus
    {
        get => _matchStatus;
        set
        {
            if (SetField(ref _matchStatus, value))
            {
                OnPropertyChanged(nameof(MatchSymbol));
                OnPropertyChanged(nameof(ShowMatchSymbol));
                OnPropertyChanged(nameof(MatchColor));
                OnPropertyChanged(nameof(MatchTooltip));
                OnPropertyChanged(nameof(IsMultipleFound));
                OnPropertyChanged(nameof(IsNotFound));
            }
        }
    }

    public List<string> CandidateFolderPaths { get; set; } = new();

    public string? SelectedFolderPath
    {
        get => _selectedFolderPath;
        set
        {
            if (SetField(ref _selectedFolderPath, value))
            {
                MatchedFolderPath = value;
            }
        }
    }

    public bool CanOpenOdoo
    {
        get => _canOpenOdoo;
        set => SetField(ref _canOpenOdoo, value);
    }

    public bool CanOpenFolder
    {
        get => _canOpenFolder;
        set => SetField(ref _canOpenFolder, value);
    }

    public bool CanOpenAkte
    {
        get => _canOpenAkte;
        set => SetField(ref _canOpenAkte, value);
    }

    public bool CanOpenAkteBu
    {
        get => _canOpenAkteBu;
        set => SetField(ref _canOpenAkteBu, value);
    }

    public bool CanInsertEntry
    {
        get => _canInsertEntry;
        set => SetField(ref _canInsertEntry, value);
    }

    public bool CanOpenAkteBi
    {
        get => _canOpenAkteBi;
        set => SetField(ref _canOpenAkteBi, value);
    }

    public bool CanOpenAkteBe
    {
        get => _canOpenAkteBe;
        set => SetField(ref _canOpenAkteBe, value);
    }

    public bool CanInsertEntryBi
    {
        get => _canInsertEntryBi;
        set => SetField(ref _canInsertEntryBi, value);
    }

    public IReadOnlyList<ParticipantHintDisplay> ActiveHints
    {
        get => _activeHints;
        set
        {
            var normalized = value.ToList();
            _activeHints = normalized;
            _hintMarkers = normalized.Take(4).ToList();
            _hintOverflowText = string.Empty;
            OnPropertyChanged();
            OnPropertyChanged(nameof(HintMarkers));
            OnPropertyChanged(nameof(HasActiveHints));
            OnPropertyChanged(nameof(HasHintOverflow));
            OnPropertyChanged(nameof(HintOverflowText));
        }
    }

    public IReadOnlyList<ParticipantHintDisplay> HintMarkers => _hintMarkers;
    public bool HasActiveHints => ActiveHints.Count > 0;
    public bool HasHintOverflow => !string.IsNullOrWhiteSpace(HintOverflowText);
    public string HintOverflowText => _hintOverflowText;

    public double DisplayOpacity => IsPresent ? 1.0 : 0.35;
    public bool ShowAbsenceStatus => !IsPresent;
    public bool IsMultipleFound => MatchStatus == MatchStatus.MultipleFound;
    public bool IsNotFound => MatchStatus == MatchStatus.NotFound;
    public string PresenceSymbol => IsPresent ? "●" : "○";
    public string AbsenceStatusText => IsPresent
        ? string.Empty
        : (string.IsNullOrWhiteSpace(Status) ? "Abwesend" : Status);

    public bool ShowAbsenceDetailSuffix => !IsPresent
                                           && HasAbsenceDetail
                                           && !string.Equals(
                                               AbsenceDetail.Trim(),
                                               Status.Trim(),
                                               StringComparison.OrdinalIgnoreCase);

    public string MatchSymbol => MatchStatus switch
    {
        MatchStatus.Found => "",
        MatchStatus.NotFound => "✗",
        MatchStatus.MultipleFound => "⚠",
        _ => ""
    };

    public bool ShowMatchSymbol => MatchStatus != MatchStatus.Found;

    public string MatchColor => MatchStatus switch
    {
        MatchStatus.NotFound => "#D17878",
        MatchStatus.MultipleFound => "#C8A96C",
        _ => "#888899"
    };

    public string MatchTooltip => MatchStatus switch
    {
        MatchStatus.Found => "",
        MatchStatus.NotFound => "Kein Ordner gefunden - Ordnername prüfen",
        MatchStatus.MultipleFound => "Mehrere Treffer - bitte manuell auswählen",
        _ => string.Empty
    };

    public event PropertyChangedEventHandler? PropertyChanged;

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

public enum MatchStatus
{
    Found,
    NotFound,
    MultipleFound
}
