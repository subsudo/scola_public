using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp;

public sealed class SettingsWindowModel
{
    public string ServerPath { get; set; } = string.Empty;
    public bool UseSecondaryServerPath { get; set; }
    public string SecondaryServerPath { get; set; } = string.Empty;
    public bool UseTertiaryServerPath { get; set; }
    public string TertiaryServerPath { get; set; } = string.Empty;
    public string ScheduleRootPath { get; set; } = string.Empty;
    public bool ShowParticipantInitials { get; set; }
    public bool ShowBtnOdoo { get; set; }
    public bool ShowBtnOrdner { get; set; }
    public bool ShowBtnAkte { get; set; }
    public bool ShowBtnBu { get; set; }
    public bool ShowBtnEintrag { get; set; }
    public bool ShowBtnBi { get; set; }
    public bool ShowBtnBe { get; set; }
    public bool ShowBtnEintragBi { get; set; }
    public bool IsDarkTheme { get; set; }
    public string DisplayDensity { get; set; } = DisplayDensityMode.Standard;
    public bool AutoPrefillOnEmptyClipboard { get; set; }
    public string DefaultEntryInitials { get; set; } = string.Empty;
    public bool EnableDebugLogging { get; set; }
    public bool EnableWordLifecycleLogging { get; set; }
}

public sealed class SettingsWindowResult
{
    public string ServerPath { get; set; } = string.Empty;
    public bool UseSecondaryServerPath { get; set; }
    public string SecondaryServerPath { get; set; } = string.Empty;
    public bool UseTertiaryServerPath { get; set; }
    public string TertiaryServerPath { get; set; } = string.Empty;
    public string ScheduleRootPath { get; set; } = string.Empty;
    public bool ShowParticipantInitials { get; set; }
    public bool ShowBtnOdoo { get; set; }
    public bool ShowBtnOrdner { get; set; }
    public bool ShowBtnAkte { get; set; }
    public bool ShowBtnBu { get; set; }
    public bool ShowBtnEintrag { get; set; }
    public bool ShowBtnBi { get; set; }
    public bool ShowBtnBe { get; set; }
    public bool ShowBtnEintragBi { get; set; }
    public bool IsDarkTheme { get; set; }
    public string DisplayDensity { get; set; } = DisplayDensityMode.Standard;
    public bool AutoPrefillOnEmptyClipboard { get; set; }
    public string DefaultEntryInitials { get; set; } = string.Empty;
    public bool EnableDebugLogging { get; set; }
    public bool EnableWordLifecycleLogging { get; set; }
}

public partial class SettingsWindow : Window, INotifyPropertyChanged
{
    private const string DefaultPrimaryServerPath = @"K:\FuturX\20_TNinnen";
    private const string DefaultSecondaryServerPath = @"K:\FuturX\20_TNinnen\02_Lehrbegleitung";
    private const string DefaultScheduleRootPath = @"K:\FuturX\10_Arbeitsplanung\20_Planung\22_Wochenplanung\Einteilung TN";
    private readonly bool _initialIsDarkTheme;
    private readonly SettingsWindowModel _initialModel;

    private string _serverPath = string.Empty;
    private bool _useSecondaryServerPath;
    private string _secondaryServerPath = string.Empty;
    private bool _useTertiaryServerPath;
    private string _tertiaryServerPath = string.Empty;
    private string _scheduleRootPath = string.Empty;
    private bool _showParticipantInitials;
    private bool _showBtnOdoo;
    private bool _showBtnOrdner;
    private bool _showBtnAkte;
    private bool _showBtnBu;
    private bool _showBtnEintrag;
    private bool _showBtnBi;
    private bool _showBtnBe;
    private bool _showBtnEintragBi;
    private bool _autoPrefillOnEmptyClipboard;
    private string _defaultEntryInitials = string.Empty;
    private bool _enableDebugLogging;
    private bool _enableWordLifecycleLogging;
    private bool _isDarkTheme;
    private string _displayDensity = DisplayDensityMode.Standard;
    private bool _isSaving;
    private bool _isInitializing;
    private bool _isDirty;

    public SettingsWindowResult Result { get; private set; } = new();

    public string ServerPath
    {
        get => _serverPath;
        set => SetField(ref _serverPath, value);
    }

    public bool UseSecondaryServerPath
    {
        get => _useSecondaryServerPath;
        set => SetField(ref _useSecondaryServerPath, value);
    }

    public string SecondaryServerPath
    {
        get => _secondaryServerPath;
        set => SetField(ref _secondaryServerPath, value);
    }

    public bool UseTertiaryServerPath
    {
        get => _useTertiaryServerPath;
        set => SetField(ref _useTertiaryServerPath, value);
    }

    public string TertiaryServerPath
    {
        get => _tertiaryServerPath;
        set => SetField(ref _tertiaryServerPath, value);
    }

    public string ScheduleRootPath
    {
        get => _scheduleRootPath;
        set => SetField(ref _scheduleRootPath, value);
    }

    public bool ShowParticipantInitials
    {
        get => _showParticipantInitials;
        set => SetField(ref _showParticipantInitials, value);
    }

    public bool ShowBtnOdoo
    {
        get => _showBtnOdoo;
        set => SetField(ref _showBtnOdoo, value);
    }

    public bool ShowBtnOrdner
    {
        get => _showBtnOrdner;
        set => SetField(ref _showBtnOrdner, value);
    }

    public bool ShowBtnAkte
    {
        get => _showBtnAkte;
        set => SetField(ref _showBtnAkte, value);
    }

    public bool ShowBtnBu
    {
        get => _showBtnBu;
        set => SetField(ref _showBtnBu, value);
    }

    public bool ShowBtnEintrag
    {
        get => _showBtnEintrag;
        set => SetField(ref _showBtnEintrag, value);
    }

    public bool ShowBtnBi
    {
        get => _showBtnBi;
        set => SetField(ref _showBtnBi, value);
    }

    public bool ShowBtnBe
    {
        get => _showBtnBe;
        set => SetField(ref _showBtnBe, value);
    }

    public bool ShowBtnEintragBi
    {
        get => _showBtnEintragBi;
        set => SetField(ref _showBtnEintragBi, value);
    }

    public bool AutoPrefillOnEmptyClipboard
    {
        get => _autoPrefillOnEmptyClipboard;
        set => SetField(ref _autoPrefillOnEmptyClipboard, value);
    }

    public string DefaultEntryInitials
    {
        get => _defaultEntryInitials;
        set => SetField(ref _defaultEntryInitials, value);
    }

    public bool EnableDebugLogging
    {
        get => _enableDebugLogging;
        set => SetField(ref _enableDebugLogging, value);
    }

    public bool EnableWordLifecycleLogging
    {
        get => _enableWordLifecycleLogging;
        set
        {
            if (!SetField(ref _enableWordLifecycleLogging, value))
            {
                return;
            }

            if (value && !EnableDebugLogging)
            {
                EnableDebugLogging = true;
            }
        }
    }

    public bool IsDarkTheme
    {
        get => _isDarkTheme;
        set
        {
            if (!SetField(ref _isDarkTheme, value))
            {
                return;
            }

            OnPropertyChanged(nameof(IsLightTheme));
        }
    }

    public bool IsLightTheme
    {
        get => !IsDarkTheme;
        set
        {
            if (value)
            {
                IsDarkTheme = false;
            }
        }
    }

    public string DisplayDensity
    {
        get => _displayDensity;
        set
        {
            var normalized = DisplayDensityMode.Normalize(value);
            if (!SetField(ref _displayDensity, normalized))
            {
                return;
            }

            OnPropertyChanged(nameof(IsStandardDensity));
            OnPropertyChanged(nameof(IsCompactDensity));
        }
    }

    public bool IsStandardDensity
    {
        get => DisplayDensity == DisplayDensityMode.Standard;
        set
        {
            if (value)
            {
                DisplayDensity = DisplayDensityMode.Standard;
            }
        }
    }

    public bool IsCompactDensity
    {
        get => DisplayDensity == DisplayDensityMode.Compact;
        set
        {
            if (value)
            {
                DisplayDensity = DisplayDensityMode.Compact;
            }
        }
    }

    public bool IsDirty
    {
        get => _isDirty;
        private set
        {
            if (_isDirty == value)
            {
                return;
            }

            _isDirty = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(DirtyStateText));
        }
    }

    public string DirtyStateText => IsDirty
        ? "Änderungen nicht gespeichert"
        : "Keine ausstehenden Änderungen";

    public string PrimaryPathStatusText => IsPrimaryPathValid
        ? "Status: OK"
        : "Status: Ungültig";

    public bool ShowPrimaryPathValidation => !IsPrimaryPathValid;

    public string PrimaryPathValidationText => string.IsNullOrWhiteSpace(NormalizePath(ServerPath))
        ? "Primärer Pfad fehlt."
        : "Primärer Pfad ist nicht erreichbar.";

    public string SecondaryPathStatusText
    {
        get
        {
            if (!UseSecondaryServerPath)
            {
                return "Status: Deaktiviert";
            }

            return IsSecondaryPathValid ? "Status: Aktiv (OK)" : "Status: Aktiv (ungültig)";
        }
    }

    public bool ShowSecondaryPathValidation => UseSecondaryServerPath && !IsSecondaryPathValid;

    public string SecondaryPathValidationText => string.IsNullOrWhiteSpace(NormalizePath(SecondaryServerPath))
        ? "Sekundärer Pfad fehlt."
        : "Sekundärer Pfad ist nicht erreichbar.";

    public string TertiaryPathStatusText
    {
        get
        {
            if (!UseTertiaryServerPath)
            {
                return "Status: Deaktiviert";
            }

            return IsTertiaryPathValid ? "Status: Aktiv (OK)" : "Status: Aktiv (ungültig)";
        }
    }

    public bool ShowTertiaryPathValidation => UseTertiaryServerPath && !IsTertiaryPathValid;

    public string TertiaryPathValidationText => string.IsNullOrWhiteSpace(NormalizePath(TertiaryServerPath))
        ? "Dritter Pfad fehlt."
        : "Dritter Pfad ist nicht erreichbar.";

    public string SchedulePathStatusText
    {
        get
        {
            if (string.IsNullOrWhiteSpace(NormalizePath(ScheduleRootPath)))
            {
                return "Status: Nicht gesetzt";
            }

            return IsSchedulePathValid ? "Status: OK" : "Status: Ungültig";
        }
    }

    public bool ShowSchedulePathValidation =>
        !string.IsNullOrWhiteSpace(NormalizePath(ScheduleRootPath)) && !IsSchedulePathValid;

    public string SchedulePathValidationText =>
        "Stundenplanpfad muss eine vorhandene DOCX-Datei oder ein vorhandener Ordner sein.";

    public SettingsWindow(SettingsWindowModel model)
    {
        InitializeComponent();
        DataContext = this;

        _initialIsDarkTheme = model.IsDarkTheme;
        _initialModel = CloneModel(ApplyDefaultPaths(model));

        _isInitializing = true;
        var normalizedModel = ApplyDefaultPaths(model);
        ServerPath = normalizedModel.ServerPath;
        UseSecondaryServerPath = normalizedModel.UseSecondaryServerPath;
        SecondaryServerPath = normalizedModel.SecondaryServerPath;
        UseTertiaryServerPath = model.UseTertiaryServerPath;
        TertiaryServerPath = model.TertiaryServerPath;
        ScheduleRootPath = normalizedModel.ScheduleRootPath;
        ShowParticipantInitials = model.ShowParticipantInitials;
        ShowBtnOdoo = model.ShowBtnOdoo;
        ShowBtnOrdner = model.ShowBtnOrdner;
        ShowBtnAkte = model.ShowBtnAkte;
        ShowBtnBu = model.ShowBtnBu;
        ShowBtnEintrag = model.ShowBtnEintrag;
        ShowBtnBi = model.ShowBtnBi;
        ShowBtnBe = model.ShowBtnBe;
        ShowBtnEintragBi = model.ShowBtnEintragBi;
        AutoPrefillOnEmptyClipboard = model.AutoPrefillOnEmptyClipboard;
        DefaultEntryInitials = model.DefaultEntryInitials;
        EnableDebugLogging = model.EnableDebugLogging;
        EnableWordLifecycleLogging = model.EnableWordLifecycleLogging;
        IsDarkTheme = model.IsDarkTheme;
        DisplayDensity = model.DisplayDensity;
        _isInitializing = false;

        NotifyPathStateChanged();
        UpdateDirtyState();
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private bool IsPrimaryPathValid => Directory.Exists(NormalizePath(ServerPath));

    private bool IsSecondaryPathValid => Directory.Exists(NormalizePath(SecondaryServerPath));

    private bool IsTertiaryPathValid => Directory.Exists(NormalizePath(TertiaryServerPath));

    private bool IsSchedulePathValid
    {
        get
        {
            var normalized = NormalizePath(ScheduleRootPath);
            return string.IsNullOrWhiteSpace(normalized)
                   || Directory.Exists(normalized)
                   || (File.Exists(normalized) && string.Equals(Path.GetExtension(normalized), ".docx", StringComparison.OrdinalIgnoreCase));
        }
    }

    private static string NormalizePath(string? value)
    {
        return (value ?? string.Empty).Trim();
    }

    private static SettingsWindowModel CloneModel(SettingsWindowModel source)
    {
        return new SettingsWindowModel
        {
            ServerPath = NormalizePath(source.ServerPath),
            UseSecondaryServerPath = source.UseSecondaryServerPath,
            SecondaryServerPath = NormalizePath(source.SecondaryServerPath),
            UseTertiaryServerPath = source.UseTertiaryServerPath,
            TertiaryServerPath = NormalizePath(source.TertiaryServerPath),
            ScheduleRootPath = NormalizePath(source.ScheduleRootPath),
            ShowParticipantInitials = source.ShowParticipantInitials,
            ShowBtnOdoo = source.ShowBtnOdoo,
            ShowBtnOrdner = source.ShowBtnOrdner,
            ShowBtnAkte = source.ShowBtnAkte,
            ShowBtnBu = source.ShowBtnBu,
            ShowBtnEintrag = source.ShowBtnEintrag,
            ShowBtnBi = source.ShowBtnBi,
            ShowBtnBe = source.ShowBtnBe,
            ShowBtnEintragBi = source.ShowBtnEintragBi,
            IsDarkTheme = source.IsDarkTheme,
            DisplayDensity = DisplayDensityMode.Normalize(source.DisplayDensity),
            AutoPrefillOnEmptyClipboard = source.AutoPrefillOnEmptyClipboard,
            DefaultEntryInitials = (source.DefaultEntryInitials ?? string.Empty).Trim(),
            EnableDebugLogging = source.EnableDebugLogging,
            EnableWordLifecycleLogging = source.EnableWordLifecycleLogging
        };
    }

    private static SettingsWindowModel ApplyDefaultPaths(SettingsWindowModel source)
    {
        source.ServerPath = string.IsNullOrWhiteSpace(NormalizePath(source.ServerPath))
            ? DefaultPrimaryServerPath
            : NormalizePath(source.ServerPath);

        if (string.IsNullOrWhiteSpace(NormalizePath(source.SecondaryServerPath)))
        {
            source.SecondaryServerPath = DefaultSecondaryServerPath;
            source.UseSecondaryServerPath = true;
        }
        else
        {
            source.SecondaryServerPath = NormalizePath(source.SecondaryServerPath);
        }

        source.ScheduleRootPath = string.IsNullOrWhiteSpace(NormalizePath(source.ScheduleRootPath))
            ? DefaultScheduleRootPath
            : NormalizePath(source.ScheduleRootPath);

        source.TertiaryServerPath = NormalizePath(source.TertiaryServerPath);
        return source;
    }

    private SettingsWindowModel SnapshotCurrentModel()
    {
        return new SettingsWindowModel
        {
            ServerPath = NormalizePath(ServerPath),
            UseSecondaryServerPath = UseSecondaryServerPath,
            SecondaryServerPath = NormalizePath(SecondaryServerPath),
            UseTertiaryServerPath = UseTertiaryServerPath,
            TertiaryServerPath = NormalizePath(TertiaryServerPath),
            ScheduleRootPath = NormalizePath(ScheduleRootPath),
            ShowParticipantInitials = ShowParticipantInitials,
            ShowBtnOdoo = ShowBtnOdoo,

            ShowBtnOrdner = ShowBtnOrdner,
            ShowBtnAkte = ShowBtnAkte,
            ShowBtnBu = ShowBtnBu,
            ShowBtnEintrag = ShowBtnEintrag,
            ShowBtnBi = ShowBtnBi,
            ShowBtnBe = ShowBtnBe,
            ShowBtnEintragBi = ShowBtnEintragBi,
            IsDarkTheme = IsDarkTheme,
            DisplayDensity = DisplayDensity,
            AutoPrefillOnEmptyClipboard = AutoPrefillOnEmptyClipboard,
            DefaultEntryInitials = (DefaultEntryInitials ?? string.Empty).Trim(),
            EnableDebugLogging = EnableDebugLogging,
            EnableWordLifecycleLogging = EnableWordLifecycleLogging
        };
    }

    private static bool AreEquivalent(SettingsWindowModel left, SettingsWindowModel right)
    {
        return string.Equals(NormalizePath(left.ServerPath), NormalizePath(right.ServerPath), StringComparison.OrdinalIgnoreCase)
               && left.UseSecondaryServerPath == right.UseSecondaryServerPath
               && string.Equals(NormalizePath(left.SecondaryServerPath), NormalizePath(right.SecondaryServerPath), StringComparison.OrdinalIgnoreCase)
               && left.UseTertiaryServerPath == right.UseTertiaryServerPath
               && string.Equals(NormalizePath(left.TertiaryServerPath), NormalizePath(right.TertiaryServerPath), StringComparison.OrdinalIgnoreCase)
               && string.Equals(NormalizePath(left.ScheduleRootPath), NormalizePath(right.ScheduleRootPath), StringComparison.OrdinalIgnoreCase)
               && left.ShowParticipantInitials == right.ShowParticipantInitials
               && left.ShowBtnOdoo == right.ShowBtnOdoo
               && left.ShowBtnOrdner == right.ShowBtnOrdner
               && left.ShowBtnAkte == right.ShowBtnAkte
               && left.ShowBtnBu == right.ShowBtnBu
               && left.ShowBtnEintrag == right.ShowBtnEintrag
               && left.ShowBtnBi == right.ShowBtnBi
               && left.ShowBtnBe == right.ShowBtnBe
               && left.ShowBtnEintragBi == right.ShowBtnEintragBi
               && left.IsDarkTheme == right.IsDarkTheme
               && DisplayDensityMode.Normalize(left.DisplayDensity) == DisplayDensityMode.Normalize(right.DisplayDensity)
               && left.AutoPrefillOnEmptyClipboard == right.AutoPrefillOnEmptyClipboard
               && string.Equals((left.DefaultEntryInitials ?? string.Empty).Trim(), (right.DefaultEntryInitials ?? string.Empty).Trim(), StringComparison.Ordinal)
               && left.EnableDebugLogging == right.EnableDebugLogging
               && left.EnableWordLifecycleLogging == right.EnableWordLifecycleLogging;
    }

    private void UpdateDirtyState()
    {
        if (_isInitializing)
        {
            return;
        }

        IsDirty = !AreEquivalent(_initialModel, SnapshotCurrentModel());
    }

    private void NotifyPathStateChanged()
    {
        OnPropertyChanged(nameof(PrimaryPathStatusText));
        OnPropertyChanged(nameof(ShowPrimaryPathValidation));
        OnPropertyChanged(nameof(PrimaryPathValidationText));
        OnPropertyChanged(nameof(SecondaryPathStatusText));
        OnPropertyChanged(nameof(ShowSecondaryPathValidation));
        OnPropertyChanged(nameof(SecondaryPathValidationText));
        OnPropertyChanged(nameof(TertiaryPathStatusText));
        OnPropertyChanged(nameof(ShowTertiaryPathValidation));
        OnPropertyChanged(nameof(TertiaryPathValidationText));
        OnPropertyChanged(nameof(SchedulePathStatusText));
        OnPropertyChanged(nameof(ShowSchedulePathValidation));
        OnPropertyChanged(nameof(SchedulePathValidationText));
    }

    private void BrowseServerPath_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog
        {
            Title = "Serverpfad auswählen",
            FolderName = ServerPath
        };

        if (dialog.ShowDialog() == true)
        {
            ServerPath = dialog.FolderName;
        }
    }

    private void DarkThemeRadio_OnChecked(object sender, RoutedEventArgs e)
    {
        IsDarkTheme = true;
    }

    private void LightThemeRadio_OnChecked(object sender, RoutedEventArgs e)
    {
        IsDarkTheme = false;
    }

    private void BrowseSecondaryServerPath_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog
        {
            Title = "Zusätzlichen Serverpfad auswählen",
            FolderName = SecondaryServerPath
        };

        if (dialog.ShowDialog() == true)
        {
            SecondaryServerPath = dialog.FolderName;
        }
    }

    private void BrowseTertiaryServerPath_OnClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog
        {
            Title = "Dritten Serverpfad auswählen",
            FolderName = TertiaryServerPath
        };

        if (dialog.ShowDialog() == true)
        {
            TertiaryServerPath = dialog.FolderName;
        }
    }

    private void BrowseScheduleRootPath_OnClick(object sender, RoutedEventArgs e)
    {
        var currentPath = NormalizePath(ScheduleRootPath);
        if (File.Exists(currentPath))
        {
            var dialog = new OpenFileDialog
            {
                Title = "Stundenplan auswählen",
                Filter = "Word-Dokumente (*.docx)|*.docx|Alle Dateien (*.*)|*.*",
                FileName = currentPath
            };

            if (dialog.ShowDialog(this) == true)
            {
                ScheduleRootPath = dialog.FileName;
            }

            return;
        }

        var folderDialog = new OpenFolderDialog
        {
            Title = "Stundenplan-Ordner auswählen",
            FolderName = currentPath
        };

        if (folderDialog.ShowDialog() == true)
        {
            ScheduleRootPath = folderDialog.FolderName;
        }
    }

    private void CancelButton_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private void OpenLogFolder_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var logDirectory = Services.AppLogger.LogDirectoryPath;
            Directory.CreateDirectory(logDirectory);
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = logDirectory,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Log-Ordner konnte nicht geöffnet werden: {ex.Message}", "Einstellungen", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void SaveButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (!IsPrimaryPathValid)
        {
            MessageBox.Show(this, PrimaryPathValidationText, "Einstellungen", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (UseSecondaryServerPath && !IsSecondaryPathValid)
        {
            MessageBox.Show(this, SecondaryPathValidationText, "Einstellungen", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (UseTertiaryServerPath && !IsTertiaryPathValid)
        {
            MessageBox.Show(this, TertiaryPathValidationText, "Einstellungen", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!IsSchedulePathValid)
        {
            MessageBox.Show(this, SchedulePathValidationText, "Einstellungen", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        Result = new SettingsWindowResult
        {
            ServerPath = NormalizePath(ServerPath),
            UseSecondaryServerPath = UseSecondaryServerPath,
            SecondaryServerPath = NormalizePath(SecondaryServerPath),
            UseTertiaryServerPath = UseTertiaryServerPath,
            TertiaryServerPath = NormalizePath(TertiaryServerPath),
            ScheduleRootPath = NormalizePath(ScheduleRootPath),
            ShowParticipantInitials = ShowParticipantInitials,
            ShowBtnOdoo = ShowBtnOdoo,

            ShowBtnOrdner = ShowBtnOrdner,
            ShowBtnAkte = ShowBtnAkte,
            ShowBtnBu = ShowBtnBu,
            ShowBtnEintrag = ShowBtnEintrag,
            ShowBtnBi = ShowBtnBi,
            ShowBtnBe = ShowBtnBe,
            ShowBtnEintragBi = ShowBtnEintragBi,
            IsDarkTheme = IsDarkTheme,
            DisplayDensity = DisplayDensity,
            AutoPrefillOnEmptyClipboard = AutoPrefillOnEmptyClipboard,
            DefaultEntryInitials = (DefaultEntryInitials ?? string.Empty).Trim(),
            EnableDebugLogging = EnableDebugLogging,
            EnableWordLifecycleLogging = EnableWordLifecycleLogging
        };

        _isSaving = true;
        DialogResult = true;
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (!_isSaving)
        {
            App.ApplyTheme(_initialIsDarkTheme);
        }

        base.OnClosing(e);
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);

        if (propertyName == nameof(IsDarkTheme))
        {
            App.ApplyTheme(_isDarkTheme);
        }

        if (propertyName is nameof(ServerPath) or nameof(SecondaryServerPath) or nameof(UseSecondaryServerPath) or nameof(TertiaryServerPath) or nameof(UseTertiaryServerPath) or nameof(ScheduleRootPath))
        {
            NotifyPathStateChanged();
        }

        UpdateDirtyState();
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}





