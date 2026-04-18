using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Threading;
using VerlaufsakteApp.Models;
using VerlaufsakteApp.Services;

namespace VerlaufsakteApp;

public partial class MainWindow : Window
{
    private readonly record struct MonitorDescriptor(string DeviceName, Rect WorkArea, bool IsPrimary);
    private readonly record struct DisplayDensityProfile(
        string Mode,
        double WindowFontSize,
        double NameFontSize,
        double SecondaryFontSize,
        double ActionButtonFontSize,
        double PrimaryButtonFontSize,
        double ActionButtonMinHeight,
        double PrimaryButtonMinHeight,
        double ActionRowItemHeight,
        double SubtleIconButtonSize,
        double SubtleIconFontSize,
        Thickness CardPadding,
        Thickness CardMargin,
        Thickness ActionButtonPadding,
        Thickness ActionButtonMargin,
        Thickness PrimaryButtonPadding,
        Thickness PresenceCheckBoxMargin,
        Thickness CollapsedStripPadding,
        Thickness AbsenceStatusMargin,
        Thickness NotFoundHintMargin,
        Thickness WideActionRowMargin,
        Thickness BatchToggleMargin,
        Thickness BatchPanelPadding,
        Thickness BatchDescriptionMargin,
        Thickness BatchButtonsMargin,
        Thickness BatchProgressMargin);

    private const double MinExpandedInputHeight = 190;
    private const double MaxExpandedInputHeight = 340;
    private const double InputHeightScaleFactor = 1.2;
    private const double LeftNamePrefixWidth = 44;
    private const double WideLayoutContainerPaddingWidth = 36;
    private const double WideLayoutSafetyWidth = 12;
    private const double TitleBarHeight = 36;
    private const double StatusBarHeight = 24;
    private const double OuterBorderHeight = 2;
    private const double MinResultsWindowHeightAllowance = 120;
    private const double CollapseAnimationDurationMilliseconds = 170;
    private static readonly TimeSpan WeeklyScheduleRetryInterval = TimeSpan.FromMinutes(3);
    private static readonly string ApplicationVersionText = BuildApplicationVersionText();

    private readonly AppConfig _config;
    private readonly ParticipantParser _parser;
    private FolderMatcher _matcher;
    private readonly WordService _wordService;
    private readonly WordStaHost _wordStaHost;
    private readonly InitialsResolver _initialsResolver;
    private readonly DocxHeaderMetadataService _headerMetadataService;
    private readonly WeeklyScheduleService _weeklyScheduleService;
    private readonly AppUpdateService _appUpdateService;
    private readonly Dictionary<Participant, Border> _miniScheduleTrayHosts = new();
    private bool _isInputCollapsed;
    private double _expandedInputHeight = 250;
    private string _lastActionText = "Bereit";
    private bool _isBatchCollapsed = true;
    private bool _isBiTodoCollapsed = true;
    private readonly DispatcherTimer _layoutDebounceTimer;
    private readonly DispatcherTimer _weeklyScheduleRetryTimer;
    private CancellationTokenSource? _batchCancellation;
    private bool _isBatchRunning;
    private bool _isBiTodoRunning;
    private bool _isWordActionRunning;
    private bool _isEvaluating;
    private bool _isWindowCollapsed;
    private bool _isCollapseAnimationRunning;
    private int _collapseAnimationVersion;
    private bool _collapseAutoExpandArmed;
    private Participant? _expandedMiniScheduleParticipant;
    private readonly double _expandedWindowMinHeight;
    private readonly ResizeMode _defaultResizeMode;
    private double _preferredExpandedWindowHeight;
    private DisplayDensityProfile _displayDensityProfile = CreateDisplayDensityProfile(DisplayDensityMode.Standard);
    private bool _startupUpdateCheckStarted;
    private bool _isUpdateShutdownRequested;
    private bool _isWeeklyScheduleRetryRunning;

    public static readonly DependencyProperty ShowParticipantInitialsProperty = RegisterLayoutProperty(nameof(ShowParticipantInitials), true);
    public static readonly DependencyProperty ShowBtnOdooProperty = RegisterLayoutProperty(nameof(ShowBtnOdoo), false);
    public static readonly DependencyProperty ShowBtnOrdnerProperty = RegisterLayoutProperty(nameof(ShowBtnOrdner), true);
    public static readonly DependencyProperty ShowBtnAkteProperty = RegisterLayoutProperty(nameof(ShowBtnAkte), true);
    public static readonly DependencyProperty ShowBtnBuProperty = RegisterLayoutProperty(nameof(ShowBtnBu), true);
    public static readonly DependencyProperty ShowBtnEintragProperty = RegisterLayoutProperty(nameof(ShowBtnEintrag), true);
    public static readonly DependencyProperty ShowBtnBiProperty = RegisterLayoutProperty(nameof(ShowBtnBi), false);
    public static readonly DependencyProperty ShowBtnBeProperty = RegisterLayoutProperty(nameof(ShowBtnBe), false);
    public static readonly DependencyProperty ShowBtnEintragBiProperty = RegisterLayoutProperty(nameof(ShowBtnEintragBi), false);
    public static readonly DependencyProperty IsDarkThemeProperty = RegisterLayoutProperty(nameof(IsDarkTheme), true);

    public bool ShowParticipantInitials
    {
        get => (bool)GetValue(ShowParticipantInitialsProperty);
        set => SetValue(ShowParticipantInitialsProperty, value);
    }

    public bool ShowBtnOdoo
    {
        get => (bool)GetValue(ShowBtnOdooProperty);
        set => SetValue(ShowBtnOdooProperty, value);
    }

    public bool ShowBtnOrdner
    {
        get => (bool)GetValue(ShowBtnOrdnerProperty);
        set => SetValue(ShowBtnOrdnerProperty, value);
    }

    public bool ShowBtnAkte
    {
        get => (bool)GetValue(ShowBtnAkteProperty);
        set => SetValue(ShowBtnAkteProperty, value);
    }

    public bool ShowBtnBu
    {
        get => (bool)GetValue(ShowBtnBuProperty);
        set => SetValue(ShowBtnBuProperty, value);
    }

    public bool ShowBtnEintrag
    {
        get => (bool)GetValue(ShowBtnEintragProperty);
        set => SetValue(ShowBtnEintragProperty, value);
    }

    public bool ShowBtnBi
    {
        get => (bool)GetValue(ShowBtnBiProperty);
        set => SetValue(ShowBtnBiProperty, value);
    }

    public bool ShowBtnBe
    {
        get => (bool)GetValue(ShowBtnBeProperty);
        set => SetValue(ShowBtnBeProperty, value);
    }

    public bool ShowBtnEintragBi
    {
        get => (bool)GetValue(ShowBtnEintragBiProperty);
        set => SetValue(ShowBtnEintragBiProperty, value);
    }

    public bool IsDarkTheme
    {
        get => (bool)GetValue(IsDarkThemeProperty);
        set => SetValue(IsDarkThemeProperty, value);
    }

    public MainWindow()
    {
        InitializeComponent();
        _expandedWindowMinHeight = MinHeight;
        _defaultResizeMode = ResizeMode;

        _config = App.Config ?? throw new InvalidOperationException("Konfiguration nicht geladen.");
        _parser = new ParticipantParser(_config.AbsenceValues, _config.PresenceValues);
        _initialsResolver = new InitialsResolver();
        _matcher = CreateMatcher();
        _wordService = new WordService();
        _wordStaHost = new WordStaHost();
        _headerMetadataService = new DocxHeaderMetadataService(
            Path.Combine(App.AppDataDirectoryPath, "header-metadata-cache.json"),
            Path.Combine(App.AppDataDirectoryPath, "header-metadata-cache.bak"));
        _weeklyScheduleService = new WeeklyScheduleService(
            App.WeeklyScheduleCachePath,
            App.WeeklyScheduleCacheBackupPath);
        _appUpdateService = new AppUpdateService();
        _layoutDebounceTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(60)
        };
        _layoutDebounceTimer.Tick += (_, _) =>
        {
            _layoutDebounceTimer.Stop();
            UpdateDynamicMinWidth();
        };
        _weeklyScheduleRetryTimer = new DispatcherTimer
        {
            Interval = WeeklyScheduleRetryInterval
        };
        _weeklyScheduleRetryTimer.Tick += WeeklyScheduleRetryTimer_OnTick;

        Participants = new ObservableCollection<Participant>();
        BatchResults = new ObservableCollection<BatchResult>();
        BiTodoResults = new ObservableCollection<BiTodoCollectResult>();
        Participants.CollectionChanged += Participants_OnCollectionChanged;
        BatchResults.CollectionChanged += BatchResults_OnCollectionChanged;
        BiTodoResults.CollectionChanged += BiTodoResults_OnCollectionChanged;

        ShowParticipantInitials = App.UserPrefs.ShowParticipantInitials;
        ShowBtnOdoo = App.UserPrefs.ShowBtnOdoo;
        ShowBtnOrdner = App.UserPrefs.ShowBtnOrdner;
        ShowBtnAkte = App.UserPrefs.ShowBtnAkte;
        ShowBtnBu = App.UserPrefs.ShowBtnBu;
        ShowBtnEintrag = App.UserPrefs.ShowBtnEintrag;
        ShowBtnBi = App.UserPrefs.ShowBtnBi;
        ShowBtnBe = App.UserPrefs.ShowBtnBe;
        ShowBtnEintragBi = App.UserPrefs.ShowBtnEintragBi;
        IsDarkTheme = App.UserPrefs.IsDarkTheme;

        DataContext = this;
        RestoreWindowBoundsFromPrefs();
        _preferredExpandedWindowHeight = Height;
        _isWindowCollapsed = App.UserPrefs.IsCollapsed;
        ApplyDisplayDensity(App.UserPrefs.DisplayDensity);

        UpdateExpandedInputHeight();
        InputContainer.Height = _expandedInputHeight;
        InputContainer.Visibility = Visibility.Visible;
        CollapsedStrip.Visibility = Visibility.Collapsed;
        SetResultsAreaVisible(false);
        EvaluateButton.IsEnabled = false;
        ApplyCompactStartupWindowState();
        ApplyWindowCollapseState(animated: false);
        StatusBarVersion.Text = ApplicationVersionText;
        Loaded += MainWindow_OnLoaded;
        WordService.TryCleanupBiTodoTempArtifactsOnStartup();
        _appUpdateService.TryCleanupSuccessfulUpdateArtifactsOnStartup();

        AppLogger.Info($"MainWindow init. IsWordAvailable={IsWordAvailable}, ServerBasePath='{_config.ServerBasePath}', UseSecondaryServerBasePath={_config.UseSecondaryServerBasePath}, SecondaryServerBasePath='{_config.SecondaryServerBasePath}'");

        if (!IsWordAvailable)
        {
            ShowToast("Microsoft Word nicht gefunden. Stufe 3 nicht verfügbar.", ToastType.Info);
            SetLastAction("Microsoft Word nicht gefunden. Stufe 3 deaktiviert.");
            AppLogger.Warn("Microsoft Word nicht gefunden. Stufe 3 deaktiviert.");
        }

        UpdateCollapseToggleButton();
        RequestDynamicMinWidthRefresh();
    }

    public ObservableCollection<Participant> Participants { get; }
    public ObservableCollection<BatchResult> BatchResults { get; }
    public ObservableCollection<BiTodoCollectResult> BiTodoResults { get; }
    public bool IsWordAvailable => _wordService.IsWordAvailable;

    // --- Window chrome ---

    private void MainWindow_OnSizeChanged(object sender, SizeChangedEventArgs e)
    {
        UpdateExpandedInputHeight();
        if (!_isInputCollapsed)
        {
            InputContainer.Height = _expandedInputHeight;
        }

        if (WindowState == WindowState.Normal
            && SizeToContent == SizeToContent.Manual
            && !_isWindowCollapsed
            && !_isCollapseAnimationRunning
            && Height > GetCollapsedWindowHeight())
        {
            _preferredExpandedWindowHeight = Height;
        }

        SyncBodyHeightToWindow();
    }

    private void TitleBar_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2)
        {
            ToggleWindowCollapse();
            return;
        }

        if (e.LeftButton != MouseButtonState.Pressed)
        {
            return;
        }

        if (WindowState == WindowState.Maximized)
        {
            // Beim Ziehen aus maximiertem Zustand sauber auf Fensterzustand wechseln.
            var mouseInWindow = e.GetPosition(this);
            var horizontalPercent = mouseInWindow.X / Math.Max(1, ActualWidth);

            var mouseOnScreenDevice = PointToScreen(mouseInWindow);
            var source = PresentationSource.FromVisual(this);
            var mouseOnScreen = source?.CompositionTarget is not null
                ? source.CompositionTarget.TransformFromDevice.Transform(mouseOnScreenDevice)
                : mouseOnScreenDevice;

            WindowState = WindowState.Normal;
            Left = mouseOnScreen.X - (RestoreBounds.Width * horizontalPercent);
            Top = Math.Max(0, mouseOnScreen.Y - 16);
        }

        DragMove();
    }

    private void MinimizeButton_OnClick(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
    private void CollapseToggleButton_OnClick(object sender, RoutedEventArgs e) => ToggleWindowCollapse();
    private void CloseButton_OnClick(object sender, RoutedEventArgs e) => Close();

    private void MainWindow_OnActivated(object? sender, EventArgs e)
    {
        if (_collapseAutoExpandArmed
            && _isWindowCollapsed
            && WindowState == WindowState.Normal)
        {
            _collapseAutoExpandArmed = false;
            SetWindowCollapsed(false, animated: true);
        }
    }

    private void MainWindow_OnDeactivated(object? sender, EventArgs e)
    {
        if (_isWindowCollapsed && WindowState == WindowState.Normal)
        {
            _collapseAutoExpandArmed = true;
        }
    }

    private void MainWindow_OnClosing(object? sender, CancelEventArgs e)
    {
        if (_isBatchRunning)
        {
            _batchCancellation?.Cancel();
            e.Cancel = true;
            ShowToast("Batch wird abgebrochen. Danach kann das Fenster geschlossen werden.", ToastType.Warning);
            SetLastAction("Schliessen angehalten: laufender Batch wird abgebrochen.");
            return;
        }

        if (_isBiTodoRunning)
        {
            e.Cancel = true;
            ShowToast("BI: To-dos läuft noch. Bitte kurz warten.", ToastType.Warning);
            SetLastAction("Schliessen angehalten: BI: To-dos läuft noch.");
            return;
        }

        if (_isWordActionRunning)
        {
            e.Cancel = true;
            ShowToast("Word-Aktion läuft noch. Bitte kurz warten.", ToastType.Warning);
            SetLastAction("Schliessen angehalten: Word-Aktion läuft noch.");
            return;
        }

        _weeklyScheduleRetryTimer.Stop();
        SaveWindowBoundsToPrefs();
        _wordStaHost.Dispose();
    }

    private async void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
    {
        RequestDynamicMinWidthRefresh();

        if (_startupUpdateCheckStarted)
        {
            return;
        }

        _startupUpdateCheckStarted = true;

        try
        {
            await Task.Yield();
            await BeginStartupUpdateCheckAsync();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Updater: Startup-Check fehlgeschlagen: {ex.Message}");
        }

        UpdateWeeklyScheduleRetryTimerState();
        await EnsureCurrentWeekScheduleLoadedAsync("Startup");
    }

    private async void WeeklyScheduleRetryTimer_OnTick(object? sender, EventArgs e)
    {
        await EnsureCurrentWeekScheduleLoadedAsync("Timer");
    }

    private void UpdateWeeklyScheduleRetryTimerState()
    {
        if (!IsLoaded || _isUpdateShutdownRequested)
        {
            _weeklyScheduleRetryTimer.Stop();
            return;
        }

        if (string.IsNullOrWhiteSpace(_config.ScheduleRootPath))
        {
            if (_weeklyScheduleRetryTimer.IsEnabled)
            {
                _weeklyScheduleRetryTimer.Stop();
                AppLogger.Debug("MiniScheduleRetry: Timer gestoppt, kein Stundenplanpfad konfiguriert.");
            }

            return;
        }

        if (!_weeklyScheduleRetryTimer.IsEnabled)
        {
            _weeklyScheduleRetryTimer.Start();
            AppLogger.Debug($"MiniScheduleRetry: Timer gestartet. Intervall='{WeeklyScheduleRetryInterval}'.");
        }
    }

    private async Task EnsureCurrentWeekScheduleLoadedAsync(string trigger)
    {
        if (!IsLoaded || _isUpdateShutdownRequested || _isWeeklyScheduleRetryRunning)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(_config.ScheduleRootPath))
        {
            UpdateWeeklyScheduleRetryTimerState();
            return;
        }

        if (Participants.Any(participant => participant.IsMiniScheduleLoading))
        {
            AppLogger.Debug($"MiniScheduleRetry: Uebersprungen ({trigger}), anderer Stundenplan-Ladevorgang aktiv.");
            return;
        }

        var statusBefore = _weeklyScheduleService.GetCurrentWeekLoadStatus(_config.ScheduleRootPath);
        if (statusBefore.IsCurrentWeekSuccessfullyLoaded)
        {
            UpdateWeeklyScheduleRetryTimerState();
            return;
        }

        _isWeeklyScheduleRetryRunning = true;

        try
        {
            AppLogger.Debug($"MiniScheduleRetry: Versuch gestartet ({trigger}). Path='{statusBefore.ResolvedDocumentPath ?? string.Empty}', LastFailureUtc='{statusBefore.LastFailureUtc?.ToString("O") ?? string.Empty}'.");

            var statusAfter = await Task.Run(() => _weeklyScheduleService.TryWarmCurrentWeekDocument(_config.ScheduleRootPath));
            if (statusAfter.IsCurrentWeekSuccessfullyLoaded)
            {
                AppLogger.Info($"MiniScheduleRetry: Aktuelle Wochenplanung erfolgreich geladen. Path='{statusAfter.ResolvedDocumentPath ?? string.Empty}'.");
            }
            else
            {
                AppLogger.Debug($"MiniScheduleRetry: Noch kein erfolgreicher Lesezugriff ({trigger}). Path='{statusAfter.ResolvedDocumentPath ?? string.Empty}', LastFailureUtc='{statusAfter.LastFailureUtc?.ToString("O") ?? string.Empty}'.");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"MiniScheduleRetry: Fehler beim Hintergrund-Check ({trigger}): {ex.Message}");
        }
        finally
        {
            _isWeeklyScheduleRetryRunning = false;
            UpdateWeeklyScheduleRetryTimerState();
        }
    }

    private void MainWindow_OnStateChanged(object? sender, EventArgs e)
    {
        if (WindowState == WindowState.Maximized && _isWindowCollapsed)
        {
            _isWindowCollapsed = false;
            _collapseAutoExpandArmed = false;
        }

        ApplyWindowCollapseState(animated: false);
    }

    private void ToggleWindowCollapse()
    {
        if (WindowState == WindowState.Minimized)
        {
            return;
        }

        SetWindowCollapsed(!_isWindowCollapsed, animated: true);
    }

    private void SetWindowCollapsed(bool collapsed, bool animated)
    {
        if (WindowState == WindowState.Minimized)
        {
            return;
        }

        if (!collapsed && WindowState == WindowState.Maximized)
        {
            _isWindowCollapsed = false;
            ApplyWindowCollapseState(animated: false);
            return;
        }

        if (collapsed && WindowState == WindowState.Maximized)
        {
            WindowState = WindowState.Normal;
        }

        if (!collapsed)
        {
            _preferredExpandedWindowHeight = Math.Max(GetExpandedWindowHeightTarget(), _expandedWindowMinHeight);
        }

        _collapseAutoExpandArmed = false;
        _isWindowCollapsed = collapsed;
        ApplyWindowCollapseState(animated);
    }

    private void ApplyWindowCollapseState(bool animated)
    {
        UpdateCollapseToggleButton();
        UpdateMinimumWindowHeightForCurrentContent();

        if (WindowBodyContainer is null)
        {
            return;
        }

        if (WindowState == WindowState.Maximized)
        {
            StopCollapseAnimations();
            _isCollapseAnimationRunning = false;
            ResizeMode = _defaultResizeMode;
            SizeToContent = SizeToContent.Manual;
            WindowBodyContainer.BeginAnimation(FrameworkElement.HeightProperty, null);
            WindowBodyContainer.Height = double.NaN;
            return;
        }

        var targetWindowHeight = _isWindowCollapsed
            ? GetCollapsedWindowHeight()
            : GetExpandedWindowHeightTarget();
        var targetBodyHeight = _isWindowCollapsed
            ? 0
            : Math.Max(0, targetWindowHeight - GetCollapsedWindowHeight());

        if (!animated || !IsLoaded)
        {
            StopCollapseAnimations();
            _isCollapseAnimationRunning = false;
            SizeToContent = SizeToContent.Manual;
            ResizeMode = _isWindowCollapsed ? ResizeMode.NoResize : _defaultResizeMode;
            if (WindowState == WindowState.Normal)
            {
                Height = targetWindowHeight;
            }

            WindowBodyContainer.BeginAnimation(FrameworkElement.HeightProperty, null);
            WindowBodyContainer.Height = targetBodyHeight;
            return;
        }

        StopCollapseAnimations();
        _isCollapseAnimationRunning = true;
        SizeToContent = SizeToContent.Manual;
        ResizeMode = _isWindowCollapsed ? ResizeMode.NoResize : _defaultResizeMode;

        var animationVersion = ++_collapseAnimationVersion;
        var startWindowHeight = Math.Max(ActualHeight, Height);
        var startBodyHeight = GetCurrentBodyHeight();

        Height = startWindowHeight;
        WindowBodyContainer.Height = startBodyHeight;

        var easing = new CubicEase { EasingMode = EasingMode.EaseOut };
        var bodyAnimation = new DoubleAnimation
        {
            From = startBodyHeight,
            To = targetBodyHeight,
            Duration = TimeSpan.FromMilliseconds(CollapseAnimationDurationMilliseconds),
            EasingFunction = easing
        };
        bodyAnimation.Completed += (_, _) =>
        {
            if (animationVersion != _collapseAnimationVersion)
            {
                return;
            }

            _isCollapseAnimationRunning = false;
            SizeToContent = SizeToContent.Manual;
            WindowBodyContainer.BeginAnimation(FrameworkElement.HeightProperty, null);
            WindowBodyContainer.Height = targetBodyHeight;
            if (WindowState == WindowState.Normal)
            {
                Height = targetWindowHeight;
            }
        };

        var windowAnimation = new DoubleAnimation
        {
            From = startWindowHeight,
            To = targetWindowHeight,
            Duration = TimeSpan.FromMilliseconds(CollapseAnimationDurationMilliseconds),
            EasingFunction = easing
        };

        WindowBodyContainer.BeginAnimation(FrameworkElement.HeightProperty, bodyAnimation);
        BeginAnimation(HeightProperty, windowAnimation);
    }

    private void StopCollapseAnimations()
    {
        _collapseAnimationVersion++;
        BeginAnimation(HeightProperty, null);
        WindowBodyContainer?.BeginAnimation(FrameworkElement.HeightProperty, null);
    }

    private void UpdateCollapseToggleButton()
    {
        if (CollapseToggleButton is null)
        {
            return;
        }

        CollapseToggleButton.Content = _isWindowCollapsed ? "▼" : "▲";
        CollapseToggleButton.ToolTip = _isWindowCollapsed ? "Ausklappen" : "Einklappen";
        CollapseToggleButton.SetResourceReference(Control.BackgroundProperty, "Brush.PanelBg");
        CollapseToggleButton.SetResourceReference(Control.BorderBrushProperty, "Brush.Border");
        CollapseToggleButton.SetResourceReference(Control.ForegroundProperty, "Brush.SecondaryText");
    }

    private void SyncBodyHeightToWindow()
    {
        if (WindowBodyContainer is null || _isCollapseAnimationRunning)
        {
            return;
        }

        if (WindowState == WindowState.Maximized)
        {
            WindowBodyContainer.BeginAnimation(FrameworkElement.HeightProperty, null);
            WindowBodyContainer.Height = double.NaN;
            return;
        }

        WindowBodyContainer.BeginAnimation(FrameworkElement.HeightProperty, null);
        WindowBodyContainer.Height = _isWindowCollapsed
            ? 0
            : Math.Max(0, Math.Max(ActualHeight, Height) - GetCollapsedWindowHeight());
    }

    private void UpdateMinimumWindowHeightForCurrentContent()
    {
        if (_isWindowCollapsed)
        {
            MinHeight = GetCollapsedWindowHeight();
            return;
        }

        MinHeight = ResultsContainer.Visibility == Visibility.Visible
            ? _expandedWindowMinHeight
            : 0;
    }

    private double GetCollapsedWindowHeight() => TitleBarHeight + OuterBorderHeight;

    private double GetExpandedWindowHeightTarget()
    {
        var target = CalculateCompactWindowHeight();
        if (ResultsContainer.Visibility == Visibility.Visible)
        {
            target = Math.Max(
                Math.Max(Math.Max(target, _preferredExpandedWindowHeight), _expandedWindowMinHeight),
                CalculateCompactWindowHeight() + CalculateResultsWindowAllowance());
        }

        return Math.Ceiling(target);
    }

    private double GetCurrentBodyHeight()
    {
        if (WindowBodyContainer is null)
        {
            return 0;
        }

        if (!double.IsNaN(WindowBodyContainer.Height))
        {
            return Math.Max(0, WindowBodyContainer.Height);
        }

        return Math.Max(0, Math.Max(ActualHeight, Height) - GetCollapsedWindowHeight());
    }

    // --- Input area ---

    private void InputTextBox_OnTextChanged(object sender, TextChangedEventArgs e)
    {
        PlaceholderText.Visibility = string.IsNullOrWhiteSpace(InputTextBox.Text)
            ? Visibility.Visible
            : Visibility.Collapsed;
        EvaluateButton.IsEnabled = !_isEvaluating && !string.IsNullOrWhiteSpace(InputTextBox.Text);
    }

    private void ClearInputButton_OnClick(object sender, RoutedEventArgs e)
    {
        InputTextBox.Clear();
        InputTextBox.Focus();
        SetLastAction("Eingabefeld geleert");
    }

    private void ClearParticipantsButton_OnClick(object sender, RoutedEventArgs e)
    {
        InputContainer.BeginAnimation(HeightProperty, null);
        ResultsContainer.BeginAnimation(OpacityProperty, null);
        ResetMiniScheduleForAllParticipants(immediate: true);
        Participants.Clear();
        BatchResults.Clear();
        BiTodoResults.Clear();
        BatchPanelContainer.Visibility = Visibility.Collapsed;
        BiTodoPanelContainer.Visibility = Visibility.Collapsed;
        BatchProgressText.Text = "Fortschritt: —";
        BiTodoProgressText.Text = "Fortschritt: —";
        BiTodoProgressPanel.Visibility = Visibility.Collapsed;
        SetResultsAreaVisible(false);
        CollapsedStrip.Visibility = Visibility.Collapsed;
        InputContainer.Visibility = Visibility.Visible;
        InputContainer.Height = _expandedInputHeight;
        ResultsContainer.Opacity = 1;
        _isInputCollapsed = false;
        _isBatchCollapsed = true;
        _isBiTodoCollapsed = true;
        BatchStripText.Text = "▶ BU: Batch-Eintrag ⚡︎";
        BiTodoStripText.Text = "▶ BI: To-dos ⚡︎";
        EvaluateButton.IsEnabled = !string.IsNullOrWhiteSpace(InputTextBox.Text);
        ApplyCompactStartupWindowState();
        InputTextBox.Focus();
        SetLastAction("Namensliste gelöscht");
        ShowToast("Namensliste gelöscht", ToastType.Info);
    }

    private void SetEvaluateBusy(bool isBusy)
    {
        _isEvaluating = isBusy;
        EvaluateButton.Content = isBusy ? "Auswerten..." : "Auswerten";
        EvaluateProgressPanel.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;
        EvaluateProgressText.Text = isBusy ? "Wird ausgewertet..." : "Wird ausgewertet...";
        EvaluateButton.IsEnabled = !isBusy && !string.IsNullOrWhiteSpace(InputTextBox.Text);
        ClearInputButton.IsEnabled = !isBusy;
        InputTextBox.IsReadOnly = isBusy;

        if (isBusy)
        {
            StatusBarText.Text = "Liste wird ausgewertet...";
            StatusBarTimestamp.Text = DateTime.Now.ToString("HH:mm:ss");
        }
    }

    private async void EvaluateButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isEvaluating)
        {
            return;
        }

        var text = InputTextBox.Text;
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        try
        {
            SetEvaluateBusy(true);

            if (!HasAtLeastOneReachableServerPath())
            {
                var message = BuildServerPathErrorMessage();
                AppLogger.Warn(message);
                ShowToast(message, ToastType.Error);
                SetLastAction(message);
                return;
            }

            var evaluation = await Task.Run(() =>
            {
                var matcher = CreateMatcher();
                matcher.BuildIndex();
                var parsedParticipants = _parser.Parse(text, rawLine => matcher.ResolveLikelyNameFromRawLine(rawLine));
                foreach (var participant in parsedParticipants)
                {
                    matcher.MatchParticipant(participant);
                }

                return (Matcher: matcher, Participants: parsedParticipants);
            });

            _matcher = evaluation.Matcher;
            var parsedParticipants = evaluation.Participants;
            if (parsedParticipants.Count == 0)
            {
                ShowToast("Keine Teilnehmer erkannt.", ToastType.Warning);
                SetLastAction("Keine Teilnehmer erkannt.");
                return;
            }

            foreach (var participant in parsedParticipants)
            {
                UpdateActionState(participant);
            }

            ResetMiniScheduleForAllParticipants(immediate: true);
            Participants.Clear();
            foreach (var participant in parsedParticipants)
            {
                Participants.Add(participant);
            }
            RequestDynamicMinWidthRefresh();
            BeginOdooMetadataWarmupForCurrentParticipants();

            var presentCount = Participants.Count(p => p.IsPresent);
            var absentCount = Participants.Count - presentCount;

            InputContainer.Visibility = Visibility.Collapsed;
            EvaluateButton.IsEnabled = false;
            CollapsedStrip.Visibility = Visibility.Visible;
            CollapsedStripText.Text = "▶ Liste neu einfügen";
            _isInputCollapsed = true;

            // Fade in results
            ResultsContainer.Opacity = 0;
            SetResultsAreaVisible(true);
            var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(300))
            {
                EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
            };
            ResultsContainer.BeginAnimation(OpacityProperty, fadeIn);
            ExpandWindowForResults();
            _ = Dispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(EnsureActionStripsVisibleAfterEvaluate));

            SetLastAction($"Auswertung: {Participants.Count} Einträge ({presentCount} anwesend, {absentCount} abwesend)");
            ShowToast($"{Participants.Count} Einträge verarbeitet", ToastType.Success);
            AppLogger.Info($"Auswertung abgeschlossen. Eintraege={Participants.Count}, Present={presentCount}, Absent={absentCount}");

            BatchResults.Clear();
            BiTodoResults.Clear();
            BatchProgressText.Text = "Fortschritt: —";
            BiTodoProgressText.Text = "Fortschritt: —";
            if (_isBatchCollapsed)
            {
                BatchPanelContainer.Visibility = Visibility.Collapsed;
                BatchStripText.Text = "▶ BU: Batch-Eintrag ⚡︎";
            }

            if (_isBiTodoCollapsed)
            {
                BiTodoPanelContainer.Visibility = Visibility.Collapsed;
                BiTodoStripText.Text = "▶ BI: To-dos ⚡︎";
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("Fehler bei Auswertung", ex);
            ShowToast($"Fehler bei Auswertung: {ex.Message}", ToastType.Error);
            SetLastAction($"Fehler bei Auswertung: {ex.Message}");
        }
        finally
        {
            SetEvaluateBusy(false);
        }
    }

    // --- Collapse/Expand ---

    private void CollapsedStrip_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (_isInputCollapsed)
            ExpandInputArea();
        else
            CollapseInputArea();
    }

    private void CollapseInputArea()
    {
        if (_isInputCollapsed)
            return;

        _isInputCollapsed = true;
        CollapsedStrip.Visibility = Visibility.Visible;
        CollapsedStripText.Text = "▶ Liste neu einfügen";

        var from = InputContainer.ActualHeight > 0 ? InputContainer.ActualHeight : _expandedInputHeight;
        AnimateInputContainerHeight(from, 0, () => InputContainer.Visibility = Visibility.Collapsed);
    }

    private void ExpandInputArea()
    {
        if (!_isInputCollapsed)
            return;

        _isInputCollapsed = false;
        InputContainer.Visibility = Visibility.Visible;
        CollapsedStripText.Text = "▼ Liste neu einfügen";

        AnimateInputContainerHeight(0, _expandedInputHeight, null);
    }

    private void ApplyCompactStartupWindowState()
    {
        UpdateMinimumWindowHeightForCurrentContent();

        if (WindowState == WindowState.Maximized)
        {
            SyncBodyHeightToWindow();
            return;
        }

        SizeToContent = SizeToContent.Manual;
        UpdateLayout();
        Height = CalculateCompactWindowHeight();
        SyncBodyHeightToWindow();
    }

    private void RequestDynamicMinWidthRefresh()
    {
        _layoutDebounceTimer.Stop();
        _layoutDebounceTimer.Start();
    }

    private void UpdateDynamicMinWidth()
    {
        var baselineMinWidth = 280d;
        var targetMinWidth = baselineMinWidth;

        if (Participants.Count > 0)
        {
            var bounds = GetCurrentWindowBoundsForMonitorReference();
            var monitor = GetMonitorContainingPoint(
                bounds.Left + (bounds.Width / 2.0),
                bounds.Top + (bounds.Height / 2.0));
            targetMinWidth = Math.Min(
                Math.Ceiling(CalculateRequiredParticipantContentWidth()),
                Math.Max(baselineMinWidth, monitor.WorkArea.Width));
        }

        MinWidth = Math.Max(baselineMinWidth, targetMinWidth);
    }

    private Rect GetCurrentWindowBoundsForMonitorReference()
    {
        var bounds = WindowState == WindowState.Normal
            ? new Rect(Left, Top, Width, Height)
            : RestoreBounds;
        if (!bounds.IsEmpty)
        {
            return bounds;
        }

        return new Rect(Left, Top, Math.Max(Width, ActualWidth), Math.Max(Height, ActualHeight));
    }

    private double CalculateRequiredParticipantContentWidth()
    {
        return MeasureMaxNameWidth()
               + MeasureActionButtonsWidth()
               + LeftNamePrefixWidth
               + WideLayoutContainerPaddingWidth
               + WideLayoutSafetyWidth;
    }

    private void ExpandWindowForResults()
    {
        if (_isWindowCollapsed)
        {
            return;
        }

        UpdateMinimumWindowHeightForCurrentContent();
        SizeToContent = SizeToContent.Manual;

        if (WindowState == WindowState.Maximized)
        {
            return;
        }

        var targetHeight = Math.Max(
            Math.Max(_preferredExpandedWindowHeight, _expandedWindowMinHeight),
            CalculateCompactWindowHeight() + CalculateResultsWindowAllowance());
        if (ActualHeight < targetHeight - 1)
        {
            AnimateWindowHeight(Math.Max(ActualHeight, Height), targetHeight);
        }
    }

    private void EnsureActionStripsVisibleAfterEvaluate()
    {
        if (_isWindowCollapsed
            || WindowState == WindowState.Maximized
            || ResultsScrollViewer is null
            || BatchStripBorder is null
            || BatchStripBorder.Visibility != Visibility.Visible)
        {
            return;
        }

        UpdateLayout();

        var targetElement = BiTodoStripBorder is not null && BiTodoStripBorder.Visibility == Visibility.Visible
            ? (FrameworkElement)BiTodoStripBorder
            : BatchStripBorder;

        AdjustWindowToResultsElement(targetElement, 22, allowShrink: true);
    }

    private double CalculateResultsWindowAllowance()
    {
        UpdateLayout();

        var footerHeight = GetActualOrFallbackHeight(ResultsActionsFooter, 34)
                           + GetActualOrFallbackHeight(BatchStripBorder, 34)
                           + GetActualOrFallbackHeight(BiTodoStripBorder, 34);
        var paddingAllowance = _displayDensityProfile.Mode == DisplayDensityMode.Compact ? 14d : 20d;
        return Math.Ceiling(Math.Max(MinResultsWindowHeightAllowance, footerHeight + paddingAllowance));
    }

    private void AdjustWindowToResultsElement(FrameworkElement targetElement, double bottomPadding, bool allowShrink)
    {
        if (WindowState == WindowState.Maximized
            || ResultsScrollViewer is null
            || ResultsScrollViewer.ViewportHeight <= 0
            || targetElement.ActualHeight <= 0)
        {
            return;
        }

        try
        {
            var transform = targetElement.TransformToAncestor(ResultsScrollViewer);
            var bottom = transform.Transform(new Point(0, targetElement.ActualHeight)).Y;
            var delta = bottom - ResultsScrollViewer.ViewportHeight + bottomPadding;
            if (!allowShrink && delta <= 0)
            {
                return;
            }

            if (Math.Abs(delta) <= 1)
            {
                return;
            }

            var centerX = Left + (ActualWidth / 2.0);
            var centerY = Top + (ActualHeight / 2.0);
            var workArea = GetMonitorContainingPoint(centerX, centerY).WorkArea;
            var maxHeight = Math.Max(_expandedWindowMinHeight, workArea.Height);
            var minTargetHeight = Math.Max(
                _expandedWindowMinHeight,
                CalculateCompactWindowHeight() + CalculateResultsWindowAllowance());
            var targetHeight = Math.Min(
                maxHeight,
                Math.Ceiling(Math.Max(ActualHeight, Height) + delta));
            targetHeight = Math.Max(minTargetHeight, targetHeight);

            if (allowShrink)
            {
                if (Math.Abs(targetHeight - Math.Max(ActualHeight, Height)) > 1)
                {
                    AnimateWindowHeight(Math.Max(ActualHeight, Height), targetHeight);
                }
            }
            else if (targetHeight > ActualHeight + 1)
            {
                AnimateWindowHeight(Math.Max(ActualHeight, Height), targetHeight);
            }
        }
        catch
        {
            // If layout transforms are temporarily unavailable, keep the existing expanded height.
        }
    }

    private void ExpandWindowForActionPanel(FrameworkElement panelContainer)
    {
        if (_isWindowCollapsed)
        {
            return;
        }

        UpdateMinimumWindowHeightForCurrentContent();
        SizeToContent = SizeToContent.Manual;

        if (WindowState == WindowState.Maximized)
        {
            return;
        }

        ExpandWindowForResults();
        _ = Dispatcher.BeginInvoke(
            DispatcherPriority.Loaded,
            new Action(() => AdjustWindowToResultsElement(panelContainer, 18, allowShrink: false)));
    }

    private static double GetActualOrFallbackHeight(FrameworkElement? element, double fallbackHeight)
    {
        return element is not null && element.ActualHeight > 0
            ? element.ActualHeight
            : fallbackHeight;
    }

    private double CalculateCompactWindowHeight()
    {
        var mainMargin = MainContentGrid.Margin;
        var inputMargin = InputContainer.Margin;
        var inputHeight = InputContainer.Height > 0 ? InputContainer.Height : _expandedInputHeight;

        var targetHeight = TitleBarHeight
                           + StatusBarHeight
                           + OuterBorderHeight
                           + mainMargin.Top
                           + inputHeight
                           + inputMargin.Bottom
                           + mainMargin.Bottom;

        return Math.Ceiling(targetHeight);
    }

    private void SetResultsAreaVisible(bool visible)
    {
        ResultsContainer.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        UpdateMinimumWindowHeightForCurrentContent();

        if (MainContentGrid.RowDefinitions.Count > 2)
        {
            MainContentGrid.RowDefinitions[2].Height = visible
                ? new GridLength(1, GridUnitType.Star)
                : new GridLength(0);
        }
    }

    private void AnimateInputContainerHeight(double from, double to, Action? onCompleted)
    {
        InputContainer.Height = from;
        var animation = new DoubleAnimation
        {
            From = from,
            To = to,
            Duration = TimeSpan.FromMilliseconds(200),
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };

        if (onCompleted is not null)
            animation.Completed += (_, _) => onCompleted();

        InputContainer.BeginAnimation(HeightProperty, animation);
    }

    private void AnimateWindowHeight(double from, double to)
    {
        BeginAnimation(HeightProperty, null);
        Height = from;

        var animation = new DoubleAnimation
        {
            From = from,
            To = to,
            Duration = TimeSpan.FromMilliseconds(220),
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };

        BeginAnimation(HeightProperty, animation);
    }

    private void UpdateExpandedInputHeight()
    {
        var referenceHeight = ResultsContainer.Visibility == Visibility.Visible
            ? ActualHeight
            : Math.Max(_preferredExpandedWindowHeight, ActualHeight);
        var desired = referenceHeight * 0.32 * InputHeightScaleFactor;
        _expandedInputHeight = Math.Clamp(desired, MinExpandedInputHeight, MaxExpandedInputHeight);
    }

    // --- Responsive layout ---

    private double MeasureMaxNameWidth()
    {
        var max = 0.0;
        var dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        var typeface = new Typeface(
            new FontFamily("Segoe UI"),
            FontStyles.Normal,
            FontWeights.SemiBold,
            FontStretches.Normal);

        foreach (var participant in Participants)
        {
            var text = participant.FullName ?? string.Empty;
            var formatted = new FormattedText(
                text,
                CultureInfo.CurrentUICulture,
                FlowDirection.LeftToRight,
                typeface,
                _displayDensityProfile.NameFontSize,
                Brushes.White,
                dpi);
            max = Math.Max(max, formatted.WidthIncludingTrailingWhitespace);
        }

        return max;
    }

    private double MeasureActionButtonsWidth()
    {
        var labels = new List<string>();
        if (ShowBtnOdoo) labels.Add("Odoo");
        if (ShowBtnOrdner) labels.Add("Ordner");
        if (ShowBtnAkte) labels.Add("Akte");
        if (ShowBtnBu) labels.Add("BU");
        if (ShowBtnEintrag) labels.Add("Eintrag BU");
        if (ShowBtnBi) labels.Add("BI");
        if (ShowBtnBe) labels.Add("BE");
        if (ShowBtnEintragBi) labels.Add("Eintrag BI");
        var total = 0.0;
        var dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        var typeface = new Typeface(
            new FontFamily("Segoe UI"),
            FontStyles.Normal,
            FontWeights.Normal,
            FontStretches.Normal);

        foreach (var label in labels)
        {
            var formatted = new FormattedText(
                label,
                CultureInfo.CurrentUICulture,
                FlowDirection.LeftToRight,
                typeface,
                _displayDensityProfile.ActionButtonFontSize,
                Brushes.White,
                dpi);

            var horizontalPadding = _displayDensityProfile.ActionButtonPadding.Left + _displayDensityProfile.ActionButtonPadding.Right;
            var buttonWidth = Math.Max(44, Math.Ceiling(formatted.WidthIncludingTrailingWhitespace + horizontalPadding));
            total += buttonWidth + _displayDensityProfile.ActionButtonMargin.Right;
        }

        return Math.Max(0, total - _displayDensityProfile.ActionButtonMargin.Right);
    }

    // --- Participant actions ---

    private void PresenceCheckBox_OnChecked(object sender, RoutedEventArgs e)
    {
        if (sender is not CheckBox { DataContext: Participant participant })
            return;

        UpdateActionState(participant);
        SetLastAction($"{participant.FullName}: {(participant.IsPresent ? "anwesend" : "abwesend")} gesetzt");
    }

    private static DependencyProperty RegisterLayoutProperty(string name, object defaultValue)
    {
        return DependencyProperty.Register(
            name,
            defaultValue.GetType(),
            typeof(MainWindow),
            new PropertyMetadata(defaultValue, OnLayoutRelevantPropertyChanged));
    }

    private static void OnLayoutRelevantPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is MainWindow window)
        {
            window.RequestDynamicMinWidthRefresh();
        }
    }

    private void ApplyDisplayDensity(string? mode)
    {
        var normalized = DisplayDensityMode.Normalize(mode);
        _displayDensityProfile = CreateDisplayDensityProfile(normalized);

        SetDensityResource("Density.WindowFontSize", _displayDensityProfile.WindowFontSize);
        SetDensityResource("Density.NameFontSize", _displayDensityProfile.NameFontSize);
        SetDensityResource("Density.SecondaryFontSize", _displayDensityProfile.SecondaryFontSize);
        SetDensityResource("Density.ActionButtonFontSize", _displayDensityProfile.ActionButtonFontSize);
        SetDensityResource("Density.PrimaryButtonFontSize", _displayDensityProfile.PrimaryButtonFontSize);
        SetDensityResource("Density.ActionButtonMinHeight", _displayDensityProfile.ActionButtonMinHeight);
        SetDensityResource("Density.PrimaryButtonMinHeight", _displayDensityProfile.PrimaryButtonMinHeight);
        SetDensityResource("Density.ActionRowItemHeight", _displayDensityProfile.ActionRowItemHeight);
        SetDensityResource("Density.SubtleIconButtonSize", _displayDensityProfile.SubtleIconButtonSize);
        SetDensityResource("Density.SubtleIconFontSize", _displayDensityProfile.SubtleIconFontSize);
        SetDensityResource("Density.CardPadding", _displayDensityProfile.CardPadding);
        SetDensityResource("Density.CardMargin", _displayDensityProfile.CardMargin);
        SetDensityResource("Density.ActionButtonPadding", _displayDensityProfile.ActionButtonPadding);
        SetDensityResource("Density.ActionButtonMargin", _displayDensityProfile.ActionButtonMargin);
        SetDensityResource("Density.PrimaryButtonPadding", _displayDensityProfile.PrimaryButtonPadding);
        SetDensityResource("Density.PresenceCheckBoxMargin", _displayDensityProfile.PresenceCheckBoxMargin);
        SetDensityResource("Density.CollapsedStripPadding", _displayDensityProfile.CollapsedStripPadding);
        SetDensityResource("Density.AbsenceStatusMargin", _displayDensityProfile.AbsenceStatusMargin);
        SetDensityResource("Density.NotFoundHintMargin", _displayDensityProfile.NotFoundHintMargin);
        SetDensityResource("Density.WideActionRowMargin", _displayDensityProfile.WideActionRowMargin);
        SetDensityResource("Density.BatchToggleMargin", _displayDensityProfile.BatchToggleMargin);
        SetDensityResource("Density.BatchPanelPadding", _displayDensityProfile.BatchPanelPadding);
        SetDensityResource("Density.BatchDescriptionMargin", _displayDensityProfile.BatchDescriptionMargin);
        SetDensityResource("Density.BatchButtonsMargin", _displayDensityProfile.BatchButtonsMargin);
        SetDensityResource("Density.BatchProgressMargin", _displayDensityProfile.BatchProgressMargin);
        SetDensityResource("Density.MiniScheduleScale", normalized == DisplayDensityMode.Compact ? 0.9 : 1.0);
        SetDensityResource("Density.ActionPanelResultsMinHeight", normalized == DisplayDensityMode.Compact ? 32.0 : 36.0);
        SetDensityResource("Density.ActionPanelResultsMaxHeight", normalized == DisplayDensityMode.Compact ? 84.0 : 96.0);

        FontSize = _displayDensityProfile.WindowFontSize;
        UpdateExpandedInputHeight();
        if (!_isInputCollapsed && InputContainer is not null)
        {
            InputContainer.BeginAnimation(HeightProperty, null);
            InputContainer.Height = _expandedInputHeight;
        }

        RequestDynamicMinWidthRefresh();
        SyncBodyHeightToWindow();
    }

    private void SetDensityResource(string key, object value)
    {
        Resources[key] = value;
    }

    private static DisplayDensityProfile CreateDisplayDensityProfile(string? mode)
    {
        return DisplayDensityMode.Normalize(mode) switch
        {
            DisplayDensityMode.Standard => new DisplayDensityProfile(
                DisplayDensityMode.Standard,
                12.5,
                12.5,
                10.5,
                10.5,
                12.5,
                26,
                32,
                28,
                26,
                12,
                new Thickness(9, 6, 9, 6),
                new Thickness(0, 0, 0, 3),
                new Thickness(9, 4, 9, 4),
                new Thickness(0, 0, 4, 0),
                new Thickness(16, 7, 16, 7),
                new Thickness(0, 0, 6, 0),
                new Thickness(9, 7, 9, 7),
                new Thickness(10, 0, 0, 0),
                new Thickness(0, 3, 0, 0),
                new Thickness(10, 0, 0, 0),
                new Thickness(0, 6, 0, 0),
                new Thickness(9),
                new Thickness(0, 0, 0, 6),
                new Thickness(0, 8, 0, 0),
                new Thickness(0, 6, 0, 0)),
            DisplayDensityMode.Compact => new DisplayDensityProfile(
                DisplayDensityMode.Compact,
                11.5,
                11.5,
                9.5,
                9.5,
                11.5,
                22,
                28,
                24,
                22,
                10,
                new Thickness(7, 4, 7, 4),
                new Thickness(0, 0, 0, 2),
                new Thickness(7, 2.5, 7, 2.5),
                new Thickness(0, 0, 3, 0),
                new Thickness(12, 5, 12, 5),
                new Thickness(0, 0, 5, 0),
                new Thickness(8, 5, 8, 5),
                new Thickness(8, 0, 0, 0),
                new Thickness(0, 1, 0, 0),
                new Thickness(8, 0, 0, 0),
                new Thickness(0, 5, 0, 0),
                new Thickness(7),
                new Thickness(0, 0, 0, 5),
                new Thickness(0, 6, 0, 0),
                new Thickness(0, 5, 0, 0)),
            _ => new DisplayDensityProfile(
                DisplayDensityMode.Standard,
                12.5,
                12.5,
                10.5,
                10.5,
                12.5,
                26,
                32,
                28,
                26,
                12,
                new Thickness(9, 6, 9, 6),
                new Thickness(0, 0, 0, 3),
                new Thickness(9, 4, 9, 4),
                new Thickness(0, 0, 4, 0),
                new Thickness(16, 7, 16, 7),
                new Thickness(0, 0, 6, 0),
                new Thickness(9, 7, 9, 7),
                new Thickness(10, 0, 0, 0),
                new Thickness(0, 3, 0, 0),
                new Thickness(10, 0, 0, 0),
                new Thickness(0, 6, 0, 0),
                new Thickness(9),
                new Thickness(0, 0, 0, 6),
                new Thickness(0, 8, 0, 0),
                new Thickness(0, 6, 0, 0))
        };
    }

    private void SettingsButton_OnClick(object sender, RoutedEventArgs e)
    {
        var model = new SettingsWindowModel
        {
            ServerPath = _config.ServerBasePath,
            UseSecondaryServerPath = _config.UseSecondaryServerBasePath,
            SecondaryServerPath = _config.SecondaryServerBasePath,
            UseTertiaryServerPath = _config.UseTertiaryServerBasePath,
            TertiaryServerPath = _config.TertiaryServerBasePath,
            ScheduleRootPath = _config.ScheduleRootPath,
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
            AutoPrefillOnEmptyClipboard = App.UserPrefs.AutoPrefillOnEmptyClipboard,
            DefaultEntryInitials = App.UserPrefs.DefaultEntryInitials,
            EnableDebugLogging = App.UserPrefs.EnableDebugLogging,
            EnableWordLifecycleLogging = App.UserPrefs.EnableWordLifecycleLogging,
            DisplayDensity = App.UserPrefs.DisplayDensity
        };

        var dialog = new SettingsWindow(model)
        {
            Owner = this,
            WindowStartupLocation = WindowStartupLocation.CenterOwner
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        ApplySettings(dialog.Result);
    }

    private void ApplySettings(SettingsWindowResult result)
    {
        _config.ServerBasePath = result.ServerPath;
        _config.UseSecondaryServerBasePath = result.UseSecondaryServerPath;
        _config.SecondaryServerBasePath = result.SecondaryServerPath;
        _config.UseTertiaryServerBasePath = result.UseTertiaryServerPath;
        _config.TertiaryServerBasePath = result.TertiaryServerPath;
        _config.ScheduleRootPath = result.ScheduleRootPath;
        SaveSettingsToDisk();
        _matcher = CreateMatcher();
        ResetMiniScheduleForAllParticipants(immediate: true);
        UpdateWeeklyScheduleRetryTimerState();
        if (IsLoaded)
        {
            _ = EnsureCurrentWeekScheduleLoadedAsync("SettingsChanged");
        }

        ShowParticipantInitials = result.ShowParticipantInitials;
        ShowBtnOdoo = result.ShowBtnOdoo;
        ShowBtnOrdner = result.ShowBtnOrdner;
        ShowBtnAkte = result.ShowBtnAkte;
        ShowBtnBu = result.ShowBtnBu;
        ShowBtnEintrag = result.ShowBtnEintrag;
        ShowBtnBi = result.ShowBtnBi;
        ShowBtnBe = result.ShowBtnBe;
        ShowBtnEintragBi = result.ShowBtnEintragBi;
        IsDarkTheme = result.IsDarkTheme;
        ApplyDisplayDensity(result.DisplayDensity);

        App.UserPrefs.ShowParticipantInitials = ShowParticipantInitials;
        App.UserPrefs.ShowBtnOdoo = ShowBtnOdoo;
        App.UserPrefs.ShowBtnOrdner = ShowBtnOrdner;
        App.UserPrefs.ShowBtnAkte = ShowBtnAkte;
        App.UserPrefs.ShowBtnBu = ShowBtnBu;
        App.UserPrefs.ShowBtnEintrag = ShowBtnEintrag;
        App.UserPrefs.ShowBtnBi = ShowBtnBi;
        App.UserPrefs.ShowBtnBe = ShowBtnBe;
        App.UserPrefs.ShowBtnEintragBi = ShowBtnEintragBi;
        App.UserPrefs.IsDarkTheme = IsDarkTheme;
        App.UserPrefs.DisplayDensity = DisplayDensityMode.Normalize(result.DisplayDensity);
        App.UserPrefs.AutoPrefillOnEmptyClipboard = result.AutoPrefillOnEmptyClipboard;
        App.UserPrefs.DefaultEntryInitials = result.DefaultEntryInitials;
        App.UserPrefs.EnableDebugLogging = result.EnableDebugLogging;
        App.UserPrefs.EnableWordLifecycleLogging = result.EnableWordLifecycleLogging;
        AppLogger.SetDebugEnabled(result.EnableDebugLogging);
        App.ApplyTheme(IsDarkTheme);
        App.SaveUserPrefs();

        if (Participants.Count > 0)
        {
            try
            {
                if (HasAtLeastOneReachableServerPath())
                {
                    _matcher.BuildIndex();
                    foreach (var participant in Participants)
                    {
                        _matcher.MatchParticipant(participant);
                        UpdateActionState(participant);
                    }
                }
                else
                {
                    foreach (var participant in Participants)
                    {
                        ResetParticipantMatch(participant);
                        UpdateActionState(participant);
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error("Fehler beim Re-Match nach Speichern der Einstellungen", ex);
                foreach (var participant in Participants)
                {
                    ResetParticipantMatch(participant);
                    UpdateActionState(participant);
                }
            }
        }

        BeginOdooMetadataWarmupForCurrentParticipants();
        SetLastAction("Einstellungen gespeichert");
        ShowToast("Einstellungen gespeichert", ToastType.Success);
    }

    private FolderMatcher CreateMatcher()
    {
        return new FolderMatcher(
            _config.ServerBasePath,
            _config.UseSecondaryServerBasePath,
            _config.SecondaryServerBasePath,
            _config.UseTertiaryServerBasePath,
            _config.TertiaryServerBasePath,
            _config.VerlaufsakteKeyword,
            _initialsResolver);
    }

    private bool HasAtLeastOneReachableServerPath()
    {
        var primaryOk = Directory.Exists(_config.ServerBasePath);
        var secondaryEnabled = _config.UseSecondaryServerBasePath && !string.IsNullOrWhiteSpace(_config.SecondaryServerBasePath);
        var secondaryOk = secondaryEnabled && Directory.Exists(_config.SecondaryServerBasePath);
        var tertiaryEnabled = _config.UseTertiaryServerBasePath && !string.IsNullOrWhiteSpace(_config.TertiaryServerBasePath);
        var tertiaryOk = tertiaryEnabled && Directory.Exists(_config.TertiaryServerBasePath);
        return primaryOk || secondaryOk || tertiaryOk;
    }

    private string BuildServerPathErrorMessage()
    {
        var secondaryEnabled = _config.UseSecondaryServerBasePath && !string.IsNullOrWhiteSpace(_config.SecondaryServerBasePath);
        var tertiaryEnabled = _config.UseTertiaryServerBasePath && !string.IsNullOrWhiteSpace(_config.TertiaryServerBasePath);
        if (!secondaryEnabled && !tertiaryEnabled)
        {
            return $"Serverpfad nicht erreichbar:\n{_config.ServerBasePath}";
        }

        if (!tertiaryEnabled)
        {
            return $"Kein erreichbarer Serverpfad.\nPrimär: {_config.ServerBasePath}\nSekundär: {_config.SecondaryServerBasePath}";
        }

        return $"Kein erreichbarer Serverpfad.\nPrimär: {_config.ServerBasePath}\nSekundär: {_config.SecondaryServerBasePath}\nDrittpfad: {_config.TertiaryServerBasePath}";
    }

    private void SaveSettingsToDisk()
    {
        var json = JsonSerializer.Serialize(_config, new JsonSerializerOptions
        {
            WriteIndented = true
        });
        File.WriteAllText(App.SettingsPath, json);
    }

    private async void ParticipantCard_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount != 2)
        {
            return;
        }

        if (sender is not Border { DataContext: Participant participant })
        {
            return;
        }

        if (FindVisualAncestor<Button>(e.OriginalSource as DependencyObject) is not null
            || FindVisualAncestor<CheckBox>(e.OriginalSource as DependencyObject) is not null)
        {
            return;
        }

        await ToggleParticipantMiniScheduleAsync(participant);
    }

    private void NameButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { CommandParameter: Participant participant })
            return;

        if (!participant.IsMultipleFound || participant.CandidateFolderPaths.Count == 0)
        {
            _ = ToggleParticipantMiniScheduleAsync(participant);
            return;
        }

        var menu = new ContextMenu
        {
            Background = BrushFromHex("#2A2B31"),
            Foreground = BrushFromHex("#E0E0E0"),
            BorderBrush = BrushFromHex("#3A3B42"),
            BorderThickness = new Thickness(1),
            PlacementTarget = sender as Button
        };

        foreach (var candidatePath in participant.CandidateFolderPaths)
        {
            var menuItem = new MenuItem
            {
                Header = Path.GetFileName(candidatePath),
                ToolTip = candidatePath
            };
            menuItem.Click += (_, _) =>
            {
                participant.SelectedFolderPath = candidatePath;
                UpdateActionState(participant);
                BeginOdooMetadataWarmupForCurrentParticipants();
                SetLastAction($"Ordnerkandidat gewählt für {participant.FullName}");
            };
            menu.Items.Add(menuItem);
        }

        menu.IsOpen = true;
    }

    private async Task ToggleParticipantMiniScheduleAsync(Participant participant)
    {
        if (_expandedMiniScheduleParticipant is not null && ReferenceEquals(_expandedMiniScheduleParticipant, participant))
        {
            participant.IsMiniScheduleExpanded = false;
            if (TryGetMiniScheduleTrayHost(participant, out var existingTray))
            {
                AnimateMiniScheduleTray(existingTray, expand: false);
            }

            _expandedMiniScheduleParticipant = null;
            return;
        }

        if (_expandedMiniScheduleParticipant is not null)
        {
            var previous = _expandedMiniScheduleParticipant;
            previous.IsMiniScheduleExpanded = false;
            if (TryGetMiniScheduleTrayHost(previous, out var previousTray))
            {
                AnimateMiniScheduleTray(previousTray, expand: false);
            }
        }

        _expandedMiniScheduleParticipant = participant;
        participant.IsMiniScheduleExpanded = true;

        if (!TryGetMiniScheduleTrayHost(participant, out var trayHost))
        {
            return;
        }

        AnimateMiniScheduleTray(trayHost, expand: true);

        if (participant.MiniScheduleState == ParticipantMiniScheduleState.Ready
            || participant.IsMiniScheduleLoading)
        {
            return;
        }

        participant.IsMiniScheduleLoading = true;
        participant.MiniScheduleState = ParticipantMiniScheduleState.Unavailable;
        participant.MiniScheduleSummary = ParticipantMiniScheduleSummary.Create(ParticipantMiniScheduleState.Unavailable);

        try
        {
            var participantsSnapshot = Participants.ToList();
            var summary = await Task.Run(() =>
                _weeklyScheduleService.BuildSummary(
                    _config.ScheduleRootPath,
                    participant,
                    participantsSnapshot,
                    _config.UseSecondaryServerBasePath ? _config.SecondaryServerBasePath : null));

            participant.MiniScheduleSummary = summary;
            participant.MiniScheduleState = summary.State;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"MiniSchedule: Fehler fuer '{participant.FullName}': {ex.Message}");
            participant.MiniScheduleSummary = ParticipantMiniScheduleSummary.Create(ParticipantMiniScheduleState.Unavailable);
            participant.MiniScheduleState = ParticipantMiniScheduleState.Unavailable;
        }
        finally
        {
            participant.IsMiniScheduleLoading = false;
            UpdateWeeklyScheduleRetryTimerState();
        }
    }

    private void MiniScheduleTrayHost_OnLoaded(object sender, RoutedEventArgs e)
    {
        if (sender is not Border trayHost || trayHost.DataContext is not Participant participant)
        {
            return;
        }

        _miniScheduleTrayHosts[participant] = trayHost;
        var scaleTransform = EnsureMutableMiniScheduleScaleTransform(trayHost);
        if (participant.IsMiniScheduleExpanded)
        {
            trayHost.Visibility = Visibility.Visible;
            trayHost.MaxHeight = 240;
            trayHost.Opacity = 1;
            scaleTransform.ScaleY = 1;
        }
        else
        {
            scaleTransform.ScaleY = 0.96;
        }
    }

    private void MiniScheduleTrayHost_OnUnloaded(object sender, RoutedEventArgs e)
    {
        if (sender is not Border trayHost || trayHost.DataContext is not Participant participant)
        {
            return;
        }

        if (_miniScheduleTrayHosts.TryGetValue(participant, out var existing) && ReferenceEquals(existing, trayHost))
        {
            _miniScheduleTrayHosts.Remove(participant);
        }
    }

    private bool TryGetMiniScheduleTrayHost(Participant participant, out Border trayHost)
    {
        return _miniScheduleTrayHosts.TryGetValue(participant, out trayHost!);
    }

    private void AnimateMiniScheduleTray(Border trayHost, bool expand)
    {
        trayHost.BeginAnimation(Border.MaxHeightProperty, null);
        trayHost.BeginAnimation(UIElement.OpacityProperty, null);
        var scaleTransform = EnsureMutableMiniScheduleScaleTransform(trayHost);

        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, null);

        var duration = TimeSpan.FromMilliseconds(180);
        var easing = new QuadraticEase { EasingMode = EasingMode.EaseOut };

        if (expand)
        {
            trayHost.Visibility = Visibility.Visible;

            trayHost.BeginAnimation(Border.MaxHeightProperty, new DoubleAnimation(0, 240, duration)
            {
                EasingFunction = easing
            });
            trayHost.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(150)));
            scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, new DoubleAnimation(0.96, 1, duration)
            {
                EasingFunction = easing
            });
            return;
        }

        var heightAnimation = new DoubleAnimation(trayHost.ActualHeight <= 0 ? 240 : trayHost.ActualHeight, 0, duration)
        {
            EasingFunction = easing
        };
        heightAnimation.Completed += (_, _) =>
        {
            trayHost.Visibility = Visibility.Collapsed;
            trayHost.MaxHeight = 0;
            trayHost.Opacity = 0;
            if (trayHost.RenderTransform is ScaleTransform transform)
            {
                transform.ScaleY = 0.96;
            }
        };

        trayHost.BeginAnimation(Border.MaxHeightProperty, heightAnimation);
        trayHost.BeginAnimation(UIElement.OpacityProperty, new DoubleAnimation(trayHost.Opacity, 0, TimeSpan.FromMilliseconds(120)));
        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, new DoubleAnimation(1, 0.96, duration)
        {
            EasingFunction = easing
        });
    }

    private void ResetMiniScheduleForAllParticipants(bool immediate)
    {
        foreach (var participant in Participants)
        {
            ResetMiniScheduleState(participant, immediate);
        }

        _expandedMiniScheduleParticipant = null;
    }

    private void ResetMiniScheduleState(Participant participant, bool immediate)
    {
        participant.IsMiniScheduleExpanded = false;
        participant.IsMiniScheduleLoading = false;
        participant.MiniScheduleState = ParticipantMiniScheduleState.Hidden;
        participant.MiniScheduleSummary = ParticipantMiniScheduleSummary.Create(ParticipantMiniScheduleState.Hidden);

        if (!immediate || !TryGetMiniScheduleTrayHost(participant, out var trayHost))
        {
            return;
        }

        trayHost.BeginAnimation(Border.MaxHeightProperty, null);
        trayHost.BeginAnimation(UIElement.OpacityProperty, null);
        var scaleTransform = EnsureMutableMiniScheduleScaleTransform(trayHost);
        scaleTransform.BeginAnimation(ScaleTransform.ScaleYProperty, null);
        scaleTransform.ScaleY = 0.96;

        trayHost.MaxHeight = 0;
        trayHost.Opacity = 0;
        trayHost.Visibility = Visibility.Collapsed;
    }

    private static ScaleTransform EnsureMutableMiniScheduleScaleTransform(Border trayHost)
    {
        if (trayHost.RenderTransform is ScaleTransform existing && !existing.IsFrozen)
        {
            return existing;
        }

        var currentScaleY = trayHost.RenderTransform is ScaleTransform frozenExisting
            ? frozenExisting.ScaleY
            : 0.96;

        var replacement = new ScaleTransform(1, currentScaleY);
        trayHost.RenderTransform = replacement;
        return replacement;
    }

    private static T? FindVisualAncestor<T>(DependencyObject? source) where T : DependencyObject
    {
        var current = source;
        while (current is not null)
        {
            if (current is T match)
            {
                return match;
            }

            current = VisualTreeHelper.GetParent(current);
        }

        return null;
    }

    private void OpenOdooButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { CommandParameter: Participant participant })
            return;

        if (!EnsureParticipantFolder(participant))
            return;

        try
        {
            if (!EnsureOdooMetadataLoaded(participant))
            {
                SetLastAction($"Kein Odoo-Link gefunden: {participant.FullName}");
                ShowImportantAlert(
                    "Kein Odoo-Link gefunden",
                    $"Für {participant.FullName} wurde kein Odoo-Link gefunden.",
                    "In der ausgewählten Akte konnte im Kopfbereich kein gültiger Odoo-Link erkannt werden.",
                    AppAlertKind.Warning,
                    "Die Akte selbst bleibt unverändert. Du kannst den TN weiterhin normal über Ordner oder Akte öffnen.");
                return;
            }

            if (!Uri.TryCreate(participant.OdooUrl, UriKind.Absolute, out var uri) ||
                (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                SetLastAction("Odoo-Link ist ungültig.");
                ShowImportantAlert(
                    "Odoo-Link ungültig",
                    "Der gefundene Odoo-Link konnte nicht geöffnet werden.",
                    "In der Akte wurde zwar ein Link gefunden, er ist aber nicht als gültige http- oder https-Adresse verwendbar.",
                    AppAlertKind.Warning);
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = participant.OdooUrl,
                UseShellExecute = true
            });
            SetLastAction($"Odoo geöffnet: {participant.FullName}");
            ShowToast($"Odoo geöffnet: {participant.FullName}", ToastType.Success);
            AppLogger.Info($"Odoo-Link geoeffnet fuer {participant.FullName}: {participant.OdooUrl}");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Fehler beim Oeffnen des Odoo-Links fuer {participant.FullName}", ex);
            SetLastAction($"Fehler beim Öffnen des Odoo-Links: {ex.Message}");
            ShowImportantAlert(
                "Odoo konnte nicht geöffnet werden",
                $"Der Odoo-Link für {participant.FullName} konnte nicht geöffnet werden.",
                ex.Message,
                AppAlertKind.Error);
        }
    }

    private void OpenFolderButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { CommandParameter: Participant participant })
            return;

        if (string.IsNullOrWhiteSpace(participant.MatchedFolderPath) || !Directory.Exists(participant.MatchedFolderPath))
        {
            AppLogger.Warn($"Ordner konnte nicht geoeffnet werden (kein Pfad): {participant.FullName}");
            ShowToast($"Kein passender Ordner für: {participant.FullName}", ToastType.Warning);
            SetLastAction($"Kein passender Ordner für: {participant.FullName}");
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = participant.MatchedFolderPath,
                UseShellExecute = true
            });
            SetLastAction($"Ordner geöffnet: {participant.FullName}");
            ShowToast($"Ordner geöffnet: {participant.FullName}", ToastType.Success);
            AppLogger.Info($"Ordner geoeffnet: {participant.MatchedFolderPath}");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Fehler beim Oeffnen des Ordners fuer {participant.FullName}", ex);
            ShowToast($"Ordner konnte nicht geöffnet werden: {ex.Message}", ToastType.Error);
            SetLastAction($"Fehler beim Öffnen des Ordners: {ex.Message}");
        }
    }

    private async void OpenAkteButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { CommandParameter: Participant participant })
            return;

        if (!EnsureWordAvailable() || !EnsureParticipantFolder(participant))
            return;

        if (!TryBeginWordAction("Word-Aktion läuft bereits. Bitte kurz warten."))
            return;

        try
        {
            AppLogger.Info($"Akte oeffnen angefordert fuer {participant.FullName}");
            var docPath = ResolveDocumentPathForParticipant(participant);
            if (string.IsNullOrWhiteSpace(docPath))
                return;

            await _wordStaHost.RunAsync("OpenDocument", service => service.OpenDocument(docPath));

            SetLastAction($"Akte geöffnet: {participant.FullName}");
            ShowToast($"Akte geöffnet: {participant.FullName}", ToastType.Success);
            AppLogger.Info($"Akte geoeffnet: {docPath}");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Keine Verlaufsakte", StringComparison.OrdinalIgnoreCase))
        {
            ShowToast($"Keine Verlaufsakte in: {participant.MatchedFolderPath}", ToastType.Warning);
            SetLastAction($"Keine Verlaufsakte in: {participant.MatchedFolderPath}");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("gesperrt", StringComparison.OrdinalIgnoreCase))
        {
            SetLastAction(ex.Message);
            ShowImportantAlert(
                "Akte gesperrt",
                $"Die Akte von {participant.FullName} ist aktuell nicht schreibbar.",
                ex.Message,
                AppAlertKind.Warning);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Fehler im OpenAkte-Flow fuer {participant.FullName}", ex);
            ShowToast(ex.Message, ToastType.Error);
            SetLastAction($"Fehler beim Öffnen der Akte: {ex.Message}");
        }
        finally
        {
            EndWordAction();
        }
    }

    private async void OpenAkteBuButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { CommandParameter: Participant participant })
            return;

        if (!EnsureWordAvailable() || !EnsureParticipantFolder(participant))
            return;

        if (!TryBeginWordAction("Word-Aktion läuft bereits. Bitte kurz warten."))
            return;

        try
        {
            AppLogger.Info($"Akte oeffnen BU angefordert fuer {participant.FullName} (Bookmark={_config.WordBuBookmarkName})");
            var docPath = ResolveDocumentPathForParticipant(participant);
            if (string.IsNullOrWhiteSpace(docPath))
                return;

            await _wordStaHost.RunAsync("OpenDocumentAtBookmark-BU", service => service.OpenDocumentAtBookmark(docPath, _config.WordBuBookmarkName));
            SetLastAction($"Akte geöffnet (BU): {participant.FullName}");
            ShowToast($"Akte BU geöffnet: {participant.FullName}", ToastType.Success);
            AppLogger.Info($"Akte BU geoeffnet: {docPath}");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Keine Verlaufsakte", StringComparison.OrdinalIgnoreCase))
        {
            ShowToast($"Keine Verlaufsakte in: {participant.MatchedFolderPath}", ToastType.Warning);
            SetLastAction($"Keine Verlaufsakte in: {participant.MatchedFolderPath}");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Bookmark", StringComparison.OrdinalIgnoreCase))
        {
            SetLastAction($"Bookmark '{_config.WordBuBookmarkName}' nicht gefunden");
            ShowImportantAlert(
                "BU-Textmarke fehlt",
                $"Die BU-Stelle konnte für {participant.FullName} nicht geöffnet werden.",
                $"Die erwartete Textmarke '{_config.WordBuBookmarkName}' wurde in der Akte nicht gefunden.",
                AppAlertKind.Warning);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("gesperrt", StringComparison.OrdinalIgnoreCase))
        {
            SetLastAction(ex.Message);
            ShowImportantAlert(
                "Akte gesperrt",
                $"Die Akte von {participant.FullName} ist aktuell nicht schreibbar.",
                ex.Message,
                AppAlertKind.Warning);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Fehler im Akte-oeffnen-BU-Flow fuer {participant.FullName}", ex);
            ShowToast(ex.Message, ToastType.Error);
            SetLastAction($"Fehler beim Öffnen der BU-Stelle: {ex.Message}");
        }
        finally
        {
            EndWordAction();
        }
    }

    private async void OpenAkteBiButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { CommandParameter: Participant participant })
            return;

        if (!EnsureWordAvailable() || !EnsureParticipantFolder(participant))
            return;

        if (!TryBeginWordAction("Word-Aktion läuft bereits. Bitte kurz warten."))
            return;

        try
        {
            var docPath = ResolveDocumentPathForParticipant(participant);
            if (string.IsNullOrWhiteSpace(docPath))
                return;

            await _wordStaHost.RunAsync("OpenDocumentAtBookmark-BI", service => service.OpenDocumentAtBookmark(docPath, _config.WordBiBookmarkName));
            SetLastAction($"Akte geöffnet (BI): {participant.FullName}");
            ShowToast($"Akte BI geöffnet: {participant.FullName}", ToastType.Success);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Bookmark", StringComparison.OrdinalIgnoreCase))
        {
            SetLastAction($"Bookmark '{_config.WordBiBookmarkName}' nicht gefunden");
            ShowImportantAlert(
                "BI-Textmarke fehlt",
                $"Die BI-Stelle konnte für {participant.FullName} nicht geöffnet werden.",
                $"Die erwartete Textmarke '{_config.WordBiBookmarkName}' wurde in der Akte nicht gefunden.",
                AppAlertKind.Warning);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Keine Verlaufsakte", StringComparison.OrdinalIgnoreCase))
        {
            SetLastAction(ex.Message);
            ShowImportantAlert(
                "Keine Akte gefunden",
                $"Für {participant.FullName} konnte keine BI-Akte gefunden werden.",
                ex.Message,
                AppAlertKind.Warning);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("gesperrt", StringComparison.OrdinalIgnoreCase))
        {
            SetLastAction(ex.Message);
            ShowImportantAlert(
                "Akte gesperrt",
                $"Die Akte von {participant.FullName} ist aktuell nicht schreibbar.",
                ex.Message,
                AppAlertKind.Warning);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Fehler im Akte-oeffnen-BI-Flow fuer {participant.FullName}", ex);
            ShowToast(ex.Message, ToastType.Error);
            SetLastAction($"Fehler beim Öffnen der BI-Stelle: {ex.Message}");
        }
        finally
        {
            EndWordAction();
        }
    }

    private async void OpenAkteBeButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { CommandParameter: Participant participant })
            return;

        if (!EnsureWordAvailable() || !EnsureParticipantFolder(participant))
            return;

        if (!TryBeginWordAction("Word-Aktion läuft bereits. Bitte kurz warten."))
            return;

        try
        {
            var docPath = ResolveDocumentPathForParticipant(participant);
            if (string.IsNullOrWhiteSpace(docPath))
                return;

            await _wordStaHost.RunAsync("OpenDocumentAtBookmark-BE", service => service.OpenDocumentAtBookmark(docPath, _config.WordBeBookmarkName));
            SetLastAction($"Akte geöffnet (BE): {participant.FullName}");
            ShowToast($"Akte BE geöffnet: {participant.FullName}", ToastType.Success);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Keine Verlaufsakte", StringComparison.OrdinalIgnoreCase))
        {
            ShowToast($"Keine Verlaufsakte in: {participant.MatchedFolderPath}", ToastType.Warning);
            SetLastAction($"Keine Verlaufsakte in: {participant.MatchedFolderPath}");
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Bookmark", StringComparison.OrdinalIgnoreCase))
        {
            SetLastAction($"Bookmark '{_config.WordBeBookmarkName}' nicht gefunden");
            ShowImportantAlert(
                "BE-Textmarke fehlt",
                $"Die BE-Stelle konnte für {participant.FullName} nicht geöffnet werden.",
                $"Die erwartete Textmarke '{_config.WordBeBookmarkName}' wurde in der Akte nicht gefunden.",
                AppAlertKind.Warning);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("gesperrt", StringComparison.OrdinalIgnoreCase))
        {
            SetLastAction(ex.Message);
            ShowImportantAlert(
                "Akte gesperrt",
                $"Die Akte von {participant.FullName} ist aktuell nicht schreibbar.",
                ex.Message,
                AppAlertKind.Warning);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Fehler im Akte-oeffnen-BE-Flow fuer {participant.FullName}", ex);
            ShowToast(ex.Message, ToastType.Error);
            SetLastAction($"Fehler beim Öffnen der BE-Stelle: {ex.Message}");
        }
        finally
        {
            EndWordAction();
        }
    }

    private async void InsertEntryButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { CommandParameter: Participant participant })
            return;

        if (!EnsureWordAvailable() || !EnsureParticipantFolder(participant))
            return;

        if (!TryBeginWordAction("Word-Aktion läuft bereits. Bitte kurz warten."))
            return;

        try
        {
            AppLogger.Info($"Eintrag einfuegen angefordert fuer {participant.FullName}");
            var docPath = ResolveDocumentPathForParticipant(participant);
            if (string.IsNullOrWhiteSpace(docPath))
                return;

            var fallbackFields = BuildFallbackEntryFieldsIfEnabled();
            var clipboardText = _wordService.ReadClipboardTextWithRetry();
            var insertedFromClipboard = await _wordStaHost.RunAsync(
                "InsertClipboardToTable-BU",
                service => service.InsertClipboardToTable(
                    docPath,
                    _config.WordBookmarkName,
                    2,
                    fallbackFields,
                    clipboardText));
            if (insertedFromClipboard)
            {
                SetLastAction($"Eintrag eingefügt: {participant.FullName}");
                ShowToast($"Eintrag eingefügt: {participant.FullName}", ToastType.Success);
                AppLogger.Info($"Eintrag aus Clipboard eingefuegt fuer {participant.FullName} in {docPath}");
            }
            else
            {
                var usedPrefill = fallbackFields is not null;
                var actionText = usedPrefill
                    ? $"Zeile mit Datum/Kürzel vorbereitet: {participant.FullName}"
                    : $"Leere Zeile vorbereitet: {participant.FullName}";
                SetLastAction(actionText);
                ShowToast(actionText, ToastType.Info);
                AppLogger.Info($"Zeile vorbereitet (Clipboard leer/ungueltig, Prefill={usedPrefill}) fuer {participant.FullName} in {docPath}");
            }
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("Keine Verlaufsakte", StringComparison.OrdinalIgnoreCase))
        {
            ShowToast($"Keine Verlaufsakte in: {participant.MatchedFolderPath}", ToastType.Warning);
            SetLastAction($"Keine Verlaufsakte in: {participant.MatchedFolderPath}");
        }
        catch (WordTemplateValidationException ex) when (ex.Kind == WordTemplateValidationErrorKind.BookmarkMissing)
        {
            SetLastAction($"Bookmark '{_config.WordBookmarkName}' nicht gefunden");
            ShowImportantAlert(
                "BU-Tabelle fehlt",
                $"Der Eintrag für {participant.FullName} konnte nicht eingefügt werden.",
                $"Die erwartete Textmarke '{_config.WordBookmarkName}' wurde in der Akte nicht gefunden.",
                AppAlertKind.Warning);
        }
        catch (WordTemplateValidationException ex) when (ex.Kind == WordTemplateValidationErrorKind.StructuredEntryTableInvalid)
        {
            SetLastAction(ex.UserMessage);
            ShowImportantAlert(
                "BU-Tabelle ungültig",
                $"Der Eintrag für {participant.FullName} konnte nicht eingefügt werden.",
                ex.UserMessage,
                AppAlertKind.Warning);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("gesperrt", StringComparison.OrdinalIgnoreCase))
        {
            SetLastAction(ex.Message);
            ShowImportantAlert(
                "Akte gesperrt",
                $"Die Akte von {participant.FullName} ist aktuell nicht schreibbar.",
                ex.Message,
                AppAlertKind.Warning);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Fehler im Einfuegen-Flow fuer {participant.FullName}", ex);
            ShowToast(ex.Message, ToastType.Error);
            SetLastAction($"Fehler beim Einfügen: {ex.Message}");
        }
        finally
        {
            EndWordAction();
        }
    }

    private async void InsertEntryBiButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { CommandParameter: Participant participant })
            return;

        if (!EnsureWordAvailable() || !EnsureParticipantFolder(participant))
            return;

        if (!TryBeginWordAction("Word-Aktion läuft bereits. Bitte kurz warten."))
            return;

        try
        {
            var docPath = ResolveDocumentPathForParticipant(participant);
            if (string.IsNullOrWhiteSpace(docPath))
                return;

            var fallbackFields = BuildFallbackEntryFieldsIfEnabled();
            var clipboardText = _wordService.ReadClipboardTextWithRetry();
            var inserted = await _wordStaHost.RunAsync(
                "InsertClipboardToTable-BI",
                service => service.InsertClipboardToTable(
                    docPath,
                    _config.WordBiTableBookmarkName,
                    2,
                    fallbackFields,
                    clipboardText));
            if (inserted)
            {
                SetLastAction($"Eintrag BI eingefügt: {participant.FullName}");
                ShowToast($"Eintrag BI eingefügt: {participant.FullName}", ToastType.Success);
            }
            else
            {
                var usedPrefill = fallbackFields is not null;
                var actionText = usedPrefill
                    ? $"BI-Zeile mit Datum/Kürzel vorbereitet: {participant.FullName}"
                    : $"Leere BI-Zeile vorbereitet: {participant.FullName}";
                SetLastAction(actionText);
                ShowToast(actionText, ToastType.Info);
            }
        }
        catch (WordTemplateValidationException ex) when (ex.Kind == WordTemplateValidationErrorKind.BookmarkMissing)
        {
            SetLastAction($"Bookmark '{_config.WordBiTableBookmarkName}' nicht gefunden");
            ShowImportantAlert(
                "BI-Tabelle fehlt",
                $"Der BI-Eintrag für {participant.FullName} konnte nicht eingefügt werden.",
                $"Die erwartete Textmarke '{_config.WordBiTableBookmarkName}' wurde in der Akte nicht gefunden.",
                AppAlertKind.Warning);
        }
        catch (WordTemplateValidationException ex) when (ex.Kind == WordTemplateValidationErrorKind.StructuredEntryTableInvalid)
        {
            SetLastAction(ex.UserMessage);
            ShowImportantAlert(
                "BI-Tabelle ungültig",
                $"Der BI-Eintrag für {participant.FullName} konnte nicht eingefügt werden.",
                ex.UserMessage,
                AppAlertKind.Warning);
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("gesperrt", StringComparison.OrdinalIgnoreCase))
        {
            SetLastAction(ex.Message);
            ShowImportantAlert(
                "Akte gesperrt",
                $"Die Akte von {participant.FullName} ist aktuell nicht schreibbar.",
                ex.Message,
                AppAlertKind.Warning);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Fehler im Einfuegen-BI-Flow fuer {participant.FullName}", ex);
            ShowToast(ex.Message, ToastType.Error);
            SetLastAction($"Fehler beim Einfügen BI: {ex.Message}");
        }
        finally
        {
            EndWordAction();
        }
    }

    private void ToggleBatchPanel_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _isBatchCollapsed = !_isBatchCollapsed;
        BatchPanelContainer.Visibility = _isBatchCollapsed ? Visibility.Collapsed : Visibility.Visible;
        BatchStripText.Text = _isBatchCollapsed ? "▶ BU: Batch-Eintrag ⚡︎" : "▼ BU: Batch-Eintrag ⚡︎";

        if (!_isBatchCollapsed)
        {
            ExpandWindowForActionPanel(BatchPanelContainer);
            return;
        }

        ScheduleResultsHeightRefresh(allowShrink: true);
    }

    private void ClearBatchInputButton_OnClick(object sender, RoutedEventArgs e)
    {
        BatchInputTextBox.Clear();
        BatchResults.Clear();
        BatchProgressText.Text = "Fortschritt: —";
        SetLastAction("Batch-Eingabe geleert");
        ShowToast("Batch-Eingabe geleert", ToastType.Info);
    }

    private void ToggleBiTodoPanel_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        _isBiTodoCollapsed = !_isBiTodoCollapsed;
        BiTodoPanelContainer.Visibility = _isBiTodoCollapsed ? Visibility.Collapsed : Visibility.Visible;
        BiTodoStripText.Text = _isBiTodoCollapsed ? "▶ BI: To-dos ⚡︎" : "▼ BI: To-dos ⚡︎";

        if (!_isBiTodoCollapsed)
        {
            ExpandWindowForActionPanel(BiTodoPanelContainer);
            return;
        }

        ScheduleResultsHeightRefresh(allowShrink: true);
    }

    private void ClearBiTodoResultsButton_OnClick(object sender, RoutedEventArgs e)
    {
        BiTodoResults.Clear();
        BiTodoProgressText.Text = "Fortschritt: —";
        SetLastAction("BI: To-do-Status geleert");
        ShowToast("BI: To-do-Status geleert", ToastType.Info);
    }

    private void BatchResults_OnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (!_isBatchCollapsed)
        {
            ScheduleResultsHeightRefresh(allowShrink: false);
        }
    }

    private void BiTodoResults_OnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (!_isBiTodoCollapsed)
        {
            ScheduleResultsHeightRefresh(allowShrink: false);
        }
    }

    private void Participants_OnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        RequestDynamicMinWidthRefresh();
    }

    private void ScheduleResultsHeightRefresh(bool allowShrink)
    {
        if (_isWindowCollapsed
            || WindowState == WindowState.Maximized
            || ResultsContainer.Visibility != Visibility.Visible)
        {
            return;
        }

        _ = Dispatcher.BeginInvoke(
            DispatcherPriority.Loaded,
            new Action(() =>
            {
                var target = GetLowestVisibleResultsElement();
                if (target is null)
                {
                    return;
                }

                AdjustWindowToResultsElement(target, 18, allowShrink);
            }));
    }

    private FrameworkElement? GetLowestVisibleResultsElement()
    {
        if (!_isBiTodoCollapsed && BiTodoPanelContainer.Visibility == Visibility.Visible)
        {
            return BiTodoPanelContainer;
        }

        if (BiTodoStripBorder.Visibility == Visibility.Visible)
        {
            return BiTodoStripBorder;
        }

        if (!_isBatchCollapsed && BatchPanelContainer.Visibility == Visibility.Visible)
        {
            return BatchPanelContainer;
        }

        if (BatchStripBorder.Visibility == Visibility.Visible)
        {
            return BatchStripBorder;
        }

        return ResultsActionsFooter;
    }

    private void RestoreWindowBoundsFromPrefs()
    {
        var prefs = App.UserPrefs;
        var requestedWidth = prefs.WindowWidth is > 0 && prefs.WindowWidth.Value >= MinWidth
            ? prefs.WindowWidth.Value
            : Width;
        var preferredStoredHeight = prefs.ExpandedWindowHeight is > 0
            ? prefs.ExpandedWindowHeight.Value
            : prefs.WindowHeight ?? Height;
        var requestedHeight = preferredStoredHeight > GetCollapsedWindowHeight()
            ? preferredStoredHeight
            : Height;
        var adjustedRect = CreateCenteredWindowRectOnPrimaryMonitor(requestedWidth, requestedHeight);
        ApplyRestoredWindowBounds(adjustedRect);
        _preferredExpandedWindowHeight = Math.Max(requestedHeight, CalculateCompactWindowHeight());
        _isWindowCollapsed = prefs.IsCollapsed;

        if (prefs.WindowWasMaximized && !_isWindowCollapsed)
        {
            WindowState = WindowState.Maximized;
        }
    }

    private void SaveWindowBoundsToPrefs()
    {
        try
        {
            var bounds = WindowState == WindowState.Normal
                ? new Rect(Left, Top, Width, Height)
                : RestoreBounds;
            var storedHeight = Math.Max(_preferredExpandedWindowHeight, _expandedWindowMinHeight);

            if (!_isWindowCollapsed
                && WindowState == WindowState.Normal
                && Height > GetCollapsedWindowHeight())
            {
                storedHeight = Height;
            }
            else if (bounds.Height >= _expandedWindowMinHeight)
            {
                storedHeight = Math.Max(storedHeight, bounds.Height);
            }

            if (bounds.Width >= MinWidth && storedHeight > GetCollapsedWindowHeight())
            {
                App.UserPrefs.WindowWidth = bounds.Width;
                App.UserPrefs.WindowHeight = storedHeight;
                App.UserPrefs.ExpandedWindowHeight = storedHeight;
            }

            App.UserPrefs.IsCollapsed = _isWindowCollapsed;
            App.UserPrefs.WindowWasMaximized = WindowState == WindowState.Maximized && !_isWindowCollapsed;
            App.SaveUserPrefs();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Fensterzustand konnte nicht gespeichert werden: {ex.Message}");
        }
    }

    private void ApplyRestoredWindowBounds(Rect bounds)
    {
        if (bounds.IsEmpty)
        {
            return;
        }

        Width = Math.Max(MinWidth, bounds.Width);
        Height = Math.Max(MinHeight, bounds.Height);
        Left = bounds.Left;
        Top = bounds.Top;
    }

    private static Rect ClampWindowRectToWorkArea(Rect rect, Rect workArea)
    {
        var width = Math.Min(Math.Max(1, rect.Width), workArea.Width);
        var height = Math.Min(Math.Max(1, rect.Height), workArea.Height);
        var left = Math.Clamp(rect.Left, workArea.Left, Math.Max(workArea.Left, workArea.Right - width));
        var top = Math.Clamp(rect.Top, workArea.Top, Math.Max(workArea.Top, workArea.Bottom - height));
        return new Rect(left, top, width, height);
    }

    private static Rect CreateCenteredWindowRectOnPrimaryMonitor(double requestedWidth, double requestedHeight)
    {
        var workArea = SystemParameters.WorkArea;
        var width = Math.Min(Math.Max(1, requestedWidth), workArea.Width);
        var height = Math.Min(Math.Max(1, requestedHeight), workArea.Height);
        var left = workArea.Left + Math.Max(0, (workArea.Width - width) / 2.0);
        var freeVerticalSpace = Math.Max(0, workArea.Height - height);
        var top = workArea.Top + (freeVerticalSpace * 0.38);
        return new Rect(left, top, width, height);
    }

    private static MonitorDescriptor GetMonitorContainingPoint(double x, double y)
    {
        var monitors = MonitorNative.EnumerateMonitors();
        if (monitors.Count == 0)
        {
            return new MonitorDescriptor(
                string.Empty,
                new Rect(
                    SystemParameters.VirtualScreenLeft,
                    SystemParameters.VirtualScreenTop,
                    SystemParameters.VirtualScreenWidth,
                    SystemParameters.VirtualScreenHeight),
                false);
        }

        var probeX = double.IsFinite(x) ? x : 0;
        var probeY = double.IsFinite(y) ? y : 0;
        foreach (var monitor in monitors)
        {
            if (monitor.WorkArea.Contains(probeX, probeY))
            {
                return monitor;
            }
        }

        return GetPrimaryRestoreMonitor(monitors);
    }

    private static MonitorDescriptor GetPrimaryRestoreMonitor()
    {
        var monitors = MonitorNative.EnumerateMonitors();
        if (monitors.Count == 0)
        {
            return new MonitorDescriptor(
                string.Empty,
                new Rect(
                    SystemParameters.WorkArea.Left,
                    SystemParameters.WorkArea.Top,
                    SystemParameters.WorkArea.Width,
                    SystemParameters.WorkArea.Height),
                true);
        }

        return GetPrimaryRestoreMonitor(monitors);
    }

    private static MonitorDescriptor GetPrimaryRestoreMonitor(IReadOnlyList<MonitorDescriptor> monitors)
    {
        foreach (var monitor in monitors)
        {
            if (monitor.IsPrimary)
            {
                return monitor;
            }
        }

        return monitors[0];
    }

    private async void ExecuteBatchButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBatchRunning)
        {
            ShowToast("Batch läuft bereits.", ToastType.Warning);
            return;
        }

        if (_isWordActionRunning)
        {
            ShowToast("Word-Aktion läuft bereits. Bitte kurz warten.", ToastType.Warning);
            return;
        }

        var eligible = Participants.Where(p => p.CanInsertEntry).ToList();
        if (eligible.Count == 0)
        {
            ShowToast("Keine aktiven Teilnehmer für Batch-Eintrag.", ToastType.Warning);
            return;
        }

        var rawRows = BatchInputTextBox.Text
            .Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .Select(l => l.Trim())
            .ToList();

        if (rawRows.Count != eligible.Count)
        {
            ShowToast($"Batch-Zeilen ({rawRows.Count}) passen nicht zu aktiven TN ({eligible.Count}).", ToastType.Warning);
            return;
        }

        var normalizedRows = new List<string>();
        for (var i = 0; i < rawRows.Count; i++)
        {
            if (!TryNormalizeBatchRow(rawRows[i], out var normalized, out var error))
            {
                ShowToast($"Batch-Zeile {i + 1} ungültig: {error}", ToastType.Error);
                return;
            }

            normalizedRows.Add(normalized);
        }

        var mapping = eligible.Zip(normalizedRows, (participant, row) => (participant, row)).ToList();
        if (!ShowBatchMappingConfirmation(mapping))
        {
            SetLastAction("Batch abgebrochen (Zuordnung nicht bestätigt).");
            return;
        }

        _batchCancellation?.Dispose();
        _batchCancellation = new CancellationTokenSource();
        var token = _batchCancellation.Token;

        _isBatchRunning = true;
        SetWordActionBusy(true);
        ExecuteBatchButton.IsEnabled = false;
        BatchResults.Clear();

        try
        {
            foreach (var (participant, row) in mapping)
            {
                token.ThrowIfCancellationRequested();
                BatchProgressText.Text = $"Fortschritt: Verarbeite {participant.FullName}...";
                await Task.Yield();
                token.ThrowIfCancellationRequested();

                try
                {
                    var docPath = ResolveDocumentPathForParticipant(participant);
                    if (string.IsNullOrWhiteSpace(docPath))
                    {
                        BatchResults.Add(new BatchResult { Name = participant.FullName, IsSuccess = false, Message = "keine Akte" });
                        continue;
                    }

                    await _wordStaHost.RunAsync("InsertTextRowToTable-Batch", service => service.InsertTextRowToTable(docPath, _config.WordBookmarkName, row));
                    BatchResults.Add(new BatchResult { Name = participant.FullName, IsSuccess = true, Message = "Eintrag eingefügt" });
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (WordTemplateValidationException ex) when (ex.Kind == WordTemplateValidationErrorKind.BookmarkMissing)
                {
                    BatchResults.Add(new BatchResult { Name = participant.FullName, IsSuccess = false, Message = "Bookmark fehlt" });
                }
                catch (WordTemplateValidationException ex) when (ex.Kind == WordTemplateValidationErrorKind.StructuredEntryTableInvalid)
                {
                    BatchResults.Add(new BatchResult { Name = participant.FullName, IsSuccess = false, Message = ex.UserMessage });
                }
                catch (InvalidOperationException ex) when (ex.Message.Contains("gesperrt", StringComparison.OrdinalIgnoreCase))
                {
                    BatchResults.Add(new BatchResult { Name = participant.FullName, IsSuccess = false, Message = "Akte gesperrt" });
                }
                catch (InvalidOperationException ex) when (ex.Message.Contains("Keine Verlaufsakte", StringComparison.OrdinalIgnoreCase))
                {
                    BatchResults.Add(new BatchResult { Name = participant.FullName, IsSuccess = false, Message = "keine Akte" });
                }
                catch (InvalidOperationException ex) when (ex.Message.Contains("Zeilenformat", StringComparison.OrdinalIgnoreCase))
                {
                    BatchResults.Add(new BatchResult { Name = participant.FullName, IsSuccess = false, Message = "Zeilenformat ungültig" });
                }
                catch (Exception ex)
                {
                    BatchResults.Add(new BatchResult { Name = participant.FullName, IsSuccess = false, Message = ex.Message });
                }
            }

            var success = BatchResults.Count(r => r.IsSuccess);
            var failed = BatchResults.Count - success;
            BatchProgressText.Text = $"Fortschritt: abgeschlossen ({success} ✓ / {failed} ✗)";
            SetLastAction($"Batch abgeschlossen: {success} erfolgreich, {failed} fehlgeschlagen.");

            if (failed > 0)
            {
                ShowToast($"Batch abgeschlossen mit {failed} Fehler{(failed == 1 ? string.Empty : "n") }.", ToastType.Warning);
                ShowBatchFailureSummary(BatchResults.Where(r => !r.IsSuccess).ToList());
            }
            else
            {
                ShowToast($"Batch erfolgreich abgeschlossen ({success} Einträge).", ToastType.Success);
            }
        }
        catch (OperationCanceledException)
        {
            BatchProgressText.Text = "Fortschritt: abgebrochen";
            SetLastAction("Batch abgebrochen.");
            ShowToast("Batch wurde abgebrochen.", ToastType.Warning);
        }
        finally
        {
            _isBatchRunning = false;
            SetWordActionBusy(false);
            ExecuteBatchButton.IsEnabled = true;
            _batchCancellation?.Dispose();
            _batchCancellation = null;
        }
    }

    private async void ExecuteBiTodoButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isBiTodoRunning)
        {
            ShowToast("BI: To-dos läuft bereits.", ToastType.Warning);
            return;
        }

        if (!EnsureWordAvailable())
        {
            return;
        }

        var selectedParticipants = Participants
            .Where(p => p.IsPresent)
            .ToList();

        if (selectedParticipants.Count == 0)
        {
            ShowToast("Keine angehakten Teilnehmer für BI: To-dos.", ToastType.Warning);
            SetLastAction("Keine angehakten Teilnehmer für BI: To-dos.");
            return;
        }

        var requests = new List<BiTodoCollectRequest>();

        foreach (var participant in selectedParticipants)
        {
            try
            {
                requests.Add(new BiTodoCollectRequest
                {
                    FullName = participant.FullName,
                    Initials = participant.Initials,
                    DocumentPath = string.IsNullOrWhiteSpace(participant.MatchedFolderPath)
                        ? string.Empty
                        : ResolveDocumentPathForParticipant(participant) ?? string.Empty,
                    FailureMessage = string.IsNullOrWhiteSpace(participant.MatchedFolderPath)
                        ? "kein passender Ordner"
                        : string.Empty
                });
            }
            catch (InvalidOperationException ex)
            {
                requests.Add(new BiTodoCollectRequest
                {
                    FullName = participant.FullName,
                    Initials = participant.Initials,
                    DocumentPath = string.Empty,
                    FailureMessage = ex.Message
                });
            }
            catch (Exception ex)
            {
                requests.Add(new BiTodoCollectRequest
                {
                    FullName = participant.FullName,
                    Initials = participant.Initials,
                    DocumentPath = string.Empty,
                    FailureMessage = $"Vorprüfung fehlgeschlagen: {ex.Message}"
                });
            }
        }

        if (requests.Count == 0)
        {
            BiTodoResults.Clear();
            BiTodoProgressText.Text = "Fortschritt: keine To-dos gesammelt";
            SetLastAction("BI: To-dos konnte nicht gestartet werden.");
            return;
        }

        var wordActionStarted = false;
        try
        {
            if (!TryBeginWordAction("Word-Aktion läuft bereits. Bitte kurz warten."))
            {
                return;
            }
            wordActionStarted = true;

            SetBiTodoBusy(true);
            BiTodoResults.Clear();
            BiTodoProgressText.Text = $"Fortschritt: sammle {requests.Count} Teilnehmer...";

            var title = $"BI, {DateTime.Now.ToString("dddd, dd.MM.yy", CultureInfo.GetCultureInfo("de-CH"))}";
            var summary = await _wordStaHost.RunAsync(
                "CollectBiTodoDocument",
                service => service.CollectBiTodoDocument(requests, _config.WordBiTodoBookmarkName, title));

            BiTodoResults.Clear();
            foreach (var result in summary.Results)
            {
                BiTodoResults.Add(result);
            }

            BiTodoProgressText.Text = $"Fortschritt: abgeschlossen ({summary.SuccessCount} ✓ / {summary.FailureCount} ✗)";
            if (FindResource(summary.FailureCount == 0
                    ? "Brush.Success"
                    : summary.SuccessCount > 0
                        ? "Brush.Warning"
                        : "Brush.Error") is Brush progressBrush)
            {
                BiTodoProgressText.Foreground = progressBrush;
            }
            SetLastAction($"BI: To-dos abgeschlossen: {summary.SuccessCount} erfolgreich, {summary.FailureCount} fehlgeschlagen.");

            if (summary.FailureCount > 0)
            {
                ShowToast($"BI: To-dos abgeschlossen mit {summary.FailureCount} Hinweis{(summary.FailureCount == 1 ? string.Empty : "en")}.", ToastType.Warning);
            }
            else
            {
                ShowToast($"BI: To-dos erfolgreich abgeschlossen ({summary.SuccessCount} Teilnehmer).", ToastType.Success);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("BI: To-dos Fehler", ex);
            BiTodoProgressText.Text = "Fortschritt: Fehler";
            if (FindResource("Brush.Error") is Brush errorBrush)
            {
                BiTodoProgressText.Foreground = errorBrush;
            }
            SetLastAction($"BI: To-dos Fehler: {ex.Message}");
            ShowImportantAlert(
                "BI: To-dos fehlgeschlagen",
                "Der Sammellauf konnte nicht abgeschlossen werden.",
                ex.Message,
                AppAlertKind.Error);
        }
        finally
        {
            SetBiTodoBusy(false);
            if (wordActionStarted)
            {
                EndWordAction();
            }
        }
    }

    private void SetBiTodoBusy(bool isBusy)
    {
        _isBiTodoRunning = isBusy;
        ExecuteBiTodoButton.IsEnabled = !isBusy;
        BiTodoProgressPanel.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;
        if (isBusy && FindResource("Brush.SecondaryText") is Brush secondaryBrush)
        {
            BiTodoProgressText.Foreground = secondaryBrush;
        }
    }

    private bool TryBeginWordAction(string busyMessage)
    {
        if (_isWordActionRunning)
        {
            ShowToast(busyMessage, ToastType.Warning);
            return false;
        }

        SetWordActionBusy(true);
        return true;
    }

    private void EndWordAction()
    {
        SetWordActionBusy(false);
    }

    private void SetWordActionBusy(bool isBusy)
    {
        if (_isWordActionRunning == isBusy)
        {
            return;
        }

        _isWordActionRunning = isBusy;

        foreach (var participant in Participants)
        {
            UpdateActionState(participant);
        }

        if (ExecuteBatchButton is not null)
        {
            ExecuteBatchButton.IsEnabled = !_isWordActionRunning && !_isBatchRunning;
        }

        if (ExecuteBiTodoButton is not null)
        {
            ExecuteBiTodoButton.IsEnabled = !_isWordActionRunning && !_isBiTodoRunning;
        }
    }

    private static bool TryNormalizeBatchRow(string rawRow, out string normalizedRow, out string error)
    {
        normalizedRow = string.Empty;
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(rawRow))
        {
            error = "Zeile ist leer.";
            return false;
        }

        var line = rawRow
            .Replace('\u00A0', ' ')
            .Replace('\u2007', ' ')
            .Replace('\u202F', ' ')
            .Trim();

        if (line.Contains('\t'))
        {
            var parts = line.Split('\t')
                .Select(p => p.Trim())
                .ToList();

            if (parts.Count < 4)
            {
                error = "Zu wenige Spalten (mindestens 4 erwartet).";
                return false;
            }

            var first = parts[0];
            var second = parts[1];
            var third = parts[2];
            var fourth = string.Join(" ", parts.Skip(3).Where(p => !string.IsNullOrWhiteSpace(p)));
            if (string.IsNullOrWhiteSpace(fourth))
            {
                fourth = "-";
            }

            normalizedRow = $"{first}\t{second}\t{third}\t{fourth}";
            return true;
        }

        // Fallback fuer Copy/Paste ohne echte Tabs:
        // Erwartetes Muster: Datum <space> Kuersel <space> Unterrichtsart <space> Resttext
        var tokens = Regex.Split(line, @"\s+")
            .Where(t => !string.IsNullOrWhiteSpace(t))
            .ToList();

        if (tokens.Count < 4)
        {
            error = "Nicht genügend Felder (Format: Datum Kürzel Unterrichtsart Text).";
            return false;
        }

        var c1 = tokens[0];
        var c2 = tokens[1];
        var c3 = tokens[2];
        var c4 = string.Join(" ", tokens.Skip(3)).Trim();

        if (string.IsNullOrWhiteSpace(c4))
        {
            error = "Spalte 4 (Eintragstext) fehlt.";
            return false;
        }

        normalizedRow = $"{c1}\t{c2}\t{c3}\t{c4}";
        return true;
    }

    private bool ShowBatchMappingConfirmation(IReadOnlyList<(Participant participant, string row)> mapping)
    {
        var list = new ListBox
        {
            Margin = new Thickness(0, 0, 0, 10),
            MinWidth = 720,
            Height = 280
        };

        foreach (var pair in mapping)
        {
            list.Items.Add($"{pair.participant.FullName}  <=  {pair.row.Replace("\t", " | ")}");
        }

        var okButton = new Button
        {
            Content = "Bestätigen",
            Width = 120,
            Height = 32,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 0, 8, 0)
        };
        var cancelButton = new Button
        {
            Content = "Abbrechen",
            Width = 120,
            Height = 32,
            HorizontalAlignment = HorizontalAlignment.Right
        };

        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        buttonPanel.Children.Add(okButton);
        buttonPanel.Children.Add(cancelButton);

        var root = new StackPanel { Margin = new Thickness(12) };
        root.Children.Add(new TextBlock
        {
            Text = "Bitte Zuordnung prüfen: Zeile 1 -> TN 1, Zeile 2 -> TN 2 ...",
            Margin = new Thickness(0, 0, 0, 8)
        });
        root.Children.Add(list);
        root.Children.Add(buttonPanel);

        var dlg = new Window
        {
            Owner = this,
            Title = "Batch-Zuordnung bestätigen",
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Width = 860,
            Height = 430,
            ResizeMode = ResizeMode.CanResize,
            Content = root
        };

        okButton.Click += (_, _) => dlg.DialogResult = true;
        cancelButton.Click += (_, _) => dlg.DialogResult = false;
        return dlg.ShowDialog() == true;
    }

    // --- Helpers ---

    private void ShowBatchFailureSummary(IReadOnlyList<BatchResult> failures)
    {
        if (failures.Count == 0)
        {
            return;
        }

        const int maxItems = 8;
        var lines = failures
            .Take(maxItems)
            .Select(f => $"{f.Name}: {f.Message}")
            .ToList();

        if (failures.Count > maxItems)
        {
            lines.Add($"... und {failures.Count - maxItems} weitere");
        }

        AppLogger.Warn($"Batch mit Fehlern abgeschlossen: {string.Join(" | ", failures.Select(f => $"{f.Name}={f.Message}"))}");

        var dialog = new BatchFailureSummaryWindow(lines, failures.Count)
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private void ShowImportantAlert(string title, string lead, string body, AppAlertKind kind = AppAlertKind.Warning, string? footnote = null)
    {
        var dialog = new AppAlertWindow(title, lead, body, kind, footnote)
        {
            Owner = this
        };
        dialog.ShowDialog();
    }

    private async Task BeginStartupUpdateCheckAsync()
    {
        var availableRelease = await _appUpdateService.GetAvailableUpdateAsync(CancellationToken.None);
        if (availableRelease is null || !IsLoaded || _isUpdateShutdownRequested)
        {
            return;
        }

        await Dispatcher.InvokeAsync(() =>
        {
            var updateDialog = new AppUpdateWindow(_appUpdateService, availableRelease)
            {
                Owner = this
            };

            var dialogResult = updateDialog.ShowDialog();
            if (dialogResult == true && updateDialog.DownloadedUpdate is not null)
            {
                TryStartDownloadedUpdate(updateDialog.DownloadedUpdate);
                return;
            }

            if (updateDialog.WasDeferred)
            {
                _appUpdateService.SnoozeRelease(availableRelease);
            }
        });
    }

    private bool TryStartDownloadedUpdate(DownloadedUpdateInfo downloadedUpdate)
    {
        if (_isBatchRunning || _isBiTodoRunning || _isWordActionRunning)
        {
            ShowImportantAlert(
                "Update momentan nicht möglich",
                "Scola arbeitet gerade noch.",
                "Bitte warte, bis Batch, BI: To-dos oder eine Word-Aktion fertig sind. Danach kann das Update beim nächsten Start erneut durchgeführt werden.",
                AppAlertKind.Warning);
            return false;
        }

        try
        {
            _appUpdateService.LaunchUpdater(downloadedUpdate);
            AppLogger.Info("Updater: Update-Shutdown gestartet.");
            _isUpdateShutdownRequested = true;
            (Application.Current as App)?.PrepareForUpdateShutdown();
            Close();
            return true;
        }
        catch (Exception ex)
        {
            AppLogger.Error("Updater: Updater konnte nicht gestartet werden.", ex);
            ShowImportantAlert(
                "Update konnte nicht gestartet werden",
                "Der Updater liess sich nicht starten.",
                $"Das Update wurde bereits heruntergeladen, konnte aber nicht uebernommen werden: {ex.Message}",
                AppAlertKind.Error);
            return false;
        }
    }

    private static string[]? BuildFallbackEntryFieldsIfEnabled()
    {
        if (!App.UserPrefs.AutoPrefillOnEmptyClipboard)
        {
            return null;
        }

        var date = DateTime.Now.ToString("dd.MM.yy");
        var initials = (App.UserPrefs.DefaultEntryInitials ?? string.Empty).Trim();
        return new[] { date, initials, string.Empty, string.Empty };
    }

    private bool EnsureWordAvailable()
    {
        if (IsWordAvailable)
            return true;

        ShowToast("Microsoft Word nicht gefunden", ToastType.Error);
        SetLastAction("Microsoft Word nicht gefunden");
        return false;
    }

    private bool EnsureParticipantFolder(Participant participant)
    {
        if (!string.IsNullOrWhiteSpace(participant.MatchedFolderPath))
            return true;

        ShowToast($"Kein passender Ordner für: {participant.FullName}", ToastType.Warning);
        SetLastAction($"Kein passender Ordner für: {participant.FullName}");
        return false;
    }

    private string? ResolveDocumentPathForParticipant(Participant participant)
    {
        if (!string.IsNullOrWhiteSpace(participant.DocumentPath) && File.Exists(participant.DocumentPath))
        {
            return participant.DocumentPath;
        }

        var docs = _wordService.FindVerlaufsakteCandidates(participant.MatchedFolderPath!, _config.VerlaufsakteKeyword);
        string? selected = null;

        if (docs.Count == 1)
        {
            selected = docs[0];
        }
        else if (docs.Count > 1)
        {
            selected = SelectFromCandidates($"Mehrere Treffer für {participant.FullName} - bitte manuell auswählen", docs);
        }

        if (string.IsNullOrWhiteSpace(selected))
        {
            participant.DocumentPath = string.Empty;
            participant.Initials = string.Empty;
            participant.OdooUrl = string.Empty;
            return null;
        }

        participant.DocumentPath = selected;
        participant.Initials = _initialsResolver.TryResolveFromDocumentPath(selected);
        participant.OdooUrl = string.Empty;
        return selected;
    }

    private string? SelectFromCandidates(string title, IReadOnlyList<string> options)
    {
        if (options.Count == 0)
            return null;

        var dialog = new Window
        {
            Owner = this,
            Title = title,
            Width = 640,
            Height = 340,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ResizeMode = ResizeMode.CanResize,
            Content = BuildCandidateDialogContent(options)
        };

        if (dialog.Content is not Grid grid)
            return null;

        var listBox = (ListBox)grid.Children[0];
        var okButton = (Button)grid.Children[1];
        okButton.Click += (_, _) => dialog.DialogResult = true;

        var result = dialog.ShowDialog();
        if (result == true)
            return listBox.SelectedItem as string ?? options[0];

        return null;
    }

    private static Grid BuildCandidateDialogContent(IReadOnlyList<string> options)
    {
        var grid = new Grid { Margin = new Thickness(12) };

        grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var listBox = new ListBox
        {
            ItemsSource = options,
            SelectedIndex = 0,
            Margin = new Thickness(0, 0, 0, 12)
        };
        Grid.SetRow(listBox, 0);
        grid.Children.Add(listBox);

        var okButton = new Button
        {
            Content = "Auswählen",
            Width = 100,
            Height = 32,
            HorizontalAlignment = HorizontalAlignment.Right
        };
        Grid.SetRow(okButton, 1);
        grid.Children.Add(okButton);

        return grid;
    }

    private void UpdateActionState(Participant participant)
    {
        RefreshParticipantDocumentData(participant);
        var allowWordActions = !_isWordActionRunning;
        participant.CanOpenFolder = participant.IsPresent && !string.IsNullOrWhiteSpace(participant.MatchedFolderPath);
        participant.CanOpenOdoo = participant.CanOpenFolder && participant.HasOdooUrl;
        participant.CanOpenAkte = participant.CanOpenFolder && IsWordAvailable && allowWordActions;
        participant.CanOpenAkteBu = participant.CanOpenFolder && IsWordAvailable && allowWordActions;
        participant.CanInsertEntry = participant.CanOpenFolder && IsWordAvailable && allowWordActions;
        participant.CanOpenAkteBi = participant.CanOpenFolder && IsWordAvailable && allowWordActions;
        participant.CanOpenAkteBe = participant.CanOpenFolder && IsWordAvailable && allowWordActions;
        participant.CanInsertEntryBi = participant.CanOpenFolder && IsWordAvailable && allowWordActions;
    }

    private void RefreshParticipantDocumentData(Participant participant)
    {
        if (string.IsNullOrWhiteSpace(participant.MatchedFolderPath) || !Directory.Exists(participant.MatchedFolderPath))
        {
            participant.DocumentPath = string.Empty;
            participant.Initials = string.Empty;
            participant.OdooUrl = string.Empty;
            return;
        }

        if (!string.IsNullOrWhiteSpace(participant.DocumentPath)
            && File.Exists(participant.DocumentPath)
            && string.Equals(Path.GetDirectoryName(participant.DocumentPath), participant.MatchedFolderPath, StringComparison.OrdinalIgnoreCase))
        {
            participant.Initials = _initialsResolver.TryResolveFromDocumentPath(participant.DocumentPath);
            return;
        }

        var preferredPath = TryResolvePreferredDocumentPath(participant.MatchedFolderPath);
        if (string.IsNullOrWhiteSpace(preferredPath))
        {
            participant.DocumentPath = string.Empty;
            participant.Initials = string.Empty;
            participant.OdooUrl = string.Empty;
            return;
        }

        if (!string.Equals(participant.DocumentPath, preferredPath, StringComparison.OrdinalIgnoreCase))
        {
            participant.DocumentPath = preferredPath;
            participant.OdooUrl = string.Empty;
        }

        participant.Initials = _initialsResolver.TryResolveFromDocumentPath(participant.DocumentPath);
    }

    private string? TryResolvePreferredDocumentPath(string folderPath)
    {
        var docs = _wordService.FindVerlaufsakteCandidates(folderPath, _config.VerlaufsakteKeyword);
        if (docs.Count == 0)
        {
            return null;
        }

        var preferred = docs
            .OrderBy(path => path, StringComparer.CurrentCultureIgnoreCase)
            .First();

        if (docs.Count > 1)
        {
            AppLogger.Debug($"Mehrere Verlaufsakten gefunden, verwende bevorzugten Default-Pfad '{preferred}'. Count={docs.Count}.");
        }

        return preferred;
    }

    private void BeginOdooMetadataWarmupForCurrentParticipants()
    {
        if (!ShowBtnOdoo || Participants.Count == 0)
        {
            return;
        }

        var targets = Participants
            .Where(participant => !string.IsNullOrWhiteSpace(participant.DocumentPath))
            .Select(participant => participant)
            .ToList();

        if (targets.Count == 0)
        {
            return;
        }

        _ = Task.Run(() =>
        {
            foreach (var participant in targets)
            {
                var documentPath = participant.DocumentPath;
                if (string.IsNullOrWhiteSpace(documentPath) || !File.Exists(documentPath))
                {
                    continue;
                }

                var metadata = _headerMetadataService.Read(documentPath);
                Dispatcher.Invoke(() =>
                {
                    if (!string.Equals(participant.DocumentPath, documentPath, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    participant.OdooUrl = metadata.OdooUrl;
                    UpdateActionState(participant);
                });
            }
        }).ContinueWith(t => AppLogger.Error($"Odoo-Metadata-Warmup fehlgeschlagen: {t.Exception?.GetBaseException().Message}", t.Exception?.GetBaseException()), TaskContinuationOptions.OnlyOnFaulted);
    }

    private bool EnsureOdooMetadataLoaded(Participant participant)
    {
        var documentPath = ResolveDocumentPathForParticipant(participant);
        if (string.IsNullOrWhiteSpace(documentPath) || !File.Exists(documentPath))
        {
            return false;
        }

        var metadata = _headerMetadataService.Read(documentPath);
        participant.OdooUrl = metadata.OdooUrl;
        UpdateActionState(participant);
        return participant.HasOdooUrl;
    }

    private static void ResetParticipantMatch(Participant participant)
    {
        participant.MatchStatus = MatchStatus.NotFound;
        participant.CandidateFolderPaths = new List<string>();
        participant.SelectedFolderPath = null;
        participant.MatchedFolderPath = null;
        participant.DocumentPath = string.Empty;
        participant.Initials = string.Empty;
        participant.OdooUrl = string.Empty;
    }

    // --- Status bar ---

    private void SetLastAction(string message)
    {
        _lastActionText = message;
        StatusBarText.Text = message;
        StatusBarTimestamp.Text = DateTime.Now.ToString("HH:mm:ss");
        AppLogger.Info($"LastAction: {_lastActionText}");
    }

    private static string BuildApplicationVersionText()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var productName = assembly.GetName().Name ?? "Scola";
        var version = assembly.GetName().Version;
        if (version is null)
        {
            return $"{productName} 0.0.0";
        }

        if (version.Revision > 0)
        {
            return $"{productName} {version.Major}.{Math.Max(0, version.Minor)}.{Math.Max(0, version.Build)}.{version.Revision}";
        }

        return $"{productName} {version.Major}.{Math.Max(0, version.Minor)}.{Math.Max(0, version.Build)}";
    }

    // --- Toast notifications ---

    private enum ToastType { Success, Info, Warning, Error }

    private void ShowToast(string message, ToastType type = ToastType.Info)
    {
        var isDark = App.UserPrefs.IsDarkTheme;
        var (accentColor, bgColor) = type switch
        {
            ToastType.Success => ("#4CAF50", isDark ? "#1B3A1B" : "#E7F6E7"),
            ToastType.Warning => ("#C8A96C", isDark ? "#3A331B" : "#FAF2E1"),
            ToastType.Error   => ("#D17878", isDark ? "#3A1B1B" : "#FBEAEA"),
            _                 => ("#5B9BD5", isDark ? "#1B2A3A" : "#E8F1FB"),
        };

        var border = new Border
        {
            Background = BrushFromHex(bgColor),
            BorderBrush = BrushFromHex(accentColor),
            BorderThickness = new Thickness(3, 0, 0, 0),
            CornerRadius = new CornerRadius(6),
            Padding = new Thickness(12, 8, 12, 8),
            Margin = new Thickness(0, 4, 0, 0),
            MinWidth = 200,
            MaxWidth = 320,
            Opacity = 0,
            RenderTransform = new TranslateTransform(20, 0),
            Effect = new DropShadowEffect
            {
                BlurRadius = 8,
                ShadowDepth = 2,
                Opacity = 0.3,
                Color = Colors.Black
            },
            Child = new TextBlock
            {
                Text = message,
                Foreground = BrushFromHex(isDark ? "#E0E0E0" : "#1A1A1F"),
                FontSize = 12,
                TextWrapping = TextWrapping.Wrap
            }
        };

        // Limit visible toasts
        while (ToastContainer.Items.Count >= 3)
            ToastContainer.Items.RemoveAt(0);

        ToastContainer.Items.Add(border);

        // Slide-in + fade-in
        var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(250))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };
        var slideIn = new DoubleAnimation(20, 0, TimeSpan.FromMilliseconds(250))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };
        border.BeginAnimation(OpacityProperty, fadeIn);
        ((TranslateTransform)border.RenderTransform).BeginAnimation(TranslateTransform.XProperty, slideIn);

        // Auto-dismiss
        var timer = new System.Windows.Threading.DispatcherTimer
        {
            Interval = TimeSpan.FromSeconds(4)
        };
        timer.Tick += (_, _) =>
        {
            timer.Stop();
            var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(300));
            fadeOut.Completed += (_, _) =>
            {
                if (ToastContainer.Items.Contains(border))
                    ToastContainer.Items.Remove(border);
            };
            border.BeginAnimation(OpacityProperty, fadeOut);
        };
        timer.Start();
    }

    private static SolidColorBrush BrushFromHex(string hex)
    {
        var color = (Color)ColorConverter.ConvertFromString(hex);
        return new SolidColorBrush(color);
    }

    private static class MonitorNative
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct MONITORINFOEX
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public int dwFlags;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szDevice;
        }

        private delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, IntPtr lprcMonitor, IntPtr dwData);

        [DllImport("user32.dll")]
        private static extern bool EnumDisplayMonitors(
            IntPtr hdc,
            IntPtr lprcClip,
            MonitorEnumProc lpfnEnum,
            IntPtr dwData);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

        private const int MONITORINFOF_PRIMARY = 0x00000001;

        public static List<MonitorDescriptor> EnumerateMonitors()
        {
            var monitors = new List<MonitorDescriptor>();

            EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, (hMonitor, _, _, _) =>
            {
                var info = new MONITORINFOEX
                {
                    cbSize = Marshal.SizeOf<MONITORINFOEX>(),
                    szDevice = string.Empty
                };

                if (GetMonitorInfo(hMonitor, ref info))
                {
                    var workArea = new Rect(
                        info.rcWork.Left,
                        info.rcWork.Top,
                        Math.Max(1, info.rcWork.Right - info.rcWork.Left),
                        Math.Max(1, info.rcWork.Bottom - info.rcWork.Top));

                    var deviceName = (info.szDevice ?? string.Empty).TrimEnd('\0');
                    var isPrimary = (info.dwFlags & MONITORINFOF_PRIMARY) != 0;
                    monitors.Add(new MonitorDescriptor(deviceName, workArea, isPrimary));
                }

                return true;
            }, IntPtr.Zero);

            return monitors;
        }
    }
}
