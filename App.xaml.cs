using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using VerlaufsakteApp.Models;
using VerlaufsakteApp.Services;
using Forms = System.Windows.Forms;

namespace VerlaufsakteApp;

public partial class App : System.Windows.Application
{
    private const string DefaultServerBasePath = @"K:\FuturX\20_TNinnen";
    private const string DefaultSecondaryServerBasePath = @"K:\FuturX\20_TNinnen\02_Lehrbegleitung";
    private const string DefaultScheduleRootPath = @"K:\FuturX\10_Arbeitsplanung\20_Planung\22_Wochenplanung\Einteilung TN";
    private const string SingleInstanceMutexName = @"Local\Scola.SingleInstance";
    private const string SingleInstanceActivateEventName = @"Local\Scola.SingleInstance.Activate";
    private const int SwRestore = 9;
    private static readonly JsonSerializerOptions ConfigReadSerializerOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };
    private static readonly JsonSerializerOptions ConfigWriteSerializerOptions = new()
    {
        WriteIndented = true
    };

    private Mutex? _singleInstanceMutex;
    private EventWaitHandle? _singleInstanceActivateEvent;
    private RegisteredWaitHandle? _singleInstanceActivateRegistration;
    private bool _pendingForegroundRequest;
    private bool _singleInstanceInfrastructureReleased;

    public static AppConfig? Config { get; private set; }
    public static UserPrefs UserPrefs { get; private set; } = new();
    public static string AppDataDirectoryPath { get; private set; } = string.Empty;
    public static string SettingsPath { get; private set; } = string.Empty;
    public static string UserPrefsPath { get; private set; } = string.Empty;
    public static string WeeklyScheduleCachePath { get; private set; } = string.Empty;
    public static string WeeklyScheduleCacheBackupPath { get; private set; } = string.Empty;
    public static string WeeklyScheduleDiagnosticsPath { get; private set; } = string.Empty;
    public static string WeeklyScheduleDiagnosticsBackupPath { get; private set; } = string.Empty;

    protected override void OnStartup(StartupEventArgs e)
    {
        RegisterGlobalExceptionLogging();
        InitializeSingleInstanceInfrastructure();

        if (!TryAcquireSingleInstance())
        {
            Shutdown();
            return;
        }

        AppDataDirectoryPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AkteX");
        SettingsPath = Path.Combine(AppDataDirectoryPath, "settings.json");
        UserPrefsPath = Path.Combine(AppDataDirectoryPath, "user-prefs.json");
        WeeklyScheduleCachePath = Path.Combine(AppDataDirectoryPath, "weekly-schedule-cache.json");
        WeeklyScheduleCacheBackupPath = Path.Combine(AppDataDirectoryPath, "weekly-schedule-cache.bak");
        var diagnosticsDirectory = ResolveDiagnosticsDirectory(AppDataDirectoryPath);
        WeeklyScheduleDiagnosticsPath = Path.Combine(diagnosticsDirectory, "weekly-schedule-diagnostics.json");
        WeeklyScheduleDiagnosticsBackupPath = Path.Combine(diagnosticsDirectory, "weekly-schedule-diagnostics.bak");
        Directory.CreateDirectory(AppDataDirectoryPath);

        EnsureSettingsFileExists(Path.Combine(System.AppContext.BaseDirectory, "settings.json"));
        MigrateLegacyUserPrefsIfNeeded();

        AppLogger.Info($"App-Start. settings.json Pfad: {SettingsPath}");
        if (!File.Exists(SettingsPath))
        {
            AppLogger.Error($"settings.json nicht gefunden: {SettingsPath}");
            MessageBox.Show("settings.json nicht gefunden. Bitte App neu starten oder Einstellungen prüfen.", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown();
            return;
        }

        try
        {
            var json = File.ReadAllText(SettingsPath);
            Config = JsonSerializer.Deserialize<AppConfig>(json, ConfigReadSerializerOptions);

            if (Config is null)
            {
                throw new InvalidDataException("settings.json ist leer oder ungültig.");
            }

            if (string.IsNullOrWhiteSpace(Config.ServerBasePath))
            {
                throw new InvalidDataException("ServerBasePath fehlt in settings.json.");
            }

            Config.SecondaryServerBasePath ??= string.Empty;
            Config.TertiaryServerBasePath ??= string.Empty;
            Config.ScheduleRootPath ??= string.Empty;
            Config.AbsenceValues ??= new List<string>();
            Config.PresenceValues ??= new List<string> { "anwesend" };
            if (string.IsNullOrWhiteSpace(Config.WordBookmarkName))
            {
                Config.WordBookmarkName = "BU_BILDUNG_TABELLE";
            }

            if (string.IsNullOrWhiteSpace(Config.WordBuBookmarkName))
            {
                Config.WordBuBookmarkName = "_Bildung";
            }

            if (string.IsNullOrWhiteSpace(Config.WordBiBookmarkName))
            {
                Config.WordBiBookmarkName = "_Berufsintegration";
            }

            if (string.IsNullOrWhiteSpace(Config.WordBeBookmarkName))
            {
                Config.WordBeBookmarkName = "_Beratung";
            }

            if (string.IsNullOrWhiteSpace(Config.WordBiTableBookmarkName))
            {
                Config.WordBiTableBookmarkName = "BI_BERUFSINTEGRATION_TABELLE";
            }

            if (string.IsNullOrWhiteSpace(Config.WordBiTodoBookmarkName))
            {
                Config.WordBiTodoBookmarkName = "BI_BERUFSINTEGRATION_TODO";
            }

            if (ShouldForceDefaultServerPath(Config.ServerBasePath))
            {
                Config.ServerBasePath = DefaultServerBasePath;
                PersistConfig(Config, "ServerBasePath auf neuen Standard gesetzt.");
            }

            if (ApplyDefaultConfigPaths(Config))
            {
                PersistConfig(Config, "Standardpfade ergänzt.");
            }

            AppLogger.Info($"Konfiguration geladen. ServerBasePath='{Config.ServerBasePath}', UseSecondaryServerBasePath={Config.UseSecondaryServerBasePath}, SecondaryServerBasePath='{Config.SecondaryServerBasePath}', UseTertiaryServerBasePath={Config.UseTertiaryServerBasePath}, TertiaryServerBasePath='{Config.TertiaryServerBasePath}', WordBookmarkName='{Config.WordBookmarkName}', WordBuBookmarkName='{Config.WordBuBookmarkName}', WordBiBookmarkName='{Config.WordBiBookmarkName}', WordBeBookmarkName='{Config.WordBeBookmarkName}', WordBiTableBookmarkName='{Config.WordBiTableBookmarkName}', WordBiTodoBookmarkName='{Config.WordBiTodoBookmarkName}'");
        }
        catch (System.Exception ex)
        {
            AppLogger.Error("Fehler beim Laden von settings.json", ex);
            MessageBox.Show($"Fehler beim Laden von settings.json: {ex.Message}", "Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown();
            return;
        }

        var prefsService = new UserPrefsService(UserPrefsPath);
        UserPrefs = prefsService.Load();
        UserPrefs.DisplayDensity = Models.DisplayDensityMode.Normalize(UserPrefs.DisplayDensity);
        UserPrefs.PreferredWordMonitorId = NormalizePreferredWordMonitorId(UserPrefs.PreferredWordMonitorId);
        AppLogger.SetDebugEnabled(UserPrefs.EnableDebugLogging);
        ApplyTheme(UserPrefs.IsDarkTheme);

        base.OnStartup(e);

        var window = new MainWindow();
        MainWindow = window;
        window.Show();

        if (_pendingForegroundRequest)
        {
            _pendingForegroundRequest = false;
            BringWindowToFront(window);
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        ReleaseSingleInstanceInfrastructure();
        base.OnExit(e);
    }

    public void PrepareForUpdateShutdown()
    {
        AppLogger.Info("Update-Shutdown: Single-Instance-Infrastruktur wird vorzeitig freigegeben.");
        ReleaseSingleInstanceInfrastructure();
    }

    public static void SaveUserPrefs()
    {
        var service = new UserPrefsService(UserPrefsPath);
        service.Save(UserPrefs);
    }

    public static void ApplyTheme(bool isDark)
    {
        UserPrefs.IsDarkTheme = isDark;
        var res = Current.Resources;

        SetBrush(res, "Brush.WindowBg", isDark ? "#1F2024" : "#F0F0F2");
        SetBrush(res, "Brush.PanelBg", isDark ? "#2A2B31" : "#FFFFFF");
        SetBrush(res, "Brush.CardBg", isDark ? "#2E2F36" : "#F5F5F7");
        SetBrush(res, "Brush.CardHover", isDark ? "#34353D" : "#E8E8EC");
        SetBrush(res, "Brush.PrimaryText", isDark ? "#E0E0E0" : "#1A1A1F");
        SetBrush(res, "Brush.SecondaryText", isDark ? "#9B9EA8" : "#6B6E78");
        SetBrush(res, "Brush.SubtleText", isDark ? "#6F7480" : "#7A7D88");
        SetBrush(res, "Brush.Border", isDark ? "#3A3B42" : "#D0D1D8");
        SetBrush(res, "Brush.Accent", isDark ? "#8B7D6B" : "#7A6E60");
        SetBrush(res, "Brush.AccentHover", isDark ? "#A0917D" : "#8E8376");
        SetBrush(res, "Brush.AccentPressed", isDark ? "#6E6356" : "#6A5F52");
        SetBrush(res, "Brush.TitleBarBg", isDark ? "#2A2B31" : "#EDEDF1");
        SetBrush(res, "Brush.StatusBarBg", isDark ? "#252630" : "#E8E8EC");
        SetBrush(res, "Brush.InputBg", isDark ? "#23242A" : "#F7F7FA");
        SetBrush(res, "Brush.InputHoverBorder", isDark ? "#5A5C66" : "#B8BCC8");
        SetBrush(res, "Brush.ScrollThumb", isDark ? "#3A3B42" : "#C0C3CC");
        SetBrush(res, "Brush.ScrollThumbHover", isDark ? "#5A5C66" : "#A0A5B2");
        SetBrush(res, "Brush.Success", "#4CAF50");
        SetBrush(res, "Brush.Warning", "#C8A96C");
        SetBrush(res, "Brush.Error", "#D17878");
        SetBrush(res, "Brush.Info", "#5B9BD5");
    }

    private static void SetBrush(ResourceDictionary resources, string key, string colorHex)
    {
        var color = (Color)ColorConverter.ConvertFromString(colorHex);
        resources[key] = new SolidColorBrush(color);
    }

    private void InitializeSingleInstanceInfrastructure()
    {
        _singleInstanceActivateEvent = new EventWaitHandle(
            false,
            EventResetMode.AutoReset,
            SingleInstanceActivateEventName);

        _singleInstanceActivateRegistration = ThreadPool.RegisterWaitForSingleObject(
            _singleInstanceActivateEvent,
            static (state, _) =>
            {
                if (state is not App app)
                {
                    return;
                }

                app.Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (Current.MainWindow is null)
                    {
                        app._pendingForegroundRequest = true;
                        return;
                    }

                    app._pendingForegroundRequest = false;
                    BringWindowToFront(Current.MainWindow);
                }));
            },
            this,
            Timeout.Infinite,
            false);
    }

    private bool TryAcquireSingleInstance()
    {
        _singleInstanceMutex = new Mutex(true, SingleInstanceMutexName, out var createdNew);
        if (createdNew)
        {
            return true;
        }

        AppLogger.Info("Zweite Instanz erkannt. Bestehende Instanz wird aktiviert.");

        try
        {
            _singleInstanceActivateEvent?.Set();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Aktivierungssignal an bestehende Instanz fehlgeschlagen: {ex.Message}");
        }

        TryBringExistingProcessWindowToFront();
        return false;
    }

    private void ReleaseSingleInstanceInfrastructure()
    {
        if (_singleInstanceInfrastructureReleased)
        {
            return;
        }

        _singleInstanceInfrastructureReleased = true;

        try
        {
            _singleInstanceActivateRegistration?.Unregister(null);
        }
        catch
        {
        }
        finally
        {
            _singleInstanceActivateRegistration = null;
        }

        try
        {
            _singleInstanceActivateEvent?.Dispose();
        }
        catch
        {
        }
        finally
        {
            _singleInstanceActivateEvent = null;
        }

        if (_singleInstanceMutex is not null)
        {
            try
            {
                _singleInstanceMutex.ReleaseMutex();
            }
            catch
            {
            }

            try
            {
                _singleInstanceMutex.Dispose();
            }
            catch
            {
            }

            _singleInstanceMutex = null;
        }
    }

    private static void BringWindowToFront(Window? window)
    {
        if (window is null)
        {
            return;
        }

        if (window.WindowState == WindowState.Minimized)
        {
            window.WindowState = WindowState.Normal;
        }

        window.Show();
        window.Activate();

        try
        {
            window.Topmost = true;
            window.Topmost = false;
        }
        catch
        {
        }

        try
        {
            var handle = new WindowInteropHelper(window).Handle;
            if (handle != IntPtr.Zero)
            {
                NativeMethods.ShowWindowAsync(handle, SwRestore);
                NativeMethods.SetForegroundWindow(handle);
            }
        }
        catch
        {
        }

        try
        {
            window.Focus();
        }
        catch
        {
        }
    }

    private static void TryBringExistingProcessWindowToFront()
    {
        try
        {
            var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
            var currentPath = currentProcess.MainModule?.FileName;
            if (string.IsNullOrWhiteSpace(currentPath))
            {
                return;
            }

            var otherInstance = System.Diagnostics.Process
                .GetProcessesByName(currentProcess.ProcessName)
                .Where(process => process.Id != currentProcess.Id)
                .FirstOrDefault(process =>
                {
                    try
                    {
                        return string.Equals(
                            process.MainModule?.FileName,
                            currentPath,
                            StringComparison.OrdinalIgnoreCase);
                    }
                    catch
                    {
                        return false;
                    }
                });

            if (otherInstance is null)
            {
                return;
            }

            otherInstance.Refresh();
            var handle = otherInstance.MainWindowHandle;
            if (handle == IntPtr.Zero)
            {
                return;
            }

            NativeMethods.ShowWindowAsync(handle, SwRestore);
            NativeMethods.SetForegroundWindow(handle);
        }
        catch
        {
        }
    }

    private static void RegisterGlobalExceptionLogging()
    {
        Current.DispatcherUnhandledException += (_, args) =>
        {
            AppLogger.Error("Application.DispatcherUnhandledException", args.Exception);
        };

        AppDomain.CurrentDomain.UnhandledException += (_, args) =>
        {
            var ex = args.ExceptionObject as Exception;
            AppLogger.Error($"AppDomain.UnhandledException (IsTerminating={args.IsTerminating})", ex);
        };

        TaskScheduler.UnobservedTaskException += (_, args) =>
        {
            AppLogger.Error("TaskScheduler.UnobservedTaskException", args.Exception);
            args.SetObserved();
        };
    }

    private static string ResolveDiagnosticsDirectory(string appDataDirectory)
    {
        var localDiagnosticsDirectory = Path.Combine(AppContext.BaseDirectory, "diagnostics");
        if (TryEnsureWritableDirectory(localDiagnosticsDirectory))
        {
            return localDiagnosticsDirectory;
        }

        var appDataDiagnosticsDirectory = Path.Combine(appDataDirectory, "diagnostics");
        Directory.CreateDirectory(appDataDiagnosticsDirectory);
        return appDataDiagnosticsDirectory;
    }

    private static bool TryEnsureWritableDirectory(string directoryPath)
    {
        try
        {
            Directory.CreateDirectory(directoryPath);
            var probePath = Path.Combine(directoryPath, ".write-test.tmp");
            File.WriteAllText(probePath, string.Empty);
            File.Delete(probePath);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void MigrateLegacyUserPrefsIfNeeded()
    {
        try
        {
            if (File.Exists(UserPrefsPath))
            {
                return;
            }

            var candidates = new[]
            {
                Path.Combine(System.AppContext.BaseDirectory, "user-prefs.json"),
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "AkteX",
                    "user-prefs.json"),
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "AktenManager",
                    "user-prefs.json"),
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "AktenTracker",
                    "user-prefs.json"),
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "VerlaufsakteApp",
                    "user-prefs.json")
            };

            var source = candidates.FirstOrDefault(File.Exists);
            if (string.IsNullOrWhiteSpace(source))
            {
                return;
            }

            File.Copy(source, UserPrefsPath, overwrite: false);
            AppLogger.Info($"Legacy user-prefs migriert von '{source}' nach '{UserPrefsPath}'.");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Legacy user-prefs konnten nicht migriert werden: {ex.Message}");
        }
    }

    private static void EnsureSettingsFileExists(string bundledSettingsPath)
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                return;
            }

            var candidates = new[]
            {
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "AktenManager",
                    "settings.json"),
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "AktenTracker",
                    "settings.json"),
                Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "VerlaufsakteApp",
                    "settings.json"),
                bundledSettingsPath
            };

            var source = candidates.FirstOrDefault(File.Exists);
            if (!string.IsNullOrWhiteSpace(source))
            {
                File.Copy(source, SettingsPath, overwrite: false);
                AppLogger.Info($"settings.json initialisiert von '{source}' nach '{SettingsPath}'.");
                return;
            }

            var defaultConfig = new AppConfig
            {
                ServerBasePath = DefaultServerBasePath,
                UseSecondaryServerBasePath = true,
                SecondaryServerBasePath = DefaultSecondaryServerBasePath,
                UseTertiaryServerBasePath = false,
                TertiaryServerBasePath = string.Empty,
                ScheduleRootPath = DefaultScheduleRootPath,
                AbsenceValues = new List<string> { "Abwesend (entschuldigt)", "Abwesend (unentschuldigt)", "abwesend" },
                PresenceValues = new List<string> { "Anwesend", "anwesend", "Verspätet", "verspätet" },
                VerlaufsakteKeyword = "Verlaufsakte",
                WordBookmarkName = "BU_BILDUNG_TABELLE",
                WordBuBookmarkName = "_Bildung",
                WordBiBookmarkName = "_Berufsintegration",
                WordBeBookmarkName = "_Beratung",
                WordBiTableBookmarkName = "BI_BERUFSINTEGRATION_TABELLE"
                ,
                WordBiTodoBookmarkName = "BI_BERUFSINTEGRATION_TODO"
            };

            var json = JsonSerializer.Serialize(defaultConfig, ConfigWriteSerializerOptions);
            File.WriteAllText(SettingsPath, json);
            AppLogger.Warn($"settings.json neu mit Standardwerten erstellt: {SettingsPath}");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"settings.json konnte nicht initialisiert werden: {ex.Message}", ex);
        }
    }

    private static bool ShouldForceDefaultServerPath(string? serverBasePath)
    {
        var normalized = (serverBasePath ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return true;
        }

        var normalizedLower = normalized.ToLowerInvariant();
        if (normalizedLower.Contains(@"\testsetup\mock-env\mockserver"))
        {
            return true;
        }

        return normalizedLower is
            "c:\\pfad\\zu\\teilnehmenden" or
            "\\\\server\\share\\teilnehmende";
    }

    private static bool ApplyDefaultConfigPaths(AppConfig config)
    {
        var changed = false;

        if (string.IsNullOrWhiteSpace(config.SecondaryServerBasePath))
        {
            config.SecondaryServerBasePath = DefaultSecondaryServerBasePath;
            changed = true;
        }

        if (string.IsNullOrWhiteSpace(config.ScheduleRootPath))
        {
            config.ScheduleRootPath = DefaultScheduleRootPath;
            changed = true;
        }

        if (string.IsNullOrWhiteSpace(config.SecondaryServerBasePath))
        {
            return changed;
        }

        if (!config.UseSecondaryServerBasePath)
        {
            config.UseSecondaryServerBasePath = true;
            changed = true;
        }

        return changed;
    }

    private static void PersistConfig(AppConfig config, string reason)
    {
        try
        {
            var json = JsonSerializer.Serialize(config, ConfigWriteSerializerOptions);
            File.WriteAllText(SettingsPath, json);
            AppLogger.Info($"{reason} settings.json aktualisiert: {SettingsPath}");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"settings.json konnte nach Standard-Pfad-Korrektur nicht gespeichert werden: {ex.Message}");
        }
    }

    private static string NormalizePreferredWordMonitorId(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return SettingsWindowModel.PrimaryWordMonitorId;
        }

        var normalized = value.Trim();
        if (string.Equals(normalized, SettingsWindowModel.PrimaryWordMonitorId, StringComparison.OrdinalIgnoreCase))
        {
            return SettingsWindowModel.PrimaryWordMonitorId;
        }

        var primaryScreen = Forms.Screen.PrimaryScreen;
        if (primaryScreen is not null &&
            string.Equals(normalized, primaryScreen.DeviceName, StringComparison.OrdinalIgnoreCase))
        {
            return SettingsWindowModel.PrimaryWordMonitorId;
        }

        return normalized;
    }

    private static class NativeMethods
    {
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
}




