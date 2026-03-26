using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using Forms = System.Windows.Forms;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp.Services;

public class WordService
{
    private const string PrimaryMonitorId = "__PRIMARY__";
    private const string DocumentLockedMessage = "Akte ist bereits offen oder gesperrt (evtl. durch anderen Benutzer). Bitte später erneut versuchen.";
    private const int WordForegroundRetryDelayMs = 80;
    private const int ClipboardReadRetryCount = 3;
    private const int ClipboardReadRetryBaseDelayMs = 100;
    private const string BiTodoTemplateFileName = "Bi Vorlage 2.docx";
    private const string BiTodoTemplateResourceName = "VerlaufsakteApp.Templates.BiVorlage.docx";
    private const int BiTodoInitialsColor = 0x45B0E1;
    private static readonly Regex MultiWhitespaceRegex = new(@"\s+", RegexOptions.Compiled);

    private sealed class BiTodoTemplateDefinition
    {
        public required dynamic TitleParagraph { get; init; }
        public required dynamic HeaderParagraph { get; init; }
        public required dynamic CareerChoiceParagraph { get; init; }
        public required dynamic BulletParagraph { get; init; }
        public required dynamic SeparatorParagraph { get; init; }
        public dynamic? BlankParagraph { get; init; }
    }

    private sealed class BiTodoParagraphContent
    {
        public string Text { get; init; } = string.Empty;
        public bool IsBullet { get; init; }
    }

    private sealed class BiTodoParticipantContent
    {
        public string Initials { get; init; } = string.Empty;
        public string FullName { get; init; } = string.Empty;
        public string CareerChoice { get; init; } = "-";
        public IReadOnlyList<BiTodoParagraphContent> Paragraphs { get; init; } = Array.Empty<BiTodoParagraphContent>();
    }

    private sealed class WordLifecycleOperationContext
    {
        public required int OperationId { get; init; }
        public required string OperationName { get; init; }
        public string DocumentPath { get; init; } = string.Empty;
        public string BookmarkName { get; init; } = string.Empty;
    }

    private sealed class WordDocumentSnapshot
    {
        public required int Index { get; init; }
        public string Name { get; init; } = string.Empty;
        public string FullName { get; init; } = string.Empty;
        public string Path { get; init; } = string.Empty;
        public bool? Saved { get; init; }
        public bool? ReadOnly { get; init; }
        public bool IsUnsaved { get; init; }
    }

    private sealed class WordProcessSnapshot
    {
        public required int ProcessId { get; init; }
        public string MainWindowTitle { get; init; } = string.Empty;
        public long MainWindowHandle { get; init; }
        public bool? Responding { get; init; }
        public DateTime? StartTimeUtc { get; init; }
    }

    private static int _wordLifecycleOperationSequence;

    public bool IsWordAvailable => Type.GetTypeFromProgID("Word.Application") is not null;

    private static bool IsWordLifecycleLoggingEnabled()
    {
        var prefs = App.UserPrefs;
        return prefs is not null && prefs.EnableDebugLogging && prefs.EnableWordLifecycleLogging;
    }

    private static WordLifecycleOperationContext BeginWordLifecycleOperation(string operationName, string? documentPath, string? bookmarkName = null)
    {
        var context = new WordLifecycleOperationContext
        {
            OperationId = Interlocked.Increment(ref _wordLifecycleOperationSequence),
            OperationName = operationName,
            DocumentPath = documentPath ?? string.Empty,
            BookmarkName = bookmarkName ?? string.Empty
        };

        if (IsWordLifecycleLoggingEnabled())
        {
            AppLogger.Debug(
                $"WordLifecycle[{context.OperationId}] Begin Operation='{operationName}', Doc='{SanitizeForLog(context.DocumentPath)}', Bookmark='{SanitizeForLog(context.BookmarkName)}'.");
            LogWordLifecycleProcessSnapshot(context, "BeforeOperation");
        }

        return context;
    }

    private static void LogWordLifecycle(WordLifecycleOperationContext? context, string message)
    {
        if (!IsWordLifecycleLoggingEnabled())
        {
            return;
        }

        if (context is null)
        {
            AppLogger.Debug($"WordLifecycle {message}");
            return;
        }

        AppLogger.Debug($"WordLifecycle[{context.OperationId}:{context.OperationName}] {message}");
    }

    private static void LogWordLifecycleAppState(
        dynamic? app,
        WordLifecycleOperationContext? context,
        string stage,
        bool? wasCreatedHere = null,
        int? initialUnsavedDocumentCount = null)
    {
        if (!IsWordLifecycleLoggingEnabled() || app is null)
        {
            return;
        }

        dynamic? docs = null;
        try
        {
            docs = app.Documents;
            var documentCount = TryReadInt(() => (int)docs.Count);
            var isVisible = TryReadBool(() => (bool)app.Visible);
            var isUserControl = TryReadBool(() => (bool)app.UserControl);
            var hwnd = TryReadInt64(() => Convert.ToInt64(app.Hwnd));
            var processId = TryGetProcessIdFromHwnd(hwnd);

            LogWordLifecycle(
                context,
                $"Stage='{stage}', WasCreatedHere={FormatOptional(wasCreatedHere)}, InitialUnsaved={FormatOptional(initialUnsavedDocumentCount)}, Visible={FormatOptional(isVisible)}, UserControl={FormatOptional(isUserControl)}, Hwnd={FormatOptional(hwnd)}, Pid={FormatOptional(processId)}, OpenDocs={FormatOptional(documentCount)}.");
        }
        catch (Exception ex)
        {
            LogWordLifecycle(context, $"Stage='{stage}', AppStateReadFailed='{SanitizeForLog(ex.GetType().Name)}: {SanitizeForLog(ex.Message)}'.");
        }
        finally
        {
            SafeReleaseCom(docs);
        }

        LogWordLifecycleProcessSnapshot(context, $"{stage}-WinWordProcesses", app);
    }

    private static void LogWordLifecycleDocumentSnapshot(dynamic? app, WordLifecycleOperationContext? context, string stage)
    {
        if (!IsWordLifecycleLoggingEnabled() || app is null)
        {
            return;
        }

        var documents = SnapshotOpenDocuments(app);
        LogWordLifecycle(context, $"Stage='{stage}', DocumentSnapshotCount={documents.Count}.");
        foreach (var document in documents)
        {
            LogWordLifecycle(
                context,
                $"Stage='{stage}', DocIndex={document.Index}, Name='{SanitizeForLog(document.Name)}', FullName='{SanitizeForLog(document.FullName)}', Path='{SanitizeForLog(document.Path)}', Saved={FormatOptional(document.Saved)}, ReadOnly={FormatOptional(document.ReadOnly)}, IsUnsaved={document.IsUnsaved}.");
        }
    }

    private static IReadOnlyList<WordDocumentSnapshot> SnapshotOpenDocuments(dynamic app)
    {
        var documents = new List<WordDocumentSnapshot>();
        dynamic? docs = null;
        try
        {
            docs = app.Documents;
            var count = (int)docs.Count;
            for (var i = 1; i <= count; i++)
            {
                dynamic? doc = null;
                try
                {
                    doc = docs[i];
                    var fullName = TryReadString(() => doc.FullName as string);
                    var path = TryReadString(() => doc.Path as string);
                    var name = TryReadString(() => doc.Name as string);
                    var saved = TryReadBool(() => (bool)doc.Saved);
                    var readOnly = TryReadBool(() => (bool)doc.ReadOnly);

                    documents.Add(new WordDocumentSnapshot
                    {
                        Index = i,
                        Name = name ?? string.Empty,
                        FullName = fullName ?? string.Empty,
                        Path = path ?? string.Empty,
                        Saved = saved,
                        ReadOnly = readOnly,
                        IsUnsaved = string.IsNullOrWhiteSpace(path)
                    });
                }
                finally
                {
                    SafeReleaseCom(doc);
                }
            }
        }
        catch
        {
            // Diagnose-Logging darf den Ablauf nicht beeinflussen.
        }
        finally
        {
            SafeReleaseCom(docs);
        }

        return documents;
    }

    private static void LogWordLifecycleFinally(
        WordLifecycleOperationContext? context,
        dynamic? app,
        bool openedHere,
        bool operationSucceeded,
        bool shouldQuitCreatedApp,
        string finalStage)
    {
        if (!IsWordLifecycleLoggingEnabled())
        {
            return;
        }

        LogWordLifecycle(
            context,
            $"Stage='{finalStage}', OperationSucceeded={operationSucceeded}, OpenedHere={openedHere}, ShouldQuitCreatedApp={shouldQuitCreatedApp}.");
        LogWordLifecycleAppState(app, context, $"{finalStage}-AppState");
        LogWordLifecycleDocumentSnapshot(app, context, $"{finalStage}-Documents");
    }

    private static void LogWordLifecycleProcessSnapshot(
        WordLifecycleOperationContext? context,
        string stage,
        dynamic? app = null)
    {
        if (!IsWordLifecycleLoggingEnabled())
        {
            return;
        }

        var attachedProcessId = TryGetWordApplicationProcessId(app);
        var processes = SnapshotRunningWordProcesses();
        LogWordLifecycle(
            context,
            $"Stage='{stage}', WinWordProcessCount={processes.Count}, AttachedPid={FormatOptional(attachedProcessId)}.");

        foreach (var process in processes)
        {
            var isAttached = attachedProcessId is not null && process.ProcessId == attachedProcessId.Value;
            var startTimeUtc = process.StartTimeUtc?.ToString("O") ?? "n/a";
            LogWordLifecycle(
                context,
                $"Stage='{stage}', WinWordPid={process.ProcessId}, Attached={isAttached}, MainWindowHandle={process.MainWindowHandle}, Responding={FormatOptional(process.Responding)}, StartTimeUtc={startTimeUtc}, MainWindowTitle='{SanitizeForLog(process.MainWindowTitle)}'.");
        }
    }

    private static IReadOnlyList<WordProcessSnapshot> SnapshotRunningWordProcesses()
    {
        var snapshots = new List<WordProcessSnapshot>();
        Process[] processes;
        try
        {
            processes = Process.GetProcessesByName("WINWORD");
        }
        catch
        {
            return snapshots;
        }

        foreach (var process in processes)
        {
            try
            {
                var processId = process.Id;
                var mainWindowTitle = TryGetProcessString(() => process.MainWindowTitle);
                var mainWindowHandle = TryGetProcessInt64(() => process.MainWindowHandle.ToInt64()) ?? 0L;
                var responding = TryGetProcessBool(() => process.Responding);
                var startTimeUtc = TryGetProcessDateTime(() => process.StartTime.ToUniversalTime());

                snapshots.Add(new WordProcessSnapshot
                {
                    ProcessId = processId,
                    MainWindowTitle = mainWindowTitle ?? string.Empty,
                    MainWindowHandle = mainWindowHandle,
                    Responding = responding,
                    StartTimeUtc = startTimeUtc
                });
            }
            catch
            {
                // Diagnose-Logging darf den Ablauf nicht beeinflussen.
            }
            finally
            {
                process.Dispose();
            }
        }

        snapshots.Sort(static (left, right) => left.ProcessId.CompareTo(right.ProcessId));
        return snapshots;
    }

    private static string SanitizeForLog(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return value
            .Replace("\r", " ", StringComparison.Ordinal)
            .Replace("\n", " ", StringComparison.Ordinal)
            .Replace("\t", " ", StringComparison.Ordinal)
            .Trim();
    }

    private static string FormatOptional<T>(T? value)
    {
        return value?.ToString() ?? "n/a";
    }

    private static string? TryReadString(Func<string?> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static bool? TryReadBool(Func<bool> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static int? TryReadInt(Func<int> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static long? TryReadInt64(Func<long> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static int? TryGetProcessIdFromHwnd(long? hwnd)
    {
        if (hwnd is null || hwnd.Value == 0)
        {
            return null;
        }

        try
        {
            NativeMethods.GetWindowThreadProcessId(new IntPtr(hwnd.Value), out var processId);
            return processId == 0 ? null : (int)processId;
        }
        catch
        {
            return null;
        }
    }

    private static int? TryGetWordApplicationProcessId(dynamic? app)
    {
        if (app is null)
        {
            return null;
        }

        var hwnd = TryReadInt64(() => Convert.ToInt64(app.Hwnd))
                   ?? TryReadComInt64Property(app, "Hwnd");
        if (hwnd is null)
        {
            dynamic? activeWindow = null;
            try
            {
                activeWindow = app.ActiveWindow;
                hwnd = TryReadInt64(() => Convert.ToInt64(activeWindow.Hwnd))
                       ?? TryReadComInt64Property(activeWindow, "Hwnd");
            }
            catch
            {
                hwnd = null;
            }
            finally
            {
                SafeReleaseCom(activeWindow);
            }
        }

        return TryGetProcessIdFromHwnd(hwnd);
    }

    private static long? TryReadComInt64Property(object? comObject, string propertyName)
    {
        if (comObject is null)
        {
            return null;
        }

        try
        {
            var value = comObject.GetType().InvokeMember(
                propertyName,
                BindingFlags.GetProperty,
                binder: null,
                target: comObject,
                args: null);

            return value is null ? null : Convert.ToInt64(value);
        }
        catch
        {
            return null;
        }
    }

    private static string? TryGetProcessString(Func<string> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static bool? TryGetProcessBool(Func<bool> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static long? TryGetProcessInt64(Func<long> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    private static DateTime? TryGetProcessDateTime(Func<DateTime> getter)
    {
        try
        {
            return getter();
        }
        catch
        {
            return null;
        }
    }

    public string FindVerlaufsakte(string folderPath, string keyword)
    {
        var matches = FindVerlaufsakteCandidates(folderPath, keyword);
        if (matches.Count != 1)
        {
            throw new InvalidOperationException(matches.Count == 0
                ? "Keine Verlaufsakte gefunden"
                : "Mehrere Verlaufsakten gefunden");
        }

        return matches[0];
    }

    public List<string> FindVerlaufsakteCandidates(string folderPath, string keyword)
    {
        if (!Directory.Exists(folderPath))
        {
            throw new DirectoryNotFoundException($"Ordner nicht gefunden: {folderPath}");
        }

        List<string> files;
        try
        {
            files = Directory
                .GetFiles(folderPath, "*.docx", SearchOption.TopDirectoryOnly)
                .Where(f => System.IO.Path.GetFileName(f).Contains(keyword, StringComparison.OrdinalIgnoreCase))
                .OrderBy(f => f, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new InvalidOperationException($"Kein Zugriff auf Ordner: {folderPath}", ex);
        }
        catch (IOException ex)
        {
            throw new InvalidOperationException($"Fehler beim Lesen von Ordner: {folderPath}", ex);
        }

        if (files.Count == 0)
        {
            throw new InvalidOperationException("Keine Verlaufsakte gefunden");
        }

        return files;
    }

    public void OpenDocumentAtBookmark(string docPath, string bookmarkName)
    {
        AppLogger.Info($"Word.OpenDocumentAtBookmark start. Doc='{docPath}', Bookmark='{bookmarkName}'");

        if (!IsWordAvailable)
        {
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }

        if (!File.Exists(docPath))
        {
            throw new FileNotFoundException("Dokumentdatei nicht gefunden", docPath);
        }

        dynamic? app = null;
        dynamic? doc = null;
        var shouldQuitCreatedApp = false;
        var openedHere = false;
        var operationSucceeded = false;
        var lifecycle = BeginWordLifecycleOperation(nameof(OpenDocumentAtBookmark), docPath, bookmarkName);

        try
        {
            var wordApp = CreateOrAttachWordApplication(lifecycle);
            app = wordApp.App;
            shouldQuitCreatedApp = wordApp.WasCreatedHere;
            LogWordLifecycleAppState(app, lifecycle, "AfterCreateOrAttach", wordApp.WasCreatedHere, wordApp.InitialUnsavedDocumentCount);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterCreateOrAttach");

            doc = OpenOrGetDocument(app, docPath, out openedHere, lifecycle);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterOpenOrGetDocument");
            CloseTransientEmptyDocuments(app, docPath, wordApp.InitialUnsavedDocumentCount, lifecycle);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterCloseTransientEmptyDocuments");
            EnsureWordUiState(app, lifecycle);
            LogWordLifecycleAppState(app, lifecycle, "AfterEnsureWordUiState");
            EnsureDocumentNotLocked(doc, openedHere);

            if (!doc.Bookmarks.Exists(bookmarkName))
            {
                throw new InvalidOperationException($"Bookmark '{bookmarkName}' nicht gefunden. Bitte Vorlage prüfen.");
            }

            FocusBookmarkAtTop(app, doc, bookmarkName);
            operationSucceeded = true;
            shouldQuitCreatedApp = false;
            AppLogger.Info($"Word.OpenDocumentAtBookmark ok. Doc='{docPath}', Bookmark='{bookmarkName}'");
        }
        catch (COMException ex) when ((uint)ex.HResult == 0x80040154)
        {
            AppLogger.Error("Word.OpenDocumentAtBookmark COM Class not registered", ex);
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }
        catch (COMException ex)
        {
            AppLogger.Error("Word.OpenDocumentAtBookmark COM Fehler", ex);
            throw new InvalidOperationException($"Fehler beim Zugriff auf Word: {ex.Message}", ex);
        }
        catch (Exception ex)
        {
            LogWordLifecycle(lifecycle, $"Exception='{SanitizeForLog(ex.GetType().Name)}', Message='{SanitizeForLog(ex.Message)}'.");
            AppLogger.Error("Word.OpenDocumentAtBookmark Fehler", ex);
            throw;
        }
        finally
        {
            LogWordLifecycleFinally(lifecycle, app, openedHere, operationSucceeded, shouldQuitCreatedApp, "Finally-BeforeCleanup");
            if (!operationSucceeded && openedHere && !shouldQuitCreatedApp)
            {
                LogWordLifecycle(lifecycle, "Finally: Dokument wird wegen Fehler wieder geschlossen.");
                TryCloseDocument(doc);
            }

            if (doc is not null && Marshal.IsComObject(doc))
            {
                Marshal.ReleaseComObject(doc);
            }

            if (shouldQuitCreatedApp)
            {
                LogWordLifecycle(lifecycle, "Finally: Selbst gestartete Word-Instanz wird wegen Fehler beendet.");
                TryQuitWordApplication(app);
            }

            if (app is not null && Marshal.IsComObject(app))
            {
                Marshal.ReleaseComObject(app);
            }

            LogWordLifecycleProcessSnapshot(lifecycle, "Finally-AfterRelease");
            LogWordLifecycle(lifecycle, "Ende der Operation, COM-Referenzen freigegeben.");
        }
    }

    public void OpenDocument(string docPath)
    {
        AppLogger.Info($"Word.OpenDocument start. Doc='{docPath}'");

        if (!IsWordAvailable)
        {
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }

        if (!File.Exists(docPath))
        {
            throw new FileNotFoundException("Dokumentdatei nicht gefunden", docPath);
        }

        dynamic? app = null;
        dynamic? doc = null;
        var shouldQuitCreatedApp = false;
        var openedHere = false;
        var operationSucceeded = false;
        var lifecycle = BeginWordLifecycleOperation(nameof(OpenDocument), docPath);

        try
        {
            var wordApp = CreateOrAttachWordApplication(lifecycle);
            app = wordApp.App;
            shouldQuitCreatedApp = wordApp.WasCreatedHere;
            LogWordLifecycleAppState(app, lifecycle, "AfterCreateOrAttach", wordApp.WasCreatedHere, wordApp.InitialUnsavedDocumentCount);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterCreateOrAttach");

            doc = OpenOrGetDocument(app, docPath, out openedHere, lifecycle);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterOpenOrGetDocument");
            CloseTransientEmptyDocuments(app, docPath, wordApp.InitialUnsavedDocumentCount, lifecycle);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterCloseTransientEmptyDocuments");
            EnsureWordUiState(app, lifecycle);
            LogWordLifecycleAppState(app, lifecycle, "AfterEnsureWordUiState");
            EnsureDocumentNotLocked(doc, openedHere);

            FocusDocument(app, doc);
            operationSucceeded = true;
            shouldQuitCreatedApp = false;
            AppLogger.Info($"Word.OpenDocument ok. Doc='{docPath}'");
        }
        catch (COMException ex) when ((uint)ex.HResult == 0x80040154)
        {
            AppLogger.Error("Word.OpenDocument COM Class not registered", ex);
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }
        catch (COMException ex)
        {
            AppLogger.Error("Word.OpenDocument COM Fehler", ex);
            throw new InvalidOperationException($"Fehler beim Zugriff auf Word: {ex.Message}", ex);
        }
        catch (Exception ex)
        {
            LogWordLifecycle(lifecycle, $"Exception='{SanitizeForLog(ex.GetType().Name)}', Message='{SanitizeForLog(ex.Message)}'.");
            AppLogger.Error("Word.OpenDocument Fehler", ex);
            throw;
        }
        finally
        {
            LogWordLifecycleFinally(lifecycle, app, openedHere, operationSucceeded, shouldQuitCreatedApp, "Finally-BeforeCleanup");
            if (!operationSucceeded && openedHere && !shouldQuitCreatedApp)
            {
                LogWordLifecycle(lifecycle, "Finally: Dokument wird wegen Fehler wieder geschlossen.");
                TryCloseDocument(doc);
            }

            if (doc is not null && Marshal.IsComObject(doc))
            {
                Marshal.ReleaseComObject(doc);
            }

            if (shouldQuitCreatedApp)
            {
                LogWordLifecycle(lifecycle, "Finally: Selbst gestartete Word-Instanz wird wegen Fehler beendet.");
                TryQuitWordApplication(app);
            }

            if (app is not null && Marshal.IsComObject(app))
            {
                Marshal.ReleaseComObject(app);
            }

            LogWordLifecycleProcessSnapshot(lifecycle, "Finally-AfterRelease");
            LogWordLifecycle(lifecycle, "Ende der Operation, COM-Referenzen freigegeben.");
        }
    }

    public bool InsertClipboardToTable(string docPath, string bookmarkName)
    {
        return InsertClipboardToTable(docPath, bookmarkName, 2, null, null);
    }

    public bool InsertClipboardToTable(
        string docPath,
        string bookmarkName,
        int firstDataRowIndex,
        string[]? fallbackFieldsWhenClipboardInvalid = null,
        string? preReadClipboardText = null)
    {
        AppLogger.Info($"Word.InsertClipboardToTable start. Doc='{docPath}', Bookmark='{bookmarkName}'");

        if (!IsWordAvailable)
        {
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }

        if (firstDataRowIndex < 2)
        {
            throw new ArgumentOutOfRangeException(nameof(firstDataRowIndex), "firstDataRowIndex muss >= 2 sein.");
        }

        if (!File.Exists(docPath))
        {
            throw new FileNotFoundException("Dokumentdatei nicht gefunden", docPath);
        }

        dynamic? app = null;
        dynamic? doc = null;
        var shouldQuitCreatedApp = false;
        var openedHere = false;
        var operationSucceeded = false;
        var lifecycle = BeginWordLifecycleOperation(nameof(InsertClipboardToTable), docPath, bookmarkName);

        try
        {
            var wordApp = CreateOrAttachWordApplication(lifecycle);
            app = wordApp.App;
            shouldQuitCreatedApp = wordApp.WasCreatedHere;
            LogWordLifecycleAppState(app, lifecycle, "AfterCreateOrAttach", wordApp.WasCreatedHere, wordApp.InitialUnsavedDocumentCount);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterCreateOrAttach");

            doc = OpenOrGetDocument(app, docPath, out openedHere, lifecycle);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterOpenOrGetDocument");
            CloseTransientEmptyDocuments(app, docPath, wordApp.InitialUnsavedDocumentCount, lifecycle);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterCloseTransientEmptyDocuments");
            EnsureWordUiState(app, lifecycle);
            LogWordLifecycleAppState(app, lifecycle, "AfterEnsureWordUiState");
            EnsureDocumentNotLocked(doc, openedHere);

            dynamic targetTable = ResolveStructuredEntryTableForWrite(
                doc,
                bookmarkName,
                4,
                GetStructuredEntryTableDisplayName(bookmarkName));

            var clipboardText = preReadClipboardText ?? GetClipboardTextWithRetry();
            var hasClipboardContent = !string.IsNullOrWhiteSpace(clipboardText);
            string[] clipboardFields = Array.Empty<string>();
            var hasValidClipboardRow = hasClipboardContent && TryParseClipboardFields(clipboardText, out clipboardFields);
            if (hasClipboardContent && !hasValidClipboardRow)
            {
                AppLogger.Warn("Word.InsertClipboardToTable: Clipboard-Format ungueltig, leere Zeile wird vorbereitet.");
            }

            var hasFallbackFields = fallbackFieldsWhenClipboardInvalid is not null &&
                                    fallbackFieldsWhenClipboardInvalid.Length == 4;
            if (fallbackFieldsWhenClipboardInvalid is not null && !hasFallbackFields)
            {
                AppLogger.Warn("Word.InsertClipboardToTable: fallbackFieldsWhenClipboardInvalid hat nicht 4 Spalten und wird ignoriert.");
            }

            dynamic? insertedRow = null;
            try
            {
                var existingRowCount = (int)targetTable.Rows.Count;
                if (existingRowCount < firstDataRowIndex)
                {
                    // Tabelle auf die gewuenschte Startzeile erweitern (z. B. BI: Start erst ab Zeile 3).
                    while ((int)targetTable.Rows.Count < firstDataRowIndex)
                    {
                        targetTable.Rows.Add();
                    }

                    insertedRow = targetTable.Rows[firstDataRowIndex];
                }
                else
                {
                    // Neue Eintraege oben im Datenbereich einfuegen (an erster Datenzeile).
                    insertedRow = targetTable.Rows.Add(targetTable.Rows[firstDataRowIndex]);
                }

                var rowIndex = (int)insertedRow.Index;

                if (hasValidClipboardRow)
                {
                    for (var i = 1; i <= 4; i++)
                    {
                        dynamic? cell = null;
                        dynamic? range = null;
                        try
                        {
                            cell = targetTable.Cell(rowIndex, i);
                            range = cell.Range;
                            range.Text = clipboardFields[i - 1];
                        }
                        finally
                        {
                            SafeReleaseCom(range, cell);
                        }
                    }
                }
                else if (hasFallbackFields)
                {
                    for (var i = 1; i <= 4; i++)
                    {
                        dynamic? cell = null;
                        dynamic? range = null;
                        try
                        {
                            cell = targetTable.Cell(rowIndex, i);
                            range = cell.Range;
                            range.Text = fallbackFieldsWhenClipboardInvalid![i - 1] ?? string.Empty;
                        }
                        finally
                        {
                            SafeReleaseCom(range, cell);
                        }
                    }
                }
            }
            catch
            {
                TryDeleteRow(insertedRow);
                throw;
            }

            var preferredEditColumn = hasValidClipboardRow ? 1 : 3;
            var editColumn = GetSafeEditColumn(targetTable, preferredEditColumn);
            dynamic? editCell = null;
            dynamic? editRange = null;
            try
            {
                editCell = targetTable.Cell((int)insertedRow!.Index, editColumn);
                editRange = editCell.Range;
                FocusRangeAtTop(app, editRange);
            }
            finally
            {
                SafeReleaseCom(editRange, editCell);
            }

            operationSucceeded = true;
            shouldQuitCreatedApp = false;
            AppLogger.Info($"Word.InsertClipboardToTable ok. Doc='{docPath}', ClipboardUsed={hasValidClipboardRow}, FocusColumn={editColumn}");
            return hasValidClipboardRow;
        }
        catch (COMException ex) when ((uint)ex.HResult == 0x80040154)
        {
            AppLogger.Error("Word.InsertClipboardToTable COM Class not registered", ex);
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }
        catch (COMException ex)
        {
            AppLogger.Error("Word.InsertClipboardToTable COM Fehler", ex);
            throw new InvalidOperationException($"Fehler beim Zugriff auf Word: {ex.Message}", ex);
        }
        catch (Exception ex)
        {
            LogWordLifecycle(lifecycle, $"Exception='{SanitizeForLog(ex.GetType().Name)}', Message='{SanitizeForLog(ex.Message)}'.");
            AppLogger.Error("Word.InsertClipboardToTable Fehler", ex);
            throw;
        }
        finally
        {
            LogWordLifecycleFinally(lifecycle, app, openedHere, operationSucceeded, shouldQuitCreatedApp, "Finally-BeforeCleanup");
            if (!operationSucceeded && openedHere && !shouldQuitCreatedApp)
            {
                LogWordLifecycle(lifecycle, "Finally: Dokument wird wegen Fehler wieder geschlossen.");
                TryCloseDocument(doc);
            }

            if (doc is not null && Marshal.IsComObject(doc))
            {
                Marshal.ReleaseComObject(doc);
            }

            if (shouldQuitCreatedApp)
            {
                LogWordLifecycle(lifecycle, "Finally: Selbst gestartete Word-Instanz wird wegen Fehler beendet.");
                TryQuitWordApplication(app);
            }

            if (app is not null && Marshal.IsComObject(app))
            {
                Marshal.ReleaseComObject(app);
            }

            LogWordLifecycleProcessSnapshot(lifecycle, "Finally-AfterRelease");
            LogWordLifecycle(lifecycle, "Ende der Operation, COM-Referenzen freigegeben.");
        }
    }

    public string ReadClipboardTextWithRetry()
    {
        return GetClipboardTextWithRetry();
    }

    public void InsertTextRowToTable(string docPath, string bookmarkName, string tabSeparatedRow)
    {
        AppLogger.Info($"Word.InsertTextRowToTable start. Doc='{docPath}', Bookmark='{bookmarkName}'");

        if (!IsWordAvailable)
        {
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }

        if (!File.Exists(docPath))
        {
            throw new FileNotFoundException("Dokumentdatei nicht gefunden", docPath);
        }

        if (!TryParseSingleTabSeparatedRow(tabSeparatedRow, out var fields))
        {
            throw new InvalidOperationException("Ungültiges Zeilenformat. Erwartet wird genau 1 Zeile mit 4 tab-getrennten Spalten.");
        }

        dynamic? app = null;
        dynamic? doc = null;
        var shouldQuitCreatedApp = false;
        var openedHere = false;
        var operationSucceeded = false;
        var lifecycle = BeginWordLifecycleOperation(nameof(InsertTextRowToTable), docPath, bookmarkName);

        try
        {
            var wordApp = CreateOrAttachWordApplication(lifecycle);
            app = wordApp.App;
            shouldQuitCreatedApp = wordApp.WasCreatedHere;
            LogWordLifecycleAppState(app, lifecycle, "AfterCreateOrAttach", wordApp.WasCreatedHere, wordApp.InitialUnsavedDocumentCount);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterCreateOrAttach");

            doc = OpenOrGetDocument(app, docPath, out openedHere, lifecycle);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterOpenOrGetDocument");
            CloseTransientEmptyDocuments(app, docPath, wordApp.InitialUnsavedDocumentCount, lifecycle);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterCloseTransientEmptyDocuments");
            EnsureWordUiState(app, lifecycle);
            LogWordLifecycleAppState(app, lifecycle, "AfterEnsureWordUiState");
            EnsureDocumentNotLocked(doc, openedHere);

            dynamic targetTable = ResolveStructuredEntryTableForWrite(
                doc,
                bookmarkName,
                4,
                GetStructuredEntryTableDisplayName(bookmarkName));
            dynamic? insertedRow = null;
            try
            {
                insertedRow = InsertRowAtTopOfDataArea(targetTable);
                var rowIndex = (int)insertedRow.Index;

                for (var i = 1; i <= 4; i++)
                {
                    dynamic? cell = null;
                    dynamic? range = null;
                    try
                    {
                        cell = targetTable.Cell(rowIndex, i);
                        range = cell.Range;
                        range.Text = fields[i - 1];
                    }
                    finally
                    {
                        SafeReleaseCom(range, cell);
                    }
                }
            }
            catch
            {
                TryDeleteRow(insertedRow);
                throw;
            }

            dynamic? editCell = null;
            dynamic? editRange = null;
            try
            {
                editCell = targetTable.Cell((int)insertedRow.Index, 1);
                editRange = editCell.Range;
                FocusRangeAtTop(app, editRange);
            }
            finally
            {
                SafeReleaseCom(editRange, editCell);
            }
            operationSucceeded = true;
            shouldQuitCreatedApp = false;
            AppLogger.Info($"Word.InsertTextRowToTable ok. Doc='{docPath}'");
        }
        catch (COMException ex) when ((uint)ex.HResult == 0x80040154)
        {
            AppLogger.Error("Word.InsertTextRowToTable COM Class not registered", ex);
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }
        catch (COMException ex)
        {
            AppLogger.Error("Word.InsertTextRowToTable COM Fehler", ex);
            throw new InvalidOperationException($"Fehler beim Zugriff auf Word: {ex.Message}", ex);
        }
        catch (Exception ex)
        {
            LogWordLifecycle(lifecycle, $"Exception='{SanitizeForLog(ex.GetType().Name)}', Message='{SanitizeForLog(ex.Message)}'.");
            AppLogger.Error("Word.InsertTextRowToTable Fehler", ex);
            throw;
        }
        finally
        {
            LogWordLifecycleFinally(lifecycle, app, openedHere, operationSucceeded, shouldQuitCreatedApp, "Finally-BeforeCleanup");
            if (!operationSucceeded && openedHere && !shouldQuitCreatedApp)
            {
                LogWordLifecycle(lifecycle, "Finally: Dokument wird wegen Fehler wieder geschlossen.");
                TryCloseDocument(doc);
            }

            if (doc is not null && Marshal.IsComObject(doc))
            {
                Marshal.ReleaseComObject(doc);
            }

            if (shouldQuitCreatedApp)
            {
                LogWordLifecycle(lifecycle, "Finally: Selbst gestartete Word-Instanz wird wegen Fehler beendet.");
                TryQuitWordApplication(app);
            }

            if (app is not null && Marshal.IsComObject(app))
            {
                Marshal.ReleaseComObject(app);
            }

            LogWordLifecycleProcessSnapshot(lifecycle, "Finally-AfterRelease");
            LogWordLifecycle(lifecycle, "Ende der Operation, COM-Referenzen freigegeben.");
        }
    }

    public BiTodoCollectSummary CollectBiTodoDocument(
        IReadOnlyList<BiTodoCollectRequest> requests,
        string bookmarkName,
        string documentTitle)
    {
        AppLogger.Info($"Word.CollectBiTodoDocument start. Count={requests.Count}, Bookmark='{bookmarkName}', Title='{documentTitle}'");

        if (!IsWordAvailable)
        {
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }

        var summary = new BiTodoCollectSummary();
        if (requests.Count == 0)
        {
            return summary;
        }

        dynamic? app = null;
        dynamic? templateDoc = null;
        dynamic? resultDoc = null;
        BiTodoTemplateDefinition? template = null;
        string? templatePath = null;
        var keepResultDocumentOpen = false;
        var hasWrittenAnyBlock = false;
        var lifecycle = BeginWordLifecycleOperation(nameof(CollectBiTodoDocument), documentTitle, bookmarkName);

        try
        {
            var setupStopwatch = Stopwatch.StartNew();
            app = CreateDedicatedHiddenWordApplication(lifecycle);
            LogWordLifecycleAppState(app, lifecycle, "AfterCreateDedicatedHiddenWordApplication", wasCreatedHere: true);
            templatePath = ExtractEmbeddedBiTodoTemplateToTempFile();
            templateDoc = OpenReadOnlyHiddenDocument(app, templatePath, lifecycle);
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterOpenReadOnlyHiddenTemplate");
            template = ResolveBiTodoTemplate(templateDoc);
            resultDoc = app.Documents.Add();
            LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterCreateResultDocument");
            AppendBiTodoDocumentTitle(resultDoc, template, documentTitle);
            setupStopwatch.Stop();
            LogBiTodoTiming("Setup", setupStopwatch.ElapsedMilliseconds);

            foreach (var request in requests)
            {
                var participantStopwatch = Stopwatch.StartNew();

                try
                {
                    if (!string.IsNullOrWhiteSpace(request.FailureMessage))
                    {
                        AppendBiTodoParticipantFailureBlock(resultDoc, template, request, request.FailureMessage, hasWrittenAnyBlock);
                        hasWrittenAnyBlock = true;

                        summary.Results.Add(new BiTodoCollectResult
                        {
                            Name = request.FullName,
                            Initials = request.Initials,
                            IsSuccess = false,
                            Message = request.FailureMessage
                        });
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(request.DocumentPath) || !File.Exists(request.DocumentPath))
                    {
                        AppendBiTodoParticipantFailureBlock(resultDoc, template, request, "keine Akte", hasWrittenAnyBlock);
                        hasWrittenAnyBlock = true;

                        summary.Results.Add(new BiTodoCollectResult
                        {
                            Name = request.FullName,
                            Initials = request.Initials,
                            IsSuccess = false,
                            Message = "keine Akte"
                        });
                        continue;
                    }

                    var extractStopwatch = Stopwatch.StartNew();
                    var extracted = BiDocxExtractionService.Extract(request.DocumentPath, bookmarkName, request.FullName);
                    extractStopwatch.Stop();
                    LogBiTodoParticipantTiming(request.FullName, "ExtractBiTodoParticipantContent", extractStopwatch.ElapsedMilliseconds);

                    var content = CreateBiTodoParticipantContent(request, extracted);

                    var appendStopwatch = Stopwatch.StartNew();
                    AppendBiTodoParticipantBlock(resultDoc, template, content, hasWrittenAnyBlock);
                    appendStopwatch.Stop();
                    LogBiTodoParticipantTiming(request.FullName, "AppendBiTodoParticipantBlock", appendStopwatch.ElapsedMilliseconds);

                    hasWrittenAnyBlock = true;

                    summary.Results.Add(new BiTodoCollectResult
                    {
                        Name = request.FullName,
                        Initials = request.Initials,
                        IsSuccess = true,
                        Message = "BI-Inhalte übernommen"
                    });
                }
                catch (WordTemplateValidationException ex) when (
                    ex.Kind == WordTemplateValidationErrorKind.BookmarkMissing ||
                    ex.Kind == WordTemplateValidationErrorKind.BiTodoTableInvalid ||
                    ex.Kind == WordTemplateValidationErrorKind.BiTodoContentInvalid)
                {
                    AppendBiTodoParticipantFailureBlock(resultDoc, template, request, ex.UserMessage, hasWrittenAnyBlock);
                    hasWrittenAnyBlock = true;

                    summary.Results.Add(new BiTodoCollectResult
                    {
                        Name = request.FullName,
                        Initials = request.Initials,
                        IsSuccess = false,
                        Message = ex.UserMessage
                    });
                }
                catch (FileNotFoundException)
                {
                    AppendBiTodoParticipantFailureBlock(resultDoc, template, request, "keine Akte", hasWrittenAnyBlock);
                    hasWrittenAnyBlock = true;

                    summary.Results.Add(new BiTodoCollectResult
                    {
                        Name = request.FullName,
                        Initials = request.Initials,
                        IsSuccess = false,
                        Message = "keine Akte"
                    });
                }
                catch (COMException ex)
                {
                    AppLogger.Warn($"Word.CollectBiTodoDocument COM-Hinweis für '{request.FullName}': {ex.Message}");
                    AppendBiTodoParticipantFailureBlock(resultDoc, template, request, "Akte gesperrt oder bereits geöffnet", hasWrittenAnyBlock);
                    hasWrittenAnyBlock = true;

                    summary.Results.Add(new BiTodoCollectResult
                    {
                        Name = request.FullName,
                        Initials = request.Initials,
                        IsSuccess = false,
                        Message = "BI-Inhalte konnten nicht gelesen werden"
                    });
                }
                catch (Exception ex)
                {
                    AppLogger.Warn($"Word.CollectBiTodoDocument Fehler für '{request.FullName}': {ex.Message}");
                    var userMessage = GetBiTodoUserMessage(ex);
                    AppendBiTodoParticipantFailureBlock(resultDoc, template, request, userMessage, hasWrittenAnyBlock);
                    hasWrittenAnyBlock = true;

                    summary.Results.Add(new BiTodoCollectResult
                    {
                        Name = request.FullName,
                        Initials = request.Initials,
                        IsSuccess = false,
                        Message = userMessage
                    });
                }
                finally
                {
                    participantStopwatch.Stop();
                    LogBiTodoParticipantTiming(request.FullName, "Total", participantStopwatch.ElapsedMilliseconds);
                }
            }

            if (hasWrittenAnyBlock)
            {
                EnsureWordUiState(app, lifecycle);
                LogWordLifecycleAppState(app, lifecycle, "AfterEnsureWordUiState");
                LogWordLifecycleDocumentSnapshot(app, lifecycle, "AfterEnsureWordUiState");
                FocusDocument(app, resultDoc);
                keepResultDocumentOpen = true;
                summary.DocumentOpened = true;
            }

            AppLogger.Info($"Word.CollectBiTodoDocument ok. Success={summary.SuccessCount}, Failed={summary.FailureCount}, Opened={summary.DocumentOpened}");
            return summary;
        }
        catch (COMException ex) when ((uint)ex.HResult == 0x80040154)
        {
            AppLogger.Error("Word.CollectBiTodoDocument COM Class not registered", ex);
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }
        catch (COMException ex)
        {
            AppLogger.Error("Word.CollectBiTodoDocument COM Fehler", ex);
            throw new InvalidOperationException($"Fehler beim Zugriff auf Word: {ex.Message}", ex);
        }
        catch (Exception ex)
        {
            LogWordLifecycle(lifecycle, $"Exception='{SanitizeForLog(ex.GetType().Name)}', Message='{SanitizeForLog(ex.Message)}'.");
            AppLogger.Error("Word.CollectBiTodoDocument Fehler", ex);
            throw;
        }
        finally
        {
            LogWordLifecycleFinally(lifecycle, app, openedHere: false, operationSucceeded: keepResultDocumentOpen || hasWrittenAnyBlock, shouldQuitCreatedApp: !keepResultDocumentOpen, finalStage: "Finally-BeforeCleanup");
            if (templateDoc is not null)
            {
                TryCloseDocumentSilently(templateDoc, "Word: BI-Vorlage geschlossen.");
            }

            if (!keepResultDocumentOpen)
            {
                LogWordLifecycle(lifecycle, "Finally: BI-Ergebnisdokument wird geschlossen und dedizierte Instanz beendet.");
                TryCloseDocumentSilently(resultDoc, "Word: Leeres BI-To-do-Ergebnisdokument geschlossen.");
                TryQuitWordApplication(app);
            }

            if (template is not null)
            {
                SafeReleaseCom(
                    template.TitleParagraph,
                    template.HeaderParagraph,
                    template.CareerChoiceParagraph,
                    template.BulletParagraph,
                    template.SeparatorParagraph,
                    template.BlankParagraph);
            }

            SafeReleaseCom(templateDoc, resultDoc, app);
            DeleteTempFileQuietly(templatePath);
            LogWordLifecycleProcessSnapshot(lifecycle, "Finally-AfterRelease");
            LogWordLifecycle(lifecycle, "Ende der Operation, COM-Referenzen freigegeben.");
        }
    }

    private static WordApplicationHandle CreateOrAttachWordApplication(WordLifecycleOperationContext? context = null)
    {
        try
        {
            var clsid = new Guid("000209FF-0000-0000-C000-000000000046");
            NativeMethods.GetActiveObject(ref clsid, IntPtr.Zero, out var runningApp);
            if (runningApp is not null)
            {
                try
                {
                    var unsavedDocumentCount = CountUnsavedDocuments(runningApp);
                    LogWordLifecycle(context, $"CreateOrAttach: Bestehende Word-Instanz gefunden. InitialUnsaved={unsavedDocumentCount}.");
                    AppLogger.Info($"Word: An bestehende Instanz angehaengt. Ungespeicherte Dokumente davor={unsavedDocumentCount}.");
                    return new WordApplicationHandle(runningApp, false, unsavedDocumentCount);
                }
                catch
                {
                    if (Marshal.IsComObject(runningApp))
                    {
                        Marshal.ReleaseComObject(runningApp);
                    }

                    throw;
                }
            }
        }
        catch (COMException ex) when ((uint)ex.HResult == 0x800401E3)
        {
            // Keine laufende Word-Instanz; neue Instanz wird erstellt.
            LogWordLifecycle(context, "CreateOrAttach: Keine laufende Word-Instanz in ROT gefunden, neue Instanz wird erstellt.");
        }

        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType is null)
        {
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }

        dynamic app = Activator.CreateInstance(wordType)!;
        if (app is null)
        {
            throw new InvalidOperationException("Microsoft Word konnte nicht gestartet werden");
        }

        var createdUnsavedDocumentCount = CountUnsavedDocuments(app);
        LogWordLifecycle(context, $"CreateOrAttach: Neue Word-Instanz erstellt. InitialUnsavedObserved={createdUnsavedDocumentCount}.");
        AppLogger.Info($"Word: Neue Instanz gestartet. Ungespeicherte Dokumente initial={createdUnsavedDocumentCount}.");
        return new WordApplicationHandle(app, true, 0);
    }

    private static dynamic OpenOrGetDocument(dynamic app, string docPath, out bool openedHere, WordLifecycleOperationContext? context = null)
    {
        var targetPath = Path.GetFullPath(docPath);
        dynamic? docs = null;
        try
        {
            docs = app.Documents;
            var documentCount = (int)docs.Count;
            AppLogger.Debug($"Word.OpenOrGetDocument: Ziel='{targetPath}', OpenDocs={documentCount}.");

            for (var i = 1; i <= documentCount; i++)
            {
                dynamic? openDoc = null;
                try
                {
                    openDoc = docs[i];
                    var openPath = string.Empty;
                    try
                    {
                        openPath = (string)openDoc.FullName;
                    }
                    catch
                    {
                        // Ignorieren und mit naechstem Dokument weitermachen.
                    }

                    if (!string.IsNullOrWhiteSpace(openPath) &&
                        Path.GetFullPath(openPath).Equals(targetPath, StringComparison.OrdinalIgnoreCase))
                    {
                        openedHere = false;
                        LogWordLifecycle(context, $"OpenOrGetDocument: Bereits offenes Dokument wiederverwendet. ExistingPath='{SanitizeForLog(openPath)}'.");
                        AppLogger.Debug($"Word.OpenOrGetDocument: Bereits offenes Dokument wiederverwendet. Path='{openPath}'.");
                        var matchedDoc = openDoc;
                        openDoc = null;
                        return matchedDoc!;
                    }
                }
                finally
                {
                    SafeReleaseCom(openDoc);
                }
            }

            if (IsFileLocked(docPath))
            {
                throw new InvalidOperationException(BuildDocumentLockedMessage());
            }

            try
            {
                openedHere = true;
                LogWordLifecycle(context, $"OpenOrGetDocument: Dokument wird neu geoeffnet. TargetPath='{SanitizeForLog(docPath)}'.");
                AppLogger.Debug($"Word.OpenOrGetDocument: Dokument wird neu geoeffnet. Path='{docPath}'.");
                return docs.Open(docPath, ReadOnly: false, AddToRecentFiles: false);
            }
            catch (COMException ex) when (IsLockRelatedHResult((uint)ex.HResult))
            {
                openedHere = false;
                LogWordLifecycle(context, $"OpenOrGetDocument: Lock beim Oeffnen. HResult=0x{(uint)ex.HResult:X8}.");
                AppLogger.Warn($"Word.OpenOrGetDocument: Dokument beim Oeffnen gesperrt (HRESULT=0x{(uint)ex.HResult:X8}). Path='{docPath}'.");
                throw new InvalidOperationException(BuildDocumentLockedMessage(), ex);
            }
        }
        finally
        {
            SafeReleaseCom(docs);
        }
    }

    private static dynamic CreateDedicatedHiddenWordApplication(WordLifecycleOperationContext? context = null)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType is null)
        {
            throw new InvalidOperationException("Microsoft Word wurde nicht gefunden");
        }

        dynamic app = Activator.CreateInstance(wordType)!;
        if (app is null)
        {
            throw new InvalidOperationException("Microsoft Word konnte nicht gestartet werden");
        }

        try
        {
            app.Visible = false;
        }
        catch
        {
        }

        try
        {
            app.DisplayAlerts = 0;
        }
        catch
        {
        }

        LogWordLifecycle(context, "CreateDedicatedHiddenWordApplication: Dedizierte versteckte BI-Instanz erstellt.");
        AppLogger.Info("Word: Dedizierte versteckte Instanz für BI: To-dos gestartet.");
        return app;
    }

    private static dynamic OpenReadOnlyHiddenDocument(dynamic app, string docPath, WordLifecycleOperationContext? context = null)
    {
        dynamic? docs = null;
        try
        {
            docs = app.Documents;
            LogWordLifecycle(context, $"OpenReadOnlyHiddenDocument: Oeffne versteckte Vorlage. Path='{SanitizeForLog(docPath)}'.");
            return docs.Open(docPath, ReadOnly: true, AddToRecentFiles: false, Visible: false);
        }
        finally
        {
            SafeReleaseCom(docs);
        }
    }

    private static string ExtractEmbeddedBiTodoTemplateToTempFile()
    {
        var assembly = typeof(WordService).Assembly;
        using var resourceStream = assembly.GetManifestResourceStream(BiTodoTemplateResourceName);
        if (resourceStream is null)
        {
            throw new InvalidOperationException("BI-Vorlage ist nicht in der Anwendung eingebettet. Bitte Build prüfen.");
        }

        return WriteBiTodoTemplateToTempFile(resourceStream);
    }

    private static string WriteBiTodoTemplateToTempFile(Stream resourceStream)
    {
        if (resourceStream is null)
        {
            throw new InvalidOperationException("BI-Vorlage konnte nicht aus der Anwendung geladen werden.");
        }

        var targetDirectory = Path.Combine(Path.GetTempPath(), "Scola", "WordTemplates");
        Directory.CreateDirectory(targetDirectory);

        var tempPath = Path.Combine(targetDirectory, $"bi-template-{Guid.NewGuid():N}.docx");
        using var fileStream = File.Create(tempPath);
        resourceStream.CopyTo(fileStream);
        return tempPath;
    }

    private static BiTodoTemplateDefinition ResolveBiTodoTemplate(dynamic templateDoc)
    {
        dynamic? titleParagraph = null;
        dynamic? nameParagraph = null;
        dynamic? careerChoiceParagraph = null;
        dynamic? bulletParagraph = null;
        dynamic? separatorParagraph = null;
        dynamic? blankParagraph = null;

        try
        {
            titleParagraph = FindTemplateParagraph(
                (object)templateDoc,
                text => text.StartsWith("BI,", StringComparison.OrdinalIgnoreCase));
            nameParagraph = FindTemplateParagraph(
                (object)templateDoc,
                text => text.Contains("Kürzel", StringComparison.OrdinalIgnoreCase) &&
                        text.Contains("Name TN", StringComparison.OrdinalIgnoreCase));
            careerChoiceParagraph = FindTemplateParagraph(
                (object)templateDoc,
                text => text.StartsWith("Berufswahl:", StringComparison.OrdinalIgnoreCase));
            bulletParagraph = FindFirstListParagraph((object)templateDoc);
            separatorParagraph = FindTemplateParagraph((object)templateDoc, IsSeparatorParagraphText);
            blankParagraph = FindTemplateParagraph((object)templateDoc, text => string.IsNullOrWhiteSpace(text), required: false);

            if (titleParagraph is null || nameParagraph is null || careerChoiceParagraph is null || bulletParagraph is null || separatorParagraph is null)
            {
                throw new InvalidOperationException("Bi Vorlage 2.docx ist strukturell nicht brauchbar. Bitte Vorlage prüfen.");
            }

            return new BiTodoTemplateDefinition
            {
                TitleParagraph = titleParagraph,
                HeaderParagraph = nameParagraph,
                CareerChoiceParagraph = careerChoiceParagraph,
                BulletParagraph = bulletParagraph,
                SeparatorParagraph = separatorParagraph,
                BlankParagraph = blankParagraph
            };
        }
        catch
        {
            SafeReleaseCom(titleParagraph, nameParagraph, careerChoiceParagraph, bulletParagraph, separatorParagraph, blankParagraph);
            throw;
        }
    }

    private static void AppendBiTodoDocumentTitle(
        dynamic doc,
        BiTodoTemplateDefinition template,
        string title)
    {
        dynamic? insertedRange = null;
        try
        {
            insertedRange = CloneTemplateParagraph(doc, template.TitleParagraph);
            if (!ReplaceTextInInsertedRange(insertedRange, "[DATUM]", title) &&
                !ReplaceTextInInsertedRange(insertedRange, "Wochentag, Datum", title.Substring("BI, ".Length)))
            {
                insertedRange.Text = title + "\r";
            }
        }
        finally
        {
            SafeReleaseCom(insertedRange);
        }

        AppendOptionalBlankParagraph(doc, template);
    }

    private static BiTodoParticipantContent ExtractBiTodoParticipantContent(
        dynamic sourceDoc,
        dynamic sourceTable,
        BiTodoCollectRequest request,
        string bookmarkName)
    {
        var extractStopwatch = Stopwatch.StartNew();
        var careerChoiceStopwatch = Stopwatch.StartNew();
        var careerChoice = ReadBiCareerChoice(sourceDoc, sourceTable, bookmarkName);
        careerChoiceStopwatch.Stop();
        LogBiTodoParticipantTiming(request.FullName, "ReadBiCareerChoice", careerChoiceStopwatch.ElapsedMilliseconds);

        var paragraphReadStopwatch = Stopwatch.StartNew();
        var paragraphs = ReadBiTodoParagraphs(sourceTable, bookmarkName);
        paragraphReadStopwatch.Stop();
        LogBiTodoParticipantTiming(request.FullName, "ReadBiTodoParagraphs", paragraphReadStopwatch.ElapsedMilliseconds);
        extractStopwatch.Stop();
        LogBiTodoParticipantTiming(request.FullName, "ExtractBiTodoParticipantContent", extractStopwatch.ElapsedMilliseconds);

        return new BiTodoParticipantContent
        {
            Initials = request.Initials,
            FullName = request.FullName,
            CareerChoice = careerChoice,
            Paragraphs = paragraphs
        };
    }

    private static BiTodoParticipantContent CreateBiTodoParticipantContent(
        BiTodoCollectRequest request,
        BiDocxExtractionResult extracted)
    {
        return new BiTodoParticipantContent
        {
            Initials = request.Initials,
            FullName = request.FullName,
            CareerChoice = extracted.CareerChoice,
            Paragraphs = extracted.Paragraphs
                .Select(paragraph => new BiTodoParagraphContent
                {
                    Text = paragraph.Text,
                    IsBullet = paragraph.IsBullet
                })
                .ToList()
        };
    }

    private static void AppendBiTodoParticipantBlock(
        dynamic resultDoc,
        BiTodoTemplateDefinition template,
        BiTodoParticipantContent content,
        bool addSeparator)
    {
        if (addSeparator)
        {
            AppendBiTodoSeparatorSection(resultDoc, template);
        }

        AppendBiTodoHeaderSection(resultDoc, template, content.Initials, content.FullName, content.CareerChoice);

        foreach (var paragraph in content.Paragraphs)
        {
            if (paragraph.IsBullet)
            {
                AppendBiTodoBulletParagraph(resultDoc, template.BulletParagraph, paragraph.Text);
            }
            else
            {
                AppendBiTodoNormalParagraph(resultDoc, template.CareerChoiceParagraph, paragraph.Text);
            }
        }

        AppendOptionalBlankParagraph(resultDoc, template);
    }

    private static void AppendBiTodoParticipantFailureBlock(
        dynamic resultDoc,
        BiTodoTemplateDefinition template,
        BiTodoCollectRequest request,
        string failureMessage,
        bool addSeparator)
    {
        if (addSeparator)
        {
            AppendBiTodoSeparatorSection(resultDoc, template);
        }

        AppendBiTodoHeaderSection(resultDoc, template, request.Initials, request.FullName, "-");
        AppendSimpleStyledParagraph(resultDoc, failureMessage, 10, true, false, 0xB00020);
        AppendOptionalBlankParagraph(resultDoc, template);
    }

    private static void AppendBiTodoSeparatorSection(dynamic doc, BiTodoTemplateDefinition template)
    {
        AppendClonedTemplateParagraph(doc, template.SeparatorParagraph);
        AppendOptionalBlankParagraph(doc, template);
    }

    private static void AppendOptionalBlankParagraph(dynamic doc, BiTodoTemplateDefinition template)
    {
        if (template.BlankParagraph is not null)
        {
            AppendClonedTemplateParagraph(doc, template.BlankParagraph);
        }
    }

    private static void AppendBiTodoHeaderSection(
        dynamic doc,
        BiTodoTemplateDefinition template,
        string initials,
        string fullName,
        string careerChoice)
    {
        dynamic? insertedRange = null;
        dynamic? paragraphs = null;
        dynamic? careerParagraph = null;
        dynamic? careerRange = null;
        try
        {
            insertedRange = CloneTemplateParagraphBlock(doc, template.HeaderParagraph, template.CareerChoiceParagraph);
            if (!ReplaceTextInInsertedRange(insertedRange, "Kürzel", initials))
            {
                insertedRange.Text = $"{initials}\v{fullName}\rBerufswahl: -\r";
            }

            ReplaceTextInInsertedRange(insertedRange, "Name TN", fullName);

            paragraphs = insertedRange.Paragraphs;
            if ((int)paragraphs.Count >= 2)
            {
                careerParagraph = paragraphs[2];
                careerRange = careerParagraph.Range;
                ApplyBiTodoCareerChoiceText(careerRange, careerChoice);
            }
        }
        finally
        {
            SafeReleaseCom(careerRange, careerParagraph, paragraphs, insertedRange);
        }
    }

    private static void AppendBiTodoBulletParagraph(dynamic doc, dynamic templateParagraph, string bulletText)
    {
        dynamic? insertedRange = null;
        dynamic? font = null;
        try
        {
            insertedRange = CloneTemplateParagraph(doc, templateParagraph);
            if (!ReplaceTextInInsertedRange(insertedRange, "Bullets", bulletText))
            {
                insertedRange.Text = bulletText + "\r";
            }

            font = insertedRange.Font;
            font.Bold = 0;
            font.Italic = 0;
            font.Color = 0;
        }
        finally
        {
            SafeReleaseCom(font, insertedRange);
        }
    }

    private static void AppendBiTodoCareerChoiceParagraph(
        dynamic doc,
        dynamic templateParagraph,
        string careerChoice)
    {
        dynamic? insertedRange = null;
        dynamic? labelRange = null;
        dynamic? valueRange = null;
        dynamic? labelFont = null;
        dynamic? valueFont = null;
        try
        {
            insertedRange = CloneTemplateParagraph(doc, templateParagraph);
            ApplyBiTodoCareerChoiceText(insertedRange, careerChoice);
        }
        finally
        {
            SafeReleaseCom(valueFont, labelFont, valueRange, labelRange, insertedRange);
        }
    }

    private static void AppendBiTodoNormalParagraph(dynamic doc, dynamic templateParagraph, string text)
    {
        AppendSimpleStyledParagraphFromTemplate(doc, templateParagraph, text, 11, false, false, 0);
    }

    private static void AppendSimpleStyledParagraphFromTemplate(
        dynamic doc,
        dynamic templateParagraph,
        string text,
        int fontSize,
        bool bold,
        bool italic,
        int rgbColor)
    {
        dynamic? insertedRange = null;
        dynamic? font = null;
        try
        {
            insertedRange = CloneTemplateParagraph(doc, templateParagraph);
            insertedRange.Text = text + "\r";
            font = insertedRange.Font;
            font.Name = "Segoe UI";
            font.Size = fontSize;
            font.Bold = bold ? 1 : 0;
            font.Italic = italic ? 1 : 0;
            font.Color = rgbColor;
        }
        finally
        {
            SafeReleaseCom(font, insertedRange);
        }
    }

    private static void AppendClonedTemplateParagraph(dynamic doc, dynamic templateParagraph)
    {
        dynamic? insertedRange = null;
        try
        {
            insertedRange = CloneTemplateParagraph(doc, templateParagraph);
        }
        finally
        {
            SafeReleaseCom(insertedRange);
        }
    }

    private static dynamic CloneTemplateParagraph(dynamic doc, dynamic templateParagraph)
    {
        dynamic? insertionRange = null;
        dynamic? templateRange = null;
        dynamic? insertedRange = null;
        try
        {
            insertionRange = GetDocumentEndRange(doc);
            var insertionStart = (int)insertionRange.Start;

            templateRange = templateParagraph.Range;
            insertionRange.FormattedText = templateRange.FormattedText;

            var contentEnd = Math.Max(insertionStart, GetDocumentContentEnd(doc));
            insertedRange = doc.Range(insertionStart, contentEnd);
            return insertedRange;
        }
        catch
        {
            SafeReleaseCom(insertedRange);
            throw;
        }
        finally
        {
            SafeReleaseCom(templateRange, insertionRange);
        }
    }

    private static void ApplyBiTodoCareerChoiceText(dynamic targetRange, string careerChoice)
    {
        dynamic? labelRange = null;
        dynamic? valueRange = null;
        dynamic? labelFont = null;
        dynamic? valueFont = null;
        try
        {
            var normalizedCareerChoice = string.IsNullOrWhiteSpace(careerChoice) ? "-" : careerChoice;
            var fullText = $"Berufswahl: {normalizedCareerChoice}";
            targetRange.Text = fullText + "\r";

            var start = (int)targetRange.Start;
            var labelText = "Berufswahl: ";
            labelRange = targetRange.Document.Range(start, start + labelText.Length);
            valueRange = targetRange.Document.Range(start + labelText.Length, start + fullText.Length);

            labelFont = labelRange.Font;
            labelFont.Name = "Segoe UI";
            labelFont.Size = 12;
            labelFont.Bold = 1;
            labelFont.Italic = 0;
            labelFont.Color = 0;

            valueFont = valueRange.Font;
            valueFont.Name = "Segoe UI";
            valueFont.Size = 12;
            valueFont.Bold = 0;
            valueFont.Italic = 0;
            valueFont.Color = 0;
        }
        finally
        {
            SafeReleaseCom(valueFont, labelFont, valueRange, labelRange);
        }
    }

    private static dynamic CloneTemplateParagraphBlock(dynamic doc, dynamic firstParagraph, dynamic lastParagraph)
    {
        dynamic? insertionRange = null;
        dynamic? firstRange = null;
        dynamic? lastRange = null;
        dynamic? templateBlockRange = null;
        dynamic? insertedRange = null;
        try
        {
            insertionRange = GetDocumentEndRange(doc);
            var insertionStart = (int)insertionRange.Start;

            firstRange = firstParagraph.Range;
            lastRange = lastParagraph.Range;
            templateBlockRange = firstRange.Document.Range((int)firstRange.Start, (int)lastRange.End);
            insertionRange.FormattedText = templateBlockRange.FormattedText;

            var contentEnd = Math.Max(insertionStart, GetDocumentContentEnd(doc));
            insertedRange = doc.Range(insertionStart, contentEnd);
            return insertedRange;
        }
        catch
        {
            SafeReleaseCom(insertedRange);
            throw;
        }
        finally
        {
            SafeReleaseCom(templateBlockRange, lastRange, firstRange, insertionRange);
        }
    }

    private static void AppendStyledParagraphFromTemplate(
        dynamic doc,
        dynamic templateParagraph,
        string text,
        double? fontSize,
        bool? bold,
        bool? italic,
        int? rgbColor)
    {
        dynamic? insertionRange = null;
        dynamic? templateRange = null;
        dynamic? insertedRange = null;
        dynamic? font = null;

        try
        {
            insertionRange = GetDocumentEndRange(doc);
            var insertionStart = (int)insertionRange.Start;

            templateRange = templateParagraph.Range;
            insertionRange.FormattedText = templateRange.FormattedText;

            var contentEnd = Math.Max(insertionStart, GetDocumentContentEnd(doc));
            insertedRange = doc.Range(insertionStart, contentEnd);
            insertedRange.Text = text + "\r";

            font = insertedRange.Font;

            if (fontSize.HasValue)
            {
                font.Size = fontSize.Value;
            }

            if (bold.HasValue)
            {
                font.Bold = bold.Value ? 1 : 0;
            }

            if (italic.HasValue)
            {
                font.Italic = italic.Value ? 1 : 0;
            }

            if (rgbColor.HasValue)
            {
                font.Color = rgbColor.Value;
            }
        }
        finally
        {
            SafeReleaseCom(font, insertedRange, templateRange, insertionRange);
        }
    }

    private static dynamic AppendTemplateParagraph(
        dynamic doc,
        dynamic templateParagraph,
        string text,
        double? fontSize = null,
        bool? bold = null,
        bool? italic = null,
        int? rgbColor = null)
    {
        dynamic? insertionRange = null;
        dynamic? templateRange = null;
        dynamic? insertedRange = null;
        dynamic? font = null;

        try
        {
            insertionRange = GetDocumentEndRange(doc);
            var insertionStart = (int)insertionRange.Start;

            templateRange = templateParagraph.Range;
            insertionRange.FormattedText = templateRange.FormattedText;

            var contentEnd = Math.Max(insertionStart, GetDocumentContentEnd(doc));
            insertedRange = doc.Range(insertionStart, contentEnd);
            insertedRange.Text = text + "\r";

            font = insertedRange.Font;

            if (fontSize.HasValue)
            {
                font.Size = fontSize.Value;
            }

            if (bold.HasValue)
            {
                font.Bold = bold.Value ? 1 : 0;
            }

            if (italic.HasValue)
            {
                font.Italic = italic.Value ? 1 : 0;
            }

            if (rgbColor.HasValue)
            {
                font.Color = rgbColor.Value;
            }

            SafeReleaseCom(font);
            font = null;

            return insertedRange;
        }
        catch
        {
            SafeReleaseCom(insertedRange);
            throw;
        }
        finally
        {
            SafeReleaseCom(font, templateRange, insertionRange);
        }
    }

    private static void AppendSimpleStyledParagraph(
        dynamic doc,
        string text,
        int fontSize,
        bool bold,
        bool italic,
        int rgbColor)
    {
        dynamic? range = null;
        dynamic? font = null;
        try
        {
            range = GetDocumentEndRange(doc);
            range.Text = text + "\r";
            font = range.Font;
            font.Name = "Segoe UI";
            font.Size = fontSize;
            font.Bold = bold ? 1 : 0;
            font.Italic = italic ? 1 : 0;
            font.Color = rgbColor;
        }
        finally
        {
            SafeReleaseCom(font, range);
        }
    }

    private static void LogBiTodoTiming(string step, long elapsedMilliseconds)
    {
        AppLogger.Debug($"Word.CollectBiTodoDocument timing: Step='{step}', Elapsed={elapsedMilliseconds} ms.");
    }

    private static void LogBiTodoParticipantTiming(string participantName, string step, long elapsedMilliseconds)
    {
        AppLogger.Debug($"Word.CollectBiTodoDocument timing: TN='{participantName}', Step='{step}', Elapsed={elapsedMilliseconds} ms.");
    }

    private static string GetBiTodoUserMessage(Exception ex)
    {
        return ex switch
        {
            WordTemplateValidationException validation => validation.UserMessage,
            FileNotFoundException => "keine Akte",
            UnauthorizedAccessException => "Kein Zugriff auf Akte.",
            IOException ioEx when IsLockRelatedFileException(ioEx) => BuildDocumentLockedMessage(),
            InvalidDataException => "Akte konnte nicht gelesen werden. Bitte Vorlage prüfen.",
            COMException comEx when IsLockRelatedHResult((uint)comEx.HResult) => BuildDocumentLockedMessage(),
            _ => "BI: To-dos konnte nicht abgeschlossen werden. Bitte Log prüfen."
        };
    }

    private static void AppendStyledParagraph(
        dynamic app,
        dynamic doc,
        string text,
        int fontSize,
        bool bold,
        bool italic,
        int rgbColor)
    {
        MoveSelectionToDocumentEnd(app, doc);
        dynamic selection = app.Selection;
        selection.Font.Name = "Segoe UI";
        selection.Font.Size = fontSize;
        selection.Font.Bold = bold ? 1 : 0;
        selection.Font.Italic = italic ? 1 : 0;

        if (rgbColor != 0)
        {
            try
            {
                selection.Font.Color = rgbColor;
            }
            catch
            {
            }
        }
        else
        {
            try
            {
                selection.Font.Color = 0;
            }
            catch
            {
            }
        }

        if (!string.IsNullOrEmpty(text))
        {
            selection.TypeText(text);
        }

        selection.TypeParagraph();
    }

    private static dynamic? FindTemplateParagraph(
        dynamic doc,
        Func<string, bool> predicate,
        bool required = true)
    {
        var paragraphCount = (int)doc.Paragraphs.Count;
        for (var index = 1; index <= paragraphCount; index++)
        {
            dynamic? paragraph = null;
            try
            {
                paragraph = doc.Paragraphs[index];
                var text = GetVisibleParagraphText(paragraph);
                if (predicate(text))
                {
                    return paragraph;
                }
            }
            catch
            {
                SafeReleaseCom(paragraph);
                throw;
            }

            SafeReleaseCom(paragraph);
        }

        if (required)
        {
            throw new InvalidOperationException("Bi Vorlage 2.docx ist strukturell nicht brauchbar. Bitte Vorlage prüfen.");
        }

        return null;
    }

    private static dynamic FindFirstListParagraph(object doc)
    {
        var paragraphCount = (int)((dynamic)doc).Paragraphs.Count;
        for (var index = 1; index <= paragraphCount; index++)
        {
            dynamic? paragraph = null;
            dynamic? range = null;
            dynamic? listFormat = null;
            try
            {
                paragraph = ((dynamic)doc).Paragraphs[index];
                range = paragraph.Range;
                listFormat = range.ListFormat;
                if ((int)listFormat.ListType != 0 && !string.IsNullOrWhiteSpace(NormalizeParagraphText(range.Text as string ?? string.Empty)))
                {
                    var match = paragraph;
                    paragraph = null;
                    return match!;
                }
            }
            finally
            {
                SafeReleaseCom(listFormat, range, paragraph);
            }
        }

        throw new InvalidOperationException("Bi Vorlage 2.docx ist strukturell nicht brauchbar. Bitte Vorlage prüfen.");
    }

    private static string GetVisibleParagraphText(dynamic paragraph)
    {
        dynamic? range = null;
        try
        {
            range = paragraph.Range;
            return NormalizeParagraphText(range.Text as string ?? string.Empty);
        }
        finally
        {
            SafeReleaseCom(range);
        }
    }

    private static string NormalizeParagraphText(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        return text
            .Replace("\r", string.Empty, StringComparison.Ordinal)
            .Replace("\a", string.Empty, StringComparison.Ordinal)
            .Trim();
    }

    internal static string NormalizeWhitespaceForBiDocx(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        return MultiWhitespaceRegex
            .Replace(text.Replace("\r", " ", StringComparison.Ordinal).Replace("\n", " ", StringComparison.Ordinal), " ")
            .Trim();
    }

    private static bool IsSeparatorParagraphText(string text)
    {
        var trimmed = text.Trim();
        return trimmed.Length >= 10 && trimmed.All(ch => ch == '-');
    }

    private static int GetDocumentContentEnd(dynamic doc)
    {
        dynamic? content = null;
        try
        {
            content = doc.Content;
            return Math.Max(0, (int)content.End - 1);
        }
        finally
        {
            SafeReleaseCom(content);
        }
    }

    private static bool ReplaceTextInInsertedRange(dynamic insertedRange, string searchText, string replacement)
    {
        var currentText = insertedRange.Text as string ?? string.Empty;
        var index = currentText.IndexOf(searchText, StringComparison.Ordinal);
        if (index < 0)
        {
            return false;
        }

        dynamic? replacementRange = null;
        try
        {
            var start = (int)insertedRange.Start + index;
            var end = start + searchText.Length;
            replacementRange = insertedRange.Document.Range(start, end);
            replacementRange.Text = replacement ?? string.Empty;
            return true;
        }
        finally
        {
            SafeReleaseCom(replacementRange);
        }
    }

    private static dynamic GetDocumentEndRange(dynamic doc)
    {
        dynamic? content = null;
        try
        {
            content = doc.Content;
            var end = Math.Max(0, (int)content.End - 1);
            return doc.Range(end, end);
        }
        finally
        {
            SafeReleaseCom(content);
        }
    }

    private static void MoveSelectionToDocumentEnd(dynamic app, dynamic doc)
    {
        dynamic? content = null;
        try
        {
            content = doc.Content;
            var end = Math.Max(0, (int)content.End - 1);
            doc.Activate();
            app.Selection.SetRange(end, end);
        }
        finally
        {
            SafeReleaseCom(content);
        }
    }

    private static void EnsureDocumentNotLocked(dynamic doc, bool openedHere)
    {
        bool isReadOnly;
        try
        {
            isReadOnly = (bool)doc.ReadOnly;
        }
        catch
        {
            return;
        }

        if (!isReadOnly)
        {
            return;
        }

        if (openedHere)
        {
            try
            {
                doc.Close(false);
            }
            catch
            {
                // Falls das Schliessen fehlschlaegt, trotzdem die lock-spezifische Meldung zeigen.
            }
        }

        throw new InvalidOperationException(BuildDocumentLockedMessage());
    }

    private static void EnsureWordUiState(dynamic app, WordLifecycleOperationContext? context = null)
    {
        try
        {
            // Bei bereits laufender/angebundenen Instanz kann UserControl schreibgeschuetzt sein.
            app.UserControl = true;
            LogWordLifecycle(context, "EnsureWordUiState: UserControl erfolgreich gesetzt.");
        }
        catch (Exception ex)
        {
            LogWordLifecycle(context, $"EnsureWordUiState: UserControl fehlgeschlagen. Type='{SanitizeForLog(ex.GetType().Name)}', Message='{SanitizeForLog(ex.Message)}'.");
            AppLogger.Warn($"Word.UserControl konnte nicht gesetzt werden: {ex.Message}");
        }

        app.Visible = true;
        LogWordLifecycle(context, "EnsureWordUiState: Visible=true gesetzt.");
        TryApplyPreferredWindowPlacement(app);
        TryBringWordToForeground(app);
    }

    private static int CountUnsavedDocuments(dynamic app)
    {
        dynamic? docs = null;
        try
        {
            docs = app.Documents;
            var documentCount = (int)docs.Count;
            AppLogger.Debug($"Word.CountUnsavedDocuments: OpenDocs={documentCount}.");
            var unsavedCount = 0;

            for (var i = 1; i <= documentCount; i++)
            {
                dynamic? doc = null;
                try
                {
                    doc = docs[i];
                    if (IsUnsavedDocument(doc))
                    {
                        unsavedCount++;
                    }
                }
                finally
                {
                    SafeReleaseCom(doc);
                }
            }

            return unsavedCount;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word: Anzahl ungespeicherter Dokumente konnte nicht ermittelt werden: {ex.Message}");
            return 0;
        }
        finally
        {
            SafeReleaseCom(docs);
        }
    }

    private static bool IsUnsavedDocument(dynamic doc)
    {
        try
        {
            var docPath = doc.Path as string;
            return string.IsNullOrWhiteSpace(docPath);
        }
        catch
        {
            return false;
        }
    }

    private static void CloseTransientEmptyDocuments(dynamic app, string targetDocPath, int initialUnsavedDocumentCount, WordLifecycleOperationContext? context = null)
    {
        dynamic? docs = null;
        var unsavedDocuments = new List<object?>();
        try
        {
            docs = app.Documents;
            var targetPath = Path.GetFullPath(targetDocPath);
            var documentCount = (int)docs.Count;
            AppLogger.Debug($"Word.CloseTransientEmptyDocuments: Target='{targetPath}', InitialUnsaved={initialUnsavedDocumentCount}, OpenDocs={documentCount}.");

            for (var i = documentCount; i >= 1; i--)
            {
                dynamic? openDoc = null;
                var keepOpenDoc = false;
                try
                {
                    openDoc = docs[i];

                    string? fullName = null;
                    try
                    {
                        fullName = openDoc.FullName as string;
                    }
                    catch
                    {
                        // Ignorieren.
                    }

                    if (!string.IsNullOrWhiteSpace(fullName) &&
                        Path.GetFullPath(fullName).Equals(targetPath, StringComparison.OrdinalIgnoreCase))
                    {
                        LogWordLifecycle(context, $"CloseTransientEmptyDocuments: Ziel-Dokument uebersprungen. Index={i}, FullName='{SanitizeForLog(fullName)}'.");
                        continue;
                    }

                    var isUnsavedDocument = IsUnsavedDocument(openDoc);
                    LogWordLifecycle(
                        context,
                        $"CloseTransientEmptyDocuments: Inspect Index={i}, FullName='{SanitizeForLog(fullName)}', IsUnsaved={isUnsavedDocument}.");

                    if (isUnsavedDocument)
                    {
                        unsavedDocuments.Add(openDoc);
                        keepOpenDoc = true;
                    }
                }
                finally
                {
                    if (!keepOpenDoc)
                    {
                        SafeReleaseCom(openDoc);
                    }
                }
            }

            var documentsToClose = Math.Max(0, unsavedDocuments.Count - Math.Max(0, initialUnsavedDocumentCount));
            AppLogger.Debug($"Word.CloseTransientEmptyDocuments: UnsavedNow={unsavedDocuments.Count}, ToClose={documentsToClose}.");
            for (var i = 0; i < unsavedDocuments.Count && documentsToClose > 0; i++, documentsToClose--)
            {
                dynamic? transientDoc = unsavedDocuments[i];
                if (transientDoc is null)
                {
                    continue;
                }

                try
                {
                    var transientName = TryReadString(() => transientDoc.Name as string) ?? string.Empty;
                    transientDoc.Close(false);
                    LogWordLifecycle(context, $"CloseTransientEmptyDocuments: Transientes Leerdokument geschlossen. Name='{SanitizeForLog(transientName)}'.");
                    AppLogger.Info("Word: Transientes Leerdokument nach Aktion geschlossen.");
                }
                catch (Exception ex)
                {
                    LogWordLifecycle(context, $"CloseTransientEmptyDocuments: Schliessen fehlgeschlagen. Type='{SanitizeForLog(ex.GetType().Name)}', Message='{SanitizeForLog(ex.Message)}'.");
                    AppLogger.Warn($"Word: Transientes Leerdokument konnte nicht geschlossen werden: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            LogWordLifecycle(context, $"CloseTransientEmptyDocuments: Gesamtpruefung fehlgeschlagen. Type='{SanitizeForLog(ex.GetType().Name)}', Message='{SanitizeForLog(ex.Message)}'.");
            AppLogger.Warn($"Word: Pruefung auf transiente Leerdokumente fehlgeschlagen: {ex.Message}");
        }
        finally
        {
            foreach (var unsavedDoc in unsavedDocuments)
            {
                SafeReleaseCom(unsavedDoc);
            }

            SafeReleaseCom(docs);
        }
    }

    private static void FocusBookmarkAtTop(dynamic app, dynamic doc, string bookmarkName)
    {
        // Word kann nach dem Oeffnen noch asynchron auf die zuletzt gespeicherte Position springen.
        // Deshalb mehrmals auf die Bookmark fokussieren.
        const int maxAttempts = 3;
        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            dynamic? bookmark = null;
            dynamic? targetRange = null;
            try
            {
                bookmark = doc.Bookmarks[bookmarkName];
                targetRange = bookmark.Range;
                FocusRangeAtTop(app, targetRange);
            }
            finally
            {
                SafeReleaseCom(targetRange, bookmark);
            }

            if (attempt < maxAttempts)
            {
                Thread.Sleep(120);
            }
        }
    }

    private static void FocusRangeAtTop(dynamic app, dynamic range)
    {
        try
        {
            TryActivateWordWindow(app);
            app.Activate();
            var start = (int)range.Start;
            app.Selection.SetRange(start, start);
            app.ActiveWindow?.ScrollIntoView(app.Selection.Range, true);
            TryBringWordToForeground(app);
            return;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word.ScrollIntoView top-focus fehlgeschlagen ({ex.GetType().Name}): {ex.Message}");
        }

        try
        {
            TryActivateWordWindow(app);
            range.Select();
            app.ActiveWindow?.ScrollIntoView(range, true);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word.Range.Select fallback fehlgeschlagen ({ex.GetType().Name}): {ex.Message}");
        }

        TryBringWordToForeground(app);
    }

    private static void FocusDocument(dynamic app, dynamic doc)
    {
        try
        {
            doc.Activate();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word.Doc.Activate fehlgeschlagen ({ex.GetType().Name}): {ex.Message}");
        }

        TryActivateWordWindow(app);
        TryBringWordToForeground(app);
    }

    private static void TryActivateWordWindow(dynamic app)
    {
        try
        {
            app.ActiveWindow?.Activate();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word.ActiveWindow.Activate fehlgeschlagen ({ex.GetType().Name}): {ex.Message}");
        }
    }

    private static void TryApplyPreferredWindowPlacement(dynamic app)
    {
        try
        {
            var prefs = App.UserPrefs;
            if (prefs is null || !prefs.OpenWordMaximized)
            {
                return;
            }

            dynamic targetWindow = TryGetWordTargetWindow(app);
            if (targetWindow is null)
            {
                return;
            }

            var screen = ResolvePreferredScreen(prefs.PreferredWordMonitorId);
            var workingArea = screen.WorkingArea;

            TrySetWordWindowState(targetWindow, WordWindowStateNormal);
            targetWindow.Left = workingArea.Left;
            targetWindow.Top = workingArea.Top;
            targetWindow.Width = workingArea.Width;
            targetWindow.Height = workingArea.Height;
            TrySetWordWindowState(targetWindow, WordWindowStateMaximized);
            AppLogger.Debug($"Word: Platzierung angewendet. Monitor='{screen.DeviceName}', Left={workingArea.Left}, Top={workingArea.Top}, Width={workingArea.Width}, Height={workingArea.Height}, Maximized=True.");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word: Gewuenschte Monitor-/Maximierungsplatzierung konnte nicht angewendet werden: {ex.Message}");
        }
    }

    private static Forms.Screen ResolvePreferredScreen(string? monitorId)
    {
        var screens = Forms.Screen.AllScreens;
        if (screens.Length == 0)
        {
            throw new InvalidOperationException("Kein Bildschirm verfügbar.");
        }

        if (string.IsNullOrWhiteSpace(monitorId) ||
            string.Equals(monitorId, PrimaryMonitorId, StringComparison.OrdinalIgnoreCase))
        {
            return Forms.Screen.PrimaryScreen ?? screens[0];
        }

        foreach (var screen in screens)
        {
            if (string.Equals(screen.DeviceName, monitorId, StringComparison.OrdinalIgnoreCase))
            {
                return screen;
            }
        }

        return Forms.Screen.PrimaryScreen ?? screens[0];
    }

    private static dynamic? TryGetWordTargetWindow(dynamic app)
    {
        try
        {
            var activeWindow = app.ActiveWindow;
            if (activeWindow is not null)
            {
                return activeWindow;
            }
        }
        catch
        {
        }

        try
        {
            return app;
        }
        catch
        {
            return null;
        }
    }

    private static void TrySetWordWindowState(dynamic targetWindow, int state)
    {
        try
        {
            if (targetWindow is not null)
            {
                targetWindow.WindowState = state;
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word: WindowState konnte nicht gesetzt werden: {ex.Message}");
        }
    }

    private static void TryBringWordToForeground(dynamic app)
    {
        var hwnd = TryGetWordMainWindowHandle(app);
        if (hwnd == IntPtr.Zero)
        {
            return;
        }

        try
        {
            NativeMethods.ShowWindowAsync(hwnd, NativeMethods.SW_RESTORE);
            NativeMethods.SetForegroundWindow(hwnd);
            Thread.Sleep(WordForegroundRetryDelayMs);
            NativeMethods.ShowWindowAsync(hwnd, NativeMethods.SW_SHOW);
            NativeMethods.SetForegroundWindow(hwnd);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word foreground fallback fehlgeschlagen ({ex.GetType().Name}): {ex.Message}");
        }
    }

    private static IntPtr TryGetWordMainWindowHandle(dynamic app)
    {
        try
        {
            var hwndRaw = (int)app.Hwnd;
            if (hwndRaw > 0)
            {
                return new IntPtr(hwndRaw);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word.Hwnd konnte nicht gelesen werden ({ex.GetType().Name}): {ex.Message}");
        }

        try
        {
            var hwndRaw = (int)app.ActiveWindow.Hwnd;
            if (hwndRaw > 0)
            {
                return new IntPtr(hwndRaw);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word.ActiveWindow.Hwnd Fallback fehlgeschlagen ({ex.GetType().Name}): {ex.Message}");
        }

        return IntPtr.Zero;
    }

    private static void TryQuitWordApplication(dynamic? app)
    {
        if (app is null)
        {
            return;
        }

        try
        {
            app.Quit(false);
            AppLogger.Info("Word: Selbst gestartete Instanz nach Fehler geschlossen.");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word: Selbst gestartete Instanz konnte nach Fehler nicht beendet werden: {ex.Message}");
        }
    }

    private static void TryCloseDocument(dynamic? doc)
    {
        if (doc is null)
        {
            return;
        }

        try
        {
            doc.Close(false);
            AppLogger.Info("Word: Dokument nach Fehler geschlossen (Leak-Schutz).");
            AppLogger.Debug("Word.TryCloseDocument: Dokument per Leak-Schutz geschlossen.");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word: Dokument konnte nach Fehler nicht geschlossen werden: {ex.Message}");
        }
    }

    private static void TryCloseDocumentSilently(dynamic? doc, string infoMessage)
    {
        if (doc is null)
        {
            return;
        }

        try
        {
            doc.Close(false);
            AppLogger.Info(infoMessage);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word: Dokument konnte nicht geschlossen werden: {ex.Message}");
        }
    }

    private static void DeleteTempFileQuietly(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return;
        }

        try
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word: Temporäre BI-Vorlage konnte nicht gelöscht werden: {ex.Message}");
        }
    }

    private static void SafeReleaseCom(params dynamic?[] comObjects)
    {
        foreach (var obj in comObjects)
        {
            if (obj is null) { continue; }
            try
            {
                if (Marshal.IsComObject(obj))
                {
                    Marshal.ReleaseComObject(obj);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"Word: COM-Objekt konnte nicht freigegeben werden: {ex.Message}");
            }
        }
    }

    private static void TryDeleteRow(dynamic? row)
    {
        if (row is null)
        {
            return;
        }

        try
        {
            row.Delete();
            AppLogger.Info("Word: Teilweise befuellte Zeile nach Fehler entfernt (Rollback).");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word: Zeile konnte nach Fehler nicht entfernt werden: {ex.Message}");
        }
    }

    private static string GetClipboardTextWithRetry()
    {
        Exception? lastException = null;

        for (var attempt = 1; attempt <= ClipboardReadRetryCount; attempt++)
        {
            try
            {
                return Clipboard.ContainsText() ? Clipboard.GetText() : string.Empty;
            }
            catch (COMException ex)
            {
                lastException = ex;
                AppLogger.Warn($"Word: Clipboard-COM-Zugriff fehlgeschlagen (Attempt {attempt}/{ClipboardReadRetryCount}): {ex.Message}");
            }
            catch (ExternalException ex)
            {
                lastException = ex;
                AppLogger.Warn($"Word: Clipboard-Zugriff fehlgeschlagen (Attempt {attempt}/{ClipboardReadRetryCount}): {ex.Message}");
            }

            if (attempt < ClipboardReadRetryCount)
            {
                Thread.Sleep(ClipboardReadRetryBaseDelayMs * (1 << (attempt - 1)));
            }
        }

        throw new InvalidOperationException(
            "Zwischenablage ist momentan blockiert. Bitte kurz warten und erneut versuchen.",
            lastException);
    }

    private static bool TryParseClipboardFields(string clipboardText, out string[] fields)
    {
        fields = Array.Empty<string>();

        if (string.IsNullOrWhiteSpace(clipboardText))
        {
            return false;
        }

        var normalized = clipboardText
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Trim('\n');

        var lines = normalized.Split('\n', StringSplitOptions.None);
        if (lines.Length != 1)
        {
            return false;
        }

        var parts = lines[0].Split('\t');
        if (parts.Length != 4)
        {
            return false;
        }

        fields = parts;
        return true;
    }

    private static bool TryParseSingleTabSeparatedRow(string rowText, out string[] fields)
    {
        fields = Array.Empty<string>();
        if (string.IsNullOrWhiteSpace(rowText))
        {
            return false;
        }

        var normalized = rowText
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Trim('\n');

        var lines = normalized.Split('\n', StringSplitOptions.None);
        if (lines.Length != 1)
        {
            return false;
        }

        var parts = lines[0].Split('\t');
        if (parts.Length != 4)
        {
            return false;
        }

        fields = parts;
        return true;
    }

    private static dynamic ResolveStructuredEntryTableForWrite(
        dynamic doc,
        string bookmarkName,
        int expectedColumnCount,
        string tableDisplayName)
    {
        if (!doc.Bookmarks.Exists(bookmarkName))
        {
            throw CreateBookmarkMissingException(bookmarkName);
        }

        dynamic? bookmark = null;
        dynamic? bookmarkRange = null;
        int bookmarkStart;
        try
        {
            bookmark = doc.Bookmarks[bookmarkName];
            bookmarkRange = bookmark.Range;
            bookmarkStart = (int)bookmarkRange.Start;

            var currentTable = GetContainingBookmarkTable(bookmarkRange);
            var returnedCurrentTable = false;
            try
            {
                if (currentTable is not null && IsStructuredEntryTable(currentTable, expectedColumnCount))
                {
                    returnedCurrentTable = true;
                    return currentTable!;
                }
            }
            finally
            {
                if (!returnedCurrentTable)
                {
                    SafeReleaseCom(currentTable);
                }
            }
        }
        finally
        {
            SafeReleaseCom(bookmarkRange, bookmark);
        }

        var tableCount = (int)doc.Tables.Count;
        for (var tableIndex = 1; tableIndex <= tableCount; tableIndex++)
        {
            dynamic table = doc.Tables[tableIndex];
            var tableStart = GetTableStart(table);
            if (tableStart < bookmarkStart)
            {
                SafeReleaseCom(table);
                continue;
            }

            if (IsStructuredEntryTable(table, expectedColumnCount))
            {
                return table;
            }

            SafeReleaseCom(table);
        }

        throw CreateStructuredEntryTableInvalidException(bookmarkName, tableDisplayName);
    }

    private static dynamic ResolveBiTodoTable(dynamic doc, string bookmarkName)
    {
        if (!doc.Bookmarks.Exists(bookmarkName))
        {
            throw CreateBookmarkMissingException(bookmarkName);
        }

        dynamic? bookmark = null;
        dynamic? bookmarkRange = null;
        dynamic? currentTable = null;
        var returnedCurrentTable = false;
        try
        {
            bookmark = doc.Bookmarks[bookmarkName];
            bookmarkRange = bookmark.Range;
            currentTable = GetContainingBookmarkTable(bookmarkRange);

            if (currentTable is not null && IsBiTodoTable(currentTable))
            {
                returnedCurrentTable = true;
                return currentTable!;
            }
        }
        finally
        {
            if (!returnedCurrentTable)
            {
                SafeReleaseCom(currentTable);
            }

            SafeReleaseCom(bookmarkRange, bookmark);
        }

        throw CreateBiTodoTableInvalidException(bookmarkName);
    }

    private static string ReadBiCareerChoice(dynamic sourceDoc, dynamic sourceTable, string bookmarkName)
    {
        var tableStart = GetTableStart(sourceTable);
        var paragraphCount = (int)sourceDoc.Paragraphs.Count;
        var berufswuenscheEnd = -1;
        var arbeitseinsaetzeStart = -1;

        for (var index = 1; index <= paragraphCount; index++)
        {
            dynamic? paragraph = null;
            dynamic? range = null;
            try
            {
                paragraph = sourceDoc.Paragraphs[index];
                range = paragraph.Range;
                var text = NormalizeTableCellText(range.Text as string ?? string.Empty);

                if (berufswuenscheEnd < 0 &&
                    string.Equals(text, "berufswünsche", StringComparison.OrdinalIgnoreCase))
                {
                    berufswuenscheEnd = (int)range.End;
                }

                if (arbeitseinsaetzeStart < 0 &&
                    string.Equals(text, "arbeitseinsätze", StringComparison.OrdinalIgnoreCase))
                {
                    arbeitseinsaetzeStart = (int)range.Start;
                }

                if (berufswuenscheEnd >= 0 && arbeitseinsaetzeStart >= 0)
                {
                    break;
                }
            }
            finally
            {
                SafeReleaseCom(range, paragraph);
            }
        }

        if (berufswuenscheEnd < 0)
        {
            throw new WordTemplateValidationException(
                WordTemplateValidationErrorKind.BiTodoContentInvalid,
                bookmarkName,
                "Berufswünsche-Bereich konnte nicht gefunden werden. Bitte Vorlage prüfen.");
        }

        dynamic? selectedControl = null;
        dynamic? selectedRange = null;
        dynamic? searchRange = null;
        try
        {
            var searchEnd = tableStart;
            if (arbeitseinsaetzeStart > berufswuenscheEnd && arbeitseinsaetzeStart < searchEnd)
            {
                searchEnd = arbeitseinsaetzeStart;
            }

            if (searchEnd <= berufswuenscheEnd)
            {
                searchEnd = tableStart;
            }

            searchRange = sourceDoc.Range(berufswuenscheEnd, searchEnd);
            var controlCount = (int)searchRange.ContentControls.Count;
            if (controlCount > 0)
            {
                selectedControl = searchRange.ContentControls[1];
                selectedRange = selectedControl.Range;
            }

            if (selectedRange is null)
            {
                throw new WordTemplateValidationException(
                    WordTemplateValidationErrorKind.BiTodoContentInvalid,
                    bookmarkName,
                    "Berufswunsch konnte nicht gelesen werden. Bitte Vorlage prüfen.");
            }

            var rawText = NormalizeParagraphText(selectedRange.Text as string ?? string.Empty);
            if (string.IsNullOrWhiteSpace(rawText) ||
                string.Equals(rawText, "[Titel]", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(NormalizeTableCellText(rawText), "eba / efz, bereiche", StringComparison.OrdinalIgnoreCase))
            {
                return "-";
            }

            return rawText;
        }
        finally
        {
            SafeReleaseCom(searchRange, selectedRange, selectedControl);
        }
    }

    private static IReadOnlyList<BiTodoParagraphContent> ReadBiTodoParagraphs(dynamic sourceTable, string bookmarkName)
    {
        dynamic? targetCell = null;
        dynamic? cellRange = null;
        try
        {
            var rowCount = (int)sourceTable.Rows.Count;
            var columnCount = (int)sourceTable.Columns.Count;
            if (rowCount < 1 || columnCount < 2)
            {
                throw CreateBiTodoTableInvalidException(bookmarkName);
            }

            targetCell = sourceTable.Cell(rowCount, columnCount);
            cellRange = targetCell.Range;

            var paragraphs = new List<BiTodoParagraphContent>();
            var paragraphCount = (int)cellRange.Paragraphs.Count;
            for (var index = 1; index <= paragraphCount; index++)
            {
                dynamic? paragraph = null;
                dynamic? paragraphRange = null;
                dynamic? listFormat = null;
                try
                {
                    paragraph = cellRange.Paragraphs[index];
                    paragraphRange = paragraph.Range;
                    var text = NormalizeParagraphText(paragraphRange.Text as string ?? string.Empty);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        listFormat = paragraphRange.ListFormat;
                        var listType = 0;
                        try
                        {
                            listType = (int)listFormat.ListType;
                        }
                        catch
                        {
                        }

                        paragraphs.Add(new BiTodoParagraphContent
                        {
                            Text = text,
                            IsBullet = listType != 0
                        });
                    }
                }
                finally
                {
                    SafeReleaseCom(listFormat, paragraphRange, paragraph);
                }
            }

            return paragraphs;
        }
        catch (WordTemplateValidationException)
        {
            throw;
        }
        catch (COMException ex)
        {
            throw new WordTemplateValidationException(
                WordTemplateValidationErrorKind.BiTodoContentInvalid,
                bookmarkName,
                $"BI-To-do-Zelle konnte nicht gelesen werden: {ex.Message}");
        }
        catch (Exception ex)
        {
            throw new WordTemplateValidationException(
                WordTemplateValidationErrorKind.BiTodoContentInvalid,
                bookmarkName,
                $"BI-To-do-Zelle konnte nicht gelesen werden: {ex.Message}");
        }
        finally
        {
            SafeReleaseCom(cellRange, targetCell);
        }
    }

    private static object? GetContainingBookmarkTable(dynamic bookmarkRange)
    {
        if ((int)bookmarkRange.Tables.Count <= 0)
        {
            return null;
        }

        return bookmarkRange.Tables[1];
    }

    private static int GetTableStart(dynamic table)
    {
        dynamic? range = null;
        try
        {
            range = table.Range;
            return (int)range.Start;
        }
        catch
        {
            return int.MaxValue;
        }
        finally
        {
            SafeReleaseCom(range);
        }
    }

    private static bool IsStructuredEntryTable(dynamic table, int expectedColumnCount)
    {
        try
        {
            if ((int)table.Rows.Count < 1)
            {
                return false;
            }

            if ((int)table.Columns.Count != expectedColumnCount)
            {
                return false;
            }

            var expectedHeaders = new[] { "datum", "eintrag von", "thematik", "beschreibung" };
            for (var column = 1; column <= expectedHeaders.Length; column++)
            {
                dynamic? cell = null;
                dynamic? range = null;
                try
                {
                    cell = table.Cell(1, column);
                    range = cell.Range;
                    var text = NormalizeTableCellText((string)range.Text);
                    if (!string.Equals(text, expectedHeaders[column - 1], StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                }
                finally
                {
                    SafeReleaseCom(range, cell);
                }
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool IsBiTodoTable(dynamic table)
    {
        try
        {
            if ((int)table.Rows.Count < 1)
            {
                return false;
            }

            return (int)table.Columns.Count == 2;
        }
        catch
        {
            return false;
        }
    }

    private static string NormalizeTableCellText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        var cleaned = text
            .Replace("\r", " ", StringComparison.Ordinal)
            .Replace("\n", " ", StringComparison.Ordinal)
            .Replace("\a", " ", StringComparison.Ordinal)
            .Trim();

        return MultiWhitespaceRegex
            .Replace(cleaned, " ")
            .Trim()
            .ToLowerInvariant();
    }

    private static string GetStructuredEntryTableDisplayName(string bookmarkName)
    {
        return bookmarkName.Contains("BI_", StringComparison.OrdinalIgnoreCase) ? "BI" : "BU";
    }

    private static WordTemplateValidationException CreateBookmarkMissingException(string bookmarkName)
    {
        return new WordTemplateValidationException(
            WordTemplateValidationErrorKind.BookmarkMissing,
            bookmarkName,
            $"Bookmark '{bookmarkName}' nicht gefunden. Bitte Vorlage prüfen.");
    }

    private static WordTemplateValidationException CreateStructuredEntryTableInvalidException(
        string bookmarkName,
        string tableDisplayName)
    {
        return new WordTemplateValidationException(
            WordTemplateValidationErrorKind.StructuredEntryTableInvalid,
            bookmarkName,
            $"{tableDisplayName}-Verlaufstabelle hat nicht das erwartete Format. Bitte Vorlage prüfen.");
    }

    private static WordTemplateValidationException CreateBiTodoTableInvalidException(string bookmarkName)
    {
        return new WordTemplateValidationException(
            WordTemplateValidationErrorKind.BiTodoTableInvalid,
            bookmarkName,
            "BI-To-do-Tabelle hat nicht das erwartete Format. Bitte Vorlage prüfen.");
    }

    private static dynamic InsertRowAtTopOfDataArea(dynamic targetTable)
    {
        var existingRowCount = (int)targetTable.Rows.Count;
        if (existingRowCount <= 1)
        {
            targetTable.Rows.Add();
            return targetTable.Rows[(int)targetTable.Rows.Count];
        }

        return targetTable.Rows.Add(targetTable.Rows[2]);
    }

    private static int GetSafeEditColumn(dynamic targetTable, int preferredColumn)
    {
        try
        {
            var columnCount = (int)targetTable.Columns.Count;
            if (columnCount <= 0)
            {
                return 1;
            }

            return Math.Clamp(preferredColumn, 1, columnCount);
        }
        catch
        {
            return Math.Max(1, preferredColumn);
        }
    }

    private static bool IsFileLocked(string path)
    {
        try
        {
            using var stream = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        }
        catch (IOException ex) when ((uint)ex.HResult == 0x80070020 || (uint)ex.HResult == 0x80070021)
        {
            return true;
        }
    }

    private static bool IsLockRelatedHResult(uint hresult)
    {
        return hresult == 0x800A175D   // wdErrorFileLocked
            || hresult == 0x80070020   // ERROR_SHARING_VIOLATION
            || hresult == 0x80070021;  // ERROR_LOCK_VIOLATION
    }

    private static bool IsLockRelatedFileException(IOException ex)
    {
        var hresult = (uint)ex.HResult;
        return hresult == 0x80070020 || hresult == 0x80070021;
    }

    private static string BuildDocumentLockedMessage()
    {
        var message = DocumentLockedMessage;
        try
        {
            var localWordProcessCount = Process.GetProcessesByName("WINWORD").Length;
            if (localWordProcessCount > 1)
            {
                message += " Hinweis: Mehrere lokale Word-Instanzen erkannt.";
                AppLogger.Warn($"Word-Lock-Hinweis: {localWordProcessCount} lokale WINWORD-Prozesse erkannt.");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Word-Lock-Hinweis konnte lokale Prozesse nicht prüfen: {ex.Message}");
        }

        return message;
    }

    private const int WordWindowStateNormal = 0;
    private const int WordWindowStateMaximized = 1;

    private static class NativeMethods
    {
        public const int SW_SHOW = 5;
        public const int SW_RESTORE = 9;

        [DllImport("oleaut32.dll", PreserveSig = false)]
        public static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    }

    private sealed class WordApplicationHandle
    {
        public WordApplicationHandle(dynamic app, bool wasCreatedHere, int initialUnsavedDocumentCount)
        {
            App = app;
            WasCreatedHere = wasCreatedHere;
            InitialUnsavedDocumentCount = initialUnsavedDocumentCount;
        }

        public dynamic App { get; }
        public bool WasCreatedHere { get; }
        public int InitialUnsavedDocumentCount { get; }
    }
}
