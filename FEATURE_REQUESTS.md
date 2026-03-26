# Feature Requests – Verlaufsakte-App

Dieses Dokument beschreibt drei geplante Features zur Implementierung.
Technische Basis: WPF, .NET 8, C#, Code-Behind (kein MVVM), late-bound COM für Word.

---

## Feature 1: Einstellungs-Panel (Zahnrad-Icon in Titelleiste)

### Ziel
Nutzer sollen die App ohne Datei-Editierung konfigurieren können: sichtbare Buttons, Serverpfad, Theme.

### UI-Einstiegspunkt
Ein Zahnrad-Button `⚙` in der custom Titelleiste, links neben `—` und `✕`.
Style analog `TitleBarButtonStyle`. Klick öffnet ein Modal-Panel **innerhalb** des Hauptfensters (kein neues Window), das die restliche UI überlagert (z. B. als Overlay-`Border` mit `Grid.Row`-Span, semi-transparenter Backdrop).

### 1A – Button-Sichtbarkeit konfigurieren

**Vorhandene Buttons (alle standardmäßig sichtbar):**
| Key | Label | Aktion |
|-----|-------|--------|
| `ShowBtnOrdner` | Ordner | Explorer öffnen |
| `ShowBtnAkte` | Akte | Word öffnen |
| `ShowBtnBu` | BU | `OpenDocumentAtBookmark` mit `WordBuBookmarkName` (`_Bildung`) |
| `ShowBtnEintrag` | Eintrag | `InsertClipboardToTable` mit `WordBookmarkName` (`BU_BILDUNG_TABELLE`) |

**Neue Buttons (standardmäßig ausgeblendet):**
| Key | Label | Aktion |
|-----|-------|--------|
| `ShowBtnBi` | BI | `OpenDocumentAtBookmark` mit `WordBiBookmarkName` (`_Berufsintegration`) |
| `ShowBtnEintragBi` | Eintrag BI | `InsertClipboardToTable` mit `WordBiBookmarkName` (`BI_BERUFSINTEGRATION_TABELLE`) |

**Implementierungshinweise:**
- Im Einstellungs-Panel: je ein `CheckBox`-Toggle pro Button.
- `ParticipantCardTemplate` (XAML): Jeder Button erhält ein `Visibility`-Binding auf eine Window-DependencyProperty (`ShowBtnOrdner`, …). Bei `false` → `Collapsed` (kein Platzbedarf).
- Die `IsWideLayout`-Logik in `UpdateLayoutMode()` muss die Breite der aktuell **sichtbaren** Buttons messen, nicht aller sechs.
- `Participant.cs` braucht zwei neue `Can`-Properties: `CanOpenAkteBi`, `CanInsertEntryBi`. Werden in `UpdateActionState()` gesetzt analog zu `CanOpenAkteBu`/`CanInsertEntry`.
- Zwei neue Click-Handler in `MainWindow.xaml.cs`: `OpenAkteBiButton_OnClick`, `InsertEntryBiButton_OnClick` – nahezu identisch zu den BU/Eintrag-Entsprechungen, nur mit den BI-Bookmark-Namen.

### 1B – Serverpfad in der UI setzen

**UI:** TextBox mit aktuellem `AppConfig.ServerBasePath`-Wert + „Übernehmen"-Button.
Optional: Folder-Browse-Dialog via `System.Windows.Forms.FolderBrowserDialog` (oder `Microsoft.Win32.OpenFolderDialog` ab .NET 8).

**Verhalten:**
- Pfad wird in `AppConfig` und in `settings.json` gespeichert (überschreibt die Datei).
- Nach Übernahme: `_matcher.BuildIndex()` wird beim nächsten Auswerten neu ausgeführt (Index ist dann stale, kein sofortiges Rebuild nötig).
- Validierung: Pfad existiert? Sonst Toast `Warning`.

**Speichern:** `JsonSerializer.Serialize(Config, ...)` → `File.WriteAllText(settingsPath, json)`.
`settingsPath` = `Path.Combine(AppContext.BaseDirectory, "settings.json")`.

### 1C – Theme: Hell / Dunkel

**Zwei Farbschemata:**

| Token | Dunkel (aktuell) | Hell |
|-------|-----------------|------|
| WindowBg | `#1F2024` | `#F0F0F2` |
| PanelBg | `#2A2B31` | `#FFFFFF` |
| CardBg | `#2E2F36` | `#F5F5F7` |
| CardHover | `#34353D` | `#E8E8EC` |
| PrimaryText | `#E0E0E0` | `#1A1A1F` |
| SecondaryText | `#9B9EA8` | `#6B6E78` |
| Border | `#3A3B42` | `#D0D1D8` |
| Accent | `#8B7D6B` | `#7A6E60` |
| StatusBarBg | `#252630` | `#E8E8EC` |

**Implementierungsansatz:**
Alle Farben als `SolidColorBrush`-Ressourcen mit festgelegten Keys in `App.xaml` (z. B. `{x:Key="Brush.WindowBg"}`). Theme-Wechsel ersetzt die Brush-Werte zur Laufzeit:

```csharp
// In App.xaml.cs oder MainWindow.xaml.cs
private void ApplyTheme(bool isDark)
{
    var res = Application.Current.Resources;
    res["Brush.WindowBg"] = BrushFromHex(isDark ? "#1F2024" : "#F0F0F2");
    // ... alle anderen Tokens
}
```

Alle XAML-Elemente binden auf `{DynamicResource Brush.WindowBg}` statt hardcodierte Hex-Werte.
**Wichtig:** Bestehende hardcodierte Hex-Strings in `MainWindow.xaml` müssen durch `DynamicResource`-Referenzen ersetzt werden – das ist die größte Umbauarbeit dieses Features.

### 1D – Persistenz der Nutzereinstellungen

Nutzereinstellungen (Button-Sichtbarkeit, Theme) werden **nicht** in `settings.json` gespeichert (das ist eine Admin/Server-Konfiguration). Stattdessen: eigene Datei `user-prefs.json` im gleichen Verzeichnis wie `settings.json`.

Neues Model `Models/UserPrefs.cs`:
```csharp
public class UserPrefs
{
    public bool ShowBtnOrdner { get; set; } = true;
    public bool ShowBtnAkte   { get; set; } = true;
    public bool ShowBtnBu     { get; set; } = true;
    public bool ShowBtnEintrag{ get; set; } = true;
    public bool ShowBtnBi     { get; set; } = false;
    public bool ShowBtnEintragBi { get; set; } = false;
    public bool IsDarkTheme   { get; set; } = true;
}
```

Neuer Service `Services/UserPrefsService.cs` mit `Load()` und `Save(UserPrefs prefs)`.
Laden in `App.OnStartup()` analog zu `AppConfig`. In `MainWindow.cs` werden die Window-DependencyProperties beim Start mit den geladenen Werten initialisiert.

### 1E – Neue AppConfig-Felder

`Models/AppConfig.cs` braucht zwei neue optionale Felder:
```csharp
public string WordBiBookmarkName { get; set; } = "_Berufsintegration";
public string WordBiTableBookmarkName { get; set; } = "BI_BERUFSINTEGRATION_TABELLE";
```

Defaults werden in `App.OnStartup()` gesetzt falls leer (analog zu `WordBookmarkName`).

---

## Feature 2: Power-User Batch-Insert

### Ziel
Nutzer bereitet Einträge für alle Teilnehmer vor (tab-getrennter Text, eine Zeile pro TN).
Die App fügt diese sequenziell und automatisch in die jeweilige Word-Akte ein, überspringt Fehler und hält am Ende alle Akten offen zur manuellen Prüfung.

### UI-Konzept

Ein **ausklappbarer Bereich** am unteren Ende der Teilnehmerliste (nach dem letzten Abwesend-Eintrag), ähnlich dem bestehenden `CollapsedStrip`. Standardmäßig eingeklappt. Ein Strip „⚡ Batch-Eintrag" öffnet ihn.

**Inhalt des aufgeklappten Panels:**

```
┌─────────────────────────────────────────────────────┐
│  ⚡ Batch-Eintrag                           [▲]     │
├─────────────────────────────────────────────────────┤
│  Hinweis: Eine Zeile pro Teilnehmer (Reihenfolge    │
│  wie in der Liste). Tab-getrennt, 4 Spalten.        │
│                                                     │
│  ┌────────────────────────────────────────────┐     │
│  │ 05.03.2026  Bildung  Ziel erreicht  ...    │     │
│  │ 05.03.2026  Bildung  Weiter so      ...    │     │
│  │ ...                                        │     │
│  └────────────────────────────────────────────┘     │
│                                                     │
│  [ Batch ausführen ]   Fortschritt: —               │
└─────────────────────────────────────────────────────┘
```

Nach Ausführung: Eine kompakte Ergebnisliste im Panel (Name + Status: ✓ / ✗ Grund).

### Datenfluss

1. Nutzer fügt Text in die `BatchTextBox` ein.
2. Klick auf „Batch ausführen":
   - Text wird zeilenweise gesplittet → `List<string> rows` (Leerzeilen ignorieren).
   - Validierung: Anzahl Zeilen == Anzahl Teilnehmer (nur **anwesende** TN mit `CanInsertEntry == true`)? Falls nicht: Toast `Warning`, Abbruch.
   - Jede Zeile wird gegen 4 tab-getrennte Spalten geprüft. Falls ungültig: Toast `Error`, Abbruch.
3. Batch-Loop: sequenziell pro TN.
4. Am Ende: Alle Akten bleiben offen. Ergebnis-Log im Panel.

### Async-Ausführung (kritisch)

Word-COM-Operationen dauern 2–6 Sekunden pro Dokument. Der Batch **muss** asynchron laufen, damit die UI nicht einfriert. Da WPF-UI-Thread STA ist und Word-COM ebenfalls STA verlangt, gibt es zwei Optionen:

**Empfohlene Option – COM auf UI-Thread, UI-Updates via `await Task.Yield()`:**
```csharp
private async void ExecuteBatch_OnClick(object sender, RoutedEventArgs e)
{
    foreach (var (participant, row) in pairs)
    {
        BatchProgressText.Text = $"Verarbeite {participant.FullName}…";
        await Task.Yield(); // UI-Thread kurz freigeben

        try
        {
            _wordService.InsertTextRowToTable(docPath, _config.WordBookmarkName, row);
            // → Eintrag OK
        }
        catch (Exception ex)
        {
            // → Eintrag fehlgeschlagen, weitermachen
        }
    }
}
```

`Task.Yield()` gibt dem UI-Thread zwischen jedem TN kurz Luft für Repaints, ohne echtes Threading. Einfachste und robusteste Lösung für STA-COM.

### Neuer WordService-Method: `InsertTextRowToTable`

Das ist die **wichtigste technische Änderung**: statt Clipboard wird der Text direkt übergeben. (Clipboard bleibt für das "normale" Eintrag BU und Eintrag BI jedoch bestehen, diese änderung betrifft nur die Power-User-Funktion)

Signatur:
```csharp
public void InsertTextRowToTable(string docPath, string bookmarkName, string tabSeparatedRow)
```

Verhalten analog zu `InsertClipboardToTable`, aber:
- Kein Clipboard-Lesen.
- `tabSeparatedRow` wird gesplittet (`\t`) → Array mit genau 4 Werten erwartet.
- Werte werden direkt per COM in die Tabellenzellen geschrieben: `row.Cells[i].Range.Text = values[i]`.
- Falls die Zeile nicht genau 4 Werte hat → `InvalidOperationException("Ungültiges Zeilenformat")`.

### Fehlerbehandlung pro TN

Jeder Eintrag wird unabhängig behandelt. Fehlertypen und Verhalten:

| Fehler | Verhalten |
|--------|-----------|
| Akte gesperrt (`gesperrt` in Message) | Skip, Ergebnis-Log: „gesperrt" |
| Keine Verlaufsakte gefunden | Skip, Log: „keine Akte" |
| Bookmark nicht gefunden | Skip, Log: „Bookmark fehlt" |
| Ungültiges Zeilenformat | Skip, Log: „Zeilenformat ungültig" |
| Sonstige Exception | Skip, Log: `ex.Message` |

Wichtig: `try/catch` umschließt **jeden einzelnen** TN-Block. Eine Exception beim TN N bricht den Loop nicht ab.

### Ergebnis-Log im Panel

Nach Abschluss wird unterhalb der `BatchTextBox` eine kompakte Liste angezeigt:

```
✓ TN A       Eintrag eingefügt
✓ TN B       Eintrag eingefügt
✗ TN C         Akte gesperrt – bitte manuell prüfen
✓ TN D         Eintrag eingefügt
```

Grüne Farbe `#4CAF50` für ✓, rote `#D17878` für ✗.
Implementierung: `ObservableCollection<BatchResult>` als ItemsControl, `BatchResult { Name, IsSuccess, Message }`.

### Zuordnung Text → Teilnehmer

Die Zeilen werden **positionsbasiert** zugeordnet: Zeile 1 → erster TN mit `CanInsertEntry == true`, Zeile 2 → zweiter, usw. (in der Reihenfolge wie in der UI-Liste angezeigt: anwesende TN, alphabetisch).

Zur Sicherheit: Vor dem Loop wird dem Nutzer eine **Vorschau-Tabelle** angezeigt, die Zuordnung (Name ↔ Zeile) explizit zeigt, mit einem „Bestätigen"-Step. So kann der Nutzer Fehler erkennen bevor etwas eingefügt wird.


### Button-Sichtbarkeit im Batch-Panel

Das Batch-Panel bezieht sich immer auf `WordBookmarkName` (BU-Tabelle). Falls später auch ein BI-Batch gewünscht ist, kann ein zweiter Strip „⚡ Batch-Eintrag BI" hinzugefügt werden. Erstmal nur BU. In der Beschreibung der Power-User funktion soll das auch deutlich gemacht werden. Das Feature ist momentan für BU- und Mo- Einträge gedacht.

---

## Technische Zusammenfassung / Implementierungs-Checkliste

### Neue Dateien
- `Models/UserPrefs.cs` – Nutzereinstellungen-Model
- `Services/UserPrefsService.cs` – Load/Save für `user-prefs.json`
- `Models/BatchResult.cs` – Ergebnis-Model für Batch-Log

### Geänderte Dateien
- `Models/AppConfig.cs` – `WordBiBookmarkName`, `WordBiTableBookmarkName` hinzufügen
- `Models/Participant.cs` – `CanOpenAkteBi`, `CanInsertEntryBi` hinzufügen
- `Services/WordService.cs` – `InsertTextRowToTable(docPath, bookmarkName, tabRow)` hinzufügen
- `App.xaml` – Alle Farben auf `DynamicResource`-Keys umstellen; Theme-Dictionaries
- `App.xaml.cs` – `UserPrefs` laden
- `MainWindow.xaml` – Zahnrad-Button, Einstellungs-Overlay, BI-Buttons im Card-Template, Batch-Panel am Listenende, alle Hardcode-Farben → `DynamicResource`
- `MainWindow.xaml.cs` – DependencyProperties für Button-Sichtbarkeit, `ApplyTheme()`, Settings-Panel-Logik, Batch-Loop async, BI-Button-Handler, `UpdateActionState()` erweitern

### Reihenfolge
1. `AppConfig` + `UserPrefs` + `UserPrefsService` (Models/Services zuerst, kein UI-Impact)
2. `Participant`-Properties + `WordService.InsertTextRowToTable` (kein UI-Impact)
3. Theme-System (App.xaml + DynamicResource-Umbau – größter Aufwand, isolierbar)
4. Settings-Panel UI + Logik (Button-Visibility, Serverpfad, Theme-Toggle)
5. BI-Buttons im Card-Template + Handler
6. Batch-Insert-Panel + Logik

### Nicht ändern
- `Services/ParticipantParser.cs` – kein Bedarf
- `Services/FolderMatcher.cs` – kein Bedarf
- `Services/AppLogger.cs` – kein Bedarf
- `App.xaml.cs` Startup-Fehler-MessageBoxen – bleiben als `MessageBox.Show` (fatale Fehler vor UI)

---

## Offene Fragen / Entscheidungen

1. **Batch-Zuordnung:** Positionsbasiert (einfach, fehleranfällig bei Reihenfolge) oder namensbasiert (robuster, leicht mehr Parsing)?
2. **Einstellungs-Overlay:** Inline im Hauptfenster (kein Fokus-Verlust, komplexer z-Order) oder separates `Window` mit `ShowDialog()` (einfacher, aber out-of-theme)?
3. **Theme-Umbau:** Vollständig auf `DynamicResource` (sauber, aber aufwändig) oder pragmatisch nur die prominenten Farben (schneller, aber inkonsistent)?
4. **Serverpfad im UI:** Darf der Nutzer ihn überschreiben und persistieren (überschreibt `settings.json`)? Oder besser in `user-prefs.json` mit Priorität über `settings.json`?

