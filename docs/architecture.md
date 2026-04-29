# Architecture

## Ueberblick
Scola ist eine WPF-Anwendung ohne separates MVVM-Framework. Die Architektur ist pragmatisch: ein starker Haupt-Controller im MainWindow, dazu mehrere spezialisierte Services und einfache Datenmodelle.

## Einstiegspunkte
- `App.xaml.cs`
  - App-Startup, lokale Pfade, Konfiguration, Prefs, Theme, Migration
- `MainWindow.xaml.cs`
  - Hauptorchestrierung der UI und fast aller Workflows
- `SettingsWindow.xaml.cs`
  - Erfassung und Rueckgabe benutzerspezifischer Einstellungen
- `AGENTS.md`
  - Arbeitsanweisung fuer neue KI-Instanzen und kleine, sichere Aenderungen

## Hauptmodule
### UI
- `MainWindow.xaml(.cs)`
  - Input, Ergebnisliste, TN-Karten, Aktionen, Batch, Toasts, Statusleiste, Mini-Stundenplan-Tray
- `SettingsWindow.xaml(.cs)`
  - Pfade, Anzeige, Buttons und Debug
- `AppAlertWindow.xaml(.cs)`
  - gestylte Einzelwarnungen/-fehler im App-Stil
- `BatchFailureSummaryWindow.xaml(.cs)`
  - Batch-Abschlussmeldung mit Fehlerliste
- `BiTodoSummaryWindow.xaml(.cs)`
  - Abschlussfenster fuer den BI-To-do-Sammellauf mit Status pro Teilnehmer

### Modelle
- `Models/Participant.cs`
  - zentrales Teilnehmermodell inklusive Match-, Dokument- und Mini-Stundenplan-Kontext
- `Models/ParticipantMiniScheduleModels.cs`
  - Mini-Stundenplan-Zustaende, Zellen und Summaries fuer den Tray
- `Models/ParsedEntry.cs`
  - Parser-Ergebnis pro Zeile
- `Models/AppConfig.cs`
  - globale Konfiguration
- `Models/UserPrefs.cs`
  - benutzerspezifische Einstellungen
- `Models/BatchResult.cs`
  - Ergebnis pro Batch-Zeile
- `Models/BiTodoCollectResult.cs`
  - Anfrage-/Ergebnis-Modelle fuer den BI-To-do-Sammellauf
- `Models/HeaderMetadata*.cs`
  - Odoo-Header-Cache

### Services
- `Services/ParticipantParser.cs`
  - Rohtext-/Listen-Parsing, Statuslogik, Freitext-Trennung
- `Services/FolderMatcher.cs`
  - Ordnerindex, Token-Matching, Fallbacks, DocumentPath/Initials-Aufloesung
- `Services/InitialsResolver.cs`
  - Kuersel aus Dateiname
- `Services/DocxHeaderMetadataService.cs`
  - Odoo-Link aus `docx`-Headern, read-only, ZIP/XML, Cache
- `Services/WeeklyScheduleService.cs`
  - Wochenplan-Datei aufloesen, `docx` auf Paragraph-/Run-Ebene lesen, XHub-inspirierte Alias-/Ambiguitaetslogik anwenden und den Mini-Stundenplan pro TN aufbauen
  - farbliche Statuslogik wird konservativ gelesen: `disp` nur bei echter roter Markierung, nicht bei bloss roter Schrift; `ext` bleibt ueber gruene Markierung
- `Services/WordService.cs`
  - COM-Zugriff auf Word, Dokumentoeffnen, Bookmarks, Tabellen-Eintrag, Lock- und Leak-Handling
- `Services/WordStaHost.cs`
  - serieller STA-Host fuer alle Word-Operationen; trennt UI-Thread und COM-Laufzeit
- `Services/UserPrefsService.cs`
  - Laden/Speichern von UserPrefs
- `Services/AppLogger.cs`
  - Logschreiben und Debug-Logging
- `Services/JsonStorage.cs`
  - atomare JSON-Speicherung / Backup-Muster

## Hauptdatenfluesse
### 1. Listen-Input -> Teilnehmerkarten
1. Nutzer fuegt Rohtext ein
2. `ParticipantParser` erzeugt `ParsedEntry`-Ergebnisse
3. daraus werden `Participant`-Objekte gebildet
4. `FolderMatcher` baut/benutzt Index und matcht Ordner + Dokumentpfad
5. `MainWindow` aktualisiert den Action-State je Teilnehmer

### 2. Teilnehmeraktion -> Word / Explorer / Browser
1. Nutzer klickt auf Aktion in Teilnehmerkarte
2. `MainWindow` prueft Voraussetzungen
3. je nach Aktion:
   - Explorer
   - Browser
   - `_wordStaHost` -> `WordService`
4. Rueckmeldung per Toast oder gestyltem Dialog

### 3. Batch-Flow
1. Nutzer fuegt Batch-Text ein
2. Zeilen werden normalisiert
3. positionsbasierte Zuordnung zu aktiven Teilnehmenden
4. Mapping-Bestaetigung
5. pro TN `WordService.InsertTextRowToTable(...)`
6. `BatchResults` + Abschlussdialog bei Fehlern

### 4. BI: To-dos Sammellauf
1. Nutzer oeffnet `BI: To-dos ⚡︎`
2. angehakte Teilnehmer werden vorgeprueft und Dokumentpfade aufgeloest
3. ein dedizierter STA-Worker startet eine versteckte Word-Instanz
4. pro TN wird die Tabelle am Bookmark `BI_BERUFSINTEGRATION_TODO` read-only aus der BI-Akte gelesen
5. ein neues ungespeichertes Sammeldokument wird aufgebaut:
   - Titel `BI dd.MM.yyyy`
   - Name + Kuersel pro TN
   - komplette Tabelle
   - sichtbare Trennlinie
6. am Schluss wird nur das Ergebnisdokument sichtbar geoeffnet
7. ein app-gestyltes Abschlussfenster zeigt den Status pro TN

### 5. Odoo-Metadaten
1. Odoo-Button / Warmup loest Metadatenbedarf aus
2. `DocxHeaderMetadataService` liest `docx` read-only
3. Ergebnis wird dokumentbasiert gecacht
4. UI konsumiert `OdooUrl`

### 6. Word-Fensterverhalten
1. `WordService` setzt fuer Word nur Sichtbarkeit und Fokus
2. minimiertes Word darf fuer den Fokus wiederhergestellt werden
3. Position und Groesse der Word-Fenster ueberlaesst Scola vollstaendig Word

### 7. Mini-Stundenplan-Tray
1. `SettingsWindow` liefert optional einen `ScheduleRootPath`
2. Nutzer oeffnet den Tray ueber Klick auf den Namen oder Doppelklick auf die Karte
3. `MainWindow` stellt sicher, dass immer nur ein Tray gleichzeitig geoeffnet ist
4. `WeeklyScheduleService` liest die aktuelle Wochenplan-`docx` per ZIP/XML und wertet Inhalte auf Paragraph-/Run-Ebene aus
5. Ein XHub-inspirierter Alias-Matcher loest Zeilen konservativ gegen die sichtbaren TN auf; Ambiguitaet fuehrt zu `Unavailable` statt zu aggressivem Raten
6. Ergebnis wird pro Datei persistent gecacht und als `ParticipantMiniScheduleSummary` an die Karte gebunden
7. Der Tray verwendet ein kompaktes 5x2-Raster mit Headerzeile, Lunch-Seperator und Status-Badges `disp` / `ext` im XHub-Stil

## Persistenzmodell
### Im Repo
- `settings.json` als projektseitige Ausgangskonfiguration
- Testdaten und Projekt-Doku

### In `%LOCALAPPDATA%\AkteX`
- `settings.json`
- `user-prefs.json`
- `header-metadata-cache.json`
- `header-metadata-cache.bak`
- `logs\...`

## Aktuelle technische Entscheidungen
- Kein MVVM-Framework
- Word via late-bound COM (`dynamic`)
- lokale JSON-Dateien statt Datenbank
- Odoo nur indirekt ueber Link aus Akten
- Mini-Stundenplan read-only direkt aus `docx`, ohne Word
- Auswertung der Eingabeliste laeuft mit sichtbarem Ladezustand und blockiert den UI-Thread nicht mehr vollstaendig
- Word-Fenstersteuerung jetzt bewusst einfach:
  - nur sichtbar machen und fokussieren
  - keine aktive Platzierungslogik

## Architektur-Stolpersteine
- `MainWindow.xaml.cs` ist gross und fachlich dicht.
- UI, Workflow und Produktlogik sind dort eng verwoben.
- `WordService` ist kritisch: kleine Aenderungen koennen COM-Nebenwirkungen erzeugen.
- Historische Debug-/Transfer-Dokumente enthalten teils Zwischenstaende.
- `settings.json` existiert sowohl als projektseitige Startdatei als auch als Laufzeitkopie in LocalAppData; bei Analysen immer beide Ebenen unterscheiden.

## Annahmen
- Diese Architektur bleibt kurzfristig bestehen; keine grosse Entkopplung ist aktuell beabsichtigt.
- Die App wird weiterhin lokal auf Windows 11 mit installiertem Word eingesetzt.
