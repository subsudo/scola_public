# Scola

Scola ist eine lokale WPF-Desktop-App fuer den operativen Alltag mit Teilnehmer-Akten.

Hauptzweck:
- Listen oder Rohtext mit Teilnehmenden einlesen
- Teilnehmende gegen Aktenordner auf Netzlaufwerken matchen
- Word-Verlaufsakten oeffnen
- zu BU / BI / BE springen
- BU- und BI-Eintraege in Word-Tabellen einfuegen
- positionsbasierte Batch-Eintraege fuer aktive Teilnehmende ausfuehren
- BI-To-do-Tabellen fuer angehakte TN in ein gemeinsames Word-Sammeldokument ziehen
- zusaetzlich Kuersel und Odoo-Links aus Akten ableiten
- optional einen Mini-Stundenplan pro TN direkt in der Karte anzeigen

## Status
Dieses Repository ist aktiv in Entwicklung. Die App ist produktnah und bereits fuer reale Arbeitsablaeufe optimiert, aber nicht als generisches Framework gedacht.

## Einstieg
Zuerst diese Dateien lesen:
1. `docs/project-spec.md`
2. `docs/architecture.md`
3. `docs/status.md`
4. `docs/decisions.md`
5. `docs/release-workflow.md`
6. `AGENTS.md`

Danach bei Bedarf die aelteren Arbeitsdokumente:
- `HANDOVER.md`
- `BUGFIXES.md`
- `MUST_DEBUG.md`
- `TESTING.md`
- `FEATURE_REQUESTS.md`

Wichtig:
- Der aktuelle Projektkontext soll kuenftig primaer in `docs/` gepflegt werden.
- Die bestehenden Handover-/Debug-/Bugfix-Dateien bleiben als Referenz erhalten und werden nicht geloescht.
- `AGENTS.md` beschreibt den bevorzugten Arbeitsstil fuer weitere KI- oder Teamarbeit.
- Fuer Build-/Publish-/Release-Arbeit ist `docs/release-workflow.md` verbindlich.

## Projektstruktur
- `AGENTS.md`
  - kurze Arbeitsanweisung fuer weitere KI-Instanzen
- `App.xaml(.cs)`
  - Startup, lokale Pfade, Theme, Settings-/Prefs-Laden
- `MainWindow.xaml(.cs)`
  - Haupt-UI und Hauptworkflows
- `SettingsWindow.xaml(.cs)`
  - benutzerspezifische Einstellungen inklusive Stundenplanquelle
- `Services/`
  - Parser, Matching, Word, Logging, UserPrefs, Odoo-Header-Cache, Mini-Stundenplan
- `Models/`
  - Konfiguration, Teilnehmer, Parse-/Batch-/Cache-Modelle

Der Mini-Stundenplan ist aktuell fachlich und visuell naeher an `XHub` ausgerichtet:
- konservatives Matching direkt aus Wochenplan-`docx`
- Status `disp` und `ext` aus Format-/Run-Informationen
- kompaktes Raster im XHub-Stil direkt im Teilnehmer-Tray
- `TestSetup/`
  - lokales Mock-Testsetting
- `Assets/`
  - App-Icons

## Lokale Persistenz
Zur Laufzeit werden Daten unter `%LOCALAPPDATA%\AkteX\` gespeichert:
- `settings.json`
- `user-prefs.json`
- `header-metadata-cache.json`
- `header-metadata-cache.bak`
- `logs\app-YYYY-MM-DD.log`

Im Repo selbst liegen nur:
- Code
- Dokumentation
- Mock-/Testdaten
- projektseitige Startkonfiguration

## Build und Start
```powershell
dotnet build .\VerlaufsakteApp.csproj
dotnet run --project .\VerlaufsakteApp.csproj
```

## Testen mit Mock-Setup
Siehe `TESTING.md` sowie `TestSetup\New-MockTestSetup.ps1`.

## Wichtige Hinweise fuer Weiterarbeit
- Nicht zuerst neu designen; erst verstehen, dann gezielt aendern.
- `MainWindow.xaml.cs` und `Services/WordService.cs` sind die risikoreichsten Dateien.
- Word-Integration ist bewusst pragmatisch und COM-basiert.
- Robuste Alltagsfunktion ist wichtiger als Architektur-Reinheit.
- Lokale, kontrollierbare Loesungen werden bevorzugt.
- Relevante Doku-Aenderungen immer in `docs/` und bei Bedarf auch in `AGENTS.md` nachziehen.

## Annahmen
- Dieses Repo bildet das aktive Arbeitsprojekt `AkteX` ab.
- Der aktuelle sichtbare App-Name ist `Scola`.
- Technische Altbezeichnungen wie `%LOCALAPPDATA%\\AkteX` bleiben vorerst bewusst bestehen.
- `TestSetup/` ist als Teil des Projektkontexts gewollt und kein reines Wegwerfmaterial.
- Produktive Benutzerdaten liegen nicht im Repo, sondern in LocalAppData.
