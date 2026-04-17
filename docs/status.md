# Status

Stand dieser Zusammenfassung: 2026-03-17

## Aktueller Funktionsstand
### Umgesetzt
- Rohtext-/Listen-Auswertung zu Teilnehmerkarten
- Checkbox-basierte Praesenzsteuerung pro TN nach Auswertung
- Ordner-Matching ueber 1-3 Pfade
- Ordner / Akte / BU / BI / BE / Eintrag BU / Eintrag BI / Odoo als Aktionen
- Odoo-Link aus `docx`-Header mit Cache
- Odoo-Headercache speichert nur erfolgreiche Leseergebnisse; die Cache-Version wurde angehoben, damit potenziell falsch leere Alt-Einträge neu aufgebaut werden
- Kuersel aus Dateinamen, Anzeige unter dem Namen
- Kuersel koennen in Darstellung/Buttons konfiguriert werden und sind in die Match-/Suchbasis eingearbeitet
- Mini-Stundenplan-Tray pro TN-Karte mit Toggle ueber Name oder Karte
- Stundenplanquelle als separater `ScheduleRootPath` in den Einstellungen
- konservatives Wochenplan-Parsen direkt aus `docx` ohne Word
- Wochenplan-Matcher jetzt deutlich naeher an XHub: Alias-/Ambiguitaetslogik, Paragraph-/Run-Parsen und interner Cache/Diagnostik-Unterbau
- Mini-Stundenplan-Layout jetzt bewusst naeher an XHub: kompaktes Raster mit Headerzeile, Lunch-Seperator und Status-Badges fuer `disp` / `ext`
- Stundenplan-Fehlversuche werden nicht mehr als leerer Wochenplan gecacht; die aktuelle KW wird bei temporaeren Locks im Hintergrund erneut versucht
- `Auswerten` zeigt waehrend laengerer Listenverarbeitung einen sichtbaren Ladezustand mit kleinem Progress-Indikator
- Batch-Eintrag mit positionsbasierter Zuordnung und Mapping-Bestaetigung
- `BI: To-dos` als eigener Sammellauf fuer angehakte TN
- unsichtbares Lesen der BI-To-do-Tabelle per Word-Instanz ohne sichtbares Oeffnen der Quellakten
- neues ungespeichertes Sammeldokument mit Titel, Name, Kuersel und kompletter To-do-Tabelle pro TN
- BI-To-do-Ergebnisse werden nach dem versteckten Aufbau per temporaerer `docx` in eine normale sichtbare Word-Instanz uebergeben; die versteckte BI-Instanz wird danach aktiv beendet
- app-gestyltes Abschlussfenster mit Status pro TN fuer den BI-To-do-Lauf
- gestylte App-Dialoge fuer wichtige Warnungen/Fehler
- Batch-Fehler-Zusammenfassung im App-Stil
- Debug-Logging per Toggle
- Word-Leak-/Lock-Absicherungen
- Hauptfenster startet deterministisch zentriert auf dem primaeren Monitor; nur Groesse und Maximiert-Zustand werden noch gemerkt
- Teilnehmerkarten bleiben immer in einer einzigen Wide-Struktur; Buttons rutschen nicht mehr unter die Namen und die Fensterbreite hat dafuer eine dynamische Mindestbreite
- Word-Fenster werden nur noch geöffnet/fokussiert; Größe und Position verwaltet Word selbst
- die alten Word-Placement-Optionen sind aus den Einstellungen entfernt; Scola bietet dafuer bewusst keine eigene Monitor-/Maximierungssteuerung mehr an
- Auto-Updater ueber GitHub Releases aus `subsudo/scola_public`
- Mock-Testsetup
- Git/GitHub-Grundlage mit `README.md`, `docs/`, `AGENTS.md` und `.gitignore`

### Produktiv relevante sensible Bereiche
- `Services/WordService.cs`
- `Services/ParticipantParser.cs`
- `Services/FolderMatcher.cs`
- `MainWindow.xaml.cs`

## Bekannte offene Themen
### Weiter beobachten
- echte Produktionslogs fuer Parser-/Matching-Ausreisser
- echte Produktionslogs fuer Word-Leerdokumente / Nebenwirkungen
- Ghost-Word-Diagnostik fuer leere `DokumentN`-Fenster bleibt produktiv relevant, jetzt mit expliziter Zombie-Attach-Abwehr und konservativer Blankodokument-Bereinigung
- Verhalten von Word bei mehreren offenen Instanzen/Fenstern
- vereinfachtes Hauptfenster-Startverhalten auf verschiedenen Monitor-Setups im Alltag weiter beobachten
- neue MinWidth-Logik der Teilnehmerkarten im Alltag weiter beobachten, vor allem bei sehr langen Namen und wechselnden sichtbaren Buttons
- Odoo-Metadaten-Warmup nur unter realer Netzlast weiter beobachten
- Mini-Stundenplan-Matching mit echten Wochenplaenen weiter beobachten, vor allem bei haeufigen Vornamen-Clustern und Grenzfaellen wie `NameA`/`NameB`
- Wochenwechsel- und Lock-Verhalten des Mini-Stundenplans im Alltag weiter beobachten, vor allem Montag/Dienstag bei frisch geoeffneten `KW_xx.docx`
- Mini-Stundenplan-Feintuning bleibt produktnah: Breite, Linien und Badge-Proportionen wurden zuletzt mehrfach iteriert und sollten nur vorsichtig weiter angepasst werden

### Dokumentationsschuld
- `HANDOVER.md` ist wertvoll, beschreibt aber teils aeltere Zwischenstaende.
- `BUGFIXES.md` und `MUST_DEBUG.md` bleiben wichtig, sind aber problem- bzw. historienorientiert.
- Der kuenftige Hauptkontext soll in `docs/` gepflegt werden.
- `AGENTS.md` und `docs/` sollen kuenftig zusammen den Einstiegsrahmen fuer weitere KI-Arbeit bilden.
- operative Build-/Publish-/Release-Regeln stehen in `docs/release-workflow.md`

### Technische Risiken
- hohe Logikdichte in `MainWindow.xaml.cs`
- COM-/Word-Verhalten ist umgebungsabhaengig
- mehrere Konfigurationsebenen (`settings.json` im Repo vs. Laufzeitdaten in LocalAppData) koennen bei Fehleranalysen verwechselt werden

## Teststand
### Lokal gut abdeckbar
- Mock-Umgebung unter `TestSetup`
- Build und Run lokal
- UI-Flows und Batch-Basis testbar
- Mini-Stundenplan-Tray lokal mit einer passenden Wochenplan-`docx` testbar
- Git-/Dokumentationskontext jetzt im Repo nachvollziehbar

### Nur im echten Umfeld belastbar pruefbar
- Netzlaufwerkverhalten
- Locks durch andere Nutzer
- echte Word-/Office-Eigenheiten
- reale Rohtext-Importe aus dem Produktivsystem

## Empfehlungen fuer naechste Aenderungen
1. erst Doku lesen, dann Code aendern
2. bei Word- oder Parser-Aenderungen immer Log-Sicht mitdenken
3. kleine Aenderungen bevorzugen
4. bei UI-Aenderungen bestehende visuelle Sprache erhalten
5. bei Verhaltensaenderungen immer `docs/status.md` und bei Entscheidungen auch `docs/decisions.md` aktualisieren

## Annahmen
- Aktuell ist kein grosser Umbau geplant.
- Produktiver Nutzen und sichere Weiterarbeit stehen ueber Architekturkosmetik.

