# Decisions

Diese Datei haelt aktuelle, bewusste Projektentscheidungen fest. Sie ist keine vollstaendige Historie, sondern eine kompakte Arbeitsgrundlage fuer Weiterarbeit.

## D1 - Keine grosse Architektur-Migration
Scola bleibt vorerst eine pragmatische WPF-App ohne MVVM-Framework.

Begruendung:
- Produkt ist bereits funktional gewachsen
- Alltagsfunktion ist wichtiger als Architektur-Reinheit
- grosse Umbauten wuerden viel Risiko in stabile Arbeitsablaeufe bringen

## D2 - Lokale Persistenz in LocalAppData
Benutzerdaten und Laufzeitkonfiguration liegen nicht neben der EXE, sondern in `%LOCALAPPDATA%\AkteX`.

Begruendung:
- sauberer fuer portable Auslieferung
- keine sichtbaren Zusatzdateien neben der EXE noetig
- benutzerspezifische Trennung bleibt erhalten

## D3 - Word bleibt COM-basiert
Die Word-Integration bleibt late-bound ueber `dynamic` und wird nicht auf eine voellig andere Strategie umgestellt.

Begruendung:
- bestehende Funktionalitaet haengt an Word-Bookmarks und Tabellenstrukturen
- die App ist bewusst auf reale Word-Akten zugeschnitten

## D4 - Odoo nur indirekt und read-only
Odoo wird in Scola aktuell nur ueber einen in der Akte hinterlegten Link genutzt.

Begruendung:
- einfache, kontrollierbare Integration
- keine harte Laufzeitabhaengigkeit zu Odoo
- Metadaten koennen lokal gecacht werden

## D5 - Kuersel frueh aus dem Dokumentnamen
Kuersel werden nicht aus dem Ordnernamen oder dem Word-Inhalt abgeleitet, sondern frueh aus dem Dateinamen der Akte.

Begruendung:
- billig
- stabil
- sofort suchbar / nutzbar

## D6 - Batch ist positionsbasiert
Batch-Zuordnung erfolgt aktuell bewusst nach Position, mit vorgeschalteter Zuordnungsbestaetigung.

Aktuelle Regel:
- BU-Batch und BI-Batch verwenden dieselbe Eingabe-, Validierungs-, Zuordnungs- und Fehlerlogik
- der Unterschied liegt nur im Ziel-Bookmark der Word-Tabelle
- beide Batch-Arten teilen sich einen gemeinsamen Laufzustand, damit sie nicht parallel gegeneinander arbeiten

Begruendung:
- einfach
- nachvollziehbar
- in der Praxis ausreichend fuer den aktuellen Use Case

## D7 - Wichtige Fehler im App-Stil, nicht nur als Toast
Kritische oder handlungsrelevante Meldungen werden in gestylten Scola-Fenstern angezeigt.

Begruendung:
- bessere Sichtbarkeit
- konsistenteres UX-Verhalten
- besser geeignet fuer Locks, fehlende Bookmarks oder Batch-Fehler

## D8 - Word-Fensterlogik bewusst vereinfacht
Scola platziert Word-Fenster nicht mehr aktiv, sondern oeffnet und fokussiert Dokumente nur noch.

Aktuelle Regel:
- Scola setzt Sichtbarkeit und Fokus fuer geoeffnete Akten
- die eigentliche Fensterposition und -groesse bleibt bei Word
- minimiertes Word darf fuer den Fokus wiederhergestellt werden
- nicht minimiertes Word wird nicht mehr aktiv verschoben, skaliert oder auf einen Monitor gezwungen
- es gibt bewusst keine benutzerseitigen Einstellungen mehr fuer Word-Maximierung oder Zielmonitor

Begruendung:
- aktive COM-Fensterplatzierung war in der Praxis zu unzuverlaessig
- Word merkt sich eigene Fensterzustaende nicht konsistent genug fuer erzwungene Nachsteuerung durch Scola
- weniger aktive Fenstersteuerung ist alltagsrobuster

## D9 - Bestehende Alt-Dokumente bleiben erhalten
Historische Projektdateien wie `HANDOVER.md`, `BUGFIXES.md`, `MUST_DEBUG.md`, `TESTING.md` bleiben bestehen.

Begruendung:
- dort steckt wertvoller Verlauf und Problemkontext
- sie sollen nicht geloescht werden
- kuenftiger Hauptkontext liegt aber primaer in `docs/`

## D10 - Git-/KI-Kontext liegt im Repo
Projektkontext fuer Weiterarbeit soll nicht primaer in Chats liegen, sondern im Repository.

Begruendung:
- Arbeit erfolgt auf mehreren Computern
- neue KI-Instanzen sollen direkt im Repo einsteigen koennen
- `README.md`, `docs/` und `AGENTS.md` bilden zusammen den aktuellen Einstieg

## D11 - `.gitignore` bleibt strikt lokal-orientiert
Build-, Publish- und lokale Tooling-Artefakte gehoeren nicht ins Repo.

Begruendung:
- Repo soll review- und KI-tauglich bleiben
- produktive Laufzeitdaten liegen ohnehin in `%LOCALAPPDATA%\AkteX`
- Build-Outputs sollen nicht den eigentlichen Projektkontext ueberdecken

## D12 - Mini-Stundenplan konservativ und read-only
Der Mini-Stundenplan wird direkt aus einer Wochenplan-`docx` gelesen und lieber leer als falsch angezeigt.

Aktuelle Regel:
- eigener `ScheduleRootPath` in `AppConfig`
- Parsing per ZIP/XML, nicht ueber Word
- nur aktuelle Kalenderwoche
- strenges Alias-Matching aus den aktuell sichtbaren TN
- bei Unsicherheit `Unavailable` statt aggressivem Raten

Begruendung:
- Wochenplan soll schnell im Alltag helfen, aber keine falsche Sicherheit erzeugen
- Word soll dafuer nicht sichtbar gestartet werden
- das Feature bleibt additive UI-Hilfe und kein kritischer Kernworkflow

## D13 - Mini-Stundenplan orientiert sich visuell an XHub
Das Layout des Mini-Stundenplans in Scola orientiert sich bewusst am kompakten XHub-Raster statt an frei variierenden Kachelmustern.

Aktuelle Regel:
- feste Headerzeile fuer `Mo` bis `Fr`
- feste Vormittags- und Nachmittagsreihe
- schmaler Lunch-Seperator
- Hauptgruppe oben zentriert
- Lehrer und Raum in einer kleinen Mittelzeile
- `disp` und `ext` als Status-Badges statt normaler Zellinhalte

Begruendung:
- das Layout ist in XHub bereits auf Lesbarkeit fuer sehr kleine Zellen optimiert
- Status-Badges sind klarer als Mischzustaende aus Text und Farben
- die visuelle Struktur soll stabil bleiben, auch wenn die Tray-Breite leicht iteriert wird

## D14 - Sichtbarer Ladezustand bei Auswertung
Die Auswertung einer eingefuegten Liste zeigt einen sichtbaren Ladezustand, statt fuer mehrere Sekunden scheinbar gar nichts zu tun.

Aktuelle Regel:
- `Auswerten` startet die schwere Listenverarbeitung im Hintergrund
- waehrenddessen zeigt die UI einen kleinen indeterminierten Progress-Indikator und Statustext
- das Eingabefeld wird waehrend der Verarbeitung nicht weiterbearbeitet

Begruendung:
- vermeidet Unsicherheit bei laengeren Netzlaufwerk-/Index-Latenzen
- passt besser zum alltagsorientierten Charakter der App
- ist bewusst leichtgewichtig und kein grosses Blocking-Overlay

## D15 - Hauptfenster merkt keine exakte Position mehr
Scola speichert fuer das Hauptfenster keine exakten `Left`/`Top`-Koordinaten mehr, sondern nur noch Groesse sowie die relevanten Fensterzustaende.

Aktuelle Regel:
- beim Start wird das Hauptfenster immer zentriert auf dem primaeren Monitor geoeffnet
- `WindowWidth`, die expandierte Fensterhoehe, `WindowWasMaximized` und `IsCollapsed` bleiben erhalten
- exakte Fensterposition und gespeicherte Monitor-DeviceNames werden nicht mehr persistiert
- auf `DisplaySettingsChanged` wird nicht mehr aktiv reagiert

Begruendung:
- eliminiert die Off-Screen-Bugklasse strukturell statt heuristisch
- vermeidet Probleme mit wechselnden Windows-DeviceNames und Monitor-/DPI-Kombinationen
- ist fuer den Alltag vorhersehbarer als eine komplexe Restore-Heuristik

## D16 - Release-Prozess bleibt bewusst manuell
Build, Publish und GitHub-Release fuer Scola werden bewusst nicht bei normalen Commits oder Pushes automatisch ausgelost.

Aktuelle Regel:
- normale Code-Aenderung = keine neue Version, keine Release
- Test-EXE darf gebaut werden, ohne direkt veroeffentlicht zu werden

## D17 - BI-Hidden-Word wird nicht an Nutzer weitergereicht
Der `BI: To-dos`-Sammellauf darf keine versteckte Automations-Word-Instanz fuer spaetere normale Word-Aktionen hinterlassen.

Aktuelle Regel:
- BI-Inhalte duerfen intern in einer dedizierten versteckten Word-Instanz aufgebaut werden
- das sichtbare BI-Ergebnis wird danach per temporaerer `docx` in eine normale user-facing Word-Instanz uebergeben
- die versteckte BI-Instanz wird anschliessend immer aktiv beendet
- Attach an versteckte/zombiehafte Word-Instanzen wird verworfen
- leere unspeicherte Blankodokumente duerfen konservativ aufgeraeumt werden, wenn sie nicht das aktive oder das Ziel-Dokument sind

Begruendung:
- offene Ghost-Dokumente wie `Dokument25` / `Dokument58` entstanden aus ueberlebenden Hidden-BI-Instanzen
- ein sichtbares BI-Ergebnis soll fuer den Nutzer normal benutzbar bleiben, aber nicht an einer versteckten Automationsinstanz haengen
- konservative Ghost-Bereinigung ist alltagsrobuster als weitere Diagnose ohne Lifecycle-Fix
- echte Release nur auf ausdruecklichen Wunsch
- operative Details stehen in `docs/release-workflow.md`

Begruendung:
- reduziert versehentliche Veroeffentlichungen
- passt zum produktnahen, kontrollierten Arbeitsstil des Projekts
- ist fuer Mensch und KI leichter verlässlich einzuhalten als ein halbautomatischer Mischprozess

## D19 - Teilnehmerkarten bleiben immer Wide, Fensterbreite nur mit Untergrenze
Scola verwendet fuer Teilnehmerkarten keine Narrow-Fallback-Struktur mit Buttons unter dem Namen mehr.

Aktuelle Regel:
- pro Teilnehmerkarte gibt es nur noch eine Button-Zeile rechts neben dem Namen
- die Abwesenheitsinfo wird in der unteren Sekundaerzeile der Karte dargestellt, nicht mehr in der Kopfzeile
- die Fensterbreite bekommt eine dynamische `MinWidth` aus laengstem sichtbaren Namen plus sichtbaren Buttons
- die `MinWidth` enthaelt eine kleine Sicherheitsreserve, damit Button-Hintergruende lange Namen nicht optisch anschneiden
- Scola vergroessert die effektive Mindestbreite bei Bedarf, verkleinert die Fensterbreite aber nie automatisch
- eine vom User bewusst groesser gewaehlte Fensterbreite bleibt erhalten
- `AutoFit` ist entfernt; Doppelklick in der Titelleiste und der neue Header-Button toggeln stattdessen den Header-only-Collapse-Modus
- Kürzel unter dem Namen bleiben dauerhaft sichtbar und sind nicht mehr als Ansichtseinstellung abschaltbar

Begruendung:
- verhindert strukturell, dass Buttons unter die Namen rutschen
- ist alltagsrobuster und leichter vorhersehbar als ein Layout-Switch zwischen Wide und Narrow
- respektiert die Nutzerpraeferenz fuer eine bewusst breitere Fensteransicht

## D20 - Header-only-Collapse statt eigenem Maximieren
Scola bietet in der Titelleiste bewusst keinen eigenen Maximieren-Button mehr an, sondern einen Collapse-Toggle fuer einen kompakten Header-only-Modus.

Aktuelle Regel:
- der Collapse-Toggle klappt den kompletten App-Koerper weich auf die reine Header-Leiste ein und wieder aus
- Windows-Minimize bleibt unveraendert und wird nicht durch Collapse ersetzt
- ein eingeklapptes Fenster darf sich bei spaeterer normaler Aktivierung wieder automatisch ausklappen, damit es auf Multi-Monitor-Setups leichter wiedergefunden wird
- `TopMost` wird dafuer bewusst nicht eingefuehrt

Begruendung:
- der praktische Bedarf ist eher `aus dem Weg, aber sofort wieder da` als klassisches Maximieren
- ein Header-only-Modus passt besser zum Werkzeug-Charakter von Scola
- ohne `TopMost` bleibt das Verhalten alltagstauglich und drängt sich nicht ueber andere Arbeitsfenster

## D17 - Wochenplan-Fehlversuche gelten nicht als gueltiger leerer Stand
Temporäre Lesefehler beim Mini-Stundenplan, zum Beispiel gesperrte `KW_xx.docx` am Wochenwechsel, werden nicht als leerer Wochenplan persistiert.

Aktuelle Regel:
- nur erfolgreich gelesene Wochenplan-Dokumente werden im JSON-Cache gespeichert
- ein Fehlversuch darf einen bestehenden guten Cache nicht mit einem leeren Dokument ueberschreiben
- solange die aktuelle Woche noch nicht einmal erfolgreich gelesen wurde, versucht Scola sie im Hintergrund periodisch erneut

Begruendung:
- Wochenplaene sind montags oft kurz gesperrt und spaeter normal lesbar
- ein einmaliger Lock soll nicht fuer Stunden oder Tage zu `Kein Stundenplan` fuehren
- das Feature bleibt konservativ, aber robuster im realen Alltag

## D18 - Odoo-Header-Fehlversuche gelten nicht als leerer Link
Temporäre Lesefehler beim Odoo-Link aus dem `docx`-Header werden nicht als leerer Odoo-Zustand persistiert.

Aktuelle Regel:
- nur erfolgreich gelesene Header-Metadaten werden im persistenten Header-Cache gespeichert
- ein Lesefehler darf keinen bestehenden guten Odoo-Link mit `leer` überschreiben
- die Cache-Version wurde bewusst erhöht, damit alte potenziell falsch leere Einträge neu aufgebaut werden

Begruendung:
- der Odoo-Link ist für eine bestehende Akte fachlich meist stabil
- temporäre Paket-/Netz-/Lock-Fehler sollen nicht zu dauerhaft „fehlenden“ Odoo-Buttons führen
- ein kompletter Cache-Neuaufbau ist hier der pragmatischste Weg, um Altlasten zu bereinigen

## Annahmen
- Diese Entscheidungen gelten fuer den aktuellen Arbeitsstand und koennen spaeter bewusst angepasst werden.
- Noch nicht jede historische Datei ist in sich vollstaendig mit dem neuesten Code synchron; bei Konflikten gilt der Code plus `docs/`-Stand.
