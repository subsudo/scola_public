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

## D15 - Hauptfenster-Position bleibt WPF-basiert mit Sichtbarkeits-Fallback
Die Wiederherstellung des Scola-Hauptfensters bleibt bei WPF-`RestoreBounds`/gespeicherten Prefs und wird nicht auf einen separaten nativen `WINDOWPLACEMENT`-Stack umgestellt.

Aktuelle Regel:
- gespeicherte Bounds werden beim Start auf einen gueltigen sichtbaren Bereich geprueft
- fehlt der gespeicherte Monitor, faellt Scola direkt auf den primaeren Monitor zurueck
- bei zur Laufzeit geaendertem Monitor-Layout wird das Fenster nur dann zurueckgeholt, wenn es nicht mehr ausreichend sichtbar ist

Begruendung:
- passt zur bestehenden WPF-Logik ohne zweiten Placement-Stack
- minimiert Risiko gegenueber einem groesseren nativen Umbau
- loest den haeufigen Docking/Undocking-Fall pragmatisch

## D16 - Release-Prozess bleibt bewusst manuell
Build, Publish und GitHub-Release fuer Scola werden bewusst nicht bei normalen Commits oder Pushes automatisch ausgelost.

Aktuelle Regel:
- normale Code-Aenderung = keine neue Version, keine Release
- Test-EXE darf gebaut werden, ohne direkt veroeffentlicht zu werden
- echte Release nur auf ausdruecklichen Wunsch
- operative Details stehen in `docs/release-workflow.md`

Begruendung:
- reduziert versehentliche Veroeffentlichungen
- passt zum produktnahen, kontrollierten Arbeitsstil des Projekts
- ist fuer Mensch und KI leichter verlässlich einzuhalten als ein halbautomatischer Mischprozess

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
