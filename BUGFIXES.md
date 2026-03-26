# BUGFIXES

Stand: 2026-03-06

Dieses Dokument dient als laufendes Arbeitsprotokoll fuer Bugs, Auffaelligkeiten, Ursachenvermutungen und moegliche Fixes, die waehrend der Analyse im Chat gesammelt wurden. Es soll spaeter von einer anderen KI oder Person weiterverwendet werden.

## Vorgehen

- Diese Datei wird waehrend weiterer Chats regelmaessig erweitert.
- Pro Punkt werden nach Moeglichkeit festgehalten:
  - Beobachtung
  - betroffene Dateien / Codepfade
  - vermutete Ursache
  - moeglicher Fix
  - optional Reproduktionshinweise

## Bisher besprochen

### 1. Allgemeiner Projektueberblick

- Projekt: WPF Desktop-App auf .NET 8
- Produktname zur Laufzeit: `AkteX`
- Hauptfunktion:
  - Teilnehmerliste einlesen
  - Teilnehmer gegen Ordner auf Serverpfaden matchen
  - Word-Verlaufsakten oeffnen
  - zu BU / BI / BE springen
  - BU- / BI-Eintraege einfuegen
  - Batch-Eintraege verarbeiten

Wichtige Einstiegsdateien:

- `App.xaml.cs`
- `MainWindow.xaml.cs`
- `SettingsWindow.xaml.cs`
- `Services/WordService.cs`
- `Services/ParticipantParser.cs`
- `Services/FolderMatcher.cs`

### 2. Ausgewertete Logs

Analysiert:

- `%LOCALAPPDATA%\\AkteX\\logs\\app-2026-03-05.log`
- `%LOCALAPPDATA%\\AkteX\\logs\\app-2026-03-06.log`

Ergebnis:

- kein Hinweis auf einen globalen App-Crash
- aber mehrere technische und fachliche Fehlerbilder

## Bugs und Auffaelligkeiten

### Bug 1: Leere Word-Dokumente wie "Dokument6" bleiben offen

Status:

- umgesetzt am 2026-03-06
- weiter beobachten in realer Word-/Office-Umgebung

Beobachtung:

- ploetzlich sind leere Word-Dokumente offen
- teilweise blockiert
- Namen wie `Dokument6`
- laut Nutzer nicht zwingend direkt nach einem Klick sichtbar
- Nutzer war teilweise eine Zeit lang weg und sah die Dokumente erst spaeter
- beim direkten erneuten Testen oeffnet sich nicht reproduzierbar sofort eine zweite Instanz

Betroffene Datei:

- `Services/WordService.cs`

Relevante Codepfade:

- `CreateOrAttachWordApplication()`
- `EnsureWordUiState(...)`
- `OpenOrGetDocument(...)`

Technische Einschaetzung:

- Die App versucht zunaechst, eine laufende Word-Instanz via `GetActiveObject(...)` zu verwenden.
- Falls keine Instanz existiert, wird per `Activator.CreateInstance("Word.Application")` eine neue Word-Anwendung erzeugt.
- Danach wird Word sofort sichtbar gemacht (`app.Visible = true`).
- Es gibt aktuell keinen klaren Lebenszyklus, der eine selbst gestartete Instanz wieder sauber beendet.
- Dadurch ist es plausibel, dass Word mit einem Standard-Leerdokument startet bzw. ein solches offen laesst.
- Dieses Leerdokument kann dann als `DokumentX` sichtbar werden.

Wichtige Praezisierung nach Codepruefung:

- Der aktuelle Code fasst Word **nicht** spaeter autonom per Hintergrundlogik an.
- Word wird nur in direkten UI-Aktionen oder im laufenden Batch verwendet:
  - `OpenAkteButton_OnClick(...)`
  - `OpenAkteBuButton_OnClick(...)`
  - `OpenAkteBiButton_OnClick(...)`
  - `OpenAkteBeButton_OnClick(...)`
  - `InsertEntryButton_OnClick(...)`
  - `InsertEntryBiButton_OnClick(...)`
  - `ExecuteBatchButton_OnClick(...)`
- Daraus folgt:
  - Wenn leere Dokumente erst spaeter auffallen, ist das wahrscheinlich **kein neuer spaeterer App-Aufruf**, sondern eine **Spaetwirkung einer frueher gestarteten Word-Instanz** oder ein Verhalten innerhalb von Word selbst.

Verfeinerte Ursachenhypothese:

1. Die App startet irgendwann frueher eine neue Word-Instanz.
2. Diese Instanz bleibt offen.
3. Das eigentliche Zieldokument wird spaeter geschlossen oder verliert den Fokus.
4. Uebrig bleibt bzw. erscheint sichtbar ein leeres Standarddokument `DokumentX`.

Alternative Erklaerung, die ebenfalls offen bleiben muss:

- Word oder ein Office-Addin erzeugt in einer bereits laufenden, von der App gestarteten Instanz spaeter ein leeres Dokument.
- Diese Alternative ist aktuell aus dem Code nicht beweisbar, aber der Code schliesst sie auch nicht aus.

Warum das problematisch ist:

- offene Leerdokumente verwirren Nutzer
- mehrere verbleibende Word-Instanzen koennen Folgeprobleme und Locks verursachen
- dadurch koennen Dokumente "blockiert" wirken oder tatsaechlich unnoetig offen bleiben

Moeglicher Fix:

1. `CreateOrAttachWordApplication()` so umbauen, dass zurueckkommt:
   - `app`
   - `wasCreatedHere`
2. Word nicht sofort sichtbar machen, wenn die Instanz gerade neu erzeugt wurde.
3. Erst das eigentliche Zieldokument oeffnen.
4. Danach gezielt nur das Zieldokument aktivieren.
5. Falls die App die Instanz selbst erzeugt hat:
   - entweder kontrolliert offen lassen, aber ohne leeres Startdokument
   - oder bei reinen technischen Operationen wieder schliessen
6. Fuer den simplen "Akte oeffnen"-Fall pruefen, ob statt COM-Automation ein normales Shell-Open reicht.

Empfohlene Richtung:

- `OpenDocument(...)` fuer reines Oeffnen eher per Shell oeffnen
- COM-Automation nur fuer:
  - Bookmark-Navigation
  - BU-/BI-Eintrag
  - Batch

Was gegen die These "App erzeugt spaeter heimlich neue Instanzen" spricht:

- kein Hintergrund-Worker fuer Word
- kein Timer, der Word-Aktionen ausloest
- keine Autologik ausser dem explizit gestarteten Batch
- kein Codepfad gefunden, der ohne Nutzeraktion spaeter `OpenDocument*` oder `Insert*` aufruft

Reproduktion:

1. Alle Word-Instanzen schliessen
2. In der App `Akte`, `BU`, `BI` oder `Eintrag` ausfuehren
3. beobachten, ob Word mit zusaetzlichem Leerdokument startet

Erweiterte Reproduktion:

1. Alle Word-Instanzen schliessen
2. Eine Word-Aktion in der App ausfuehren
3. Danach die App und Word einige Zeit offen lassen
4. pruefen, ob spaeter in derselben Word-Instanz ein leeres `DokumentX` sichtbar wird
5. parallel im Task-Manager pruefen:
   - entsteht wirklich eine neue `WINWORD`-Instanz
   - oder bleibt nur dieselbe Instanz offen und zeigt spaeter ein leeres Dokument an

Umgesetzt am 2026-03-06:

- `CreateOrAttachWordApplication()` liefert jetzt zurueck, ob Word von der App neu gestartet wurde.
- Selbst gestartete Word-Instanzen werden nicht mehr vorzeitig sichtbar gemacht.
- Das Zieldokument wird zuerst geoeffnet, danach wird Word sichtbar/fokussiert.
- In selbst gestarteten Instanzen werden transiente Leerdokumente ohne echten Dateipfad nach dem Oeffnen des Zieldokuments zu schliessen versucht.
- Wenn eine selbst gestartete Word-Instanz vor erfolgreichem Abschluss in einen Fehler laeuft, wird sie nach Moeglichkeit wieder beendet.

Rest-Risiko:

- `OpenDocument(...)` verwendet weiterhin COM und nicht Shell-Open.
- Es bleibt moeglich, dass Word- oder Office-Addins trotzdem eigene Leerdokumente erzeugen.

### Bug 2: Clipboard-Lock fuehrt zu Insert-Fehler

Status:

- umgesetzt am 2026-03-06

Beobachtung:

- `OpenClipboard fehlgeschlagen (CLIPBRD_E_CANT_OPEN)`
- BI-Einfuegen brach in mindestens einem Fall ab

Betroffene Datei:

- `Services/WordService.cs`

Relevanter Codepfad:

- `InsertClipboardToTable(...)`

Vermutete Ursache:

- `Clipboard.ContainsText()` / `Clipboard.GetText()` greifen direkt und ohne Retry auf die Zwischenablage zu
- wenn ein anderer Prozess die Zwischenablage kurz blockiert, faellt der gesamte Insert fehl

Moeglicher Fix:

1. Hilfsmethode `TryGetClipboardTextWithRetry(...)` einfuehren
2. mehrere kurze Retries mit kleinem Delay
3. bei dauerhaftem Fehlschlag:
   - klare Meldung fuer Nutzer
   - kein irrefuehrender "Word"-Fehlertext

Reproduktion:

1. Clipboard intensiv von einem anderen Tool benutzen
2. in der App `Eintrag BI` oder `Eintrag BU` ausloesen
3. beobachten, ob sporadisch `CLIPBRD_E_CANT_OPEN` auftritt

Umgesetzt am 2026-03-06:

- `WordService` liest das Clipboard jetzt ueber Retry-Logik mit kurzem Delay.
- Bei dauerhaft blockierter Zwischenablage wird eine klare Nutzerfehlermeldung geworfen:
  - `Zwischenablage ist momentan blockiert. Bitte kurz warten und erneut versuchen.`

### Bug 3: Fehlende Bookmarks in einzelnen Akten verhindern Einfuegen

Status:

- in Log nachweisbar

Beobachtung:

- mindestens ein Dokument enthielt `BU_BILDUNG_TABELLE` nicht
- der Insert brach mit Bookmark-Fehler ab

Betroffene Datei:

- `Services/WordService.cs`

Vermutete Ursache:

- Inkonsistenz zwischen App-Konfiguration und realen Word-Vorlagen / Dokumenten

Moeglicher Fix:

1. Validierung der Vorlagen / Dokumente
2. optional Admin-Check-Funktion:
   - Dokument pruefen
   - vorhandene Bookmarks auflisten
3. Fehlermeldung weiter konkret halten
4. optional Diagnose-Dialog mit:
   - Dokumentpfad
   - erwartetem Bookmark-Namen

Reproduktion:

1. Dokument ohne `BU_BILDUNG_TABELLE` verwenden
2. `Eintrag BU` ausloesen

### Bug 4: Parser verarbeitet Freitext-Zusaetze im Namen unzureichend

Status:

- umgesetzt am 2026-03-06
- bewusst konservativ implementiert

Beobachtung:

- Eingabe wie `Abdigani Adaan Kommt um 13.30 Uhr`
- Parser uebernahm Zusatz als Teil des Namens
- kein eindeutiger Fallback-Name gefunden
- Status fiel auf Default `Anwesend`

Betroffene Datei:

- `Services/ParticipantParser.cs`

Vermutete Ursache:

- Parser erkennt zwar viele Statusvarianten, aber keine allgemeinen Freitext-Anhaenge wie:
  - `kommt um ...`
  - `spaeter`
  - `ab ...`
  - organisatorische Notizen

Umgesetzt am 2026-03-06:

- Der Parser trennt jetzt konservativ typische Freitext-Anhaenge vom Namensfeld ab, wenn kein echter Status erkannt wurde.
- Beispiele fuer Marker:
  - `kommt ...`
  - `kommt um ...`
  - `spaeter` / `später`
  - `ab ...`
  - `krank`, `arzt`, `ferien`, `urlaub`, `meldung`, `notiz`
- Der erkannte Freitext wird als Remark behandelt und nicht mehr in den Namen uebernommen.

Rest-Risiko:

- Die Logik ist absichtlich vorsichtig und basiert auf Marker-Prefixen.
- Unbekannte Freitext-Formulierungen koennen weiterhin als Teil des Namens durchrutschen.

Reproduktion:

1. Zeile wie `Vorname Nachname Kommt um 13.30 Uhr`
2. auswerten
3. Matching und Status pruefen

## Warnungen mit niedriger Prioritaet

### Warnung A: `Word.UserControl` ist schreibgeschuetzt

Beobachtung:

- tritt oft im Log auf
- Folgeoperationen laufen haeufig trotzdem erfolgreich weiter

Einschaetzung:

- eher Kompatibilitaets- / COM-Eigenheit
- wahrscheinlich nicht die eigentliche Fehlerursache

Moegliche Verbesserung:

- Warnung nur einmal pro Session loggen
- oder gezielter nach Instanztyp unterscheiden

### Warnung B: `Word.Hwnd` kann nicht gelesen werden

Beobachtung:

- tritt oft im Log auf
- Bookmark-Spruenge und Dokumentoeffnungen funktionieren haeufig trotzdem

Einschaetzung:

- betrifft vor allem Foreground-/Fokus-Fallback
- eher Diagnose-Rauschen als Kernfehler

Moegliche Verbesserung:

- Fallback robuster machen
- Warnung reduzieren, wenn Word trotzdem erfolgreich fokussiert / geoeffnet wurde

## Offene technische Fragen

### Soll fuer "Akte oeffnen" ueberhaupt COM verwendet werden?

Aktueller Stand:

- Fuer `OpenDocument(...)` wird weiterhin COM verwendet.
- Der akute Word-Instanz-Bug wurde ueber verbesserten COM-Lebenszyklus entschärft, nicht ueber Shell-Open.

Frage:

- Ist fuer den reinen Oeffnen-Fall eine Shell-Ausfuehrung sauberer und risikoaermer?

Moeglicher Vorteil:

- weniger Einfluss auf Word-Instanz-Lebenszyklus
- geringeres Risiko fuer `DokumentX`-Nebeneffekte

Nachteil:

- kein direkter COM-Zugriff fuer Fokus / Navigation in derselben Methode

## Priorisierte TODO-Reihenfolge

### Prioritaet 1: Hoch

1. Diagnose fuer fehlende Bookmarks verbessern.
2. Entscheidung spaeter erneut pruefen, ob `OpenDocument(...)` langfristig bei COM bleiben oder doch auf Shell-Open wechseln soll.

### Prioritaet 2: Mittel

1. Warnungsrauschen bei `Word.UserControl` und `Word.Hwnd` reduzieren.
2. Parser-Regeln fuer weitere Freitext-Formulierungen erweitern, falls neue reale Beispiele auftauchen.

### Prioritaet 3: Niedrig

1. Clear- und Reset-UX im Alltag gegen reale Nutzung pruefen und bei Bedarf optisch feinjustieren.

## Empfohlene Umsetzungsreihenfolge fuer die naechste KI

1. `Services/WordService.cs`
   - Bookmark-Diagnose verbessern
2. `Services/WordService.cs`
   - Logging bei `UserControl` / `Hwnd` entrauschen
3. `Services/ParticipantParser.cs`
   - weitere Freitext-Beispiele nur anhand realer Inputs nachschaerfen
4. `MainWindow.xaml` / `MainWindow.xaml.cs`
   - Clear-/Reset-UX bei Bedarf optisch anpassen

## Lokale Build-Umgebung

Aktueller Stand:

- `dotnet` ist in der aktuellen Umgebung verfuegbar.
- Das Projekt konnte lokal mit `dotnet build` und `dotnet publish` gebaut werden.
- Vorhandene EXE-Dateien koennen also aktiv aus dem aktuellen Stand neu erzeugt werden.

## Neue Feature-Requests aus dem Chat

### Feature 1: Clear-Button neben `Auswerten`

Wunsch:

- Im Bereich `Liste neu einfuegen` soll neben `Auswerten` ein zusaetzlicher Button stehen, der nur den Inhalt des Eingabefeldes leert.

Umgesetzt am 2026-03-06:

- Neben `Auswerten` existiert jetzt ein separater Button `Loeschen`.
- Er leert nur den Inhalt des grossen Eingabefelds.
- Der Fokus springt danach wieder ins Eingabefeld.
- Die bestehende Auswertungsliste wird dadurch nicht automatisch geaendert.

### Feature 2: Unauffaelliger Reset der Auswertungsliste

Wunsch:

- Die ausgewaertete Teilnehmerliste soll zurueckgesetzt werden koennen
- die Funktion soll nicht zu auffaellig integriert sein

Umsetzungsidee aus dem Chat:

- kein grosser prominenter Button in der Hauptflaeche
- stattdessen ein dezenter `Reset` rechts in der bestehenden `Liste neu einfuegen`-Leiste

Umgesetzt am 2026-03-06:

- In der Leiste `Liste neu einfuegen` gibt es jetzt einen dezenten Button `Reset`.
- `Reset`:
  - leert die Teilnehmerliste
  - leert die Batch-Ergebnisse
  - blendet den Ergebnisbereich aus
  - blendet den Eingabebereich wieder ein
  - behaelt den aktuellen Rohtext im Eingabefeld
