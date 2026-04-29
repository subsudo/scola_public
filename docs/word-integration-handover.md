# Word Integration Handover

Diese Datei beschreibt die aktuelle Word-Architektur von Scola so, dass dieselbe Logik in `Akta` moeglichst verlustfrei nachgebaut werden kann.

Sie ist bewusst keine allgemeine COM-Doku, sondern eine konkrete Uebergabe:
- welche Dateien relevant sind
- welche Verantwortung wo liegt
- wie normale Open-Flows laufen
- wie der BI-Sonderfall funktioniert
- welche Teile moeglichst 1:1 uebernommen werden sollten

## Zielbild
Scola fuehrt Word-Operationen **nicht direkt aus dem UI-Thread** aus.

Stattdessen:
1. `MainWindow.xaml.cs` orchestriert nur den UI-Flow
2. `WordStaHost` besitzt einen dedizierten **STA-Worker-Thread**
3. auf diesem STA-Thread lebt genau eine `WordService`-Instanz
4. alle Word-Operationen laufen seriell ueber diesen Host

Das ist der entscheidende Umbau gegen COM-/Threading-Probleme und gegen instabile Mehrfachzugriffe.

## Relevante Dateien
- `MainWindow.xaml.cs`
  - UI-Handler, Guarding, Dokumentpfad-Aufloesung, Queueing zum STA-Host
- `Services/WordStaHost.cs`
  - dedizierter STA-Worker mit Queue
- `Services/WordService.cs`
  - gesamte Word-/COM-Logik
- `docs/word-ghost-diagnostics.md`
  - Hintergrund, warum die Ownership-/Lifecycle-Logik so gebaut wurde

## Verantwortungen
### 1. MainWindow.xaml.cs
Die Button-Handler machen bewusst nur Orchestrierung:
- Voraussetzungen pruefen
- Teilnehmer-/Aktepfad aufloesen
- parallele Word-Aktionen verhindern
- Operation an `_wordStaHost.RunAsync(...)` uebergeben
- UI-Rueckmeldung via Toast / Alert / Status

Relevante Handler:
- `OpenAkteButton_OnClick(...)`
- `OpenAkteBuButton_OnClick(...)`
- `OpenAkteBiButton_OnClick(...)`
- `OpenAkteBeButton_OnClick(...)`
- `InsertEntryButton_OnClick(...)`
- `InsertEntryBiButton_OnClick(...)`

Wichtig fuer Akta:
- keine COM-Operation direkt in den Click-Handlern
- keine eigene Word-App im UI-Thread erzeugen
- Dokumentpfad vor dem Queueing aufloesen

### 2. WordStaHost.cs
`WordStaHost` ist die Trennschicht zwischen UI und COM.

Kernidee:
- ein Hintergrund-Thread
- `ApartmentState.STA`
- genau eine `WordService`-Instanz auf diesem Thread
- serielle Verarbeitung ueber `BlockingCollection`

Wichtige Eigenschaften:
- `RunAsync(...)` nimmt eine Aktion/Funktion entgegen
- die Aktion wird spaeter auf dem STA-Thread gegen dieselbe `WordService`-Instanz ausgefuehrt
- Rueckgabe/Fehler laufen per `TaskCompletionSource` an den Aufrufer zurueck

Portierungsempfehlung fuer Akta:
- dieses Muster moeglichst direkt uebernehmen
- nicht „vereinfacht“ wieder auf ad-hoc-`Task.Run` mit COM zurueckfallen

### 3. WordService.cs
`WordService` ist die gesamte operative Word-Schicht.

Dort liegen:
- Dokument oeffnen
- an Bookmark springen
- Tabellenzeile einfuegen
- BI-To-do-Sammellauf
- sichtbare/versteckte Word-Instanzen
- COM-Lifecycle / Release / Quit
- Foreground-/Visible-Verhalten
- Ghost-/Mehrinstanz-Abwehr

## Normaler Open-Flow
### A. Ganze Akte oeffnen
UI-Seite:
1. `OpenAkteButton_OnClick(...)`
2. `ResolveDocumentPathForParticipant(...)`
3. `_wordStaHost.RunAsync("OpenDocument", service => service.OpenDocument(docPath))`

Service-Seite:
1. `WordService.OpenDocument(docPath)`
2. `CreateOrAttachWordApplication(...)`
3. `OpenOrGetDocument(...)`
4. `CloseTransientEmptyDocuments(...)`
5. `EnsureWordUiState(...)`
6. `FocusDocument(...)`

Wichtig:
- wenn Scola selbst eine neue Word-Instanz erzeugt hat, wird sie nur dann wieder beendet, wenn die Operation fehlschlaegt
- bei Erfolg bleibt die sichtbare Benutzerinstanz bestehen

### B. Akte an Bookmark oeffnen
BU / BI / BE folgen demselben Muster, nur mit anderer Service-Methode:
- `OpenDocumentAtBookmark(docPath, bookmarkName)`

Service-Seite:
1. `CreateOrAttachWordApplication(...)`
2. `OpenOrGetDocument(...)`
3. `CloseTransientEmptyDocuments(...)`
4. `EnsureWordUiState(...)`
5. `EnsureDocumentNotLocked(...)`
6. `FocusBookmarkAtTop(...)`

Damit wird:
- dieselbe Akte normal in Word gezeigt
- aber direkt zur fachlich relevanten Stelle gesprungen

### C. Tabelleneintrag
BU-/BI-Eintraege laufen ebenfalls ueber denselben STA-Host, nur mit schreibender Operation:
- `InsertClipboardToTable(...)`
- `InsertTextRowToTable(...)`

Auch hier gilt:
- Word-App attachen oder erzeugen
- Dokument holen
- transient leere Dokumente schliessen
- sichtbare User-Instanz herstellen
- dann Bookmark-/Tabellenlogik ausfuehren

## CreateOrAttach: die zentrale Ownership-Logik
`CreateOrAttachWordApplication(...)` ist der wichtigste Portierungsbaustein.

Sie entscheidet:
- an bestehende Benutzerinstanz anhaengen
- oder neue Word-Instanz erzeugen

Die Rueckgabe enthaelt mehr als nur `app`:
- COM-App
- ob die Instanz von Scola **neu erzeugt** wurde
- wie viele unsaved Dokumente initial schon offen waren

Diese Information wird spaeter gebraucht fuer:
- Cleanup leerer transienter Dokumente
- Fehlerbehandlung
- Quit-Entscheidungen

Zusatzhaertung in Scola:
- mehrere Retry-Faelle fuer ROT-/RPC-Hickups
- Ghost-/Zombie-Abwehr bei verwaisten Word-Kontexten

Portierungsempfehlung:
- wenn Akta denselben Stabilitaetsgrad will, diese Methode nicht funktional “nachbauen”, sondern strukturell moeglichst eng uebernehmen

## Sichtbarkeit und Fensterverhalten
`EnsureWordUiState(...)` ist die Stelle, die Word fuer den Nutzer sichtbar macht.

Wichtig:
- Scola setzt hier nur `Visible = true`
- danach `TryBringWordToForeground(...)`
- minimiertes Word darf wiederhergestellt werden
- Fensterplatzierung/-groesse wird **nicht** aktiv gesetzt

Das ist absichtlich so:
- Word verwaltet seine Fensterposition selbst
- Scola steuert Fokus/Sichtbarkeit, aber kein Placement

Wenn Akta dieselbe UX will, sollte sie genau diese Zurueckhaltung beibehalten.

## BI-Sonderfall: dedizierte versteckte Instanz
`CollectBiTodoDocument(...)` ist fachlich und technisch ein Sonderfall.

Hier reicht normales Attachen an User-Word nicht, weil:
- Inhalte aus mehreren BI-Akten gelesen werden
- dabei keine sichtbaren Ghost-/Zwischenfenster entstehen sollen

Deshalb nutzt Scola:
1. `CreateDedicatedHiddenWordApplication(...)`
2. verstecktes, dediziertes Word nur fuer den BI-Sammellauf
3. Ergebnisdokument in dieser Hidden-Instanz erzeugen
4. Ergebnis als Temp-`docx` speichern
5. Hidden-Instanz sauber beenden
6. auf echte User-Word-Instanz ueber `CreateOrAttachWordApplication(...)` handoffen
7. dort das fertige Dokument sichtbar oeffnen

Die zusaetzlichen Haertungen dafuer:
- Hidden-PID-Erkennung
- `WaitForExit(...)` nach `Quit(false)`
- Retry auf `GetActiveObject(...)` bei fluechtigen RPC-Fehlern
- Hidden-PID-Diff statt fragiler COM-Fensterabfrage

Das ist ein eigener Portierungsblock fuer Akta:
- nur uebernehmen, wenn Akta denselben BI-Sammeldokument-Mechanismus braucht
- nicht noetig fuer einfache Open-/Bookmark-Flows

## Was Akta moeglichst 1:1 uebernehmen sollte
Wenn Akta den grossen Umbau sauber replizieren will, dann diese Teile moeglichst direkt:

1. `WordStaHost` als Architekturprinzip
- eigener STA-Thread
- serialisierte Queue
- eine `WordService`-Instanz

2. `MainWindow`-/UI-Seite nur als Orchestrierung
- Pfad aufloesen
- Guarding
- Rueckmeldung
- keine COM-Arbeit im UI-Thread

3. `WordService`-Ownership-Logik
- `CreateOrAttachWordApplication(...)`
- `OpenOrGetDocument(...)`
- `CloseTransientEmptyDocuments(...)`
- `EnsureWordUiState(...)`

4. Fehler- und Cleanup-Muster
- bei Fehlern geoeffnetes Dokument wieder schliessen, wenn es nur fuer diese Operation geoeffnet wurde
- selbst gestartete Word-Instanz nur bei Fehler wieder beenden
- COM-Objekte im `finally` konsequent freigeben

## Was Akta anpassen muss
Diese Dinge sind nicht blind 1:1 zu kopieren, sondern an Akta zu mappen:

- konkrete Button-/Command-Handler
- Akta-spezifische Dokumentpfad-Aufloesung
- Bookmark-Namen / Tabellennamen
- Alert-/Toast-Mechanik
- Teilnehmer-/Aktenmodell

Falls Akta dieselbe Aktenbasis und dieselben Bookmarks nutzt, kann ein grosser Teil der fachlichen Parameter sogar gleich bleiben.

## Wichtige Stolpersteine
### 1. Nicht aus dem UI-Thread mit Word reden
Das ist die wichtigste Regel.

### 2. Nicht mehrere unkoordinierte WordService-Instanzen parallel erzeugen
Sobald mehrere COM-Kontexte ungeordnet parallel laufen, wird das Ghost-/Ownership-Problem schnell wieder schlechter.

### 3. Sichtbarkeit nicht mit Placement verwechseln
Scola macht Word sichtbar und fokussiert es, setzt aber keine Fensterposition.

### 4. BI-Handoff nicht mit normalen Open-Flows vermischen
Der BI-Sammellauf ist ein Spezialfall mit Hidden-Instanz und Temp-Handoff.

## Minimale Portierungsreihenfolge fuer Akta
Wenn Akta das schrittweise uebernehmen soll, ist diese Reihenfolge die stabilste:

1. `WordStaHost` in Akta einfuehren
2. normale `OpenDocument`- und `OpenDocumentAtBookmark`-Flows daran haengen
3. Guarding/Cleanup wie in Scola mitnehmen
4. Tabellen-Eintrag-Flows portieren
5. erst danach den BI-/Sammeldokument-Sonderfall portieren

## Kurzfassung fuer eine andere Instanz
Wenn eine andere Instanz nur die Kernbotschaft braucht:

- Die Word-Dateien werden in Scola nicht direkt im UI-Thread geoeffnet.
- Alle Word-Aktionen laufen ueber `WordStaHost` auf einem eigenen STA-Thread.
- Die eigentliche COM-Logik sitzt in `WordService.cs`.
- Der normale Flow ist:
  - UI-Handler -> `_wordStaHost.RunAsync(...)` -> `WordService.OpenDocument(...)` oder `OpenDocumentAtBookmark(...)`
- Der BI-Sammellauf ist ein Sonderfall mit versteckter dedizierter Word-Instanz und spaeterem Handoff in sichtbares User-Word.
