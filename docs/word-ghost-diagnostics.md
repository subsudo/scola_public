# Word Ghost Diagnostics

Stand dieser Notiz: 2026-03-26

## Ziel
Diese Datei haelt den aktuellen Erkenntnisstand zu den beobachteten Word-"Geisterdokumenten" fest.

Sie ist bewusst als Arbeitsdokument gedacht:
- bekannte Symptome
- belegte Logbefunde
- aktuelle Hypothesen
- offene Fragen
- Leitfaden fuer spaetere Log-Analysen

Die Datei soll in den naechsten Tagen/Wochen weiter ergaenzt werden, statt dieselben Schluesse immer wieder neu aus Chatverlaeufen zu rekonstruieren.

## Symptom-Bild
### Typ A - leeres Dokument ohne normale Schliessbarkeit
Beobachtung aus der Praxis:
- leere Fenster wie `Dokument 6`, `Dokument 7` oder aehnlich
- keine normale Bedienbarkeit ueber das uebliche Word-UI
- teilweise nur ueber Kontextmenue/Taskleiste schliessbar

### Typ B - echte Akte mit Inhalt, aber ohne normale Word-UI
Beobachtung aus der Praxis:
- Akteninhalt ist sichtbar und scrollbar
- Ribbon / Menue / normale Fensterchrome fehlen
- Fenster wirkt wie rohe Word-Automation statt wie normales Word

### Typ C - Geist mit echtem Aktennamen
Beobachtung aus der Praxis:
- Fenster zeigt nicht `Dokument N`, sondern den echten Aktennamen
- Beispiel: `Verlaufsakte_TN-A.docx`
- nach dem Schliessen der eigentlichen Akte bleibt scheinbar noch ein weiteres Fenster mit demselben Titel zurueck

## Relevante Logs
### Bereits ausgewertet
- `C:\Users\chris\OneDrive - Genossenschaft FuturX\Schule\Desktop\app-2026-03-24.log`
- `C:\Users\chris\OneDrive - Genossenschaft FuturX\Schule\Desktop\app-2026-03-25.log`

### Diagnose-Logging-Staende
- erstes erweitertes Word-Lifecycle-Logging:
  - Commit `4e718f7` - `Add Word lifecycle diagnostics logging`
- zusaetzliches `WINWORD`-Prozess-Snapshot-Logging:
  - Commit `603a032` - `Add WinWord process diagnostics logging`

## Was aktuell belegt ist
### B1 - BI-To-do kann ein leeres `Dokument1` in einer dedizierten Word-Instanz erzeugen
Im Log vom 2026-03-25 am Morgen ist belegt:
- `CollectBiTodoDocument` startet eine dedizierte versteckte Word-Instanz
- darin existiert danach ein unsaved `Dokument1`
- `UserControl` konnte fuer diese Instanz nicht gesetzt werden
- die Instanz wurde danach trotzdem sichtbar gemacht

Das ist ein realer Ghost-Kandidat.

Wichtig:
- dieser Befund ist echt
- er erklaert aber nicht automatisch jeden spaeteren Geist des Tages
- insbesondere nicht zwingend den am Nachmittag beobachteten TN-A-Fall

### B2 - Viele normale Open-Flows enden mit `Visible=True`, obwohl `UserControl` zuvor fehlschlug
In mehreren normalen `Akte`-/`BU`-Open-Flows vom 2026-03-25 sieht man:
- keine laufende Word-Instanz in ROT gefunden
- neue Word-Instanz erstellt
- `UserControl`-Setzen schlaegt fehl
- danach trotzdem `Visible=true`
- Endzustand teilweise:
  - `Visible=True`
  - `UserControl=False`
  - echtes Dokument offen

Das ist derzeit der staerkste allgemeine Kandidat fuer Symptom Typ B:
- echter Inhalt
- aber keine normale Word-UI

### B3 - Der TN-A-Fall war kein `Dokument N`, sondern ein echtes Dateidokument
Rueckmeldung aus der Praxis:
- das Geistfenster hiess exakt wie die TN-A-Akte
- nicht `Dokument N`

Das passt nicht gut zu einer rein spaeten Blankodokument-These.

Im Log vom 2026-03-25 ist belegt:
- `Verlaufsakte_TN-A.docx` wurde um `13:05` per `OpenDocumentAtBookmark-BU` geoeffnet
- `Verlaufsakte_TN-A.docx` wurde um `16:16` im `BU-Batch` nochmals geoeffnet
- in beiden Snapshots handelt es sich um das echte Dateidokument:
  - echter `FullName`
  - `Path` gesetzt
  - `IsUnsaved=False`

Das spricht fuer eine wichtigere Unterscheidung:
- nicht jeder Geist ist ein leeres Standarddokument
- es gibt offenbar auch Faelle, in denen eine echte Akte in einem weiteren/falschen Word-Kontext weiterlebt

### B4 - Der bisherige Snapshot sieht nur die Word-Instanz, an die Scola gerade angehaengt ist
Das aktuelle Lifecycle-Logging zeigte bis einschliesslich 2026-03-25:
- Dokument-Snapshots der aktuell benutzten Word-COM-Instanz
- aber nicht sicher alle parallel laufenden `WINWORD`-Prozesse

Dadurch blieb eine grosse Luecke:
- wenn ein weiterer verwaister Word-Prozess existiert
- und Scola sich an einen anderen Prozess anhaengt
- dann erscheint der verwaiste Prozess nicht in den bisherigen Dokument-Snapshots

Genau deshalb wurde das zweite Diagnose-Upgrade mit `WINWORD`-Prozess-Snapshots eingebaut.

## Was aktuell nicht belegt ist
### N1 - Nicht belegt: `CloseTransientEmptyDocuments(...)` ist die alleinige Hauptursache
Der Verdacht war zwischenzeitlich stark, ist aber aktuell nicht mehr der wahrscheinlichste Haupttreiber.

Warum:
- mehrere problematische Faelle passen besser zu neu erzeugten, sichtbar gelassenen Word-Instanzen
- besonders dann, wenn `UserControl` nicht sauber gesetzt werden konnte

`CloseTransientEmptyDocuments(...)` kann weiterhin ein Nebenfaktor sein, ist aber aktuell nicht die beste Erklaerung fuer alle beobachteten Symptomtypen.

### N2 - Nicht belegt: BI ist die Hauptursache fuer alle Ghosts
Die dedizierte BI-Instanz ist ein belegter Risikofaktor.

Aber:
- Ghosts wurden auch an Tagen/Faellen ohne BI-Nutzung beobachtet
- Ghosts wurden auch bei normalem `Akte`, `BU` und `BU-Batch` erlebt

Fazit:
- BI bleibt relevant
- ist aber derzeit nicht als alleinige Root Cause haltbar

### N3 - Nicht belegt: `BU-Batch` ist immer der Hauptverursacher
Es gibt gute Gruende, `BU-Batch` als Verstaerker zu sehen, weil dort viele Akten nacheinander offen bleiben koennen.

Aber:
- Ghosts treten nicht nur dort auf
- im konkreten Log vom 2026-03-25 zeigen die Batch-Snapshots selbst keine neuen `Dokument N`-Leerdokumente

Fazit:
- Batch bleibt verdaechtig
- ist aber aktuell nicht als alleinige oder sicher wichtigste Ursache bewiesen

## Aktuelle Arbeitshypothesen
### H1 - Allgemeines Word-Ownership-Problem
Staerkste Arbeitshypothese.

Scola erzeugt oder findet Word, fuehrt eine Operation aus, gibt die COM-Referenzen frei und laesst Word weiterlaufen.

Problematisch wird das besonders dann, wenn:
- die Instanz von Scola frisch erzeugt wurde
- sie nicht sauber in einen echten Benutzerzustand uebergeht
- die App danach keinen klaren Besitz mehr ueber diese Instanz hat

### H2 - Frisch erzeugte Instanzen werden teils sichtbar gelassen, obwohl sie nicht sauber `UserControl`-faehig sind
Das ist der derzeit staerkste Kandidat fuer:
- rohe Inhaltsfenster ohne Ribbon/Menue
- Word-Automationsfenster, die spaeter komisch weiterleben

Kurzform:
- neue Instanz
- `UserControl` scheitert
- `Visible=true` trotzdem
- spaeter bleibt eine schwer kontrollierbare Word-Instanz zurueck

### H3 - Es gibt wahrscheinlich Mehrinstanz-Faelle, die die alten Snapshots nicht sehen
Besonders fuer den TN-A-Fall ist das aktuell die beste Vermutung.

Wenn ein Geist exakt denselben echten Dateinamen traegt, ist eine plausible Erklaerung:
- dieselbe Akte lebt in einem zweiten Word-Prozess / zweiten Word-Kontext weiter
- waehrend Scola sich fuer spaetere Operationen an eine andere Word-Instanz anhaengt

Diese Hypothese war bisher nicht sauber pruefbar.
Das neue `WINWORD`-Prozess-Logging soll genau das klaeren.

### H4 - BI-Sonderinstanz bleibt ein Zusatzrisiko
Auch wenn BI nicht der Haupttreiber aller Ghosts ist:
- die dedizierte BI-Instanz erzeugt nachweislich einen problematischen Zustand mit `Dokument1`
- sie bleibt deshalb als separater Spezialfall relevant

## Aktueller Diagnose-Fokus
Der Fokus liegt nicht mehr primaer auf:
- nur BI
- nur `CloseTransientEmptyDocuments(...)`
- nur `Dokument N`

Der Fokus liegt jetzt auf:
1. Wie viele `WINWORD`-Prozesse existieren rund um die Operation?
2. An welchen Prozess haengt sich Scola an?
3. Bleibt nach einer Operation ein weiterer `WINWORD`-Prozess zurueck?
4. Traegt dieser Prozess einen Hauptfenstertitel, der zu einer konkreten Akte passt?
5. Entstehen Geister vor allem nach:
   - neuer Word-Instanz
   - `UserControl`-Fehlschlag
   - spaeterem manuellem Schliessen echter Aktenfenster

## Wie kuenftige Logs gelesen werden sollen
### Besonders relevante Marker
- `CreateOrAttach: Keine laufende Word-Instanz in ROT gefunden, neue Instanz wird erstellt`
- `EnsureWordUiState: UserControl fehlgeschlagen`
- `Stage='...-WinWordProcesses'`
- `Stage='Finally-AfterRelease'`
- `WinWordPid=...`
- `MainWindowTitle='...'`

### Wichtige Fragen pro Vorfall
Wenn ein neuer Ghost-Vorfall untersucht wird, sollten diese Fragen beantwortet werden:

1. Welche Aktion lief kurz davor?
- `Akte`
- `BU`
- `BE`
- `BI`
- `Eintrag BU`
- `BU-Batch`
- `BI: To-dos`

2. Wurde dabei eine neue Word-Instanz gestartet?

3. Schlug `UserControl` fehl?

4. Wie viele `WINWORD`-Prozesse gab es:
- vor der Operation
- nach `CreateOrAttach`
- nach `EnsureWordUiState`
- nach `Finally-AfterRelease`

5. Welche `MainWindowTitle` hatten diese Prozesse?

6. Passt der Geist eher zu:
- leerem `Dokument N`
- oder echter Akte mit echtem Dateinamen

7. Taucht dieselbe Akte spaeter nochmals in einer anderen Operation auf?

## Naechste sinnvolle Analyse-Schritte
### A1 - Neue Logs mit Prozess-Snapshots abwarten
Der wichtigste naechste Schritt ist nicht sofort ein weiterer Fix, sondern echte Produktionsbeobachtung mit dem neuen Prozess-Logging.

Ziel:
- Beweis oder Widerlegung der Mehrinstanz-Hypothese

### A2 - Vorfall mit Uhrzeit notieren
Wenn ein neuer Geist auftaucht, ist hilfreich:
- ungefaehre Uhrzeit
- Aktion kurz davor
- Titel des Fensters
- leer oder mit echtem Inhalt

### A3 - Erst nach mehreren neuen Logs ueber Fix-Reihenfolge entscheiden
Vor einem weiteren Eingriff sollte moeglichst klar sein:
- ob wir primaer ein Mehrinstanz-Problem haben
- oder eher ein Sichtbarkeits-/`UserControl`-Problem
- oder beides

## Moegliche spaetere Fix-Richtungen
Noch nicht umsetzen, nur als Arbeitsrahmen festhalten.

### F1 - Frisch erzeugte Word-Instanz nicht sichtbar machen, wenn `UserControl` nicht sauber gelingt
Potenziell hoher Hebel fuer Symptom Typ B.

### F2 - Ownership pro Operationstyp sauberer trennen
Beispiele:
- `Akte`/`BU`/`BI` duerfen sichtbare Benutzerfenster hinterlassen
- reine Sammel-/Hilfsinstanzen eher nicht

### F3 - BI-Sonderinstanz spaeter separat haerten
Nur falls neue Logs bestaetigen, dass BI weiterhin eigenstaendig Geister erzeugt.

### F4 - Groesseren Session-/Ownership-Umbau nur bei klarer Notwendigkeit
Aktuell bewusst noch nicht priorisieren.

## Kurzfazit
Der aktuelle Stand spricht am staerksten fuer ein allgemeines Word-Instanz-/Ownership-Problem.

Besonders verdaechtig sind Faelle, in denen:
- Scola eine neue Word-Instanz startet
- `UserControl` nicht sauber gesetzt werden kann
- die Instanz trotzdem sichtbar und spaeter weiterverwendet oder allein gelassen wird

Der TN-A-Fall schaerft die Diagnose:
- nicht jeder Geist ist ein leeres `Dokument N`
- mindestens ein Teil der Ghosts scheint echte Akten in einem falschen oder zweiten Word-Kontext zu betreffen

Der naechste Erkenntnissprung soll ueber die neuen `WINWORD`-Prozess-Snapshots kommen.

