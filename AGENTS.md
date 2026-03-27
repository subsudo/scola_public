# AGENTS.md

Dieses Projekt ist `Scola`, eine lokale WPF-.NET-8-App fuer den operativen Alltag mit Teilnehmer-Akten.

## Lesereihenfolge
Vor jeder inhaltlichen Aenderung zuerst lesen:
1. `README.md`
2. `docs/project-spec.md`
3. `docs/architecture.md`
4. `docs/status.md`
5. `docs/decisions.md`
6. `docs/release-workflow.md`

Danach bei Bedarf als Referenz:
- `HANDOVER.md`
- `BUGFIXES.md`
- `MUST_DEBUG.md`
- `TESTING.md`
- `FEATURE_REQUESTS.md`
- `AGENT_TRANSFER_ODDO_CACHE_AKTEX.md`
- `WORD_WINDOW_POSITION_FIX_TRANSFER.md`

## Arbeitsregel
Erst Kontext lesen, dann handeln.

Vor Aenderungen immer klaeren:
- welcher reale Workflow betroffen ist
- ob Word-/Parser-/Matching-Logik beruehrt wird
- welche bestehende Doku mitgepflegt werden muss

## Hard Constraints
- Keine grosse Architektur-Migration.
- Kein MVVM-Umbau als "Nebenbei-Fix".
- Keine generische Neu-Design-Idee ueber bestehende UX stuelpen.
- Keine Laufzeitdaten neben der EXE einfuehren.
- Word nicht auf eine voellig andere Integrationsstrategie umstellen.
- Keine stillen Aenderungen an kritischer Logik ohne Dokumentationsnachzug.
- Bestehende historische Markdown-Dateien nicht loeschen.

## Current Preferences
- Pragmatisch, produkt- und workfloworientiert.
- Robuste Alltagsfunktion vor Perfektion.
- Kleine, nachvollziehbare Aenderungen.
- UX zaehlt; visuelle Sprache von Scola erhalten.
- Lokale, kontrollierbare Loesungen bevorzugen.
- Wichtige Fehler sichtbar und im App-Stil behandeln.
- Kontext kuenftig primaer in `docs/` pflegen.

## Exploration Allowed
Exploration ist erlaubt, aber gezielt:
- Code lesen
- bestehende Logs und Debug-Dokumente auswerten
- Mock-Setup unter `TestSetup` nutzen
- Builds lokal pruefen

Nicht ohne guten Grund:
- grosse Umstrukturierungen
- breitflaechige Refactors
- mehrere sensible Bereiche gleichzeitig anfassen

## Sensible Bereiche
Bei Aenderungen hier besonders vorsichtig sein:
- `MainWindow.xaml.cs`
- `Services/WordService.cs`
- `Services/ParticipantParser.cs`
- `Services/FolderMatcher.cs`
- `SettingsWindow.xaml(.cs)`

## Dokumentationspflicht
Wenn sich Verhalten, Architektur, Workflows oder aktuelle Entscheidungen aendern:
- `docs/status.md` aktualisieren
- `docs/decisions.md` aktualisieren, falls eine bewusste Entscheidung betroffen ist
- bei groesseren Aenderungen auch `README.md` oder `docs/architecture.md` nachziehen

## Build- und Release-Regeln
- Normale Commits und Pushes erzeugen keine Release.
- Version nur bei echter Release oder bewusstem Release-Kandidaten aendern.
- Release nur auf ausdrueckliche Anweisung.
- Release-Asset immer exakt `Scola.exe`.
- GitHub-Repo fuer Releases ist `subsudo/scola_public`.
- Fuer operative Details immer `docs/release-workflow.md` verwenden.

## Erwarteter Aenderungsstil
- Kleine Diffs bevorzugen.
- Vorhandene Muster weiterverwenden.
- Annahmen klar markieren.
- Reale Produktfaelle und echte Logs wichtiger nehmen als theoretische Schoenheit.
