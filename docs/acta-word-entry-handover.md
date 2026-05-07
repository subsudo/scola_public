# Acta Word-Eintrag Handover

Stand: 2026-05-07

Dieses Dokument beschreibt, wie Acta die Scola-Logik fuer interaktive Word-Eintraege uebernehmen soll. Es ist bewusst kein Auftrag, Scola umzubauen. Scola bleibt in seiner aktuellen Struktur bestehen; Acta darf intern sauberer und generischer bauen, solange das Verhalten identisch bleibt.

## Ziel
Acta soll interaktive Eintragsaktionen fuer diese Ziele anbieten:

- `Eintrag BU`
- `Eintrag BI`
- `Eintrag BE`
- `Eintrag LB`

Fachlich ist das immer derselbe Ablauf:

1. Teilnehmer / Akte in der UI bestimmen.
2. finalen `docPath` vor jeder Word-Operation aufloesen.
3. Word-Operation ueber einen zentralen `WordStaHost` ausfuehren.
4. Tabelle am Ziel-Bookmark finden.
5. neue Zeile oben im Datenbereich einfuegen.
6. Clipboard-Inhalt oder vorbereitete Fallback-Zeile eintragen.
7. Cursor in die sinnvolle Bearbeitungszelle setzen.
8. bei interaktivem Einzelklick das konkrete Word-Fenster nach vorne holen.

## Scola bleibt unveraendert
Scola wird fuer diesen Port nicht refactored.

Wichtig:

- Acta soll sich nicht daran stoeren, dass Scola historisch einzelne Config-Felder und Handler nutzt.
- Acta darf eine generische Zielstruktur verwenden.
- Gleichheit bedeutet hier: gleiches Verhalten und gleiche Stabilitaetsregeln, nicht identischer Codeaufbau.

## Relevante Scola-Quellen
Die wichtigsten Referenzstellen in Scola:

- `Services/WordStaHost.cs`
  - zentraler STA-Worker fuer alle Word-COM-Operationen
- `Services/WordService.cs`
  - `InsertClipboardToTable(...)`
  - `InsertTextRowToTable(...)` als Batch-Referenz
  - `CreateOrAttachWordApplication(...)`
  - `OpenOrGetDocument(...)`
  - `CloseTransientEmptyDocuments(...)`
  - `EnsureWordUiState(...)`
  - gezielter Foreground-Fix ueber Word-Fenstertitel
- `Services/WordTemplateValidationException.cs`
  - fachliche Fehler fuer fehlendes Bookmark / ungueltige Tabelle
- `MainWindow.xaml.cs`
  - `InsertEntryButton_OnClick(...)`
  - `InsertEntryBiButton_OnClick(...)`

## Zielarchitektur fuer Acta
Acta sollte den Scola-Grundsatz uebernehmen:

- UI-Code orchestriert nur.
- Kein Word-COM direkt im UI-Thread.
- Alle Word-Auftraege laufen seriell ueber einen app-weiten `WordStaHost`.
- Auf dem STA-Thread lebt genau eine `WordService`-Instanz.
- Exceptions muessen typstabil zur UI zurueckkommen.

Der UI-Teil bleibt verantwortlich fuer:

- Dokumentpfad-Aufloesung
- Busy-Guard
- Wait-Cursor / Status
- ReadOnly-Dialog, falls Acta diesen Flow nutzt
- nutzernahe Fehlermeldungen

Der Word-Service bleibt verantwortlich fuer:

- Word attachen oder erzeugen
- Dokument oeffnen oder bereits offenes Dokument wiederverwenden
- Dokument-Lock erkennen
- Ziel-Tabelle am Bookmark validieren
- Zeile einfuegen / bei Fehler rollbacken
- Cursor setzen
- Word sichtbar machen
- konkretes Word-Fenster bei interaktiven Aktionen nach vorne holen
- COM-Objekte im `finally` freigeben

## Acta-Zielstruktur
Acta sollte die vier Eintragsziele als Konfiguration modellieren, statt vier fast identische Methoden zu bauen.

Empfohlene Form:

```csharp
internal sealed class StructuredEntryTarget
{
    public required string Key { get; init; }              // BU, BI, BE, LB
    public required string Label { get; init; }            // Eintrag BU, ...
    public required string TableBookmarkName { get; init; }
    public int FirstDataRowIndex { get; init; } = 2;
    public int ExpectedColumnCount { get; init; } = 4;
    public bool BringToForeground { get; init; } = true;
}
```

Das ist eine Acta-Abstraktion. Scola muss dafuer nicht angepasst werden.

## Ziel-Bookmarks
Scola kennt aktuell produktiv:

| Ziel | Scola-Config | Default |
| --- | --- | --- |
| BU-Eintrag | `WordBookmarkName` | `BU_BILDUNG_TABELLE` |
| BI-Eintrag | `WordBiTableBookmarkName` | `BI_BERUFSINTEGRATION_TABELLE` |

Scola nutzt ausserdem reine Sprung-Bookmarks:

| Bereich | Scola-Config | Default |
| --- | --- | --- |
| BU oeffnen | `WordBuBookmarkName` | `_Bildung` |
| BI oeffnen | `WordBiBookmarkName` | `_Berufsintegration` |
| BE oeffnen | `WordBeBookmarkName` | `_Beratung` |

Fuer Acta muessen die Tabellen-Bookmarks fuer `BE` und `LB` fachlich bestaetigt werden. Diese Namen nicht aus den Scola-Sprung-Bookmarks ableiten, ohne die Vorlage zu pruefen.

Empfohlene Acta-Konfiguration:

```csharp
BU: TableBookmarkName = "BU_BILDUNG_TABELLE"
BI: TableBookmarkName = "BI_BERUFSINTEGRATION_TABELLE"
BE: TableBookmarkName = "<in Acta/Vorlage bestaetigen>"
LB: TableBookmarkName = "<in Acta/Vorlage bestaetigen>"
```

## Eintragslogik
Acta sollte eine gemeinsame Methode fuer alle Ziele verwenden:

```csharp
bool InsertClipboardToStructuredEntryTable(
    string docPath,
    StructuredEntryTarget target,
    string? preReadClipboardText,
    string[]? fallbackFieldsWhenClipboardInvalid);
```

Verhalten wie Scola:

- Clipboard wird vor oder innerhalb der Word-Operation gelesen.
- Gueltig ist genau eine Zeile mit vier tab-getrennten Spalten.
- Bei gueltigem Clipboard werden alle vier Spalten eingefuegt.
- Bei leerem oder ungueltigem Clipboard wird trotzdem eine neue Zeile vorbereitet.
- Wenn Fallback-Felder vorhanden sind, werden diese vier Werte eingetragen.
- Wenn keine Fallback-Felder vorhanden sind, bleibt die Zeile leer.
- Bei Fehlern waehrend des Schreibens wird die teilweise eingefuegte Zeile wieder geloescht.
- Nach erfolgreichem Einfuegen wird der Cursor gesetzt:
  - bei gueltigem Clipboard bevorzugt Spalte 1
  - bei leerem/ungueltigem Clipboard bevorzugt Spalte 3
  - falls diese Spalte nicht existiert, sichere Ersatzspalte verwenden

## Tabellenvalidierung
Vor dem Schreiben muss Acta dieselben fachlichen Checks machen:

- Bookmark existiert.
- Bookmark liegt in einer Tabelle.
- Tabelle hat mindestens die erwartete Spaltenzahl.
- Ziel-Datenbereich beginnt bei `FirstDataRowIndex`.

Bei Fehlern keine generische COM-Exception in der UI zeigen, sondern fachliche Fehler:

- Bookmark fehlt
- Tabelle am Bookmark ungueltig
- Dokument gesperrt / nicht schreibbar
- Word nicht installiert

Scola-Referenz:

- `WordTemplateValidationException`
- `ResolveStructuredEntryTableForWrite(...)`
- `GetContainingBookmarkTable(...)`
- `GetSafeEditColumn(...)`
- `TryDeleteRow(...)`

## Foreground-Regel
Interaktive Eintragsaktionen muessen Word nach vorne holen.

Acta soll dabei die neue Scola-Regel uebernehmen:

- kein Word-Placement
- kein `ActiveWindow.Hwnd`
- kein `Process.MainWindowHandle` als Fokusquelle
- kein automatisches Schliessen von Ghost-Fenstern
- nach Cursor-Positionierung gezielt das konkrete Word-Fenster suchen
- Suche ueber `EnumWindows` und Titelmatch auf den Akten-Dateinamen
- `AllowSetForegroundWindow(wordPid)` verwenden
- bei Bedarf kurzer `AttachThreadInput`-Fallback

Reihenfolge ist wichtig:

1. Dokument oeffnen / aktivieren.
2. Tabelle finden.
3. Zeile einfuegen.
4. Cursor/Selection in Zielzelle setzen.
5. erst danach das konkrete Word-Fenster nach vorne holen.

Nur so kann der User nach `Eintrag BU` / `Eintrag BI` / `Eintrag BE` / `Eintrag LB` direkt tippen.

## Batch-Abgrenzung
Dieses Handover beschreibt primaer interaktive Eintragsbuttons.

Wenn Acta spaeter Batch-Eintraege bekommt:

- dieselbe Tabellenlogik verwenden
- aber `BringToForeground = false`
- kein Fokus-Sprung pro Zeile
- kein Fokus-Sprung am Ende
- Fehler in der Batch-Ergebnisliste anzeigen

## Was Acta nicht uebernehmen soll
Nicht aus Scola kopieren:

- Scola-Kachel-Layout
- Scola-Mini-Stundenplan
- Scola-Hinweis-Anzeige
- BI-To-do-Sammeldokument, ausser Acta braucht exakt diesen Sonderfall
- alte Word-Placement-Logik
- automatische Ghost-Cleanup-Experimente
- parallele Word-COM-Aufrufe aus mehreren Threads

## Failure Modes
Acta soll Fehler bewusst behandeln:

- Word nicht installiert:
  - klare Meldung, keine COM-Rohmeldung
- Dokument nicht gefunden:
  - UI-Meldung mit Teilnehmer / Pfadkontext
- Dokument gesperrt:
  - bestehenden Acta-ReadOnly-Fallback beibehalten, falls vorhanden
- Bookmark fehlt:
  - Zielname nennen, Vorlage pruefen lassen
- Tabelle ungueltig:
  - Zielname nennen, keine Zeile einfuegen
- Clipboard blockiert:
  - kurzer Retry wie in Scola, danach klare Meldung
- Word-Foreground nicht moeglich:
  - Eintrag bleibt trotzdem geschrieben, nur loggen

## Testfaelle fuer Acta
Nach dem Port muessen mindestens diese Tests gruen sein:

1. `Eintrag BU` mit gueltigem Clipboard:
   - vier Spalten werden eingefuegt
   - Word kommt nach vorne
   - Cursor steht in der erwarteten Zelle
   - sofortiges Tippen landet im Word-Dokument

2. `Eintrag BI` mit gueltigem Clipboard:
   - gleicher Test im BI-Tabellenziel

3. `Eintrag BE`:
   - gleicher Test im bestaetigten BE-Tabellenziel

4. `Eintrag LB`:
   - gleicher Test im bestaetigten LB-Tabellenziel

5. Leeres Clipboard:
   - neue Zeile wird vorbereitet
   - Fallback-Felder werden verwendet, falls aktiviert
   - Cursor steht in der Eingabezelle

6. Ungueltiges Clipboard:
   - keine halbe Datenzeile
   - vorbereitete Zeile wie bei leerem Clipboard

7. Fehlendes Bookmark:
   - fachliche Fehlermeldung
   - keine neu eingefuegte Zeile

8. Ungueltige Tabelle:
   - fachliche Fehlermeldung
   - keine neu eingefuegte Zeile

9. Gesperrte Akte:
   - bestehender Acta-Lock-/ReadOnly-Flow bleibt korrekt

10. Bereits offene Akte:
    - wird wiederverwendet
    - kein zweites unnötiges Dokumentfenster

11. Mehrere Word-Dokumente offen:
    - Foreground trifft das konkrete Aktenfenster, nicht ein beliebiges Word-Fenster

12. Mehrere schnelle Klicks:
    - UI-Busy-Guard verhindert parallele Word-Operationen
    - WordStaHost bleibt seriell

## Minimaler Acta-Port in Schritten
1. App-weiten `WordStaHost` einfuehren oder bestaetigen.
2. bestehenden Acta-Word-Service auf STA-Host umhaengen.
3. `StructuredEntryTarget` fuer BU/BI/BE/LB definieren.
4. gemeinsame `InsertClipboardToStructuredEntryTable(...)` bauen.
5. BU und BI portieren und testen.
6. BE und LB mit bestaetigten Tabellen-Bookmarks aktivieren.
7. Foreground-Fix fuer interaktive Eintraege einschalten.
8. Fehler- und Lock-Flows gegen die Tests pruefen.

## Kurzfassung
Acta soll nicht Scola-Code blind kopieren. Acta soll die stabile Scola-Word-Architektur uebernehmen:

- Word-COM nur ueber einen zentralen STA-Host
- generische Eintragsziele fuer BU/BI/BE/LB
- ein gemeinsamer Tabellen-Eintragspfad
- interaktive Eintraege bringen das konkrete Word-Fenster nach vorne
- Batch/automatische Aktionen bringen Word nicht nach vorne
- keine Word-Positionierung
- keine Ghost-Fenster automatisch schliessen
