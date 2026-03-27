# Release Workflow

Diese Datei ist der verbindliche operative Workflow fuer Build, Publish, Versionierung und GitHub-Releases in `Scola`.

Sie gilt fuer:
- menschliche Mitarbeit
- Codex-/KI-Instanzen

Sie ist bewusst keine Produktdoku, sondern ein Arbeitsvertrag fuer operative Repo-Arbeit.

## Repo und Release-Ziel
- GitHub-Repo: `https://github.com/subsudo/scola_public.git`
- Updater-Quelle: GitHub Releases aus `subsudo/scola_public`
- Release-Asset muss exakt heissen:
  - `Scola.exe`
- Sichtbare Startdatei fuer User bleibt immer:
  - `Scola.exe`

## Schnellregeln
- normale Code-Aenderung = **keine** neue Version
- normaler Commit/Push = **keine** Release
- Test-EXE = bauen, aber **nicht** automatisch veroeffentlichen
- Release nur auf ausdruecklichen Wunsch
- Release-Hinweis immer:
  - kurz
  - deutsch
  - fuer das Update-Fenster geeignet

## Build-Arten
### Normaler Check
Fuer normalen Code-Check:

```powershell
dotnet build .\VerlaufsakteApp.csproj
```

Das ist kein Release-Schritt.

### Lokaler App-Start
Fuer lokale Entwicklung:

```powershell
dotnet run --project .\VerlaufsakteApp.csproj
```

### Verteilbare EXE
Fuer eine testbare oder release-faehige EXE:

```powershell
dotnet publish .\VerlaufsakteApp.csproj -c Release -r win-x64 --self-contained true -o .\publish\<klar benannter ordner>
```

Wichtige Regeln:
- der Publish-Ordner liegt unter:
  - `VerlaufsakteApp\publish\...`
- der Ordnername soll den Zweck klar zeigen, z. B.:
  - `Scola-win-x64-single-20260327-v0.8.4-window-fix`
- fuer Nutzer ist am Ende nur relevant:
  - `Scola.exe`

## Projektdefault fuer Publish
Im Projekt ist bereits als Default verdrahtet:
- one-file
- self-contained
- komprimierter Single-File-Publish

Das heisst:
- Releases sollen weiter als eine verteilbare `Scola.exe` gebaut werden
- keine alternativen Asset-Namen verwenden

## Versionierung
- Version steht in:
  - `VerlaufsakteApp.csproj`
- bei echter Release oder bewusstem Release-Kandidaten:
  - `Version`
  - `AssemblyVersion`
  - `FileVersion`
  gemeinsam anpassen
- wenn nichts anderes gesagt ist:
  - Patch-Version hochzaehlen
  - Beispiel: `0.8.4` -> `0.8.5`

## Exakter Release-Ablauf
Nur wenn der User eine echte Veroeffentlichung will:

1. Version in `VerlaufsakteApp.csproj` erhoehen
2. aenderungen committen
3. `main` pushen
4. Git-Tag setzen:
   - `vX.Y.Z`
5. Release-EXE bauen
6. GitHub Release anlegen oder bewusst aktualisieren
7. `Scola.exe` hochladen
8. kurzen deutschen Release-Hinweis setzen

## Was bei Test-EXEs gilt
Wenn der User nur testen will:
- Version darf fuer den Test bewusst erhoeht werden, wenn das fuer Update-Tests noetig ist
- aber:
  - nicht automatisch releasen
  - nicht automatisch pushen
  - nicht automatisch GitHub-Asset hochladen

Erst nach ausdruecklicher Freigabe:
- pushen
- taggen
- release erstellen/ersetzen

## GitHub-Release-Regeln
- Repo fuer Releases:
  - `subsudo/scola_public`
- Assetname exakt:
  - `Scola.exe`
- Release muss fuer den Updater sichtbar sein:
  - kein Draft
  - kein Pre-release
- jede echte Release braucht eine neue Versionsnummer

## Updater-relevante Regeln
Der Updater in Scola erwartet:
- Repo: `subsudo/scola_public`
- `releases/latest`
- Assetname: `Scola.exe`

Wenn davon abgewichen wird, kann der Update-Flow brechen.

## Was explizit nicht getan werden soll
- keine automatische Versionsaenderung bei normalen Commits
- keine Release bei jedem Build
- keinen anderen Release-Asset-Namen verwenden
- keine GitHub-Release ohne klare User-Anweisung
- keine Release-Notes in langem oder englischem Stil, wenn sie im Update-Fenster erscheinen sollen

## Erwartete Standardentscheidung fuer Codex
Wenn nichts anderes gesagt ist:
- kleine normale Code-Aenderung:
  - nur aendern, testen, committen
- Test-EXE:
  - bauen, Pfad nennen, nicht releasen
- echte Release:
  - Patch-Version hoch
  - commit
  - push
  - Tag
  - Release
  - `Scola.exe`
  - kurzer deutscher Hinweis

## Akzeptanzkriterium fuer diese Datei
Eine neue Codex-Instanz soll nach Lesen von:
- `README.md`
- `AGENTS.md`
- `docs/release-workflow.md`

korrekt unterscheiden koennen zwischen:
- normalem Build
- Test-EXE
- echter Release

und dabei denselben Workflow ausfuehren wie in diesem Projekt bereits etabliert.
