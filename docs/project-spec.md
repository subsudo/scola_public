# Project Spec

## Produktziel
Scola unterstuetzt das schnelle und robuste Arbeiten mit Teilnehmer-Akten in Word.

Im Zentrum stehen vier Hauptaufgaben:
1. Teilnehmende aus Rohtext oder Listen erkennen
2. Teilnehmende gegen Aktenordner matchen
3. die richtige Verlaufsakte oeffnen oder an BU / BI / BE anspringen
4. Eintraege einzeln oder im Batch in Word-Tabellen einfuegen

Die App ist bewusst eine lokale Arbeitsanwendung fuer reale Alltagsablaeufe, nicht ein generisches Dokumentmanagement-System.

## Primäre Use Cases
- Anwesenheits- oder Rohlisten aus externen Systemen einfuegen
- aktive Teilnehmende visuell pruefen und bei Bedarf manuell korrigieren
- schnell Ordner oder Akten oeffnen
- direkt zum BU-, BI- oder BE-Bereich springen
- BU- oder BI-Eintrag einfuegen, auch wenn das Clipboard leer oder ungueltig ist
- mehrere Eintraege positionsbasiert im Batch verarbeiten
- Odoo-Link aus der Akte oeffnen
- Kuersel unterhalb des Namens anzeigen und ueber Kuersel suchen/matchen
- Word bei Bedarf maximiert auf einem bevorzugten Monitor oeffnen

## Nicht-Ziele
Aktuell bewusst nicht im Fokus:
- Mehrbenutzer-Synchronisation innerhalb der App
- serverseitige Datenhaltung fuer UI-Zustaende
- MVVM-Framework oder groessere Architektur-Migration
- direkte Odoo-API-Integration als Hauptdatenquelle
- generisches Dokumentmanagement ausserhalb der Verlaufsaktenlogik

## Produktprinzipien
- workfloworientiert statt abstrakt
- lokale Kontrolle statt verteilter Infrastruktur
- robuste Alltagsfunktion wichtiger als Perfektion
- Word bleibt Kernbestandteil des Produkts
- UX ist wichtig, aber ohne grundsaetzlichen Redesign-Anspruch

## Fachliche Kernregeln
- Matching betrachtet nur die erste Ebene der konfigurierten Serverpfade.
- `Anwesend` und `Verspaetet` gelten als praesent.
- `Abwesend (...)` gilt als nicht praesent.
- Batch-Zuordnung ist aktuell bewusst positionsbasiert.
- Word-Eintrag erfolgt in definierte Tabellen/Bookmarks.
- Wenn Clipboard-Inhalt ungueltig oder leer ist, wird trotzdem eine neue Zeile vorbereitet.

## Konfiguration
Wichtige globale Konfiguration in `settings.json`:
- primaerer, sekundaerer, tertiaerer Serverpfad
- Keyword fuer Verlaufsakten-Dateien
- Bookmark-Namen fuer BU / BI / BE und Tabellen

Wichtige benutzerspezifische Prefs in `user-prefs.json`:
- Theme
- Darstellungsdichte
- sichtbare Buttons
- Kuerselanzeige
- Debug-Logging
- Word-Monitor / maximiertes Oeffnen
- Prefill-Verhalten fuer leeres Clipboard

## Externe Abhaengigkeiten
- Microsoft Word via COM
- Windows-Dateisystem / Netzlaufwerke
- Explorer fuer Ordneroeffnung
- Standardbrowser fuer Odoo-Links

## Testfaehigkeit
Ein lokales Mock-Testsetting ist vorhanden:
- `TestSetup\New-MockTestSetup.ps1`
- `TestSetup\mock-env\...`

Zusaetzlich sind echte Produktionslogs fuer Word-/Parser-Probleme weiterhin eine zentrale Diagnosequelle.

## Offene Punkte / Annahmen
- Annahme: Die aktuelle Ordnerstruktur und Bookmark-Konventionen bleiben vorerst fachlich stabil.
- Offener Punkt: Word-/Office-Verhalten kann je nach Umgebung leicht variieren.
- Offener Punkt: Einige historische Debug-Dokumente beschreiben aeltere Zwischenstaende und muessen bei Detailfragen mit dem Code abgeglichen werden.
- Offener Punkt: Mehrere gleichzeitig offene Word-Fenster/Instanzen bleiben ein produktiver Beobachtungspunkt.
