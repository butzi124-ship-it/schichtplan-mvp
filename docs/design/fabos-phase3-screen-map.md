# FABOS Phase 3 – tatsächliche visuelle Szenenzuordnung

Status: Vorbereitung der Bildprüfung; keine Freigabe von Referenzen, Merge oder Deployment.
Basis: v0.20.9, PR #48, `d7c6cbb3a752221d83cba36377563b8bed044547`.
Designvorgabe: PR #46, `design/fabos-reference-v1`, `6e1ba5eafac204b5aa9b15a8cf6ce3c1203ca6ba`.

Dieses eigenständige Dokument präzisiert die Zuordnung auf dem Phase-3-Branch. Das separate `docs/design/fabos-v1/SCREEN-MAP.md` und sämtliche übrigen Dateien aus PR #46 werden weder kopiert noch geändert. Seine spätere Aktualisierung ist separat vorzuschlagen.

## Aufnahmeziele – elf bestehende Szenen auf drei Geräten

| Szene | Tatsächliche Navigation und Aufnahmeziel | Erfasste Bedienelemente und Grenze |
| --- | --- | --- |
| `login` | Startseite, ganze Seite | Loginfelder und isolierte Testzugänge; kein produktiver Auth-Test |
| `production-overview` | Produktion → Stückzahl; `#view` mit Region Fertigungsübersicht | Alle fünf KPI-Karten, Abteilung/Maschine/Status/Suche, BA-Tabelle, Auftragsvorschau und Maschinenübersicht; keine einzelne BA öffnen |
| `touch-production` | Produktion → Aufträge / BA → BA-100 öffnen → Produktionsansicht; `#view` | Maschine, BA-Auswahl, Mitarbeiter, Fortschritt, Station, Mengenzähler und Abschlussbutton |
| `article-tabs` | Gleicher Touchweg; Tablist Aktive Aufträge | Tatsächlich BA-/Auftragswahl, nicht Artikelstammbearbeitung; Auswahl und aktive Markierung |
| `station-tabs` | Gleicher Touchweg; Tablist Spannungsauswahl | Stationswahl und aktive Markierung |
| `good-counter` | Gleicher Touchweg; Gutteile-Zählerbereich | Gutmenge und Plus/Minus |
| `scrap-clarification` | Gleicher Touchweg; `.production-touch-station` | Stationskarte einschließlich Ausschuss und In Abklärung, Plus/Minus und Mengen; kein separater QualityOS-Bereich |
| `6m-dialog` | Ausschuss +1 im Touchweg; letztes Overlay `.fixed.inset-0.z-50` | Ursachenwahl, Notiz, Abbrechen und Speichern im scrollbaren Dialog; bestehende Scrollkamera unverändert |
| `mandatory-checklist` | Auftrag abschließen im Touchweg; letztes Overlay | Zehn Checklistenfelder und gesperrte endgültige Fertigmeldung; keine Bestätigung ausführen |
| `completion-ready` | Synthetischer sequenzieller Auftrag, geprüfte Checkliste; letztes Overlay | Zehn Felder und aktivierte endgültige Fertigmeldung; keine Bestätigung ausführen |
| `tool-management` | Werkzeugverwaltung → Werkzeugliste; `#view` | Werkzeugbestand, vorhandene Filter und Listenaktionen; nicht sämtliche Werkzeug-Unterbereiche |

Viewports stammen ausschließlich aus `playwright.visual.config.js`: Desktop Chrome 1280×720, Tablet 1024×768 mit Touch, Smartphone Pixel 5 mit 390×844 CSS-Pixeln. Mobile PNGs verwenden den Device Scale Factor der vorhandenen Pixel-5-Konfiguration. Keine Viewports oder Bildtoleranzen geändert.

## Nachweis der fünf KPI-Karten

Die bisherige Szene `production-overview` öffnete BA-100 und fotografierte die einzelne Produktionsvorschau. Das widersprach der Designzuordnung: dort sind die KPI-Karten nicht vorhanden.

Die korrigierte Szene öffnet die tatsächliche Region Fertigungsübersicht aus `js/production-design.js`, ohne einen Auftrag zu öffnen. Vor dem Bildvergleich prüft sie Überschrift, genau fünf `.h-kpi`-Karten, die exakten Bezeichnungen **Aktive BA**, **Laufend**, **Pausiert**, **Maschinen**, **BA mit Abklärung**, sichtbare Karten und numerische Werte sowie Filterformular und Auftragsvorschau. Das Aufnahmeziel `#view` enthält die vollständige Region. Alle Zahlen stammen aus denselben synthetischen Mockzeilen wie bisher; keine neuen Daten oder Kennzahlberechnungen.

## BA-/OP-Abdeckung und Abnahmeumfang

`tests/design-reference.spec.js` bleibt bytegleich zu v0.20.9. Seine beiden funktionalen Fälle prüfen weiterhin echte Übersicht/Filter/Detailauswahl und BA-Anlage mit unveränderlichem OP-Snapshot über eine isolierte PGlite-Testdatenbank. Auch sämtliche übrigen BA-/OP- und 251 ursprünglichen Regressionstests bleiben erhalten. Die visuelle Zuordnung ändert keine fachlichen Assertions oder Testfilter.

33 Bilder reichen für die elf oben benannten Zustände aus, nicht für eine vollständige Designabnahme aller App-Bereiche. Insbesondere fehlen strikte Szenen für BA-Anlage und OP-Snapshot, das Start-Dashboard sowie Werkzeugjournal/QR-Unterbereiche. Vorhandene funktionale Zusatzscreenshots ersetzen dort keinen strikten Baselinevergleich. Nur falls diese Bereiche mit abgenommen werden sollen, werden zunächst BA-Anlage und OP-Snapshot (je drei Geräte, sechs zusätzliche Szenenbilder) sowie separat das Start-Dashboard (drei) vorgeschlagen. Keine neuen Szenen wurden hinzugefügt.

Die ersetzte Produktionsübersichtsaufnahme ist eine begründete Korrektur des Aufnahmeziels, keine Freigabe einer neuen Baseline. Der Vergleich zur alten BA-Vorschau zeigt deshalb bewusst eine größere Differenz.

## Doppelte Aufnahme und bereitgestellte Evidenz

Der unveränderte gepinnte Container `mcr.microsoft.com/playwright:v1.62.1-noble` führt dieselben 33 strikten Vergleiche zweimal aus. Der erste Bildsatz wird vor dem zweiten Lauf kopiert, da Playwright das Ausgabeverzeichnis beim Start ersetzt. Die Aufnahmen erhalten keine automatische Referenzfreigabe.

Im visuellen CI-Artefakt liegt unter `artifacts/visual/reference-proposals/`:

- `first/` und `second/`: je 33 aktuelle Bilder, Kopien der alten Referenzen, Differenzbilder und Manifest mit Datei-/Pixelhashes, Größe, Version und tatsächlichem Checkout-SHA.
- `stability.json`: alle 33 Vergleiche zwischen erstem und zweitem Lauf, einschließlich unveränderter Referenzhashes. Nur vollständige, pixelidentische Aufnahmen desselben Stands melden `STABLE_33_OF_33`. Fehlende/mehrdeutige Aufnahmen oder Abweichungen bleiben Fehler.
- `index.html`: pro Gerät/Szene bisherige Referenz, aktuelles Bild aus Lauf 1 und 2 und Differenz zur Referenz nebeneinander, Originalbilder anklickbar.

Ein stabiler neuer Bildsatz ist kein bestandener Vergleich zur alten Referenz und keine Designfreigabe. Nach CI-Abschluss sind Bildsichtung aller 33 Szenen, gegebenenfalls ausdrücklich beauftragte Korrekturen und eine gesonderte ausdrückliche Referenzfreigabe erforderlich. Merge und Deployment bleiben getrennte Entscheidungen.
