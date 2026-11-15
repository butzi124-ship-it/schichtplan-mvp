# Ansichten und Zuordnung zu aktuellem Repository

## Elf zu prüfende Szenen je Gerät = 33 Screenshots

| # | Test-Szenenschlüssel | Fachbereich | Präsentationshinweise | Bestand/Quelle |
|---|---|---|---|---|
| 1 | `login` | Anmeldung | klare Login-Karte; Test-Login nur im isolierten Mock | `schichtplan_mvp.html`, `css/app-shell.css` |
| 2 | `production-overview` | Fertigungsübersicht | KPI, Filter, Tabelle, Detail, Maschinenkarten; keine Fake-KPIs | `js/production-design.js`, `css/production-design.css` |
| 3 | `touch-production` | Touch-Zähler | Maschine, Auftrag, Mitarbeiter, Station und Zähler | `schichtplan_mvp.js`, `css/production-design.css` |
| 4 | `article-tabs` | BA-/Artikelwahl | aktive Artikelkarte, schmale Breite ohne Overflow | `schichtplan_mvp.js`, `css/production-design.css` |
| 5 | `station-tabs` | Spannungs-/Stationswahl | aktive Station unterscheidbar; Touchziele | `schichtplan_mvp.js` |
| 6 | `good-counter` | Gutteile | sichere Plus-/Minus-Schaltflächen und Mengenanzeige | `schichtplan_mvp.js` |
| 7 | `scrap-clarification` | Ausschuss/Abklärung | Rot/Orange, Pflichtursachen-Logik unverändert | `schichtplan_mvp.js`, `js/ba-lifecycle.js` |
| 8 | `6m-dialog` | 6M-Ursache | Dialog vollständig sichtbar/scrollbar, echter maximaler Scroll | `tests/smoke.spec.js`, `css/app-shell.css` |
| 9 | `mandatory-checklist` | Pflichtcheckliste | gesperrte Zustände ehrlich darstellen | `js/ba-lifecycle.js`, `schichtplan_mvp.js` |
| 10 | `completion-ready` | Abschlussbereitschaft | Prüfstatus klar, 52-px-Abschlussziel, keine Umgehung | `css/production-design.css` |
| 11 | `tool-management` | Werkzeugverwaltung | Filter, Bestand, Aktionen, Journal, QR entsprechend Rolle | bestehende Werkzeugansicht im `schichtplan_mvp.js` |

Jede Szene wird in `visual-desktop`, `visual-tablet`, `visual-mobile` aufgenommen, entsprechend den tatsächlich im Repo gepflegten Playwright-Viewports (hier **keine** parallele Definition abweichender Testgrößen).

## Visuelle SVG-Orientierungen
- `reference-desktop.svg`: Shell + Produktionsübersicht mit echten Strukturmustern
- `reference-tablet.svg`: kompakte Shell + Touch-Zähler / 6M-Dialog als Beispiel
- `reference-mobile.svg`: responsive Touch-Bedienung + Dialog-Regel

Die SVGs zeigen **exemplarische** Screens und nicht sämtliche elf Szenen. Die 33 tatsächlichen Screenshot-Baselines müssen aus der **laufenden App mit synthetischen Fixtures** erzeugt, zweimal auf Stabilität geprüft und gesondert freigegeben werden.

## Geplante FABOS-Produktmodule
Die Namen ShiftOS, PlanOS, ToolOS, QualityOS, TrackOS, CostOS, StockOS usw. sind aus dem Produktgespräch abgeleitete **Zukunftskonzepte**. Nur implementierte und erlaubte aktuelle Bereiche rendern. Keine neuen Module, Menüpunkte, Datenbankschemas oder Berechtigungsregeln im Designauftrag anlegen.

## Branding-Schutz
Nicht alle existierenden Humbel-Grafiken und Labels sind generische FABOS-Markenbestandteile. Das Designpaket zeigt neutrale FABOS-Platzhalter, aber **keine stillschweigende Massenumbenennung**. Vor kommerziellem Einsatz Rechte/Marken und die Entfernung firmenspezifischer Grafiken/Daten prüfen.
