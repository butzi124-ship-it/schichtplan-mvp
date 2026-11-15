# FABOS – verbindliches Designpaket v1.0 (Konzeption)

**Status:** Gestaltungsvorgabe zur Prüfung, keine implementierte Funktion und **keine Freigabe bestehender Screenshot-Baselines**.  
**Repository-Basis:** `design/reference-production-package1` · Humbel-App v0.20.6 · Commit `d7c0c22c6a9ac8563d4ecd48f0b2759d42e524e3` · Draft-PR #45.  
**Verwendung:** Work liest dieses Paket als technische Designanweisung; es ersetzt weder reale Screenshots noch Projektsicherheitsregeln.

## 1. Zweck und Hierarchie

FABOS ist der geplante, herstellerunabhängige Produktname. Der aktuelle Code und seine sichtbare Oberfläche tragen noch teilweise die Bezeichnung „Humbel-App“ beziehungsweise „HUMBEL GEARS“. **Dieses Dokument autorisiert ausdrücklich keine Umbenennung von Quellcode, Daten, Mandanten, Assets oder Authentifizierung.** Eine Produktumbenennung braucht einen separaten Auftrag und Prüfung möglicher Marken-/Namensrechte.

Designrichtung (bereits vereinbart): helles modernes Fertigungs-ERP, dunkelblaue linke Navigation, weiße Oberflächen, blaue Primäraktionen, kontrastreiche Ampel-/Statusfarben, klare Symbole, große Touchziele. Desktop, Tablet und Smartphone müssen dieselben fachlichen Informationen bedienbar machen; die Präsentation darf sich anpassen.

Reihenfolge bei Widersprüchen:
1. `AGENTS.md`, `docs/DEVELOPMENT.md`, Rollen-/Sicherheits-/Geschäftsanforderungen und vorhandene Tests haben Vorrang.
2. Vorliegende **ausdrücklich freigegebene** fachliche und visuelle Anforderungen gelten.
3. Dieses Paket bestimmt die **einheitliche Präsentationsrichtung**; die SVGs zeigen konzeptionelle Layoutmuster, keine pixelgenauen erwarteten Screenshots.
4. Aktuelle DOM-Struktur und vorhandene Fachmodule sind die Wahrheit über real vorhandene Funktionen. Nicht gezeigte oder geplante Funktionen dürfen nicht hinzuerfunden werden.

## 2. Paketdateien und Verwendung

| Datei | Verbindlicher Zweck |
| --- | --- |
| `tokens.json` | Designwerte, responsive Raster, Status- und Touchregeln |
| `COMPONENTS.md` | Wiederverwendbare Muster für Navigation, Tabellen, Formulare, Dialoge und Fertigungszähler |
| `SCREEN-MAP.md` | Zuordnung aller elf visuellen Szenen zu bestehenden App-Bereichen und Tests |
| `ACCEPTANCE.md` | Abnahmekriterien, Schutzregeln und Freigabe der 33 echten Bildreferenzen |
| `WORK-AUFTRAG.md` | Kopierbare, begrenzte Work-Weiterarbeit |
| `reference-desktop.svg` | Konzeptionelle Desktop-Vorschau |
| `reference-tablet.svg` | Konzeptionelle Tablet-Vorschau |
| `reference-mobile.svg` | Konzeptionelle Smartphone-Vorschau |

**SVGs sind visuelle Orientierung, keine UI-Test-Snapshots.** Alle Maß-/Status-/Auftragswerte sind synthetische Beispieldaten und kein Nachweis einer Geschäftslogik.

## 3. Stand und gezielte Integration

Verwendete bestehende Anknüpfungspunkte:
- `css/app-shell.css`: globale Token, Navigation, Fokus, responsive Shell, Login, Karten
- `css/production-design.css`: Produktionsübersicht, BA/OP, Zähler und Dialoge
- `css/home-dashboard.css`: Schnellzugriffe und Start-Dashboard
- `js/app-shell.js`: vorhandene Shell-Präsentation und Tastaturnavigation
- `js/production-design.js`: Fertigungsübersicht, Maschinenkarten, Filter und Detailansicht
- `js/dashboard-home.js`: Startansicht und bestehende Schnellzugriffe
- `schichtplan_mvp.html`: vorhandenes DOM, ohne eigenmächtige Umbenennung
- `tests/smoke.spec.js`: elf Szenen × Desktop, Tablet, Mobil = 33 visuelle Referenzen

Bitte nicht parallel ein neues Komponenten-Framework, neue CSS-Architektur, neue Auth oder neue Geschäftsprozesse bauen. Bestehende CSS-Variablen möglichst weiterverwenden; Änderungsschwerpunkte klein und rückverfolgbar halten.

## 4. Nachweisführung

Der vorherige Stand v0.20.6 bestand **853/853 funktionale Browserprüfungen**, Dokumentation, Design-Display, BA-/OP-Abnahme und Security-Workflows. **Visual Regression blieb rot**, weil die aktuellen Bilder den alten Referenzen widersprechen. Die CI erzeugte **33/33 Review-only-Bildvorschläge**, aber diese sind **nicht freigegeben**.

Für die 33 originalgetreuen Screenshots siehe GitHub Actions: https://github.com/butzi124-ship-it/schichtplan-mvp/actions/runs/43726264797/artifacts/14166521205

**Keine automatische `--update-snapshots`-Freigabe.** Erst nach stabiler doppelter Aufnahme, Sichtkontrolle und ausdrücklicher Nutzerfreigabe neue Referenzen versionieren. Die drei SVGs sind **nicht** an Stelle der echten 33 Playwright-Referenzen zu verwenden.

## 5. Offene Entscheidungen

- FABOS-Branding statt bisheriger Humbel-Bezeichnungen: gesonderte Entscheidung und rechtliche Namensprüfung.
- Pixelgenaue neue Baselines: erfordern gesonderte, ausdrückliche Freigabe.
- Geplante Module wie CostOS, PlanOS oder BuyOS sind Produktkonzepte und **nicht** als bestehende Menüfunktionen zu zeichnen oder zu implementieren.
