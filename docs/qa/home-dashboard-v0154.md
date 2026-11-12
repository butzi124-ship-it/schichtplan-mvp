# Startseite: QA-Integrität auf PR #27 / 0.15.4

Auftrag vom 12.11.2026. Bestehender Draft #27 wird auf demselben Branch `feat/home-dashboard-v0153` ergänzt. Verbindlicher Ausgangscommit `7dc05acd4d60dea62331a268438df6d7f744a5fa`, kumulative Basis #26 `cf7d3bf2b06e895e72e698857327e1486cadbf4d`. Refs #17, #25, #26, #27; keine neue Entwicklungskette. APP_VERSION **0.15.4**, VERSION_LOG **2026-11-12 06:18 Europe/Berlin**, tatsächlicher Zeitpunkt der Codekorrektur.

## Nachgewiesene Fehlerursache

[Actions-Lauf 42469909314](https://github.com/butzi124-ship-it/schichtplan-mvp/actions/runs/42469909314), Job QA Integrity `128698023265`: Die vor allen Schutzprüfungen noch nicht normalisierte Dashboard-Einbindung ergibt `d171a8029102831d2c8fcfbde87d55a31a45e2276c20fcf424184fd021fc284f` statt des geschützten ursprünglichen Hashes `a6de03aa3ce6a007a720604eaefb9a837783c4379a49fe44329271f3cb9367e1`. Lokal am unveränderten Ausgangscommit mit derselben Assertion und denselben Hashes reproduziert.

Nach exakter Rückführung der vier Dashboard-Einbindungen wurde ein zweiter, zuvor verdeckter Integritätsfehler sichtbar: Die zwei Asset-Tags aus #27 ändern auch die geschützte HTML-Datei. Auch diese Einbindungen werden jetzt exakt rückgeführt, bevor die bisherigen Moduladapter und der ursprüngliche HTML-Hash geprüft werden.

Im historischen Lauf bestanden Visual Regression zweimal sowie die separate BA/OP Isolated Acceptance. Functional Browsers war **cancelled**, Quality Gate und e2e folgerichtig rot; das ist kein erfolgreicher vollständiger Browsernachweis. Die neuen GitHub-Actions-Ergebnisse werden nach Abschluss dieses Auftrags nicht abgefragt.

## Eng begrenzte geprüfte Präsentationsausnahmen

`tests/qa/dashboard-presentation-patches.json` dokumentiert den tatsächlichen Diff #26 → #27 als Before/After-Paare. Jede Ausnahme muss exakt **einmal** vorkommen; fehlende, doppelte oder veränderte Hooks scheitern. Keine beliebige Funktions-/Bereichsentfernung, keine Regex-Ausblendung von Geschäftslogik.

| Ausnahme | Fachliche Prüfung |
|---|---|
| Body-Klasse nach bestehender autorisierter Tabauflösung | Reine CSS-Aktivierung; bestehende Rollenauswahl und Tabkorrektur bleiben unverändert. |
| Dashboard in bestehender Vollbreiten-Wrapperauswahl | Reine Präsentation, keine Änderung von Werkzeug-/Lagerfunktionen. |
| Übergabe von `tabs` an `renderDashboard` | Bereits berechnete erlaubte Navigation wird verwendet; keine neue Rechteentscheidung. |
| Übersichtszweig vor dem vorhandenen Dashboard-Renderer | Ausschließlich Leseselektion und HTML-Ausgabe; alle bestehenden Untertabs und Fallbacks bleiben erhalten. |
| Zwei HTML-Tags für CSS und Renderer | Nur die exakten Asset-Einbindungen; ursprünglicher geschützter HTML-Hash bleibt erhalten. |

Die komplette normalisierte App aus #27 wurde zusätzlich direkt mit #26 verglichen: Nach Rückführung dieser vier Änderungen und der bereits erlaubten Version/Log-Normalisierung identisch. Die vorhandenen Moduladapter sowie alle ursprünglichen Hashes in `protected-files.json` bleiben unverändert.

Die drei tatsächlich geprüften Präsentationsdateien `js/dashboard-home.js`, `css/home-dashboard.css` und `schichtplan_mvp.html` erhalten zusätzlich exakte SHA-256-Prüfungen. Das sind neu hinzugefügte Schutzprüfungen der unveränderten #27-Dateien, kein Austausch bisheriger erwarteter Hashes. Künftige Änderungen benötigen eine erneute ausdrückliche Prüfung.

## Daten, Rollen und Erreichbarkeit

- Der Renderer verwendet nur die vorhandenen sichtbaren Maschinen und Stückzahleinträge aus `getVisibleProductionMachines` / `getVisibleProductionCounts`. Abteilungsleiter bleiben auf ihre Abteilungen beschränkt; ohne vorhandenen Produktionstab werden keine entsprechenden Daten übergeben.
- Mitarbeiter erhalten nur ihre zugeordneten Schichten aus dem vorhandenen Generator und ihre offenen Aufgaben aus dem bisherigen Aufgabenfilter. Admin erhält alle offenen Aufgaben. Keine Änderung des Schichtgenerators oder der Aufgabenrechte.
- Maschinenstatus und Tagesfortschritt bleiben ausdrücklich als nicht verfügbar ausgewiesen; die Zahl sichtbarer Stückzahleinträge wird nicht als heutige Produktionsmenge ausgegeben. Keine erfundenen Produktionsdaten oder Live-Zustände.
- Alle vier Dashboard-Unterbereiche bleiben erhalten. Der bisherige Renderer bleibt bei fehlendem Präsentationsmodul verfügbar. Die neue Darstellung und die vorhandenen Moduldateien bleiben bytegleich zu #27.
- Dunkelblaue Navigation, helle Karten, blaue Akzente, große Schnellzugriffe und vorhandene responsive CSS-Regeln wurden im Quelltext geprüft. Eine aktuelle visuelle/Pointer-Abnahme auf Desktop, Tablet und Smartphone konnte lokal nicht ausgeführt werden.

Neue Node-Regressionen führen die **tatsächlichen App-Adapter und Selektoren** mit synthetischen Eingaben aus: Admin/Abteilungsleitung/Mitarbeiter, Aufgaben-/Schichtfilter, Datenunveränderlichkeit, Untertabs/Fallback, Escaping und fehlende Fantasiezustände. Manipulationsprüfungen erfassen entfernte/doppelte/veränderte Hooks, fremde Geschäftslogikänderungen und geänderte/fehlende HTML-/CSS-/JS-Dateien.

Fünf neue Browserfälle sind additiv in allen fünf bisherigen Projekten eingebunden: Layout/Navigation, Mitarbeiter, Abteilungsleitung, Scanner und fehlendes Präsentationsmodul. 733 bisherige Fälle plus 25 neue ergeben 758; keine bisherigen Filter verengt, Fälle/Assertions entfernt, Skips oder Bildreferenzänderungen eingeführt. Diese neuen Browserfälle sind mangels Laufzeit noch nicht fachlich ausgeführt.

## Lokale Prüfungen

Die Ergebnisse beziehen sich auf den lokalen finalen Dateistand; GitHub-CI-Erfolge für den neuen Commit werden nicht vorweggenommen.

| Prüfung | Ergebnis |
|---|---|
| `npm ci --ignore-scripts` | PASS |
| Beide Pflicht-`node --check`, `check:syntax`, zusätzliche Syntaxprüfung neuer JS-Dateien | PASS |
| `test:ci` | PASS, 2/2 |
| `test:qa` | PASS, 16/16, davon 7 neue Dashboard-/Schutzprüfungen; keine Skips |
| `qa:check` | PASS, ursprünglicher Geschäftslogik-/HTML-Hash, 251 Legacy-Identitäten und 33 unveränderte Bildreferenzen |
| `test:integration` | PASS, 5/5; isolierte PGlite-/Manifestverträge, kein nativer Produktivnachweis |
| `qa:agent-server` | PASS, offizieller MCP-Handshake und Werkzeugerkennung |
| Bestehende Modulverträge: Artikel/OP/Readiness/Forecast/BA-Anlage/Routing/Rüsten/Menge/Lebenszyklus | PASS, 8+15+17+41+9+10+17+14+37 = 168/168, keine Skips |
| `test:e2e` | BLOCKED_NOT_TESTED: vollständige 758 Fälle aufgerufen; Browserstart durch fehlende Playwright-Binaries blockiert |
| `test:agents` | BLOCKED_NOT_TESTED: Seed aufgerufen, gleicher fehlender Browser |
| Zusätzlicher System-Chromium-Versuch | BLOCKED_NOT_TESTED: 15 neue Chromium/Desktop-/Mobil-/Tablet-Fälle aufgerufen, Start scheitert mit `socket() failed: Operation not permitted` vor dem Öffnen einer Seite |
| `test:visual` | BLOCKED_NOT_TESTED, Exit 2: gepinnter Playwright-1.62.1-noble-Container fehlt; keine Umgebungskennung vorgetäuscht, keine Referenzen erneuert |
| `test:db:preflight` / native isolierte DB-Gesamtabnahme | BLOCKED_NOT_TESTED, Exit 2: Loopback-Teststack, Docker und geprüfte Auth-/RLS-Fixtures fehlen |
| `git diff --check` | PASS |

Der Versuch, die gepinnten Playwright-Binaries zu installieren, scheiterte an einem ungültigen/unvollständigen Download. Kein Ausweichen auf die Live-App. Die erste parallele Ausführung mehrerer speicherintensiver PGlite-Suiten beendete einzelne Testdateiprozesse ohne fachliche Assertiondiagnose. Die betroffenen vollständigen unveränderten Skripte wurden anschließend einzeln erfolgreich ausgeführt (BA 9/9, Routing 10/10, Menge 14/14, Forecast 41/41). Das ist eine dokumentierte lokale Ausführungsgrenze, keine Testanpassung.

## Grenzen und Abschluss

Die Sicherheitslücken und weiteren Freigabesperren aus [PR #25](production-release-step1-2026-11-08.md) bleiben offen; die Integrationskorrekturen aus [PR #26](production-integration-step2-2026-11-08.md) bleiben erhalten. Keine neue Gesamtsicherheitsanalyse, keine neue BA-/OP-Phase. Keine Supabase-Verbindung, SQL-/RLS-/Auth-/Rechte-/Feature-Flag- oder Deploymentänderung. Kein Merge, keine Migration und keine Produktivfreigabe.

Nach Sicherung und Aktualisierung des bestehenden Draft-PR #27 endet der Auftrag. Keine Actions-Warteschleife, kein Dispatch und kein zusätzlicher Wiederholungslauf in GitHub. Branch, endgültiger Commit, Dateiliste und Testgrenzen werden im PR dokumentiert.
