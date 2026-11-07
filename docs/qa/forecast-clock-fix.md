# PR #22 – Prognose-Testuhr, v0.13.1

## Ursache und begrenzte Korrektur

Auf dem unveränderten Phase-4-Head `8b1b8c3077da58d3fba8616053d1866ce9865a51` lokal 28/30 reproduziert: `planned` statt `expected` und fehlende erwartete Ablehnung beim Veröffentlichungs-Rollback. Die synthetischen Daten verwenden den aktuellen Berliner Tag, 00:00–23:59-Schichten und drei aufeinanderfolgende OPs mit je 5 Minuten Rüsten und 48 Minuten Bearbeitung. Die einzige qualifizierte Person begrenzt die gemeinsame Kapazität. Spät am Tag liegt die Vorgängerfertigstellung am Folgetag. `arrival_today=false` verhindert korrekt die Vorankündigung und damit auch den Arrival-Insert, dessen Trigger der Rollback-Test prüfen soll. Der grüne Terminstatus allein garantiert keine heutige Ankunft.

Nur die Testuhr wird korrigiert: Node-Date-Mock auf `2026-11-07T09:00:00Z`, vor Datenbankanlage und Beobachtungen. PGlite nutzt dieselbe JavaScript-Wanduhr für SQL `now()` und `clock_timestamp()`. Ein Sekundenschritt vor jedem bestehenden Test erhält die Reihenfolge verschiedener Veröffentlichungen; identische/ältere Läufe bleiben abgewiesen. Nach Datenbankschluss wird die Uhr zurückgesetzt. Keine Änderung an Prognoseberechnung, SQL, Veröffentlichung, Phase-4-Funktionen oder Produktionsdaten.

Alle bestehenden Prognose-Assertions bleiben bytegenau erhalten. Der Quellschutz rechnet ausschließlich die zwei exakten Import-/Testuhr-Blöcke zurück und vergleicht anschließend mit dem **unveränderten ursprünglichen SHA-256**. Sämtliche geschützten Hashes und 251 Legacy-Identitäten bleiben erhalten. Neue Zeitgrenzfälle liegen separat, damit das Verstellen der Uhr keine noch offene Datenbank beeinflusst.

## Deterministische Zeitgrenzen

| Zeitpunkt (UTC) | Berlin | Geprüftes Verhalten OP 20 |
|---|---|---|
| 07.11.2026 22:30 | 07.11. 23:30 | Ankunft am 08.11.; `arrival_today=false`, `planned`, keine Vorankündigung |
| 07.11.2026 23:30 | 08.11. 00:30 | Berliner Kalenderdatum 08.11.; heutige Ankunft, `expected` |
| 29.03.2026 00:30 | Beginn Sommerzeitwechsel | Heutige Ankunft trotz übersprungener Stunde, `expected` |
| 25.10.2026 01:30 | Wiederholte Herbststunde | Heutige Ankunft, `expected` |

Jeder Fall prüft JavaScript-/SQL-Zeitgleichheit, Kalenderdatum, grüne Prognose, tatsächlichen Ankunftstag, Readiness und Zahl der Vorankündigungen. Physische Mengen bleiben 0 und Arbeitsstatus bleibt `planned`. Die bestehenden Veröffentlichungs-/Rollback-Tests prüfen weiterhin Cache, Zähler und atomaren Rollback.

## Lokale Prüfung

- `npm ci --ignore-scripts`, beide vorgeschriebenen `node --check`, `check:syntax`, zusätzliche Syntaxprüfungen für neue Testdatei/Quellschutz und `git diff --check`: bestanden.
- Prognose: **34/34**, einschließlich unveränderter 30 Bestandsfälle; zusätzlich vollständige Wiederholung mit `TZ=UTC` und `TZ=America/Los_Angeles`: jeweils **34/34**.
- Artikel / OP / Readiness / BA / Routing / Rüsten / Phase-4-Mengen: **90/90**, einschließlich aller **14/14** Phase-4-Fälle.
- CI-Bericht **2/2**, QA-Vertragstests **9/9**, `qa:check` und `qa:agent-server`: bestanden. Exakte Rückrechnung der beiden Uhr-Blöcke auf die vollständige Originaldatei separat nachgewiesen.
- Unverengte Browsermatrix `npm run test:e2e` und `test:agents` aufgerufen: fehlende Playwright-Binaries verhindern Browserstart; keine bestandene fachliche Browserprüfung.
- `test:visual`: **BLOCKED_NOT_TESTED**, vorgeschriebener Playwright-1.62.1-noble-Container lokal nicht verfügbar; keine Referenzen geändert.
- Native Datenbankintegration lokal nicht erneut ausgeführt. Der gemeldete GitHub-Stand **42/43** wird nicht als vollständig bestanden dargestellt und ist durch diese Prognosekorrektur nicht automatisch behoben.

## Übergabe

Bestehender Branch `feature/ba-op-phase4`, bestehender Draft-PR [#22](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/22), Bezug [#17](https://github.com/butzi124-ship-it/schichtplan-mvp/issues/17). Commit/SHA siehe PR-Beschreibung. [Actions](https://github.com/butzi124-ship-it/schichtplan-mvp/actions): kein neuer Status nach dem Push abgerufen, kein Dispatch und keine Warteschleife. Kein Merge, Deployment oder produktiver Datenbankeingriff. Browser-/Visual- und native DB-Ergebnisse bleiben gesondert zu prüfen.
