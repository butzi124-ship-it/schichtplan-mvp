# PR #22 – gemeinsame native und Browser-Testzeit, v0.13.2

## Nachgewiesene Restursachen

Die native CI auf dem v0.13.1-Stand meldet 42/43; der unveränderte Edge-/SQL-Test erwartet für OP 20 `expected`, obwohl bei spätem Start die Ankunft am Folgetag liegt. Der Workflow prüft eine separat gepinnte Infrastruktur (`25e9a6914e38fc791d9ecad13d92a45e30d19bb3`). Der Node-Testuhr-Fix aus v0.13.1 war dort nicht aktiv.

Das Browserfixture hatte drei unabhängige Quellen: echte Node-/PGlite-Zeit, echte Browserzeit und eine UTC-Datenbanksitzung. Der Unterschied `red`/`orange` ist separat bei **07.11.2026 23:30 UTC / 08.11.2026 00:30 Berlin** reproduziert: `current_date` unter UTC sperrt für die synthetische Abwesenheit nur den 08.11. im Prognosehorizont; am 09.11. bleibt Kapazität, Ergebnis `orange`. Derselbe Fall mit Berliner Datenbankzone sperrt den 08. und 09.11. und ergibt `red`. Die fachlichen Statusregeln bleiben unverändert.

## Begrenzte Korrektur

- Eine gemeinsame Fixture `tests/fixtures/forecast-clock.json`: **07.11.2026 09:00:00 UTC**, Zone **Europe/Berlin**. Node-Vertragstests und alle fünf Browserprofile beziehen ihre Zeit daraus.
- Browserfixtures setzen Node/PGlite und die Browser-Date-Uhr gemeinsam vor dem Seeding. Jede Berechnungsanforderung schreitet auf beiden Seiten um eine Sekunde fort. Native Timer laufen weiter. Der bestehende Cachetest überschreitet zusätzlich ausdrücklich auf beiden Seiten die Zwei-Sekunden-Grenze; seine ursprünglichen Grün-/Grau-/Material-Assertions bleiben erhalten.
- Die PGlite-Fixture setzt ihre Datenbanksitzung explizit auf Berlin. `current_date`, Personalverfügbarkeit und Prognosekalender stimmen damit auch am UTC/Berliner Tageswechsel überein.
- Der native Workflow behält die gepinnte Infrastruktur und alle 43 Tests bei. Ein Adapter ergänzt ausschließlich zwei exakt gehashte Stellen im wegwerfbaren Checkout: Installation einer Testuhr nach dem originalen Katalognachweis, vor Anwendungsmigrationen; Materialbeobachtung aus der SQL-Uhr statt aus der Node-Wanduhr. Alle Assertions bleiben bytegenau erhalten; unbekannte Infrastrukturversionen werden abgewiesen.
- Die native Uhr darf erst nach Prüfung von isoliertem Host/Port, Datenbankname und tatsächlichem Schutzmarker installiert werden. Ihre private Hilfstabelle liegt außerhalb von `public`, erhält keine direkten Benutzerrechte und verändert die erwartete Zahl der Anwendungstabellen nicht. Die Anwendungs-SQL-Dateien bleiben unverändert. Nur beim Laden in die isolierte Instanz werden Uhrabfragen durch qualifizierte Testuhr-Funktionen ersetzt. Kommentare, Literale und alle Geschäfts-/Berechtigungsprädikate bleiben erhalten.
- Native SQL-Zeit = gemeinsame feste Basis + tatsächlich verstrichene Laufzeit. Damit bleiben Zeitreihenfolge, Transaktionszeit, Statement-Zeit, Cachealter und Ablaufprüfungen wirksam; kein Einfrieren von Auth-/JWT-Uhren. Die echte Edge-Berechnung verwendet die von SQL gelieferten `as_of`-/Beobachtungszeitpunkte und benötigt keine eigene Wanduhr-Umschreibung. Das Ergebnisartefakt kennzeichnet diesen Testuhrmodus ausdrücklich.
- Quellschutz rechnet nur die exakten autorisierten Fixture-/Uhrblöcke zurück. Sämtliche ursprünglichen geschützten Hashes, Geschäftsassertionen, 251 Legacy-Identitäten und Phase-4-Quellen bleiben erhalten. Keine Produktiv-SQL-, Edge-, Statusregel- oder Feature-Schalteränderung.

## Lokale Nachweise und Grenzen

| Prüfung | Ergebnis |
|---|---|
| `npm ci --ignore-scripts`, vorgeschriebene Syntaxprüfungen, neue Adapter-/Testsyntax, `git diff --check` | bestanden |
| Vollständige Prognose-/Fixture-/Testuhr-Verträge unter UTC | **41/41** |
| Dieselbe vollständige Prognosesuite unter America/Los_Angeles | **41/41** |
| Artikel / OP / Readiness / BA / Routing / Rüsten / Mengen | **90/90**, davon Phase 4 **14/14** |
| CI-Bericht / QA-Verträge | **2/2**, **9/9** |
| `qa:check`, `qa:agent-server` | bestanden |
| Originale native Ziel-/Runtime-Guards aus gepinntem Checkout | **2/2** |
| Exakte Rückrechnung beider nativen Quellpatches auf Original-SHA-256; Syntax der adaptierten nativen Suite | bestanden |
| Unverengte Browsermatrix | alle **703** aufgerufen; Browserstart wegen fehlender Binaries blockiert, keine fachliche Abnahme |
| Agent-Seed | Browserstart blockiert |
| Erneuter Browserdownload | defekte/leere CDN-ZIP; kein installierbares Browserpaket |
| Native PostgreSQL/GoTrue/PostgREST/Edge-Containerabnahme | **BLOCKED_NOT_TESTED** lokal, Docker nicht verfügbar |
| Gepinnte Visuals | **BLOCKED_NOT_TESTED** lokal, vorgeschriebener Container nicht verfügbar |

Die sieben zusätzlichen Verträge prüfen die beiden Kalenderzeitpunkte (einschließlich Berliner Mitternacht), Grün/Orange/Rot, Vorankündigung ohne physische Freigabe, synchronen Cacheablauf, native Transaktions-/Wandzeit über mehrere Clientadapter, UTC/Los-Angeles/Kiritimati-Sitzungen, gesperrten Ankerzugriff, Zielverweigerung sowie ausschließlich erlaubte Uhrsubstitutionen über alle acht SQL-Quellen. PGlite-Verträge sind kein Ersatz für die echte native Abnahme. Ein gleichzeitiger erster Fachtestlauf endete bei zwei Testdateien ohne fachlichen Assertion-Befund; die vollständige unverengte Wiederholung mit Dateiparallelität 1 bestand **90/90**.

Bestehender Branch `feature/ba-op-phase4`, [PR #22](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/22), Bezug [#17](https://github.com/butzi124-ship-it/schichtplan-mvp/issues/17). Commit siehe aktuelle PR-Beschreibung. [Actions](https://github.com/butzi124-ship-it/schichtplan-mvp/actions): nach Upload keine Ergebnisabfrage, kein Dispatch, keine Warteschleife. Kein Merge, Deployment oder produktiver Datenbankeingriff.

---

## Historischer Bericht v0.13.1

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
