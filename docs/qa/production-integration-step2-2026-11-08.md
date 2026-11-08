# Produktivvorbereitung Schritt 2 – technische Integrationskorrekturen

**Keine Produktivfreigabe.** Gestapelter Draft auf PR #25 (`docs/production-release-step1`, `5abcb9265483c5eac275e98a2c80248958e7d88f`, 0.15.1), Refs #17, #15, #16, #24. Auftrag ausschließlich F04 und der Testhook-/Quellenkonflikt aus Abschnitt 6.2 des [Freigabeberichts](production-release-step1-2026-11-08.md). Version **0.15.2**, VERSION_LOG **08.11.2026 09:04 Europe/Berlin**: tatsächlicher Zeitpunkt der Versionsänderung. Bestehende Fachfunktionen und Feature-Schalter unverändert.

## Konkrete Änderungen

1. `supabase/review/issue-10-production-forecast.sql` entspricht bytegenau PR #15, Head `25e9a6914e38fc791d9ecad13d92a45e30d19bb3`. `forecast_source_dirty` wird auf allen 23 gebundenen Tabellen **AFTER STATEMENT** erzeugt. Der Kommentar zur vermiedenen Sperrreihenfolge/40P01 ist mitübernommen. Keine anderen Trigger/Funktionen/Berechtigungen dieser SQL geändert. SHA-256: `fff6eed3f17d52877ff923ba8e0ae57997248b1767f76977d516ca0bd851db21`. Ausschließlich der bereits in #15 getestete Forecast-SQL-Hash wird in `protected-files.json` nachvollziehbar aktualisiert; keine Test-/Assertion-/Bildreferenz-Hashes geändert. Neue PGlite-Katalogregression prüft die tatsächlich installierten 23 Trigger, einschließlich AFTER und STATEMENT.
2. Drei unabhängige native Matrix-Jobs, `fail-fast: false`, jeweils eigener Runner und frischer temporärer Stack:

   | Modus | Unveränderte Baselineprüfungen aus #15 | Zusätzlicher zwingender Pfad |
   |---|---|---|
   | `diagnostic` | Alle 14 einschließlich negativer Parallel-/40P01-Regression und A/B-Sicherheitsdiagnose | Keiner; ausschließlich Diagnose |
   | `hardening` | Alle 14 vor dem Kandidaten | Originaler isolierter #16-Kandidat und sämtliche zugehörigen negativen JWT/REST/RLS-Assertions |
   | `ba-op` | Alle 14 vor den Phasen | Erstverantwortung → Lebenszyklus → Menge → Rüsten → Routing → BA-Anlage, unveränderte Tests und Assertions |

   #16-Testmodul von `security-extension.js` nach `hardening-extension.js` übernommen, Inhalt bytegleich zu `bf14f4ed32fd8dc8e4e8374a5faa0c51a29416c7`; Kandidaten-SQL ebenfalls bytegleich übernommen. Nur Reviewdatei, keine registrierte Migration. BA-/OP-Module bleiben unter ihren ursprünglichen Namen; kein Kopieren mehr über `security-extension.js`. Ein unbekannter/fehlender Modus oder fehlendes Pflichtmodul scheitert ausdrücklich. Die bekannte unsichere Baseline wird weiterhin mit allen Exploit-Assertions geprüft. Der Hardening-Status wird getrennt gespeichert; `securityStatus` bleibt der offene Baselinebefund. Der konservative Admin-only-Kandidat ist weiterhin keine freigegebene Rollenmatrix und wird nicht in der BA-/OP-Umgebung installiert. Eine gemeinsame fachlich freigegebene Sicherheitslösung ist nicht Gegenstand dieser Änderung.
3. Der zweite Checkout bleibt ausschließlich **gepinntes natives Testharness** aus #15, inklusive strukturellem Katalogspiegel, Laufzeit und nativen Basistests/Abhängigkeiten. Dessen Version ist offengelegt, keine verdeckte Branchabhängigkeit. Die zu installierenden Dateien stammen vollständig aus dem **konkreten aktuellen Integrationscheckout**:

   - Ganze Verzeichnisse `supabase/`, `js/`, `tests/db-integration/` werden vor dem Kopieren im disponiblen Ziel ersetzt; dazu die aktuelle `forecast-clock.json`. Somit bleiben keine alten SQL-/Edge-/Erweiterungsdateien aus dem Harness zurück.
   - Der Integrationscheckout muss für diese Eingaben exakt mit seinem HEAD übereinstimmen. Nicht versionierte Eingaben, auch ignorierte Dateien, sind verboten. Symlinks werden abgewiesen.
   - SHA-256-Manifest `artifacts/db-integration/integration-source.json` enthält tatsächlichen Checkout-SHA, Harness-SHA und jede Eingabedatei. Bei PR-Ereignissen darf das der tatsächlich ausgecheckte Merge-Commit sein; er wird nicht fälschlich als PR-Head bezeichnet.
   - Runner verifiziert das vollständige Manifest erneut **vor der ersten Test-Datenbankverbindung**. Austausch, zusätzliche Datei oder Löschen scheitert. Manifest und Testergebnisse werden je Modus mit eindeutigem Artefaktnamen 14 Tage archiviert.
   - Der bereits bestehende, exakt gepinnte Clock-Adapter bleibt unverändert. Roh-SQL und Roh-Hashes bleiben erhalten; ausschließlich Zeitreferenzen werden beim Ausführen im isolierten Test transformiert. JWT-Zeit bleibt real. Dies ist ausdrücklich keine bytegleiche Ausführung aller Zeitreferenzen ohne Testadapter.
   - Ein weiterer exakter reversibler Adapter ersetzt nur den optionalen Hook durch den expliziten Dispatcher und ergänzt die vorgelagerte Quellen-/Modusprüfung. Der bekannte Original-Runnerhash wird geprüft. Keine Basistest-Assertions entfernt oder verändert.

Die native Pipeline führt weiterhin Original-Target-Guard, Baseline, Paralleltests, Summary, garantierten Cleanup und Artefaktarchivierung aus. Der bestehende e2e-/Quality-Gate-/Deploymentpfad bleibt erhalten; lediglich `test:integration` wird als zusätzlicher Integritätscheck ergänzt. Keine neuen Paketabhängigkeiten oder Lockfileänderungen.

## Lokal tatsächlich geprüft

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| `npm ci --ignore-scripts` im Integrationscheckout und gepinnten Harness | bestanden |
| Beide Pflicht-`node --check`, `check:syntax`, `git diff --check` | bestanden |
| Artikel / OP / Readiness / Forecast / BA-Anlage / Routing / Rüsten / Menge / Lebenszyklus | **8 + 15 + 17 + 41 + 9 + 10 + 17 + 14 + 37 = 168/168 bestanden**, keine Skips |
| `test:integration` | **5/5 bestanden**, keine Skips: explizite Modi/Fehler, Baseline-Status, Quellen-/Symlink-/untracked-/ignored-Manipulationen, unbekannter Runner, 23 tatsächlich installierte AFTER-STATEMENT-Trigger, vollständige Modus-/Phasenkette |
| Originaler `test:db:guard` im zusammengesetzten Harness | **2/2 bestanden**, keine Skips; exakt unveränderte #15-Assertions |
| `test:ci` / `test:qa` | **2/2** / **9/9 bestanden** |
| `qa:check` / `qa:agent-server` | bestanden; 251 Legacy-Identitäten und 33 Bildreferenzen erhalten |
| Reale lokale Zusammenstellung aus sauberem Implementierungscommit | bestanden; 40 manifestierte Dateien einschließlich 11 SQL-Dateien, erfolgreiche Clock-/Dispatcheradapter, Manifestprüfung und Syntax von vorbereitetem Runner/Seed |
| Beide aktuellen Edge-Einstiege mit gepinnten Harness-Abhängigkeiten bündeln | bestanden; Buildnachweis, keine Runtimeabnahme |
| Ungültiger Modus im tatsächlich vorbereiteten Runner | erwartete Ablehnung vor Konfigurations-/DB-Zugriff, 0 DB-Tests; kein behaupteter DB-Erfolg |
| Vollständiges `test:e2e` | **733 Fälle aufgerufen**, alle Browserstarts blockiert: Chromium/Desktop, Mobil, Tablet, Firefox und WebKit fehlen als Binaries; keine fachliche Browserabnahme |
| `test:agents` | Seed aufgerufen, Browserstart aus demselben Grund blockiert |
| `test:visual` | **BLOCKED_NOT_TESTED**, vorgeschriebener `playwright:v1.62.1-noble`-Container fehlt; Referenzen unverändert |
| `test:db:preflight` | Exit 2, **BLOCKED_NOT_TESTED**: isolierte Laufzeit/Auth-Fixtures/URL fehlen |
| Native PostgreSQL-/GoTrue-/PostgREST-/Edge-/Parallelabnahme aller drei Modi | **BLOCKED_NOT_TESTED**, kein Docker/Podman, PostgreSQL-Client oder Docker-Socket; kein Ausweichen auf produktives Supabase |

Ein erster lokaler Aufruf hatte zusätzlich den ausschließlich zum gepinnten Harness gehörenden nativen Target-Guard im eigenständigen Featuretree gestartet; dessen zweite Prüfung konnte dort `runtime.js` nicht lesen. Die Trennung ist jetzt explizit: neue Integrationsverträge im Featuretree, **beide unveränderten nativen Guards im zusammengesetzten Harness**. Danach 5/5 plus 2/2 erfolgreich. Kein Guard entfernt oder abgeschwächt.

PGlite-/Build-/Manifestprüfungen ersetzen keine native Mehrverbindungs-, echte JWT/RLS-, Runtime- oder vollständige Originalschemaabnahme. Der historische 40P01-Test aus #15 bleibt unverändert in allen nativen Modi aktiv, ist hier aber mangels nativer Laufzeit nicht erneut ausgeführt. Auch ein später grüner Kandidatenlauf beweist keine produktive Sicherheitskorrektur.

## Unveränderte Sperren und Abschluss

Die offenen RLS-Lücken und die weiteren Befunde aus PR #25 bleiben bestehen. Kein Live-Supabase-Zugriff, keine produktive SQL-/RLS-/Auth-Änderung, kein Merge, Deployment oder Aktivieren von Funktionen. Kein pauschales Admin-only-Hardening im gemeinsamen Fachstand. Originalschema-/ACL-/Backup-/Betriebsnachweise und fachliche Rollenentscheidungen bleiben separat erforderlich.

Vor dem Push werden alle lokalen Arbeiten abgeschlossen. Anschließend nur den beauftragten gestapelten Draft-PR anlegen; keine Actions-Abfrage, Warteschleife, wiederholte Statusabfrage oder zusätzlicher Workflowdispatch. CI-Konfiguration: [BA and OP Isolated Acceptance](https://github.com/butzi124-ship-it/schichtplan-mvp/actions/workflows/ba-entry-isolated.yml) und [QA and Release](https://github.com/butzi124-ship-it/schichtplan-mvp/actions/workflows/e2e.yml). Es wird kein neuer CI-Erfolg behauptet.
