# Schritt 4: QA-Laufzeit und isolierte Berechtigungsintegration

**Keine Produktivfreigabe.** Gestapelter Draft auf PR #28, Basis `dd64a8e2e6a5c489cc2c98d0fa7862d59cada340` (0.16.0), Branch `security/permissions-integration-step4`. APP_VERSION **0.16.1**, VERSION_LOG **2026-11-12 09:13 Europe/Berlin**, tatsächlicher Änderungszeitpunkt. Refs #17, #25–#28; keine Startseiten-/BA-/OP-Neuentwicklung und keine Änderung abgeschlossener PRs.

## A. Nachgewiesene QA-Ursache

[QA and Release 43171988032](https://github.com/butzi124-ship-it/schichtplan-mvp/actions/runs/43171988032): alle verfügbaren Jobprotokolle abgerufen, Functional-/Visual-Artefakte heruntergeladen und ZIP-CRC/SHA-256 geprüft. Getesteter PR-Merge-Checkout `87b9080e893f7af7ad33c13c00f430b0727e18e9`, PR-Head `dd64a8e2e6a5c489cc2c98d0fa7862d59cada340`. Maschinenlesbarer Nachweis: `pr28-qa-evidence.json`.

| Kategorie | Tatsächlicher Befund |
|---|---|
| Fachlicher Fehler | **Kein protokollierter fehlgeschlagener Browserfall**. 726 als bestanden geloggte Fälle; von 758 sind 32 ohne abschließenden Nachweis. Das ist kein vollständig grüner Lauf. |
| Zeitlimit | Job beginnt 06:59:43.629 UTC, Abbruch 07:19:42.788 UTC: **1199.158 Sekunden**, praktisch genau das konfigurierte 20-Minuten-Joblimit. Testlauf beginnt 07:00:26, arbeitet mit zwei Workern bis zum Abbruch weiter. Das Laufzeitbudget ist die belegte zu korrigierende Ursache. |
| Runner/Container | Browser starten und absolvieren 726 Fälle; kein belegter OOM/Crash/Startdefekt. Containerbefehle, Berichtsversuch, Upload und Cleanup funktionieren danach. Docker-config-Warnungen erscheinen auch außerhalb des Abbruchpunkts; kein Beleg, dass sie den Abbruch verursachten. |
| Abhängigkeiten | Quality Gate sieht `FUNCTIONAL=cancelled` und scheitert korrekt; e2e scheitert danach korrekt. Diese Gates werden nicht verändert. |
| Konkurrenz/Doppelstarts | Kein Protokollbeleg für manuelles oder ersetzendes Concurrency-Cancel. Die vorhandenen Gruppen enthalten Workflow und Ref und trennen die genannten Workflows; keine spekulative Änderung von Triggern oder Cancel-Regeln. Vollständige externe Cancellation-Auditdaten liegen nicht vor. |
| Artefakte | Beide ZIPs intakt, Digest entspricht API-Metadaten. Functional enthält fünf Dashboard-PNGs und ci-summary.md, **kein results.json/abschließender QA-Report**. Der Prozess wurde vor Reporterabschluss beendet. Fehlender Abschlussbericht ist keine beschädigte ZIP-Datei und keine bestandene Suite. |
| Vorhandene Erfolge | QA Integrity erfolgreich; Visuals zweimal **33/33**. Native #28-Sicherheit separat erfolgreich: Run `43171988044`, Artefakt `13922053348`, **70/70** bei `authoritativeRestore:false`. BA/OP-Run `43171988237`: Diagnose, #16-Härtung und BA/OP jeweils success. Diese Ergebnisse gelten ausschließlich für den Ausgangsstand. |

### Tatsächliche QA-Korrektur

Nur das funktionale Jobbudget wird von 20 auf **40 Minuten** erhöht. Playwright erhält zusätzlich ein **30-Minuten-Gesamtlimit**, sodass vor dem Joblimit noch Zeit für Reporter/Upload/Cleanup bleibt. Die Testzeit von **30 Sekunden**, Expect-Zeit von **5 Sekunden**, vorhandene Retryzahl, alle Projekte, Assertions, geschützten Identitäten und fehlschließenden Gates bleiben erhalten. Keine Erhöhung von Einzelfall-/Assertion-Toleranzen, kein Sharding oder Ausschluss zur Abkürzung.

Ein zusätzlicher `qa-progress-reporter.js` schreibt bei Begin/Teststart/Testende atomar `artifacts/functional/progress.json`. Er erfasst geplante Fälle, alle Versuche einschließlich Fehler/Retry/Unterbrechung und noch laufende Fälle. `complete` bedeutet Reporterabschluss, **nicht**, dass alle Fälle bestanden hätten. Fehlt onEnd bei hartem Abbruch, bleibt complete=false. Er ersetzt keinen vorhandenen Reporter, überschreibt kein Endresultat und kann den Lauf nicht grün machen. Die vorhandene always()-Artefaktaufnahme umfasst diesen Pfad bereits. Die neue Laufzeit ist eine überprüfbare Korrektur des belegten Budgets, **noch kein Beweis eines anschließend erfolgreichen CI-Browserlaufs**.

## B. Konkrete Zuordnung heutiger und sicherer Zugriffe

`security-step4-access-inventory.json` enthält **152 literal adressierte Tabellenzugriffe auf 30 Tabellen**, mit Funktion, Operation und Zeile am unveränderlichen Ausgangscommit. Das ist ein Quellinventar, kein Laufzeit-/Rechtenachweis. Die neun weiteren `.from(`-Treffer sind Array.from, keine Supabase-Tabellenaufrufe. RPC-/Edge-Familien sind nachfolgend getrennt aufgeführt. Die bekannte Tabelle planner_special_day_assignments hat keinen solchen direkten App-Aufruf; ihre SQL-/Forecast-Verwendung und #28-Policy bleiben separat erhalten.

| Heutiger Zugriff / Funktionen | Sicherer Zielvertrag und Stand |
|---|---|
| bootSupabase/testSupabaseConnection/getCurrentEmployeeRecord: vor Login employees.id, danach employees SELECT* nach auth_user_id; syncSupabaseSessionToApp/getSession | Neuer parameterloser **security_actor_context** leitet Identität/Rolle/Bereich aus #28-Registry und aktueller DB-Bindung ab. Adapter kombiniert auth.getUser mit diesem Serverkontext und prüft gleiche Auth-ID. **Isoliert implementiert/getestet.** Alte Boot-/Login-/Testlogin-Oberfläche noch nicht umgestellt. Anon-Denial darf nicht zu lokalem vermeintlichem Adminbetrieb mit DB-Rechten führen. |
| loadEmployeesFromSupabase; add/update/activate/deactivate/delete/savePersonnel*/Slot-Zuordnung | #28-RLS lässt nur verifizierte Administration schreiben und eigene/admin Mitarbeiterzeilen lesen. Künftiger Admin-Client benötigt bestätigten Feldumfang und getrennte Authority-Provisionierung. Kein Ersatz-SELECT* und keine neue Admin-RPC erfunden. Rollen-/Auth-ID-Änderung darf nicht automatisch Registry-Rechte erzeugen. **Legacy-Anbindung offen.** |
| Eigene Kontakte | Bestehender geprüfter RPC **security_own_contacts** plus serverseitige Feld-Allowlist. Adapter besitzt keinen fremden Zielparameter. **Isoliert integriert**, nur synthetisch freigegebenes Feld getestet. Tatsächliche Kontaktfelder/Store-/UI-Migration bleiben D4; keine ursprünglichen Kontaktspalten vorhanden. |
| loadPlanningDataFromSupabase: neun Planner-SELECT*, Sondertage, Tasks; generateThreeMonths/persist | Mitarbeiterlesen über **security_published_schedule** mit vier minimalen Feldern und expliziter Freigabe/Scope. **Isoliert integriert.** Alte Namens-/Generatoransicht braucht D1/D5 und geprüfte ID-Darstellung; keine Zuordnung aus Namen/Schichtkürzeln erfunden. |
| assignShift/assignOptionalShift/assignSuggestedSpringer/approveSaturdayRequest und weitere planner_assignments-Upserts/Deletes | Reine eindeutige Zuordnung über **security_assignment_command** mit stabilen IDs, Abteilung und erwarteter Version; Freigabe separat **security_publish_assignment**. **Isoliert integriert.** Kombinierte Altaktionen schreiben zusätzlich Anträge/Ausfälle/Ersatz: diese dürfen nicht teilweise auf neue RPCs umgebogen werden. |
| planner_shift_cancellations, planner_absence_replacements und resetPlanCurrentFuture | #28 sperrt Raw-DML. Sichere transaktionale Storno-/Ersatz-/Reset-Kommandos und Vertrauensbindungen fehlen; **D2/D5 blockiert**, keine Admin-only-Abkürzung. |
| planner_absences/vacations/sick_leaves, availability/saturday_requests/swaps, special_days, tasks | #28 hat befristete Zeilen-/Bereichs-Lesegrants, keine sicheren vollständigen Schreibworkflows. Medizinische Sichtbarkeit, Anträge/Genehmigung und To-Do-Fremdrechte sind **D2/D3**. Unbekannte Zeilen bleiben gesperrt; keine ersatzweise breite RPC. |
| departments CRUD/find/load; shift_definitions load/save | Admin-Stammdaten und RLS-Bereichslesen getrennt; Planungsbefugnis ist keine Befugnis, Leiter oder globale Schichtmodelle zu verändern. Vollständige Original-Grants/Privacy-Projektion und UI-Umstellung fehlen. **Unverändert/offen.** |
| production_machines CRUD; production_orders/counts/order_stations/station_counts/station_events/order_employees/history/checklist/templates/qa_causes | Legacy-Direktpfade bleiben unangetastet, sind aber **nicht pauschal als gehärtet kompatibel abgenommen**. Vor späterem Betrieb Alt-BA/Neue-OP-Schreibpfade trennen; keine direkten Tabellenpatches als Ersatz für neue OP-RPCs. Originalschema/ACL-/Mischbetrieb-Nachweis bleibt Blocker. |
| Artikel/Operations/BA-Einstieg | Bestehende **article_master_***, **operation_plan**, **operations_article_command**, **order_from_article**, **operation_order_read/scan**, **ba_entry_command** beibehalten. Unabhängige Scopes und Erst-OP-Berechtigung werden nicht durch neue Admin-only-Wrapper ersetzt. Native #28-Nachweise vorhanden; erneute aktuelle Gesamtabnahme erforderlich. |
| Maschinenübergabe/Rüsten/Mengen/Lebenszyklus | Bestehende **machine_readiness_view/command/initialize**, **op_routing_view/command**, **setup_time_view/command**, **op_quantity_view/command**, **ba_lifecycle_view/command** unverändert. Abteilungsübergreifende Ausführung behält eigene Production-Authorities; Planungsvertretung ist keine zusätzliche Fertigungsfreigabe. |
| Prognose/Edge | Bestehende **production_forecast_***, Forecast- und Worker-Endpunkte unverändert und produktive Schalter unverändert. Kein Schicht-/Kontaktadapter wird Service-Role-Client oder Produktionsaktivierung. |
| loadToolMaterialsFromSupabase; Materialverwaltung | Neuer Adapter liest unter vorhandener #28-RLS nur **id,name,sort_order,is_active**, keine SELECT*-Ersetzung. Mitarbeiter/Scanner-Lesen **isoliert integriert/getestet**. Admin-Referenzschreiben bleibt #28-Policy, nicht neue pauschale Mitarbeiterfreigabe. |
| tools: load/edit/bookToolChange/restock/withdraw/QR/move/undo/markOrdered; tool_journal load/add | Heutige Bestand-/Journal-DML bleiben erhalten. Originalrechte, erlaubte Scannerbereiche und atomare Buchungs-/Journal-/Rücknahmebedingungen müssen für spätere sichere Schreibschnittstelle belegt werden. Kein erfundener Buchungs-RPC und kein behaupteter vollständiger Werkzeugnachweis. |
| inventory_sessions/entries | Getrennte Inventurfreigabe, ACLs und atomare Abnahme erforderlich; produktiver Inventurschalter bleibt aus. Kein neuer Zugriff aus diesem Adapter. |

Der Bezug zur vorhandenen Anwendung ist konkret: Adapterdatei im bestehenden js-Verzeichnis, reale #28-RPC-Signaturen und Materialtabelle, neue minimale Server-Identität, Tests mit derselben Review-SQL sowie eine lokale Browser-Testanbindung. **Die Produktions-HTML lädt diesen Adapter nicht.** Es gibt keinen neuen produktiven Feature-Schalter. Damit werden bestätigte Abläufe vorbereitet, ohne ungeklärte Altprozesse oder Rechte still zu verändern.

## C. Implementierte Integration und Schutzgrenzen

`security-step4-integration-isolated.sql` ist eine additive, sentinelgesperrte Review-Datei auf dem unveränderten #28-Kandidaten, keine Migration. Ein neuer RPC `security_actor_context()` ohne Identitätsparameter liefert eigene stabile Identität/Rolle/Bereich, verifizierte Planungsbereiche und konfigurierte Kontaktfeldnamen, keine fremden Mitarbeiterzeilen. PUBLIC/Anon-EXECUTE ist entzogen. Keine #28-Policy oder vorhandene Fachfunktion ersetzt.

`js/secure-access.js` ist standardmäßig deaktiviert. Vor **jedem** Aufruf: verifiziertes auth.getUser, frischer DB-Kontext, Konsistenz der Auth-ID. Kein getSession-/JWT-User-Metadata-/UI-Rollen-Bootstrap. Der Clientfilter ist nur zusätzliche Bedien-/Datenminimierung; die eigentliche Autorisierung bleibt im SQL-RPC/RLS. Veröffentlichung wird nicht aus AL-Rolle oder Planungsbereichen abgeleitet.

Die Integration hält nur ihre eigenen minimalen Daten im Modul-Arbeitsspeicher. Keine Speicherung, kein Offline-Fallback und kein Raw-Table-Fallback. Sessionevents einschließlich Tokenwechsel, Offlinezustand, Ablehnung und dispose verwerfen diesen Cache; ein cleared-Callback entfernt die Anzeige der isolierten Testanbindung. Generationen binden Antworten an die aktuelle Session. Eine alte Antwort darf weder alte Daten anzeigen noch bereits geladene Daten des neuen Nutzers überschreiben/leeren. Payloads enthalten ausschließlich die festgelegten RPC-Parameter; keine fremde Kontaktidentität.

**Nicht damit gelöst:** bestehender allgemeiner persist()-Snapshot enthält weiterhin Planner-/Abwesenheits-/Taskdaten, alte Datenlogs und breite Login-/Boot-Abfragen bestehen. Diese Bereiche werden nicht an unbekannten Berechtigungsentscheidungen vorbei umgeschrieben. Die neue Integration schreibt ihre Daten gerade nicht in diesen State. Vor Aktivierung braucht es gesondert geprüfte Altcache-Bereinigung, Wegfall vertraulicher Logs, sichere UI-/ID-Anbindung und bestätigte Offline-Regeln. Serverseitiger session_id-/Tokenwiderruf und sofortige Rücknahme bereits sichtbarer Daten sind nicht durch einen Modulcache bewiesen. Jeder neue Adapteraufruf prüft aktuelle DB-Befugnisse; Realtime-/In-flight-Widerruf bleibt Gesamtabnahme.

## D. Positive/negative Prüfungen und CI-Vorbereitung

- Vorhandene 17 Sicherheitsprüfungen unverändert erhalten. Neue Tests verbinden tatsächlichen Adapter mit tatsächlicher #28/#4-SQL gegen den strukturellen PGlite-Spiegel: standardmäßig aus, serverbasierte Identität, Admin/AL/Mitarbeiter/Scanner, eigene/fremde Abteilung, gültige/abgelaufene/widerrufene Vertretung, getrennte Veröffentlichung, eigene Kontakte/geschützte Felder, Materialprojektion, Offline, Sessionwechsel/Tokenrefresh/Logout, widersprüchliche Identität und verspätete Antworten.
- Drei neue lokale Browserfälle werden additiv in **allen fünf** bestehenden Projekten ausgeführt. Ihre neutrale Testseite ist keine Produkt-Designumstellung. Alle bisherigen 758 Fälle bleiben erhalten, jetzt **773**. Die Browserbrücke verwendet tatsächliche SQL mit expliziten Auth-Shims; dies ist kein GoTrue-/JWT-Beweis.
- Neue separate native Step4-Pipeline führt zuerst die unveränderte gesamte Step3-/BA-OP-/Baseline-Kette und dann drei zusätzliche Adapterprüfungen mit echten GoTrue-/PostgREST-Aufrufen aus. Exakter rückführbarer Runneradapter auf geprüftem Step3-Runner, frischer Quellenhash vor DB-Verbindung, Cleanup/Artefakte. Die vorhandenen drei nativen Akzeptanzmodi und Step3-Nachweise werden nicht überschrieben. Native Mehrverbindungsprüfung aus Step3 bleibt enthalten.
- Während der Implementierung fanden die neuen Tests einen Syntaxfehler im neuen SQL-Guard und eine falsche skalare Darstellung des tabellarischen RPC-Ergebnisses in der neuen eingebetteten Testbrücke. Beides wurde an der Quelle korrigiert; der Adapter akzeptiert weiterhin nur die wirkliche Array-/Projektionsform. Keine Bestandsassertion oder Hasherwartung angepasst.

## E. Lokale Ergebnisse

| Prüfung | Ergebnis des neuen lokalen Standes |
|---|---|
| `npm ci --ignore-scripts`, beide Pflicht-Nodechecks, `check:syntax` und alle neuen JavaScriptdateien | PASS |
| `test:ci` | PASS **2/2** |
| `test:qa` | PASS **20/20**, einschließlich Laufzeit-/Fortschrittsnachweis; bestehende Assertions unverändert |
| `qa:check` | PASS: Geschäftslogik-/Testhashes, 251 Legacy-Identitäten und 33 PNG-Referenzen erhalten |
| `test:integration` | PASS **5/5** |
| `test:security` | PASS **28/28**, keine Skips: bisherige 17 plus zehn SQL-Adapter-/Sessionfälle und eine Step4-Orchestrierungsprüfung |
| Bestehende neun Fachmodulsuiten | PASS **168/168** unverändert, keine Skips; keine Behauptung einer gemeinsamen aktuellen nativen Gesamtabnahme |
| `qa:agent-server` | PASS |
| Neue native Runnerfortsetzung / Quellenmanifest / ursprüngliche Target-Guards | PASS: exakte Clock-/Provenienz-/Step3-/Step4-Montage, Syntax und 48 Eingabedateien geprüft; ursprüngliche Guards **2/2** bestanden. Keine native DB-Ausführung |
| `test:e2e` | **BLOCKED_NOT_TESTED**: alle **773** Fälle aufgerufen, **773** fehlende Browser-Executables vor Seitenaufruf, 0 bestanden/0 Skips |
| Neues Fortschrittsartefakt im tatsächlichen vollständigen lokalen Testaufruf | Geplante Fälle 773, 773 fehlgeschlagene Versuche, 0 laufend, complete=true, runnerStatus=failed; keine falsche grüne Darstellung |
| `test:agents` | **BLOCKED_NOT_TESTED**: Seed aufgerufen, gleicher Browserstartblocker |
| `test:visual` | **BLOCKED_NOT_TESTED**, Exit 2: vorgeschriebener Playwright-1.62.1-noble-Container fehlt; Referenzen unverändert |
| Native neue Sicherheits-/Adapter-/BA-OP-/Parallelabnahme und `test:db:preflight` | **BLOCKED_NOT_TESTED**, Preflight Exit 2: kein Docker/Podman/psql oder freigegebener lokaler Stack; neue Pipeline vorbereitet, nicht bestanden |
| `git diff --check` | PASS |

Historische #28-CI-Erfolge sind kein Testnachweis des neuen Commits. Fehlende Browser-/Containerlaufzeit wird nicht durch erfolgreiche eingebettete Tests ersetzt. Das größere CI-Laufzeitbudget und die Browser-Testanbindung müssen im späteren regulären CI-Lauf beurteilt werden; nach PR-Erstellung wird dieser Auftrag keinen Lauf abfragen oder nachstarten.

## F. Verbleibende Entscheidungen und nächster kontrollierter Schritt

1. **D1/D4:** Schichtleserkreis, sichtbare Felder, Veröffentlichungszuständigkeit sowie konkrete eigene Kontaktfelder/Validierung/Speicherung bestätigen. Die Kontakt-Allowlist bleibt produktiv unentschieden; synthetic_contact ist ausschließlich Testfreigabe.
2. **D2/D3:** Storno, Ersatz, Tausch, Antrag/Genehmigung, Samstag, Sondertage, Tasks und vertrauliche Abwesenheitsrechte einschließlich Aufbewahrung/Offline-Vorgaben entscheiden. Keine halb umgestellten Mehrtabellenaktionen aktivieren.
3. **D5/D6:** unabhängig geprüfte Auth-/Mitarbeiter-/Abteilungs-/Altzeilenbindungen, Authority-Issuer und Befristungs-/Widerrufsprozess bereitstellen. Keine automatische Aufnahme aus alten Profilrollen oder Namen.
4. Autoritativer aktueller Schema-only-Restore einschließlich Owner/ACL/Default-Grants, Funktionen/Views/Trigger/Sequenzen/Versionen/API-Schemas fehlt weiterhin. Der erfolgreiche native 31-Tabellen-Spiegel ist **kein vollständiger Originalrestore**.
5. Gegen diesen Restore gesamte sichere Oberfläche, Legacy-/Neu-BA-Mischbetrieb, Werkzeug-/Journal-/Scan-/Inventurtransaktionen, Sessionwiderruf/Realtime, Daten-/Cacheminimierung und Last abnehmen. Die alte App-Boot-/SELECT*-Anbindung bleibt ein klarer Aktivierungsblocker.
6. Backup/Restoreprobe, RPO/RTO, Wartungs-/Rückfallverantwortliche und #25-Pages-/Betriebsnachweise bleiben nötig. Der kontrollierte Migrations-/Backup-/Restoreplan aus #28 gilt weiter; ein Rückfall darf keine unsicheren permissiven Policies automatisch reaktivieren. Kein Merge, Deployment oder Produktionsmigration ohne getrennte spätere ausdrückliche Entscheidungen.

**Abschluss:** neuer gestapelter Draft auf #28, nach Erstellung keine Actions-Abfrage/Warteschleife und kein Dispatch. Keine produktive Supabase-Verbindung, RLS-/Auth-/Daten-/Feature-Flag-Änderung, kein Merge und kein Deployment. Keine vorzeitige Produktivfreigabe.

Dokumentationsgrundlagen: https://playwright.dev/docs/test-timeouts#global-timeout ; https://playwright.dev/docs/api/class-reporter ; https://supabase.com/docs/reference/javascript/auth-getuser . Keine Paket-/CLI-Version angehoben.
