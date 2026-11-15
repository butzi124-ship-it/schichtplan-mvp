# Schritt 5 – Originalnachweis und kontrollierte Migration

**Keine Produktivfreigabe.** Basis PR #29, `2e46d3a63c339389f08f16db67d1a9236216e130`, Version 0.16.1. Neuer gestapelter Draft, Branch `security/original-schema-step5`, Version **0.16.2**, VERSION_LOG **12.11.2026 10:39 Europe/Berlin**. Refs #17, #15, #16 und #25–#29. Vorhandene Fachfunktionen, Sicherheitskandidaten, Assertions und Bildreferenzen bleiben unverändert. Die einzige Änderung am Anwendungscode ist Version/Log. Keine neue Startseite, keine neue BA-/OP-Implementierung.

## 1. Was heute tatsächlich verifiziert wurde

Am **12.11.2026 ab 09:35:26 UTC** wurden ausschließlich die vorhandenen Metadaten-Endpunkte für das bereits belegte Projekt `itntnpzqzyjfdnjdshvm` aufgerufen. Keine freie SQL-Ausführung, keine Geschäftsdaten, Auth-Benutzer, Sessions, Logs oder Schlüssel. Bereinigte Ergebnisse und individuelle Zeitpunkte: `security-step5-original-evidence.json`.

| Endpunkt | Heutiger Befund | Beweisgrenze |
|---|---|---|
| Projektinformationen | Schichtplan, eu-west-3, ACTIVE_HEALTHY, gemeldete Datenbankversion **17.11.0.003**, Engine 17 | Management-Metadaten, kein SELECT version(), kein Katalog-Restore |
| public Tabellen/Spalten | **BLOCKED_NOT_TESTED**: SQLSTATE 28P01, Authentifizierung des Connector-Read-only-Users fehlgeschlagen | Aktuelle Tabellen/Spalten/PK/FK/RLS **UNVERIFIED**, nicht „0 Tabellen“ |
| Erweiterungen | **BLOCKED_NOT_TESTED**, gleicher Authentifizierungsfehler | Installierte Versionen **UNVERIFIED** |
| Migrationshistorie | **BLOCKED_NOT_TESTED**, SQLSTATE 28P01 im Metadaten-Endpunkt | Keine Aussage über tatsächlich installierte Migrationen |
| Edge-Funktionsliste | Endpunkt liefert leere Liste | Nur Listenantwort zum Abrufzeitpunkt; kein Beweis für fehlende Datenbank-RPCs, alte Deployments oder API-/Auth-Konfiguration |

Es gab keine Wiederholungsversuche, keine Credential-Reparatur und keine Umgehung durch execute_sql. Die Blockade ist eine technische Connector-Authentifizierung, kein nachgewiesener neuer RLS-/App-Fehler. Der Betreiber muss den reinen Metadatenzugang wiederherstellen oder den unten spezifizierten Originalexport liefern. Keine Produktionsänderung durch diesen Auftrag.

Der Katalog vom **08.11.2026** belegt historisch 31 public Tabellen, 308 Spalten, zehn RLS-deaktivierte Tabellen sowie gemeldete Versionen. Die ältere Diagnose aus #15 enthält zusätzliche historische Constraints, Policies, Trigger, sechs Helferdefinitionen und Teilindizes. Diese Stände sind **keine heute bestätigte Originaldefinition**, kein konsistenter vollständiger Export und kein vollständiger Originalrestore. Historische Exploit-Nachweise aus #25 bleiben offen; ein aktueller Zugriffsausfall widerlegt oder behebt sie nicht.

## 2. Konkrete Abweichungsmatrix

`security-step5-deviations.json` enthält **57 Objektzeilen**: alle 31 historisch belegten Tabellen plus 26 CREATE-/CREATE-OR-REPLACE-Objekte aus den beiden Sicherheitsdateien. Jede Zeile nennt Objekt, erwarteten Vertrag/Definition, tatsächlich verifizierte Definition, Abweichung, Risiko, betroffene Funktionen, Korrektur und Status. Exakte Kandidatendefinitionen werden über Dateipfad/SHA-256 referenziert, keine erfundenen Original-DDLs. Das gesamte Repository-SQL-Inventar enthält 13 Dateien mit Hashes; daraus folgt keine Produktivinstallation. ALTER-/GRANT-/Policy-Deltas und Fach-RPC-Familien sind zusätzlich nachfolgend aufgeführt. Diese Matrix ersetzt keinen vollständigen Katalogdiff.

| Objekt / Umfang | Erwartung / vorhandener Kandidat | Originalnachweis / Abweichung | Risiko / notwendiger Nachweis |
|---|---|---|---|
| planner_assignments, planner_absences, planner_vacations, planner_sick_leaves, planner_availability, planner_swaps, planner_shift_cancellations, planner_saturday_requests, planner_absence_replacements, tool_materials | RLS aktiv und passende Least-Privilege-ACL; #28 entzieht Tabellen- **und** Spaltenrechte, ersetzt Policies und sperrt Planner-Raw-DML | Historisch RLS aus; aktuell **UNVERIFIED** | Breite anonyme/authentifizierte Rechte historisch belegt; aktuelle ACL/Policies/FORCE RLS/Owner und effektive Rechte nachweisen |
| employees | Unabhängige Identität; geschützte Rolle, auth_user_id, department_id; eigene Kontakte getrennt | Historische permissive Mitarbeiterpolicies können mit Adminpolicies OR-verknüpft sein; aktuelle Policytexte/ACL **UNVERIFIED** | Nur zusätzliche Adminpolicy genügt nicht. Vollständige Policies einschließlich Rollen/PERMISSIVE/RESTRICTIVE sowie Spaltenrechte vergleichen |
| employees.role | UI/Fachmodell kennt department_admin | Historischer CHECK enthält admin/employee/tool_scanner; aktuelle Definition **UNVERIFIED** | Kein blindes ALTER CHECK. Originaldefinition und abhängige Fachmigrationen zuerst prüfen; AL-Enrollment erst nach expliziter Freigabe |
| employees Kontaktfelder | #28 schlägt getrennten contact store vor | Historisch keine Kontaktspalten; aktuelle Definition **UNVERIFIED** | D4 entscheiden; synthetic_contact bleibt Testwert, kein Originalfeld |
| departments / Mitarbeiterbindungen | Wirklicher Leiter, unabhängige gültige Zuordnung, kein Name als Authority | Historische IDs/FK teilweise vorhanden; aktuelle Constraints/Owner **UNVERIFIED** | D5/D6, unabhängige Zuordnung statt Bootstrap aus veränderbaren Mitarbeiterrollen |
| planner_special_days, planner_special_day_assignments, planner_tasks | Feld-/Bereichsfreigabe; unbekannte Zeilen gesperrt | Historische Struktur; aktuelle Regeln **UNVERIFIED**; special_day_assignments ohne direkten App-Aufruf im #29-Inventar | SQL-/Forecast-Pfad zusätzlich prüfen, D1–D3; nicht aus fehlendem Frontend-Aufruf auf Unbenutztheit schließen |
| humbel_security.* (neun Tabellen) | Leere unabhängige Authority-/Bindungs-/Kontakt-/Auditstrukturen | **Vorschlag**, kein Nachweis vorhandener Produktivobjekte | Kollisionen, Owner/BYPASSRLS, Trusted-Issuer, Aufbewahrung und Grants klären |
| current_app_employee_id/role, current_employee_id, is_admin, can_access_production_department | #28 ersetzt fünf historische Helfer durch vertrauensgebundene Definitionen | Sechs historische Helfer inklusive set_updated_at bekannt; aktuelle Bodies/Overloads/Owner/ACL **UNVERIFIED** | Vollständige Signaturen/Abhängigkeiten/Definer-Rechte vergleichen; kein Namensvergleich allein |
| security_actor_context, assignment_command, publish_assignment, published_schedule, own_contacts und private Helfer | Exakte #28/#29 Review-SQL, leerer search_path, authentifiziert und serverseitig geprüft | **Proposed only** bzw. Ersetzung, aktuelle Existenz unbekannt | Überladungen, EXECUTE/Default Grants und Least-Privilege-Owner prüfen; keine automatische Registrierung als Migration |
| production_*, tools, tool_journal, inventory_sessions/entries | Bestehende erlaubte Produktion/Scanner/Inventur erhalten | Historische Struktur; aktuelle RLS/ACL/Trigger **UNVERIFIED**; Kandidat deckt diese Pfade nicht vollständig ab | Native Alt-/Neu-Mischbetriebs- und Transaktionsabnahme; keine Admin-only-Ersatzpolicy |
| Artikel-/Operations-/BA-/Routing-/Rüst-/Mengen-/Lifecycle-/Forecast-Objekte | Alle bestehenden Review-/Migrationsdateien und AFTER STATEMENT Forecast-Korrektur aus #26 erhalten | Im historischen 31-Tabellen-Katalog nicht als vollständig installiert belegt; heute **UNVERIFIED** | Vollständiger Originaldiff einschließlich Funktionen/Trigger/Abhängigkeiten/Versionen nötig |
| Views, Materialized Views, Trigger, Indizes, Sequenzen, Schema-ACL, Default Grants, Owner, Rollen/Mitgliedschaften | Exakte Originaldefinition plus freigegebene gezielte Deltas | Vollständige aktuelle Definitionen **UNVERIFIED** | Unerwartete Definer-/View-Bypässe, EXECUTE/CREATE/USAGE, Sequence- und Ownerrechte ausschließen |
| Auth/API/Realtime | Bekannte exponierte Schemas, JWT-Rollen/Claims, Widerrufs-/Sessionregeln, Publications | Konfiguration heute **UNVERIFIED** | Serverseitige Identität und Zugriff über REST/Realtime einschließlich alten Tokens separat beweisen |

## 3. Zugriffsinventar weitergeführt

`security-step4-access-inventory.json` bleibt bytegleich. Die Ableitung `security-step5-access-plan.json` ergänzt **alle 152 Aufrufidentitäten auf 30 Tabellen** mit Zielrollen, beobachteter Projektion und exaktem Operationsargument am gepinnten Baseline-SHA (Payloadausdruck, nicht Laufzeitinhalt), erforderlicher RLS/RPC, isoliertem Unterstützungsstatus und konkretem Änderungsbedarf. `baseline_line` bleibt ausdrücklich die ursprüngliche Referenz; keine falsche Behauptung eines Laufzeitmitschnitts oder aktueller serverseitiger Rechte. Der jeweilige Originalfunktionsname bleibt erhalten. Schreibpayloads sind an dieser Funktion zu prüfen; `observedProjection` beschreibt eine vorhandene SELECT-Projektion, keine pauschale Aussage über Schreibfelder.

| Heutige Pfade | Daten / tatsächlicher Stand | Erforderlicher isolierter Umstellungsschritt |
|---|---|---|
| testSupabaseConnection, bootSupabase, getCurrentEmployeeRecord, syncSupabaseSessionToApp | Vor Login employees.id; danach employees SELECT*, Rolle aus Mitarbeiterprofil; anschließend viele breite Ladevorgänge | Kein Personendatenprobezugriff beim Boot. getUser + security_actor_context als Authority; explizite Projektion nach Rolle. Fehlender/mehrdeutiger/inaktiver Actor muss abbrechen, kein lokales Admin-Fallback. D5/D6 und vollständige UI-Tests fehlen |
| Mitarbeiter-CRUD/Slots/Aktivierung | Stammdaten und sicherheitsrelevante Zuordnungen | Administrative Felder und Authority-Provisionierung trennen. Änderung des Profils erzeugt keine Authority. Admin-Ansicht und sichere Fehlerbehandlung zuerst isoliert umstellen |
| Eigene Kontakte | #29-Adapter nutzt security_own_contacts ohne fremde Ziel-ID | D4: reale erlaubte Felder, Validierung, Store-/UI-Migration entscheiden. Keine Selbstserviceänderung von role/auth_user_id/department_id/personnel_no/Identität. Produktive Allowlist bleibt unentschieden |
| loadPlanningDataFromSupabase / generateThreeMonths | Neun SELECT*, plus Tasks/Sondertage; medizinische Daten und Namensbindungen | Freigegebene minimale Schichtprojektion statt gesamtem Planner-State. Stabile servergeprüfte IDs; keine automatische Namens-/Slot-Zuordnung. Privatsicht gemäß D1/D3 separat laden |
| assignShift/Optional/Springer, Samstagfreigabe, resetPlanCurrentFuture, Urlaub/Krankheit/Tausch/Anträge/Ersatz | Mehrere Tabellenänderungen, lokale Änderungen und Persistenz; Lesen/Planen/Bestätigen derzeit nicht vollständig sicher integriert | Einfache Zuordnung über vorhandenen assignment_command mit erwarteter Version; Publish separat. Kombinierte Aktion erst durch **ein** fachlich bestätigtes atomares Kommando ersetzen. Keine Teil-RPC plus unsicheres Rest-DML; D2/D3 blockiert |
| Fertigung: Maschinen/Stationen/Verantwortung/Ereignisse/History | Alte production_*-DML und neue BA-/OP-RPCs koexistieren | D7: Lesepfad/Wahrheitsquelle je altem/neuem Auftrag festlegen, doppelte Mengen-/Historienbuchung verhindern. Bestehende abteilungsübergreifende OP-Verantwortung erhalten |
| Werkzeuge: Laden, QR/Scan, Entnahme/Auffüllen, Umzug, Undo, Journal | Direkte tools/tool_journal-Zugriffe; Material-SELECT*; Inventur separat | Material-Adapterprojektion schon isoliert vorhanden. Bestand+Journal/Undo atomar, idempotente Buchung und parallele Scans erst nach belegtem Originalvertrag/Rollenmatrix. Scanner ist kein Admin. D8-Werkzeugentscheidung offen |
| Sessionwechsel/Logout/Realtime | #29-Modul verwirft nur eigenen Memorycache, prüft Actor bei jeder Aktion; Alt-State existiert weiter | Alte Requests/Subscriptions bei Actorwechsel entkoppeln, DOM und sensible Speicher generationengebunden löschen, Rechtewiderruf/Tokenablauf/Offline testen. Keine Daten aus alter Session übernehmen |
| persist(), load/save state; Login-/Tasklogs | snapshot entfernt viele Produktions-/Personenlisten, **nicht** Planner-/Abwesenheits-/Taskdaten; Logs enthalten u.a. employees/Planungsdaten | Vor Aktivierung Feld-Allowlist nur UI-Präferenzen; einmalige geprüfte Altcache-Bereinigung vor Restore; vertrauliche Logs entfernen. D3-Aufbewahrung und notwendige Datenerhaltung entscheiden. Keine stille Löschung ungesicherter Altänderungen |

Zusätzlich zu Tabellen sind folgende vorhandene RPC-Familien zu vergleichen, nicht neu zu erfinden: **article_master_capabilities/list/command; operation_plan; operations_article_command; order_from_article; operation_order_read/scan; ba_entry_command; machine_readiness_view/command/initialize; op_routing_view/command; setup_time_view/command; op_quantity_view/command; ba_lifecycle_view/command; production_forecast_*** sowie Forecast-Edge/Worker. Serverautorisierung, Übergaben, Rüst-/Stückzahl-/Nachfertigungs-/Abschlussregeln bleiben gemäß vorhandener Fachspezifikation. Aktuelle Produktionssignaturen/EXECUTE/Owner sind **UNVERIFIED**. Arbeitsstand wird über die 13 gehashten SQL-Dateien und die #29-RPC-Zuordnung referenziert.

### Bereits integriert versus noch offen

**Bereits ausschließlich isoliert vorhanden:** #28-RLS/ACL-Kandidat, unabhängige Authority-Bindung, Abteilungs-/Vertretungsgrenzen, versionierte einfache Planung, getrennte Veröffentlichung, eigener allowlistgesteuerter Kontaktstore, Materialrechte; #29-Serverkontext und standardmäßig deaktivierter Memoryadapter. Neue lokale Abnahme führt diese unverändert erneut aus. Es wird keine neue Sicherheitslücke als durch Schritt 5 geschlossen ausgegeben.

**Nicht vollständig integrierte Anwendung:** normaler Boot/Login und Personalansichten, Planner-Darstellung/Mehrtabellenaktionen/Privatsichten, Legacy-Produktion, ganze Werkzeug-/Inventurabläufe, allgemeiner Cache/Logs/Realtime sowie serverseitige Sitzungssperre. Die HTML-Datei lädt den Adapter weiterhin nicht. Keine produktiven Schalter, SQL-Dateien oder Grants geändert. Ohne Originaldefinitionen/Fachentscheidungen wäre eine Aktivierung eine mögliche Funktionssperre; daher keine teilweise Umstellung.

## 4. Originalexport und späterer Betriebsablauf

Ausführlicher, gateweise vorbereiteter Ablauf: **security-step5-migration-runbook.md**. Der Auftraggeber benötigt dafür zunächst einen bereinigten administrativen Schema-/Rechteexport mit Herkunft/Zeiten/Umfang/Hashmanifest und eine ausdrücklich zugelassene isolierte Restoreumgebung. Der aktuelle strukturelle Spiegel bleibt getrennt davon.

Die Review-Dateien sind **keine ausführbare Produktivmigration**: Isolationsguard, CREATE ohne Wiederanwendungssemantik, fehlende Originalabhängigkeiten und ungeklärtes Authority-Enrollment. Sie dürfen nicht durch Entfernen des Guards oder bloßes Kopieren nach migrations produktiv gemacht werden. Ein wiederholtes CREATE darf scheitern und die Transaktion zurückrollen; das ist **kein** Nachweis erfolgreich getesteter Upgrade-Idempotenz. Erst nach vollständigem Restore wird eine separate gehashte Vorwärtsmigration mit Precondition-/Versionjournal und Fail-Closed-Abbruch erstellt und abgenommen.

## 5. Nachweise und Prüfgrenzen

Vier neue lokale Nachweistests prüfen die ehrliche Kennzeichnung der blockierten Endpunkte, lückenlose Weiterführung des Originalinventars, historische/SQL-Provenienz und Nichtaktivierung des isolierten Clients. Das sind **keine** Original-RLS- oder Migrations-/Rollbacktests. Bestehende positive und negative PGlite-Sicherheitstests bleiben erhalten. Native JWT/REST/Mehrverbindungs-Tests und Originalrestore-Abnahme bleiben getrennte Nachweisklassen.

Die endgültigen lokalen Ergebnisse werden nach Abschluss der unveränderten Prüfskripte unten ergänzt. Fehlende Browser-/Containerlaufzeit bleibt BLOCKED_NOT_TESTED, auch wenn eingebettete SQL-Tests bestanden sind. Historische #28/#29-CI-Erfolge sind kein Testresultat dieses Commits.

## 6. Was fehlt konkret vor Produktivschaltung?

1. **Originalnachweis:** Connector-Metadatenzugang reparieren oder signierten/zugeordneten Schema-/Rechteexport nach Runbook liefern; vollständigen isolierten Restore und echten Objekt-/ACL-/Abhängigkeitsdiff durchführen. Heute fehlen aktuelle Policy-/Funktions-/Trigger-/View-/Owner-/Default-Grant-Definitionen und API/Auth-Konfiguration.
2. **Fachfreigabe D1–D6:** Leser/Felder/Veröffentlichung, Planungs-/Genehmigungs-/Abwesenheitsabläufe, Kontaktfelder, geprüfte Identitäts-/Altzeilenbindungen, Issuer-/Befristungs-/Widerrufsprozess entscheiden. Keine automatische Übernahme alter Rollen/Namen.
3. **D7/D8:** Legacy-/Neu-BA-Wahrheitsquellen sowie Werkzeug-/Scanner-/Inventur-/Undo-Rechte und Atomizität bestätigen. Bestehende berechtigte Fertigung darf durch Härtung nicht ausfallen.
4. **Vollständige Appintegration:** Boot, Datenprojektionen, sichere atomare Schreibkommandos, private Ansichten, Altcache-/Logbereinigung und Session-/Realtime-Invalidierung entwickeln und gegen den Originalrestore positiv/negativ abnehmen; keine Umgehung durch Raw-SELECT*.
5. **Technische Abnahme:** aktuelle vollständige Browsermatrix, unveränderte Visuals, echte native PostgreSQL/GoTrue/PostgREST- und Paralleltests; Originalkompatibilität, Fehler-Datenerhalt, wiederholte Migration/Abbruch/Rollback und Wiederherstellungsprobe. Blockierte Prüfungen nachholen, keine Baseline-Exploit-Reproduktion als Fix zählen.
6. **Betreiberfreigabe:** RPO/RTO, Backup/PITR/Schlüsselzugriff, isoliertes Restoreziel, Wartungsfenster, Verantwortliche und Audit-/Historienerhalt festlegen; nachgewiesenes sicheres Rückfallpaket. Pages-/Branchschutz-/Deploymentzustand aus #25 bleibt separat vor Veröffentlichung nachzuweisen.
7. Erst danach **separate ausdrückliche Entscheidungen** über Migration, Merge, Deployment und Produktivaktivierung. Dieser Draft erteilt keine davon.

## 7. Tatsächliche lokale Ergebnisse

Maschinenlesbar einschließlich tatsächlicher Exitcodes, Testzahlen und Loghashes: `security-step5-validation.json`. Alle vorhandenen Definitionen vor Ausführung geprüft, keine Paket-/Runtimeversion angehoben. Die lokalen Logs sind keine Originaldatenbankartefakte.

| Prüfung | Ergebnis |
|---|---|
| npm ci --ignore-scripts, check:syntax einschließlich beider Pflicht-Nodechecks, neue Testdatei-Syntax | PASS |
| test:ci | PASS **2/2** |
| test:qa | PASS **20/20** |
| qa:check | PASS; 251 Legacy-Identitäten, Geschäftslogik und 33 Bildreferenzen erhalten |
| test:integration | PASS **5/5** |
| test:security | PASS **32/32**, keine Skips: bisherige 28 plus vier Nachweistests; kein neuer Original-RLS-Nachweis |
| Neun bestehende Fachmodulsuiten | PASS **168/168**, keine Skips |
| qa:agent-server | PASS |
| test:e2e | **BLOCKED_NOT_TESTED**, Exit 1: vollständige 773 Fälle, 773 fehlende Executables vor Seitenaufruf, 0 bestanden. Fortschrittsreport complete=true / runnerStatus=failed; kein fachlicher UI-Nachweis |
| test:agents | **BLOCKED_NOT_TESTED**, Exit 1: Seed aufgerufen, fehlendes Browser-Executable |
| test:visual | **BLOCKED_NOT_TESTED**, Exit 2: vorgeschriebener Playwright-1.62.1-noble-Container fehlt |
| Native Sicherheits-/BA-OP-/Paralleltests und Originalrestore-/Upgrade-/Rollbackprobe | **BLOCKED_NOT_TESTED**: kein Docker/Podman/psql oder autorisiertes Originalrestoreziel; vorhandene native Workflows unverändert. test:db:preflight Exit 2 |
| Originale native Target-Guards im gepinnten Harness | PASS **2/2**; ausschließlich Zielschutz, keine Datenbankabnahme |
| git diff --check | PASS |

Keine Assertions/Tests entfernt, übersprungen oder abgeschwächt, keine geschützten Hashes oder PNGs angepasst. Keine Änderungen an SQL-Objekten; alle 13 SQL-Eingaben gehasht und unverändert. Keine neue produktive oder isolierte Policyaktivierung durch Schritt 5. Der später regulär ausgelöste GitHub-CI-Lauf wird hier weder abgefragt noch vorgestartet. Nach Erstellung des gestapelten Draft-PR endet der Auftrag.
