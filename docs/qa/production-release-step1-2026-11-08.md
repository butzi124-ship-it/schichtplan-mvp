# Technischer Freigabebericht: Produktivvorbereitung, Schritt 1

**Ergebnis: Produktivfreigabe weiterhin gesperrt.** Dieser Bericht bereitet Entscheidungen und Nachweise vor. Er implementiert keine neue Berechtigungsmatrix, repariert keine produktive RLS und erteilt keine Merge-, Deployment-, Migrations- oder Aktivierungsfreigabe.

Stand: **08.11.2026, Europe/Berlin**. Entwicklungsbasis PR #24, **0.15.0**, Commit `c4b75295ab4888ac759abaeec844991a30db40db`; Dokumentationsversion **0.15.1**, tatsächliche Uhrzeit im VERSION_LOG. Dokumentationsbranch: `docs/production-release-step1`. Untersucht: AGENTS.md, DEVELOPMENT.md, Fachspezifikation, aktuelle Module/SQL/Workflows, PR #15/#16 und #18–#24 einschließlich Kommentare und Reviews. Für alle neun untersuchten offenen PRs wurden keine Review-Einreichungen oder Inline-Reviewthreads zurückgegeben. Das ersetzt keine unabhängige Abnahme. Issue #17 einschließlich seiner leeren Kommentarliste gelesen.

## 1. Priorisierte Freigabesperren

| ID / Priorität | Befund und Beleg | Erforderliche Maßnahme / überprüfbares Abschlusskriterium |
|---|---|---|
| F01 / **P0** | Aktuelle Metadaten bestätigen zehn öffentliche Tabellen ohne RLS. Historische Originalmetadaten und native Diagnose aus #15 belegen Anon-DML und unbeschränkte Planner-Policies. Aktuelle ACLs wurden nicht erneut gelesen. | DBA/Security-Verantwortlicher: vollständige effektive Rechte erheben, endgültige fachliche Matrix bestätigen, gesonderte Korrektur entwickeln. Im gehärteten isolierten Zielstand sind Anon-DML und fremde Schreib-/Lesepfade tatsächlich gesperrt, zulässige Planungsaktionen bleiben erfolgreich. |
| F02 / **P0** | Historische `employees_insert/update/delete` erlauben `true`; zusätzliche Admin-Policies beseitigen diese permissive OR-Freigabe nicht. Native JWT-Diagnose änderte eine fremde Rolle zu `admin`. Nicht behoben. | Vertrauenswürdige Identität/Rolle/Abteilung von frei änderbaren Profilen trennen; vollständige Policy-/Spalten-/RPC-Analyse, kontrollierte Provisionierung und negative Eskalations-/Rebindingtests. Keine Übernahme alter Profilrollen als neue Vertrauenswurzel. |
| F03 / **P0** | Originalschema nur teilweise belegt; kein autoritativer aktueller Restore mit Ownership, ACLs, Default-Grants und vollständigen Funktions-/Triggerabhängigkeiten. | Metadatenpaket nach Abschnitt 3, Konsistenz-/Hashabgleich und isolierter Restore ohne Testsubstitute. Keine Differenz als implizit unkritisch einstufen. |
| F04 / **P0** | **Konkrete SQL-Divergenz:** #15 behebt 40P01 durch `forecast_source_dirty AFTER STATEMENT`; #24 enthält noch `BEFORE STATEMENT`. Seine native Pipeline verwendet aber die Prognose-SQL aus #15. | Bei späterer Integration geprüften AFTER-Fix erhalten; zu installierende SQL-Artefakte nach SHA exakt mit dem Integrationstree abgleichen. Dieselbe vollständige SQL-Kombination, echte Paralleltests und Datums-/Zeitzonentests ausführen. Keine Umsetzung in diesem Bericht. |
| F05 / **P0** | #16 ist ein sentinelgesperrter **Admin-only-Testkandidat**, keine bestätigte Schichtplanmatrix. #24 führt dessen Härtung nicht aus, sondern hängt seine Featuretests an denselben optionalen Testhook. | Fachlich bestätigte Rechte als separaten Härtungs-PR entwickeln; Baseline-Reproduktion und endgültige Korrekturtests als getrennte Testphasen erhalten. Gesamten Featurestand gegen gehärteten Originalrestore prüfen. |
| F06 / **P0** | GitHub-Pages-Quelle und aktuelle Ruleset-/Environment-Einstellungen nicht administrativ nachgewiesen. Ein Branch-Publishing kann trotz manuellem Actions-Deploy automatisch veröffentlichen. | Vor jedem potenziell veröffentlichenden Merge Settings-Nachweis nach Abschnitt 7, einschließlich tatsächlich geschützter Zielbranches und getesteter SHA. Keine Settingsänderung durch diesen Auftrag. |
| F07 / **P0** | Kein belegtes aktuelles Backup, kein gemessener erfolgreicher Restore, keine verbindlichen RPO/RTO und kein freigegebener datenbewahrender Rückfallplan. Admin-UI enthält dafür nur einen Platzhalter. | Verschlüsseltes Backup und Schema-/ACL-Sicherung separat autorisieren; Rücksicherung in eigener Umgebung nachweisen; Verantwortliche, Schreibstopp, RPO/RTO und Wiederanlauf entscheiden. |
| F08 / **P0 vor Pilot** | Übergang zwischen Legacy-Direktschreiben und neuen geschützten BA-/OP-Pfaden nicht auf vollständigem Originalschema abgenommen. | Verbindliche Zuordnung von Legacy-/verknüpften BA, Autorisierung jedes alten und neuen Schreibpfads, parallele Alt-/Neu-Clients sowie serverseitige Pilotabgrenzung testen. Kein pauschales Konvertieren von Altaufträgen. |
| F09 / **P1 vor Nutzung** | Original-`employees.role`-Constraint erlaubt laut aktuellem Metadatenabruf nur `admin`, `employee`, `tool_scanner`; App und Phase 2 verwenden zusätzlich `department_admin`. | Review der bereits vorhandenen Phase-2-Constraint-Transition und des Rollen-/Authority-Provisionierungsprozesses. Kein Produktions-ALTER oder automatisches Umetikettieren bestehender Mitarbeiter. |
| F10 / **P1 vor Planungspilot** | Planner-Zuordnungen verwenden Namen/Schlüssel; viele Tabellen haben keinen vertrauenswürdigen Personen-/Abteilungs-FK. App lädt alle Abwesenheits-/Planerzeilen und `employees.*`; `persist()` entfernt diese Planungsbereiche nicht aus seinem State-Snapshot. | Personenbindung, datensparsame Planungsprojektion und getrennte vertrauliche Abwesenheitsansicht spezifizieren. Browser-Caches, Logout und bestehende Logs prüfen; keine weiteren Daten spiegeln. Fachlich erforderliche Schichtansicht erhalten. |
| F11 / **P1 vor gemeinsamer Aktivierung** | Deaktivierte Prognose arbeitet mit ursprünglichem Artikelsnapshot; dynamische BA-Zielmengen/Nachfertigung sind im Phase-5-Bericht ausdrücklich noch offen. | Prognose und Worker deaktiviert lassen, bis Phase-6-/Planungsabgleich sowie Teilmengen-/Nachfertigungsregeln gemeinsam bestätigt und getestet sind. |
| F12 / **P1 vor Pilot** | Native Spiegeltests ersetzen keine Last-, Session-/Realtime-, Hardware- und Betriebsabnahme. Original PostgreSQL-Metadaten melden 17.11.0.003, gepinnter Teststack verwendet `supabase/postgres:17.6.1.136`. | Relevante Originalversion/Erweiterungen und API-Konfiguration im isolierten Stack nachvollziehbar angleichen, Unterschiede bewerten; Last-/Sperrzeiten, Scanner und sichere Beobachtbarkeit messen. |

P0 blockiert sichere Freigabe; daraus folgt **keine** unbestellte Änderung am laufenden System. Da F01/F02 bereits den Altbestand betreffen, sollte der Eigentümer eine gesonderte zeitnahe Sicherheits-/Eindämmungsentscheidung treffen. Eine mögliche zeitweilige Einschränkung des Altbetriebs benötigt einen ausdrücklich vereinbarten Betriebsplan und ist keine stillschweigende Admin-only-Produktivmatrix.

## 2. Was tatsächlich geprüft wurde

### 2.1 Aktuelle technische Originalmetadaten

Die vorhandene Verbindung wurde ausschließlich mit lesenden **Metadatenoperationen** verwendet: Projektliste zur Identifikation, `list_tables(public, verbose=true)`, Migrationsliste und Erweiterungsliste. Projekt **Schichtplan**, Ref `itntnpzqzyjfdnjdshvm`, stimmt mit der im Repository gesetzten SUPABASE_URL überein. Kein allgemeiner `execute_sql`-Aufruf, keine Abfrage von Anwendungs-/Auth-Benutzerzeilen, keine Logs mit Nutzdaten, keine Live-App-Navigation und keine produktive Änderung.

Das nach Auswahl rein technischer Felder gespeicherte [Metadatenartefakt](production-catalog-2026-11-08.json) enthält Aufnahmezeit, Tabellen-/Spaltendefinitionen, zurückgegebene PK-/FK-Metadaten, RLS-Flags, Migrationseinträge und installierte Erweiterungen. Vom Werkzeug mitgelieferte Tabellen-Zeilenangaben wurden **nicht übernommen oder als Datenbestand interpretiert**. Keine personenbezogenen Datensätze, Auftragsnummern, Fertigungsdaten, Schlüssel oder Passwörter wurden gelesen/exportiert. Spaltenoptionen wie `updatable` sind Strukturmerkmale, kein Nachweis effektiver Client-Grants. Die getrennten Metadatenaufrufe sind kein gemeinsamer transaktionaler Schema-Snapshot.

SHA-256 des bereinigten JSON-Artefakts: `d2ad520e4f06f4099623738004b89a05dc85d5491f5ad4791ff176c3f31e1968`. Er belegt die gespeicherten Bytes, keine vollständige Schemafidelität.

Aktuell bestätigt:

- **31 Tabellen / 308 Spalten** im abgefragten `public`-Bestand; Größenordnung entspricht dem älteren Inventar, beweist aber keine vollständige Gleichheit.
- RLS deaktiviert für `planner_assignments`, `planner_absences`, `planner_vacations`, `planner_sick_leaves`, `planner_availability`, `planner_swaps`, `planner_shift_cancellations`, `planner_saturday_requests`, `planner_absence_replacements` und `tool_materials`.
- `employees` und die übrigen zurückgegebenen Tabellen haben RLS eingeschaltet. Daraus folgt **keine** sichere Policy oder effektive Rechtebegrenzung.
- `employees.auth_user_id` wird als nullable/unique gemeldet; `department_id` und die `departments.leader_employee_id`-Relation sind im Schema vorhanden. Keine vorhandene personelle Zuordnung wurde gelesen oder bestätigt.
- Aktueller Rollen-Check auf `employees.role`: `admin`, `employee`, `tool_scanner`; nicht `department_admin`.
- Keine Artikel-/neuen BA-/Routing-/Setup-/Lifecycle-/Erstzuordnungstabellen unter den zurückgegebenen öffentlichen Tabellen. Das belegt die Abwesenheit dieser Objekte im abgefragten Schema, nicht den Zustand anderer Schemas oder sämtlicher Konfigurationen.
- Migrationsliste meldet genau `20260802073617 / add_full_production_counter_schema`. Diese Liste beweist nicht, dass es keine manuell vorgenommenen Schemaänderungen gab.
- Installiert gemeldet: `plpgsql 1.0`, `supabase_vault 0.3.1`, `pg_stat_statements 1.11`, `pgcrypto 1.3`, `uuid-ossp 1.1`. Keine Vault-Inhalte gelesen.

### 2.2 Historische Nachweise sauber abgegrenzt

`docs/qa/schema-inventory.json` vom 05.11.2026 ist ausdrücklich `restorable:false`: 31 Tabellen, 308 Spalten, 134 Constraints, 127 Policies, 11 Trigger. #15 ergänzt am 06.11.2026 **sechs Originalfunktionen** und **25 eigenständige Indizes** in `scripts/db-integration/source-catalog.js`; sein erzeugter Testkatalog prüft 76 Indizes insgesamt. Die frühere zu breite Testsubstitution der Produktions-Abteilungsberechtigung ist dort korrigiert. Die sechs Definitionen und 25 Indizes sind deshalb nicht mehr generell „unbekannt“; **aktuelle Vollständigkeit, Ownership und ACLs bleiben unbestätigt**.

Die fünf dort dokumentierten Definer-Helfer sind `can_access_production_department`, `current_app_employee_id`, `current_app_employee_role`, `current_employee_id` und `is_admin`; `set_updated_at` ist der sechste, nicht als Definer deklarierte Helfer. Die dokumentierten Definer verwenden `search_path=public`. Ob daraus aktuell ein ausnutzbarer Objektersetzungspfad entsteht, hängt unter anderem von tatsächlichen CREATE-/Owner-/EXECUTE-Rechten ab und wird ohne deren Prüfung nicht behauptet.

#15 nennt für seinen letzten Stand **14/14 native Spiegeltests**, #16 **15/15** einschließlich isoliert angewandtem Sicherheitskandidaten. Frühere Kommentare zu GoTrue-Ownership 42501 und Forecast-Deadlock 40P01 sind historische Diagnose-/Korrekturbelege, keine Behauptung, diese Startfehler bestünden am letzten #15-Head weiter. Insbesondere ist ein erfolgreicher Test, der eine Rolleneskalation reproduziert, **kein** erfolgreicher Sicherheitsfix.

Für #18–#24 wurden aktuelle Beschreibungen, Basis/Heads, Kommentare und Reviewstatus ausgewertet. In diesem Auftrag wurden keine neuen GitHub-Testläufe ausgelöst und keine historischen PR-Berichte als aktuelle CI-Ergebnisse ausgegeben. Die abschließende QA erfolgt separat auf dem späteren exakten Integrationstree.

## 3. Präzise noch zu beschaffende Originalinformationen

Der allgemeine SQL-Endpunkt bietet keinen für diesen Auftrag nachgewiesenen ausschließlich lesenden Katalogmodus und wurde wegen des Verbots produktiver SQL-Ausführung nicht verwendet. Für die folgenden Lücken ist ein **separat autorisierter Katalogexport durch DBA/Administrator** mit nachweislich lesendem Zugriff nötig. Hier wird kein ausführbares Produktions-SQL bereitgestellt.

| Paket | Benötigte technische Informationen | Abnahmekriterium |
|---|---|---|
| O1: konsistentes Schema | Schema-only-Export mit Zeitpunkt, PostgreSQL-/Extensionversionen, Schema-/Objektidentität, Spalten, Defaults, Identity/Generated, Domains/Enums, vollständigen Constraints samt Validierung/Deferrability, Partitionierung, Views/Materialized Views | Kein Datenteil; Hashmanifest; konsistenter Aufnahmepunkt; keine unterdrückten Owner-/ACL-Informationen ohne gesonderten gleichwertigen Export |
| O2: RLS | `pg_policies`/`pg_policy`: Schema, Tabelle, Name, permissive/restrictive, Kommandos, Rollen, USING/WITH CHECK; je Tabelle ENABLE/FORCE und Owner | Effektive Kombination aller Policies pro Rolle/Aktion nachvollziehbar, einschließlich PUBLIC und leerer Policies |
| O3: Funktionen/Abhängigkeiten | Vollständige Definitionen und Signaturen, Owner, proacl/Default-EXECUTE, SECURITY DEFINER/INVOKER, search_path/proconfig, Volatilität, Trigger-/View-/RPC-Abhängigkeiten | Alle exponierten RPCs und internen Privilegpfade inventarisiert; keine Substitutkörper im isolierten Restore |
| O4: Rechte | Schema-USAGE/CREATE, Tabellen-/Spalten-ACLs, Sequenz-USAGE/SELECT/UPDATE, Funktions-EXECUTE, Datenbankprivilegien, `pg_default_acl` je Ersteller/Schema | NULL-ACLs korrekt als implizite Defaults behandeln; effektive Rechte einschließlich Erbschaft/PUBLIC/Ownership berechnen |
| O5: technische Rollen | Rollennamen und Flags wie LOGIN, INHERIT, SUPERUSER, BYPASSRLS; Mitgliedschaften/ADMIN-/SET-Optionen; Owner und relevante Suchpfade | `anon`, `authenticated`, API-/Service-, Auth-, Owner- und Betriebsrollen exakt abgebildet; **keine** Passwortfelder aus `pg_authid`, keine Auth-Benutzer-/Sessionzeilen |
| O6: Trigger/Indizes | Alle Triggerdefinitionen einschließlich Timing, Row/Statement, Aktivstatus und referenzierter Funktionen; Event-/Constraint-Trigger; alle Indexdefinitionen, UNIQUE/Partial/Expression, Constraintbezug, Validität/Ready-Status | Bekannte 40P01-Sperrreihenfolge und neue Guard-Adapter gegen das echte Katalogbild geprüft; keine Indizes stillschweigend ausgelassen |
| O7: API/Betrieb | Exponierte Schemas und PostgREST-Rollen/Cache, Realtime-Publikationen, vertrauenswürdige Views, Extension-/Webhook-/Jobdefinitionen, Edge-Versionen sowie JWT-/Session-Konfiguration als Metadaten | Clientzugriffe, Servicepfade und Widerruf modelliert; keine Secrets, Token, Personen-/Produktionslogs oder Konfigurationspasswörter im Bericht |
| O8: Veröffentlichung | Aktuelle Rulesets/Branch Protection, erforderliche Checks, Bypassrollen, Pages-Quelle/Branch/Ordner, Environment-Freigaben, Deployment- und Artefakt-SHA | Nachweis durch Repository-Administrator; tatsächliches ausgeliefertes Artefakt eindeutig dem genehmigten Tree zugeordnet |

SQL-Funktionskörper oder Konfiguration können eingebettete Geheimnisse enthalten. Ein unredigiertes Vollpaket gehört deshalb in einen gesondert zugriffsgeschützten Prüfbereich, nicht ungeprüft in GitHub. Öffentlich dokumentieren: nötige technische Definitionen nach Secret-Prüfung, Provenienz und Hashes; keine Zugangsdaten. Schema-only ohne Owner/ACL ist keine Berechtigungsabnahme. Alle bislang nur synthetisch angenommenen Rechte müssen explizit gegen O1–O7 ersetzt oder als blockierende Abweichung behandelt werden.

## 4. Berechtigungsmatrix zur fachlichen Entscheidung

**Keine implementierte Produktivmatrix.** `B` = durch aktuelle Fachspezifikation/ausdrücklichen Auftrag bestätigt, aber technisch noch vollständig abzunehmen. `I` = im vorhandenen UI/Code beobachtete Nutzung, keine Autorisierung durch bloße UI. `O` = Vorschlag/offene Fachentscheidung. Jede Tabellenzelle gilt nur für authentifizierte aktive Identitäten, verifizierte Objekt-/Abteilungsbindung und die jeweilige unabhängige serverseitige Capability. AL ist kein frei beschreibbarer Profilwert. Ein aktiver Stellvertreter erbt nur ausdrücklich delegierten Bereich und Zeitraum.

### 4.1 Bestehende Schichtplanung, Mitarbeiter und Werkzeuge

| Aufgabe / Datenklasse | Normaler Mitarbeiter | Zuständiger AL | Aktive Stellvertretung | Administrator | Status / positiver Abnahmetest |
|---|---|---|---|---|---|
| Eigene Anmeldung, aktive Identität, eigene zulässige Stammsicht | Eigene ID und benötigte Anzeige-/Capabilityfelder lesen | Eigene + zuständige minimale Verzeichnissicht | Gleicher expliziter Scope | Geprüfte Personalverwaltung | I/O: Feldumfang verbindlich festlegen; Login bleibt möglich, keine fremden Personalnummern durch allgemeine `employees.*`-Lesesicht |
| Eigenen Dienstplan lesen | Eigene Dienste; sonstige erforderliche Teambelegung nach Entscheidung | Zuständigen Plan lesen | Delegierten Plan lesen | Betriebsweit nach bestätigtem Auftrag | I/O: bisherige „Meine Schichten“ erhalten; Team-/Abteilungs-/Betriebslesebereich bestätigen |
| Öffentliche betriebliche Schichtübersicht | Datensparsame Besetzungs-/Lückenansicht | Zuständiger Plan | Delegierter Plan | Geprüfter Gesamtplan | O: keine Krankheits-/Urlaubsgründe über dieselbe Projektion, aber planungsnotwendige Besetzung erhalten |
| Eigene Verfügbarkeit / Samstagsanfrage | Eigene Erklärung abgeben/ändern nach Fristregel | Im bestätigten Zuständigkeitsbereich bearbeiten | Nur delegiert | Nach festgelegter Verwaltungsaufgabe | I/O: bestehende Selbstbedienung aus `setWeekendAvailability`/`requestSaturdayEvening` positiv testen; Fremd-`user_name` darf keine Person bestimmen |
| Eigener Schichttausch / Abwesenheitsmeldung | Eigene Anträge, keine eigenmächtige Fremdzuweisung | Zuständige Prüfung/Freigabe, falls bestätigt | Aktive Delegation, falls bestätigt | Definierter Verwaltungsumfang | I/O: Antrags-/Freigabeworkflow, Fristen und Stornierung beschließen; nicht automatisch Admin-only |
| Schichten zuweisen, Ausfälle/Vertretungswochen planen | Nur ausdrücklich bestätigte Selbstbedienungsaktionen | Abteilungsplanung **offen** | Planungsdelegation **offen** | Vorhandene Verwaltungsfunktion erhalten | I/O: Planungsverantwortlicher ist als eigene Capability zu entscheiden, nicht aus „AL darf alles“ ableiten |
| Urlaubs-/Krankheitsdetails | Eigene Details; fremde Details nicht pauschal | Nur ausdrücklich fachlich erforderliche eigene Bereichsdaten | Nur erforderlicher delegierter Umfang | Kein pauschales medizinisches Leserecht allein aus technischer Adminrolle | O: vertrauliche Details, Planungsstatus, Freigabeverantwortung und Aufbewahrung separat entscheiden |
| Mitarbeiter anlegen/ändern/deaktivieren | Kein Recht auf role/is_active/auth_user_id/Abteilung | Keine automatische Personaladministration | Keine automatische Personaladministration | Vorhandene geprüfte Personalverwaltung; Rechteänderung als gesonderte privilegierte Aktion | I/O: Pflege zulässiger Felder erhalten; Identitäts-/Rollenänderungen nie über allgemeinen Profil-UPDATE freigeben |
| Abteilungsleiter-/Scope-/Vertreterzuordnung | Keine Selbstzuweisung | Keine eigene Rechteausweitung | Keine Delegation an sich selbst | Separater vertrauenswürdiger Provisionierungsprozess, Vier-Augen-Vorschlag | B für unabhängige Quelle; O für organisatorische Freigabe und Verfahren |
| Werkzeugverzeichnis / Buchung / Journal | Bisher erlaubte Mitarbeiterbuchungen erhalten | Bisher erlaubte Nutzung | Bisher erlaubte Nutzung | Stammdatenverwaltung nach bisherigem Auftrag | I/O: `tool_scanner` nur beauftragter Scan-/Buchungsumfang; technische Material-Referenzdaten nicht pauschal sperren |
| Scannerkonto | Kein Planner-/Personal-/BA-Recht aufgrund Scannerrolle | – | – | Scannerkonto verwalten, nicht mit Admin-Session teilen | B/I: reiner Scanpfad positiv, Zugriff auf Personal/Schicht-/BA-Funktionen negativ testen |

Diese Matrix gewährt noch kein neues AL-Schichtplanrecht und entzieht keine bestätigte bisherige Mitarbeiterfunktion. Die Fachseite muss vor Policyentwicklung ausdrücklich entscheiden, welche der bestehenden Funktionen teamweit, abteilungsbezogen oder personenbezogen vorgesehen ist. Dabei **zulässige** bestehende Abläufe zuerst als Positivtests sichern, nicht zufällige derzeitige RLS-Lücken als fachlich berechtigte Zugriffe konservieren.

### 4.2 Bestätigte BA-/OP-Regeln der Phasen 1–5

| Aktion | Mitarbeiter | Operativer OP-AL / aktive OP-Vertretung | BA-Haupt-AL / aktive Hauptvertretung | Administrator | Status |
|---|---|---|---|---|---|
| BA an OP 1 anlegen | Bei unabhängig geprüfter Erst-OP-Berechtigung, drei Angaben | Nur mit eigener gültiger Anlageberechtigung | Kein pauschales Anlagenrecht durch Titel | Geprüfter Scope | B |
| Automatischen Haupt-AL speichern | Keine manuelle Zielvorgabe; Serverregel aus gültiger Erst-OP-Abteilung | Keine eigenständige Zuweisung durch die Anlage | Genau ein aktiver passender autoritativer AL wird gespeichert | Fehlende Zuordnung über bestehende geprüfte Erstzuordnung nachholen | B; #24 |
| Maschine auswählen / abteilungsübergreifend festlegen | Geeignete aktive Maschine im zugelassenen Bereich, ohne bindende Sperre | Fachlich erlaubte Festlegung/Rückgabe, Audit | Keine zusätzliche operative Berechtigung aus BA-Titel | Unabhängig verifizierter Scope | B |
| Rüsten / Zählen / eigene Korrektur | Eigene berechtigte Arbeit/Korrektur mit Grund | Zuständige Abteilung und aktuelle Rechte | Keine automatische OP-Korrekturbefugnis | Geprüfter Scope; historische Daten erhalten | B |
| OP ausdrücklich abschließen | Berechtigter tatsächlicher Teilnehmer; offene Fälle sperren, unter Soll Grund | Zuständige OP-Leitung / aktive Vertretung | Nur bei zusätzlicher operativer Berechtigung | Geprüfter Scope | B |
| Mengenfall erklären / auflösen | Erklären und eigene berechtigte Korrektur; kein Auflösen | Auflösen, Restabweichung nur mit Pflichtgrund | Kein zusätzliches Auflösungsrecht ohne operative Zuständigkeit | Geprüfter Bereich | B |
| BA-Menge ändern / Nachfertigung freigeben | Nicht erlaubt | Nur mit zusätzlicher BA-Hauptberechtigung | Nach bestehender geprüfter Hauptverantwortungsregel; keine Ist-Daten löschen | Unabhängiger BA-/Abteilungsscope | B |
| Gesamten BA abschließen | Nicht erlaubt | Nur mit zusätzlicher Hauptberechtigung | Alle OPs/Fälle ausdrücklich abgeschlossen | Unabhängiger BA-/Abteilungsscope | B |
| Hauptverantwortung übertragen | Nicht erlaubt | Nicht allein durch OP-Zuständigkeit | Bisheriger Haupt-AL / aktive Hauptvertretung, Kataloggrund | Geprüfte Erst-/Übertragung; Kataloggrund | B |

Für fremde Abteilung, inaktive Identität, Widerruf, abgelaufene Vertretung oder manipuliertes Objekt gilt keine entsprechende Erlaubnis. Bereits gesehene Informationen können nicht rückwirkend „ungesehen“ gemacht werden; jeder neue Zugriff und Request-Replay muss die aktuelle Leseberechtigung prüfen. Session-/Cachebereinigung ist zusätzlich nötig.

### 4.3 Verbindlich offene Entscheidungen

| Entscheidung | Konkret zu bestätigen | Zuständige Freigabe |
|---|---|---|
| D1 | Team-, Abteilungs- oder Betriebsweite der Schichtübersicht; exakt freigegebene Felder | Schichtplanung/Fachverantwortung |
| D2 | Wer darf Fremdschichten, Ausfälle, Wochenvertretungen und Tausche ändern/freigeben? | Schichtplanung + zuständige Leitung |
| D3 | Welche Urlaubs-/Krankheitsangaben sehen Kollegen, AL, Vertretung und technische Administration? | Personal-/Datenschutzverantwortung + Fachseite |
| D4 | Zulässige eigene Mitarbeiterprofilfelder und verbleibende Admin-Pflege | Personalverantwortung + Security |
| D5 | Verifizierte Identitätsbindung für historische Namensschlüssel; Umgang mit Umbenennungen, unzugeordneten/mehrdeutigen Datensätzen | Fachseite + DBA; kein Datenzugriff in diesem Schritt |
| D6 | Erteilung/Entzug von Planer-, Artikel-, Maschinen-, AL-/Vertreter-/Admin-Capabilities; Befristung und Vier-Augen-Vorgabe | Eigentümer + Security |
| D7 | Historische BA-/Spannungsübernahme, UI-Führung für Alt-/Neu-BA und serverseitige Pilotabgrenzung | Fertigungsleitung + DBA |
| D8 | RPO/RTO, Wartungsfenster, Incident Owner, Stop-Kriterien und Wiederanlaufentscheidung | Eigentümer + Betrieb |
| D9 | Aktivierungsumfang ohne Prognose, späterer Phase-6-Abgleich und internes Übertragungsgrundsortiment | Fertigungs-/Planungsverantwortung |

Die jeweils verantwortlichen Personen sind organisatorisch noch zu benennen. Dieser Bericht erfindet keine Namen und setzt keine dieser Entscheidungen technisch um.

## 5. Konkreter Sicherheitskorrekturvorschlag für einen späteren eigenen PR

1. **Positivinventar einfrieren:** anhand Abschnitt 4 alle heute fachlich zulässigen Abläufe mit Actor, Objekt, Aktion, Feldern, Abteilung und Workflowzustand festhalten. Altes UI ist Funktionsnachweis, keine Berechtigungsquelle. Zustimmung zu D1–D6 dokumentieren.
2. **Vertrauenswurzel schließen:** Identität an verifiziertes `auth.uid()` binden; Admin-/Planer-/AL-/Stellvertreterrechte aus unabhängig verwalteten, für Clients nicht beschreibbaren Zuordnungen. Rollen, `auth_user_id`, Aktivität, Abteilung und Leiterzeiger nicht über beliebige Profilupdates manipulierbar lassen. Vorhandene erlaubte Personalpflege über geprüfte feldbegrenzte Verwaltungsaktionen erhalten; eigenen Profil-Selbstservice nur nach D4.
3. **Personen-/Abteilungsbindung modellieren:** historische `user_name`, `assigned_user`, `shift_id` und optionales `created_by_employee_id` reichen nicht als Subjekt- oder Eigentumsnachweis. `created_by` ist nicht automatisch die betroffene Person. Servergeprüfte Zuordnung zu stabilen IDs/Planungsbereichen vorsehen. Erforderliche spätere Datenklärung nur separat autorisiert; mehrdeutige Fälle nicht automatisch falsch zuordnen.
4. **Planner-Lesen trennen:** für Mitarbeiter notwendige Dienst-/Teamansicht datensparsam projizieren; vertrauliche Abwesenheitsdetails in eigenen geprüften Zugriffspfad. Bei Views Invoker-Verhalten bzw. vollständige eigene Autorisierung prüfen. `SELECT *` im Frontend auf abgenommenen Feldumfang umstellen; ein bloßer Clientfilter schützt keine bereits heruntergeladenen Daten.
5. **Planner-Schreiben transaktional prüfen:** eigene Anträge/Verfügbarkeit, berechtigte Teamzuweisung und Administration als unterschiedliche Aktionen. Serverseitig Actor/Subjekt/Abteilung/Status, Allowlist, Version/Idempotenz und Alt-/Neu-Audit prüfen. Keine Freigabe fremder Namen durch Payload.
6. **Policies und Grants gemeinsam korrigieren:** bekannte offene Employee- und Planner-Policies gezielt ersetzen, Anon-Rechte soweit fachlich nicht erforderlich entfernen, RLS/Default-Grants/Spaltenrechte/Sequenzen mit abnehmen. Eine zusätzliche restriktive aussehende permissive Policy genügt nicht neben `true`. Für allgemeine Tabellen-UPDATE-Rechte reichen nur geänderte Spaltengrants nicht ohne Entzug des weiterbestehenden Tabellenrechts. Keine pauschale Admin-only-Ersetzung aller Geschäftsabläufe.
7. **Definer-/Servicepfade separat prüfen:** sichere Ownerrollen, nicht vom Client beschreibbare Suchpfade, explizites EXECUTE und eigene Auth-/Scopeprüfung; keine neue SECURITY-DEFINER-Umschaltung als Antwort auf einen 403. Service-/Worker-Schlüssel nur serverseitig. Sessionentzug bei Rechtewechsel prüfen; ein JWT kann trotz abgemeldeter Session bis zu seinem Ablauf gültig bleiben, weshalb Rechte bei Zugriff aktuell serverseitig auszuwerten sind.
8. **Vorhandene Anwendungs-Caches minimieren:** `loadPlanningDataFromSupabase`, `loadEmployeesFromSupabase`, `persist()` und Datenlogs wie `planner_tasks raw data` im gesonderten Korrektur-PR auf personenbezogene Felder untersuchen. Bestehende Löschung von Produktions-/Werkzeugdaten aus dem persistierten Snapshot beibehalten. Kein Personal-/Abwesenheitsspiegel auf gemeinsamem Terminal; benötigte UI-Präferenzen erhalten.
9. **Abnahmepaket:** Katalogdiff vor/nach, effektive Rollen-/Objektrechte, genehmigte Matrix, alle Positiv-/Negativtests, exakte SQL-/Codehashes sowie geprüfter Rückfallplan. Erst daraus einen produktionstauglichen Migrationskandidaten erstellen; #16 bleibt Diagnose-/Lernbasis, nicht unverändert anzuwendendes Produktionsskript.

## 6. Integration der PRs und tatsächliche Konflikte

### 6.1 Beobachtete PR-Kette

Aktuelles `main` wurde per Git gelesen: **`e764cf2e93382e847627d8bc7c802b02ff70c375`, APP_VERSION 0.9.0**. Es ist gemeinsamer Vorfahr von #15 und #24. Das ist **kein** Nachweis der tatsächlich ausgelieferten Pages-Version. #2 und #11 sind laut aktuellem GitHub-Metadatenstand bereits gemergt; alte DEVELOPMENT-Hinweise zu ihrer damaligen Reihenfolge sind historische Bedingungen, keine erneut auszuführenden offenen Merges.

| PR | Aktueller Head | Basis | Einordnung |
|---|---|---|---|
| #15 | `25e9a6914e38fc791d9ecad13d92a45e30d19bb3` | main | Isolierte Infrastruktur + Forecast-AFTER-Fix, 0.9.13 |
| #16 | `bf14f4ed32fd8dc8e4e8374a5faa0c51a29416c7` | #15 | Nur gesperrter Sicherheits-Testkandidat, 0.9.14 |
| #18 | `91d13e4bed23174625a808213f6f0517ea046891` | main | Fachspezifikation |
| #19 | `8348fadc7bd1f9042b4147bd47568cb484547e3a` | #18 | Phase 1, 0.10.1 |
| #20 | `e276270adc704fc9033e304ff5ad81db56593ecd` | #19 | Phase 2, 0.11.5 |
| #21 | `6e7bd8a85c06b4cf73bb2f94114200b4e3d8dbde` | #20 | Phase 3, 0.12.4 |
| #22 | `72034961668e32b7503a7980f3382c5deb4a84e6` | #21 | Phase 4, 0.13.2 |
| #23 | `05b57de2f1dfcedf5575a4064d8ff79acafa0d1b` | #22 | Phase 5, 0.14.2 |
| #24 | `c4b75295ab4888ac759abaeec844991a30db40db` | #23 | Automatische Erstverantwortung, 0.15.0 |

Alle neun stehen weiterhin offen/Draft. #15 und #24 sind **divergiert**; GitHub-Compare meldet 13 Commits nur auf #15 und 22 nur auf #24, Merge-Base `e764cf2…`. Ein grüner Test eines Stapels enthält nicht automatisch den anderen Stapel.

### 6.2 Vor einer späteren Zusammenführung aufzulösen

- **Forecast-Trigger:** #24 `supabase/review/issue-10-production-forecast.sql` erzeugt `forecast_source_dirty before ... for each statement`; [#15 am verifizierten Head](https://github.com/butzi124-ship-it/schichtplan-mvp/blob/25e9a6914e38fc791d9ecad13d92a45e30d19bb3/supabase/review/issue-10-production-forecast.sql) verwendet `after` mit dokumentiertem 40P01-Grund. Die #24-Abnahme checkt #15 separat aus und überlagert nur Phase-1–5-/Erstzuordnungs-SQL. Ihre Ergebnisse beweisen daher nicht die Prognose-SQL-Datei des #24-Trees. Zukünftige Abnahme muss alle tatsächlich zu installierenden Dateien aus einem eindeutig geprüften Kandidaten nehmen oder jede Mischung vollständig per SHA-Manifest belegen.
- **Testhook-Kollision:** #16 nutzt `tests/db-integration/security-extension.js`; #24 kopiert dorthin seine Erstzuordnungserweiterung, welche die Phasen 1–5 aufruft. Die #16-Härtungsprüfung ist darin nicht enthalten. Integrierte Harness-Reihenfolge ausdrücklich komponieren; keinen Hook still überschreiben.
- **Baseline versus Fix:** #15 erwartet im Diagnosefall absichtlich erfolgreiche Anon-Zugriffe und Mitarbeitereskalation. Diese Assertions unverändert nur auf der unsicheren synthetischen Baseline ausführen; danach eine getrennte gehärtete Umgebung mit negativen Angriffstests und allen erlaubten Geschäftsabläufen abnehmen. Ein Produktionskandidat darf nicht deshalb als sicher gelten, weil seine Exploit-Reproduktion grün ist.
- **Geschützte Tests/Versionsdateien:** package-/lock-Dateien, Versionslogs, e2e-/Native-Workflows und geschützte SQL-/Testhashes nachvollziehbar vereinigen. Originalassertions und Bildreferenzen nicht neu „grün verankern“, um Konflikte zu verstecken. Exakte gewollte Integrationsadapter dokumentieren.
- **Legacy-Schreibpfade:** `completeProductionOrderWithChecklist`, alte Stations-UPDATEs und Personal-/Planer-DML dürfen nach Aktivierung nicht an neuen Abschluss-/Zeit-/Auditregeln vorbei schreiben. Neue Guard-Fehler dürfen aber auch keine fachlich erlaubten Alt-BA-Aktionen ohne vorgesehenen Ersatz blockieren. Mischbetrieb mit noch geöffneten alten Browsern testen.
- **Altaufträge:** Legacy-BA ohne Snapshot/OP-Verknüpfung bleiben nachvollziehbar Legacy. Keine Ableitung einer ersten OP allein aus Spannung/Code `1AA`, kein synthetisches Gutteil und keine Umschreibung alter Fertigungszeiten. Eine spätere Übernahme braucht eigenen geprüften Prozess.
- **Planungssemantik:** Ursprüngliche Snapshot-Sollmenge bleibt historisch; aktuelles BA-Soll, reale Gutteile, Abschnittsbaseline und bestätigte physische Übergabe sind verschiedene Werte. Bis zum gemeinsamen Planungsabgleich Prognose/Worker nicht mit Nachfertigung gemeinsam aktivieren.
- **Erstzuordnungsdokumentation:** In älteren Berichten und in Abschnitt I.4 der Fachspezifikation steht noch „Regel offen“, während die spätere ausdrückliche Entscheidung in A.3 und #24 umgesetzt ist. Das ist ein dokumentarischer Nachzugspunkt, kein Anlass, die bereits beschlossene Regel neu zu erfinden.

### 6.3 Kontrollierte zukünftige Integrationsreihenfolge

1. F01–F07 und D1–D6 mit Verantwortlichen klären. Pages-Schutz **vor** einem Merge mit möglicher Veröffentlichung nachweisen.
2. In getrenntem Integrationsbranch auf dann aktuellem main die geprüfte #15-Infrastruktur einschließlich AFTER-Fix mit #18–#24 vereinigen. #16 nur als gesperrten Test-/Reviewbeleg behandeln; endgültige Härtung aus eigenem fachlich bestätigtem PR.
3. Entweder Stapel einzeln #18→#19→#20→#21→#22→#23→#24 integrieren, nach jedem Squash/Rebase Nachfolger explizit neu auf tatsächliches main basieren, **oder** einen überprüfbaren kumulativen Integrations-PR verwenden. Nicht beides und nicht bereits enthaltene Änderungen ein zweites Mal übernehmen. Entscheidung, Reviews und SHA-Zuordnung protokollieren.
4. Gesamten Integrationskandidaten auf autoritativem isoliertem Schema wiederherstellen und Sicherheit, Feature-SQL, native/Browser-/Visual-Regressionskette ausführen. Danach erst separate Entscheidung über Code-Merge. Alle Schalter bleiben aus.
5. Produktionstaugliche Migrationen erst aus dem abgenommenen Gesamtstand registrieren. Vorgesehene fachliche Abhängigkeit: Artikel #7 → OP-Snapshot #8 → Readiness #9 → geprüfte Forecast-Grundlage #10 (Schalter/Worker aus) → Phase 1 → Phase 2 → Phase 3 → Phase 4 → Phase 5 → automatische Erstzuordnung. Sicherheit/Identitätsbindung ist Voraussetzung der Aktivierung. Physische Reihenfolge bei etwaigen FK-/Trigger-Abhängigkeiten im Restore beweisen; keine Dateikopie als Migrationsfreigabe.
6. Codeintegration, DB-Migration, Deployment und Modulaktivierung als **getrennte** genehmigte Betriebsschritte ausführen. Jeder konkrete Schritt nennt unveränderlichen Artefakt-/Commit-SHA, Verantwortlichen und Rückfallpunkt. Dieser Auftrag führt keinen davon aus.

## 7. GitHub-Pages- und CI-Schutz prüfen lassen

Aktuelle Workflows belegen die Codekette `functional + integrity + visual → quality-gate → e2e`, jeweils fail-closed. `Deploy Tested Main` verlangt `workflow_dispatch`, `main` und `deploy_approved=true`. `scripts/stage-pages.js` kopiert nur App-Dateien, `js`, `img` und `qa-commit.txt` in das Pages-Artefakt; SQL/Test-/Metadatenberichte werden durch diesen Stagingpfad nicht absichtlich veröffentlicht.

**Nicht aktuell belegt:** tatsächliche Pages-Quelle, aktuelles Ruleset/required Checks, Bypassakteure, Environment-Reviewer und tatsächlich ausgelieferter Commit. Die Connectorwerkzeuge stellen dafür keinen geeigneten Administrationseinstellungsabruf bereit; kein Browserlogin und keine Settingsänderung erfolgt. Das früher dokumentierte Ruleset 26664844 ist ein historischer Beleg, kein aktueller Admin-Nachweis.

Repository-Administrator soll vor Freigabe dokumentieren:

1. Settings → Pages → Build and deployment: Source; falls „Deploy from a branch“, Branch und Ordner sowie Konsequenz jedes geplanten Push/Merge. Nur eine gesondert autorisierte Umstellung darf dies ändern.
2. Aktive Rulesets/Branch Protection für main: PR-Zwang, Aktualität, genehmigte Reviewpflicht, keine unerwünschten Bypässe, tatsächlich erforderliche Checks. **Die native DB-Abnahme ist aktuell ein separater Workflow und keine Abhängigkeit des e2e-Gates**; vor Produktivfreigabe muss ihr Erfolg auf demselben Kandidaten verbindliches Kriterium sein. Ob dies zusätzlich als required Check abgesichert wird, ist administrativ zu entscheiden.
3. Environment `github-pages`: Reviewer, erlaubte Branches, Schutz vor Selbstfreigabe soweit vorgesehen; Rollen für `pages:write`/`id-token:write` und weitere deployende Workflows kontrollieren.
4. Tatsächlich ausgelieferte Version/Artefakt-SHA gegen genehmigtes `qa-commit.txt`/Deployment-Artefakt zuordnen. `main=0.9.0` oder ein grüner PR-Testmerge ist nicht automatisch die ausgelieferte Version.
5. Cache-/Stale-Client-Plan: alten Tab/alte Assets nach Rechte-/Schemawechsel testen, Version sichtbar machen, laufende Buchungen nicht durch erzwungenen Reload verlieren.

## 8. Isolierte Abnahme und Regression: konkreter Nachweiskatalog

| Test-ID | Erwartung / Nachweis |
|---|---|
| SEC-01 | Anon kann keine Mitarbeiter-, Abwesenheits-, Planner- oder Fertigungsdaten lesen/ändern; ausnahmsweise bestätigte öffentliche Referenzansichten separat benennen |
| SEC-02 | Mitarbeiter kann weder eigene noch fremde `role`, `auth_user_id`, Abteilungs-/Leiter-/Adminbindung oder Aktivität über direkte Tabellen-/Spalten-/RPC-Payloads manipulieren |
| SEC-03 | Alle fachlich bestätigten Positivfälle aus D1–D4: eigene Dienste lesen, Teambelegung sehen, Verfügbarkeit/Samstagsanfrage/Tausch sowie berechtigte Planer-/Adminpflege bleiben erfolgreich |
| SEC-04 | Kein fremder Personen-/Abteilungszugriff durch Objekt-ID, Namen, Suchfilter, Join, View, RPC, Request-Replay oder Realtime; keine vertraulichen Abwesenheitsfelder in Planungsprojektionen |
| SEC-05 | Abgelaufene/widerrufene Delegation, deaktivierter Benutzer, veränderte Rolle, altes JWT und Sessionwechsel; keine Fortsetzung aus Browsercache oder manipuliertem user_metadata |
| SEC-06 | Öffentliche/private Funktionsgrenzen, EXECUTE-Defaults, Owner/BYPASSRLS und CREATE/search_path; Angriffe können keine privilegierten Helfer oder Transaktionsfähigkeiten ersetzen |
| DB-01 | Autoritativen Schema-only-Stand mit realen Owner-/ACL-/Default-/Trigger-/Indexdefinitionen isoliert restoren; Versions-/Hashmanifest vergleichen; keine stillen Substitute |
| DB-02 | Alle zu installierenden SQL-Dateien aus **demselben** freizugebenden Integrationstree anwenden; populated-Phase-4→Phase-5-Migration ohne 55006 und mit deferred-FK-Prüfung |
| DB-03 | Separate echte Verbindungen: BA-Duplikat, Idempotenz, Mengen-/OP-/BA-Abschlussrennen, Scope-Widerruf und Erstzuordnung; 40P01-Fall mit geplanter Triggerreihenfolge |
| DB-04 | Audit-/Notification-/FK-Fehler führen zum vollständigen Rollback; immutable Historie, alte Stückzahlen, Ausschuss, Klärfälle, Rüstzeiten und Abschlusssnapshots erhalten |
| DB-05 | Historische Korrektur unter richtigem Actor, getrennte OP-/BA-AL-Rechte, Restklärung, Nachfertigung nach Abschluss und reale Folgemengen |
| DB-06 | Teil-/Gesamtmigration, erneuter Start nach Fehler, Schema-/ACL-Rückfall auf isolierter Kopie; nach echten Schreibvorgängen kein Drop-/Überschreibrollback |
| QA-01 | 733 bestehende funktionale Fälle aller fünf Profile sowie 168 aktuelle lokale Modulverträge; keine Entfernung/Skip/Assertionabschwächung |
| QA-02 | 33 strikte Bildvergleiche zweimal im gepinnten Container, unveränderte Bildreferenzen/Toleranzen und 251 Legacy-Identitäten |
| QA-03 | Echte Session-/Token-/Realtime-Tests mit Positiv- und Negativnachweis, Mehrbenutzer-/Altclient-Betrieb, Zeitzonen/Tagesgrenzen und DST mit deterministischer Testuhr |
| OPS-01 | Last-/Lock-/Timeoutmessung der globalen Lifecycle- und SHARE-Sperren; Hardware-Scanner/Code128-Druck, mobile Bedienung, sichere Logs/Monitoring |
| OPS-02 | Restoreübung mit gemessener RPO/RTO, Stopp/Drain, Rollbackentscheidung, erneuter Integrityprüfung und dokumentiertem Pilotabbruch |

Die bestehende native Featurekette sieht **62** Fälle (52 bisherige + 10 Erstzuordnung) vor. Das ist kein Ersatz für zusätzliche SEC-/Originalrestore-/Betriebskriterien. Die Exploit-Reproduktion aus #15 bleibt separat erhalten; ein später gehärteter Stand muss die Angriffe verweigern. Ein leeres REST-Ergebnis allein beweist keine erfolgreiche Berechtigungsprüfung: positive Kontrollidentität, konkrete Status-/Mutationsprüfung und Zustand vor/nach gehören dazu. Nicht auf Produktionsdaten testen.

## 9. Backup und überprüfter Rückfallplan

**Noch nicht ausgeführt und noch nicht freigegeben.** Backupdienst, verfügbare PITR-Aufbewahrung und Wiederherstellungsrechte müssen vom Betreiber belegt werden; Verfügbarkeit eines bestimmten Supabase-Tarifs wird nicht angenommen.

| Zeitpunkt | Vorzubereitende Handlung | Entscheidungs-/Prüfbeleg |
|---|---|---|
| Vor Änderungen | Code-/Schema-/Policy-/Owner-/ACL-/Default-Grant- und Flagmanifest; verschlüsseltes Datenbackup oder bestätigter PITR-Punkt im Betriebsbereich | Zugriff, Integrität und Lösch-/Aufbewahrungsregel; keine Echtdaten in öffentlichen QA-Artefakten |
| Probe | Wiederherstellung in separat freigegebener abgeschotteter Umgebung; Jobs/Webhooks/E-Mail/Worker zunächst aus | Zeitmessung, Counts/Constraints/Relations-/Geschäftsinvarianten durch autorisierten Betreiber; synthetische Probe ersetzt keine Backuplesbarkeit |
| Vor Migration | Genehmigtes Wartungsfenster, Schreibstopp/Drain, aktive Buchungen behandeln, RPO/RTO und Incident Owner festlegen | Restorepunkt und geplanter maximaler Datenverlust ausdrücklich bestätigt |
| Vor Commit | Einzelne Migration atomar mit Sperr-/Statement-Timeout und Vorbedingungen | Fehler bricht Transaktion ab; keine halb angewandte Fortsetzung |
| Nach Commit, vor Schreibfreigabe | Berechtigungs-Smoke, Versions-/Kataloghash und Integritätscheck; Schalter/Worker weiterhin aus | Bekannter alter Codepfad funktioniert im genehmigten Umfang, ansonsten kein Pilotstart |
| Nach Pilot-Schreibvorgängen | Bei Fehler zuerst betroffenen Schreibpfad stoppen, Sessions/Worker kontrolliert sperren, neue Daten/Audits sichern | Keine neuen Audit-/Snapshot-/Zählertabellen löschen, keine historischen Ist-Werte überschreiben |
| Wiederanlauf | Vorzugsweise abwärtskompatibler Forward-Fix oder geprüfter UI-/Flag-Rückfall mit Datenbewahrung | Alter Frontendstand muss zum neuen Schema passen; vollständiges Restore/PITR nur separate Incident-Entscheidung mit offenem Verlustfenster |

Ein Rückfall darf die bekannten Anon-/Rollenlücken nicht unbemerkt wieder öffnen. Der Policy-Rückfall wird zwar isoliert auf technische Reversibilität getestet, aber unsichere alte ACLs werden **nicht** automatisch produktiv zurückgespielt. Bei einem solchen Konflikt: genehmigter sicherer Wartungszustand und kontrollierter Fix statt behaupteter Wiederherstellung eines sicheren Zustands. Welche vorübergehenden Funktionen im Wartungszustand verfügbar bleiben, ist Bestandteil der separaten Betriebsentscheidung.

## 10. Kontrollierte Aktivierung und Pilot – spätere Ausführung

1. Alle P0-Nachweise, Entscheidungen und unabhängiges Review vollständig; Release Owner und exakte SHAs festlegen. Erfolg einer CI ist keine fachliche oder produktive Freigabe.
2. Deaktivierten Gesamtcode integrieren und erst nach gesonderter Zustimmung veröffentlichen; vorher Pages-/Environment-Schutz nachweisen. Keine neue Produktivfunktion allein durch Frontendflag aktivieren.
3. Separat genehmigte DB-/Sicherheitsmigration auf dem vorher abgenommenen Gesamtstand durchführen; serverseitige und Frontendmodule zunächst aus. Originalprofile/Scopes nicht ungeprüft bootstrapen.
4. Vertrauenswürdige Rollen, Erst-OP-/Maschinenberechtigungen, AL-/Stellvertreterzuordnungen, Maschinen-Eignungen und Grundkataloge durch den bestätigten Verwaltungsprozess provisionieren. Fehlende Erst-AL-Zuordnung bleibt „Hauptverantwortung offen“, kein künstlicher Ersatz-AL.
5. Pilot nur für vereinbarten Bereich und geeignete neue BA. **Noch nachzuweisen ist eine serverseitig belastbare Pilotabgrenzung**; die bestehenden globalen Gates allein belegen keine Bereichs-Isolation. Kein Testschreiben in reale Aufträge als Ersatz für Staging.
6. Reihenfolge mit Voraussetzungstest: Artikel/OP-Snapshot → BA-Einstieg → Routing/Readiness → Rüstzeit → Spannungszählung → Klärfälle/Abschlüsse/Mengen → automatische Hauptverantwortung. Serverseite und UI als kompatibles Paar schalten; bei Störung keine unzulässigen alten Direktbuchungen zulassen. Prognose/Worker bis Phase-6-Abgleich aus.
7. Vereinbartes Pilotfenster mit geschulten Nutzern, eindeutigem Support-/Abbruchweg und manueller fachlicher Gegenprüfung. Nur datensparsame technische Metriken beobachten: Berechtigungsfehler, Latenz/Locks/Deadlocks, Auditfehler und Wiederholungsrate; keine Personen-/Auftragsdaten in allgemeine Diagnoseausgaben.
8. Sofortige Stop-Kriterien: unberechtigter Datenzugriff, verlorene/duplizierte Buchung, veränderliche Historie, falsche Haupt-/OP-Zuständigkeit, Auditlücke oder nicht akzeptierte Sperr-/Latenzwerte. Quantitative Latenzgrenzen und Beobachtungsdauer vom Betrieb festlegen, nicht hier als schon vereinbart ausgeben.
9. Rollout nur nach dokumentiertem erfolgreichem Pilotabschluss, unverändertem freigegebenem Artefakt und erneuter ausdrücklicher Zustimmung. Keine automatische Vollaktivierung.

## 11. Lokale Prüfungen dieses Dokumentationsstands

Keine Funktions-, SQL-, RLS-, Auth-, Test- oder Workflowänderung durch diesen PR. Nur Bericht, bereinigter technischer Metadatennachweis, Verlinkung und APP_VERSION/VERSION_LOG.

| Prüfung | Ergebnis |
|---|---|
| `npm ci --ignore-scripts` | bestanden |
| Beide Pflicht-`node --check`, `npm run check:syntax` | bestanden |
| Unverengte Artikel-/OP-/Readiness-/BA-/Routing-/Rüst-/Mengen-/Prognose-/Lifecycle-Verträge | **168/168 bestanden**, 0 Fehler/Skips |
| CI-Bericht / QA-Gates | **2/2** / **9/9 bestanden** |
| `qa:check`, `qa:agent-server`, Metadaten-JSON-/Link-/Diffprüfung | bestanden |
| `test:e2e` | Alle 733 Fälle aufgerufen; fehlende Browser-Binaries blockieren lokale fachliche Browserabnahme |
| `test:agents` | Aufgerufen; Browserstart wegen fehlender Binaries blockiert |
| `test:visual` | BLOCKED_NOT_TESTED; vorgeschriebener Playwright-1.62.1-noble-Container lokal nicht verfügbar |
| `test:db:preflight` / native Gesamtabnahme | BLOCKED_NOT_TESTED; isolierte Dockerlaufzeit fehlt; keine Produktionsersatztests |
| Produktive Daten-/Sicherheitsprüfung | Nur die beschriebenen technischen Metadaten; kein Live-SQL, DML, RPC-Geschäftsaufruf oder Exploitversuch |

Die obigen lokalen Ergebnisse wurden vor dem Push finalisiert. Für den neu hochgeladenen Commit werden keine GitHub-Resultate vorausgesetzt. Konfiguriert sind **QA and Release** und **BA and OP Isolated Acceptance**; nach Push/Erstellung des Dokumentations-Draft-PR unmittelbar beenden, keine CI-Abfragen oder Warteschleifen.

## Quellen und Nachweisgrenzen

- Verbindliche [Fachspezifikation](../requirements/ba-op-shopfloor-2026-11-07.md), [Entwicklungsregeln](../DEVELOPMENT.md), [Artikel-/RPC-Sicherheitsgrenzen](../articles/SECURITY.md), [Phase-5-Bericht](lifecycle-phase5.md), [Erstzuordnungsbericht](ba-initial-responsibility.md).
- [PR #15](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/15) / [Originalfunktions- und Indexmetadaten](https://github.com/butzi124-ship-it/schichtplan-mvp/blob/25e9a6914e38fc791d9ecad13d92a45e30d19bb3/scripts/db-integration/source-catalog.js) / [Isolations- und Restoregrenzen](https://github.com/butzi124-ship-it/schichtplan-mvp/blob/25e9a6914e38fc791d9ecad13d92a45e30d19bb3/docs/qa/isolated-database-acceptance.md).
- [PR #16](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/16) / [technisch produktivgesperrter Härtungskandidat](https://github.com/butzi124-ship-it/schichtplan-mvp/blob/bf14f4ed32fd8dc8e4e8374a5faa0c51a29416c7/supabase/review/legacy-security-hardening-isolated.sql). Alte Kommentarstände sind nicht die aktuellen Heads.
- Entwicklungs-PRs [#18](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/18), [#19](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/19), [#20](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/20), [#21](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/21), [#22](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/22), [#23](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/23), [#24](https://github.com/butzi124-ship-it/schichtplan-mvp/pull/24); genaue Heads in Abschnitt 6.
- Fachliche Bewertung und Korrekturplan in diesem Bericht sind Schlussfolgerungen aus diesen Quellen, **keine** Implementierungs- oder Freigabenachweise.
- Technische Primärreferenzen, am 08.11.2026 gelesen: [PostgreSQL 17: CREATE POLICY und permissive OR-Kombination](https://www.postgresql.org/docs/17/sql-createpolicy.html), [Supabase RLS, Zugriffsebenen und Authgrenzen](https://supabase.com/docs/guides/database/postgres/row-level-security), [GitHub Pages: Branch- oder Actions-Publishing](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site).

**Nächster Arbeitsschritt nach separater Beauftragung:** O1–O8 beschaffen und D1–D9 bestätigen, danach eigener isolierter Sicherheits-/Integrations-PR. Nicht weiter produktiv schalten, bevor die beschriebenen Sperren nachweislich aufgehoben sind.
