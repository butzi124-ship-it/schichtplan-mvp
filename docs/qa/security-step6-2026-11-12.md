# Sicherheitsfreigabe Schritt 6 – bestätigte Originalrechte und isoliertes Reviewpaket

**Keine Produktivfreigabe. Kein vollständiger Originalrestore.** Gestapelter Draft auf PR #30, verbindlicher Parent `35a56bb07e58b7c1a12fbc0654af064740dd292e`, neuer Branch `security/rls-migration-step6`, APP_VERSION **0.16.3**, VERSION_LOG **12.11.2026 12:22 Europe/Berlin** (tatsächlicher Änderungszeitpunkt). Refs #17, #15, #16 und #25–#30. Keine Änderungen an alten PRs, Startseite, Schichtplanlogik, Werkzeug-/BA-/OP-Modulen oder produktiven Schaltern. Die Haupt-App erhält ausschließlich Version/Log.

## 1. Fünf getrennte Nachweisklassen

| Klasse | Befund / Beweisgrenze |
|---|---|
| **Direkt bestätigtes Original, manuell geliefert** | Der Auftraggeber hat am 12.11.2026 im Supabase-Dashboard „Schichtplan“, Branch main, lesende Katalogabfragen ausgeführt. Verbindung postgres/postgres, PostgreSQL 17.11; 31 public Tabellen, 21 RLS an, zehn RLS aus; auf allen zehn haben anon und authenticated SELECT/INSERT/UPDATE/DELETE. **20 Kombinationen, 80 positive Privilegergebnisse.** Vom Agenten nicht erneut ausgeführt; kein maschinenlesbarer Originalexport oder genauer Aufnahmezeitpunkt geliefert. |
| **Historische Originalbefunde** | #15/#25: 31 Tabellen/308 Spalten, ältere Constraints/127 Policies/11 Trigger, sechs Originalhelper und Teilindizes; #25-Katalog vom 08.11. und #30-Connectorbefund 28P01. Diese Definitionen bleiben historisch, nicht heute vollständig bestätigt. |
| **Isolierter korrigierter Kandidat** | Unveränderte #28/#29-SQL plus neues additives Material-ACL-Delta; neues atomisches Installations-/Wiederholungs-/Rollbackpaket gegen strukturellen Spiegel und synthetische Identitäten. Kein Originalrestore, keine produktive Autorisierung. |
| **Originaldefinitionen UNVERIFIED** | Aktuelle vollständige Spalten/Constraints/FKs, konkrete Policytexte, Spalten-ACL, Ursprung geerbter/PUBLIC-Grants, Owner/Default-Grants, Funktionen/Overloads, Trigger/Views/Indizes/Sequenzen/Abhängigkeiten sowie API/Auth/Realtime-Konfiguration. Die neuen manuellen Ergebnisse schließen diese Lücken nicht. |
| **Vollständige Anwendungskompatibilität nicht abgenommen** | Breiter Boot/Login/SELECT*, Planner-Mehrtabellenprozesse, Privatsichten, allgemeine Logs/Offlinecaches, Werkzeug-/Scan-/Journal-/Inventurtransaktionen und Legacy-/Neu-BA-Mischbetrieb sind nicht durch dieses Reviewpaket vollständig sicher integriert. |

Die vom Auftraggeber gelieferten Originalbefunde gelten ausdrücklich als **bestätigte Sicherheitsbefunde**, nicht länger nur als historische Vermutung. Herkunft und Reichweite: `security-step6-original-evidence.json`. Es wird weder eine neue eigene Liveprüfung noch ein Screenshot-/Exportnachweis behauptet. Keine produktiven Datenabfragen, keine Connector-Wiederholung/Passwortänderung oder Umgehung der Authentifizierung.

Die Prüfung von `has_table_privilege` bestätigt **effektive Tabellenrechte**, nicht deren direkten ACL-Ursprung. RLS ist deaktiviert, sodass dessen Zeilenfilter auf diesen Tabellen nicht schützen. **Extern erreichbarer Data-API-Zugriff wurde nicht unabhängig getestet**: Schema-USAGE, API-Exposure, Netzwerk und weitere Konfiguration sind separat nachzuweisen. Ein erfolgreicher SQL-Editorzugriff als postgres ist kein positiver Anwendungstest mit anon/authenticated und keine RLS-Abnahme; Owner/Superuser können RLS umgehen. Keine Geschäfts-/Personal-/Werkzeug-/Fertigungszeilen sind aus dem Original gelesen oder exportiert worden.

## 2. Abdeckung der zehn Tabellen und vorhandenen 152 Zugriffe

`security-step6-coverage.json` enthält jede der zehn Tabellen, den bestätigten Originalbefund, #28/#29-Abdeckung, neue Zielrechte, Originaldefinitionsstatus, Kompatibilitätslücken und **alle 59 betroffenen Aufrufe** aus dem unveränderten 152-Aufruf-Inventar. Stabile A-Identitäten, Funktion, Operation, Baselinezeile und Payloadausdruck bleiben rückführbar. Die anderen 93 Aufrufe und die RPC-Familien bleiben in der #30-Matrix erhalten; das Inventar wird nicht neu erfunden.

| Tabelle | #28/#29-Kandidat | Schritt 6 / noch notwendiger Vertrag |
|---|---|---|
| planner_assignments | RLS, keine Anonrechte, befristete Bereichs-/Zeilenbindung, Raw-Schreiben gesperrt; serverseitiger versionierter assignment_command, getrennte publish/read RPCs und isolierter #29-Adapter | Einfache erlaubte Zuordnung/Veröffentlichung getestet. Alte direkte Upsert-/Delete-/Namensanzeige nicht kompatibel; Mehrtabellenaktionen nicht teilweise umstellen |
| planner_absences | RLS/Zeilenbindung + explizite befristete Lesegrants; Raw-DML entzogen | Abwesenheitsmeldung, Zuständigkeit, Felder und Bestätigungs-/Schreib-RPC fehlen, D2/D3/D5 |
| planner_vacations | Gleicher Schutz; keine allgemeine Mitarbeiter-/Admin-Privatsicht | Beantragen/Bewilligen/Ändern/Storno und Urlaubssicht offen; keine erfundene Admin-only-Ersatzregel |
| planner_sick_leaves | Gleicher Schutz; ungebundene Zeilen quarantiniert | Krankmeldung, medizinische Details/Planungsstatus und Leser getrennt entscheiden. Synthetischer Privatlesegrant ist keine fachliche Freigabe |
| planner_absence_replacements | RLS/Bindung, Raw-DML entzogen | Ersatz + Ursprungsabwesenheit + Zuordnung/Storno atomar in einem bestätigten Fachkommando erforderlich |
| planner_swaps | RLS/Bindung, Raw-DML entzogen | Eigener Antrag, Zustimmung/Genehmigung, Fristen/Rücknahme und zwei Parteien serverseitig zuordnen; D2/D5 |
| planner_saturday_requests | RLS/Bindung, Raw-DML entzogen | Eigener Antrag und berechtigte Genehmigung getrennt; vorhandene Samstag-Mehrtabellenaktion nicht durch Einzel-RPC ersetzen |
| planner_shift_cancellations | RLS/Bindung, Planungsbereich kann lesen, Raw-DML entzogen | Storno/Aufhebung/Veröffentlichungsinvalidierung mit Versions-/Audit-/Ersatzfolgen fehlt |
| planner_availability | RLS/Bindung, Raw-DML entzogen | Zulässige eigene Erklärung mit Frist, Bereichsleser und Verwaltungsänderung fachlich bestätigen; kein fremder user_name im Payload als Identität |
| tool_materials | RLS: aktive verifizierte Mitarbeiter/Scanner lesen, verifizierter Admin schreibt; #29 liest id,name,sort_order,is_active. **Authenticated hatte bisher dennoch alle vier Tabellenrechte** hinter Policies | Neu: nur Tabellen-SELECT; INSERT-Spalten name,sort_order,is_active; UPDATE-Spalten name,is_active,updated_at; kein API-DELETE/kein Schreiben von id/created_at. Bestehende Adminpayloads Anlegen/Umbenennen/Deaktivieren positiv geprüft |

Alle zehn verlieren im #28-Kandidaten PUBLIC-/Anon-/Authenticated-Altgrants auf Tabellen **und Spalten**. Alte Policies werden auf dem **bekannten isolierten Spiegel** ersetzt; eine zusätzliche permissive Adminpolicy würde vorhandenes true nicht schließen. Planner behält Tabellen-SELECT mit servergeprüften RLS-Prädikaten, keine direkten INSERT/UPDATE/DELETE-Rechte. Wer nur authenticated ist, erhält dadurch keine Zeilen: unabhängige vertrauenswürdige Identität und erlaubte Bindung/Grants sind zusätzlich nötig.

Zusätzliche abhängige Tabellen **employees, departments, planner_tasks, planner_special_days und planner_special_day_assignments** bleiben unter dem unveränderten #28-Schutz (15 bestehende Tabellen insgesamt). Tasks-/Sondertage-/Freigabeworkflows benötigen eigene D2/D3-Abnahme. Abteilungen/Personalstammdaten schreiben nur unabhängig verifizierte Administratoren. AL/Vertretung erhalten keine Personalrollen-/auth_user_id-/Abteilungsadministration. Der historische employees.role-CHECK ohne department_admin bleibt ein **UNVERIFIED** Originalkompatibilitätsproblem; die bestehende synthetische Fixture führt den bekannten Phase-2-Übergang ausdrücklich nur im Test aus.

## 3. Werkzeug-, Auth- und Fertigungskompatibilität

Die aktuellen Funktionen **addToolMaterial**, **renameToolMaterial**, **deactivateToolMaterial** sind am gepinnten #30-Code geprüft. INSERT verwendet name/sort_order/is_active und RETURNING; Umbenennen schreibt name/updated_at; Deaktivieren schreibt is_active/updated_at. Es existiert dort **kein physischer Material-DELETE-Pfad** und kein sort_order-UPDATE. Darauf begrenzen sich die neuen Spaltengrants; UUID/created_at werden serverseitig erzeugt, Referenzzeilen bleiben bei Deaktivierung erhalten. Eine spätere zusätzliche Sortier-/Löschfunktion erfordert ihren eigenen belegten Vertrag. Kein alter Test und keine betriebliche Buchungsfunktion wird dafür entfernt.

Die RLS-Prüfung bleibt für alle authentifizierten Rollen wirksam: Nichtadmins können Referenzen lesen, aber nicht anlegen/ändern; Scanner erhalten keine zusätzlichen Planungs-/Personalrechte. SELECT auf allen sechs historischen Materialspalten bleibt für bestehendes RETURNING/Legacy-Lesen erhalten; die vorhandene isolierte sichere Lesefunktion projiziert vier Felder. Eine SQL-Spaltenerlaubnis ist keine fachliche Vollmacht; die Adminpolicy prüft die unabhängige Authority.

**Nicht daraus abgeleitet:** Werkzeugbestand, Werkzeugstammdaten, Entnahme/Auffüllen, Umbuchen, QR/Scan, Journal, Undo und Inventur wären vollständig gehärtet oder gemeinsam abgenommen. tools, tool_journal, inventory_sessions/entries bleiben unverändert. Originalgrants, Scannerbereiche und transaktionales Bestand+Journal/Undo sind D8. Geometrische Werkzeugsuche ist nicht Bestandteil des Auftrags.

Die bestehende unabhängig bestätigte BA-/OP-Verantwortung, abteilungsübergreifende Maschinenzuordnung, Rüst-/Mengen-/Nachfertigungs-/Abschlusslogik und AFTER-STATEMENT-Forecastkorrektur aus #26 bleiben unverändert. Bestehende Fach-RPCs werden nicht mit neuen Admin-only-Wrappern ersetzt. D7-Mischbetrieb und aktuelle Original-Owner/Definer-/EXECUTE-Abhängigkeiten bleiben offen. Keine neue View oder SECURITY-DEFINER-Funktion als Umgehung fehlender Rechte.

Login/Boot und Rollenanzeige der normalen App bleiben unverändert. `js/secure-access.js` wird weiterhin nicht von der Produkt-HTML geladen. Nur der isolierte #29-Adapter verwendet getUser + aktuellen serverseitigen Actor, sichere RPCs und Memorycache-Invalidierung. Normale Alt-SELECT*, Planner-/Abwesenheits-/Task-Persistenz und vertrauliche Logs werden **nicht** als behoben ausgegeben. Eine Aktivierung des isolierten Pakets bei heutiger Alt-App könnte berechtigte Abläufe sperren und ist verboten, bevor sichere Verträge/Adapter und positive Abnahme vorliegen.

## 4. Inhalt des Review-Migrationspakets

**Review-/Testartefakt, keine registrierte Produktivmigration:**

- `supabase/review/security-step6-isolated.sql`: ausschließlich additives Material-ACL-Delta mit Isolationsguard, Voraussetzung security_actor_context, Lock-/Statementlimit und Transaktionsrahmen. Kein Backfill, keine neue Fachberechtigung, kein DROP/Datenumbau.
- `tests/db-integration/security-step6-package.json`: feste SHA-256 der unveränderten #28-/#29-SQL und neuen Step6-SQL, Ausgangscommit, Status NOT_PRODUCTION_READY.
- `tests/db-integration/security-step6-review.js`: netzwerkfreier Orchestrator, der ausschließlich ein vom Aufrufer bereitgestelltes freigegebenes isoliertes DB-Objekt verarbeitet. Keine URL/Connection-/Credentialbeschaffung und keine CLI-Aktivierung.

### Geprüfte Reihenfolge und Fehlerverhalten

1. Rohdateien gegen exakte Paketchecksummen prüfen. Nur jeweils äußeres BEGIN/COMMIT wird nach Prüfung der genauen Begrenzungen entfernt; alle SQL-Guards und Rechteassertionen bleiben erhalten. Die drei Schritte bilden **eine gemeinsame Transaktion**.
2. Vor Mutation: Testdatenbankname/PGlite-Nachweis und langer Test-Sentinel, begrenzte Wartezeiten, Transaktions-Advisory-Lock sowie explizite ACCESS-EXCLUSIVE-Locks auf allen 15 betroffenen bestehenden Tabellen. Keine DDL zwischen ungesperrten Teilcommits.
3. Frischer isolierter Ausgangsstand erforderlich. Ein unbekanntes schon vorhandenes humbel_security-Schema oder Teilpaket wird zurückgewiesen, nicht überschrieben. Originalbaseline und korrigierter Kandidat bleiben getrennte Testzustände.
4. Unveränderte #28-Vertrauenswurzel/Helfer/Policies/RPCs installieren, unveränderten #29-Serverkontext ergänzen, Step6-Materialrechte reduzieren. Die Authority startet leer, keine Aufnahme heutiger Rollen oder Namensschlüssel.
5. RLS und effektive Tabellen-/Spaltenrechte **für jedes einzelne Recht** einschließlich TRUNCATE, REFERENCES, TRIGGER, MAINTAIN und Grant Options prüfen; nicht nur Komma-OR-Abfrage. Anon ohne SELECT/INSERT/UPDATE/DELETE und ohne Spalten-SELECT/INSERT/UPDATE/REFERENCES; Authenticated Planner nur SELECT, Material nur dokumentierte Spaltenwrites. Geerbte Restrechte führen zum Fehler und vollständigen Rollback.
6. Privates RLS-geschütztes, für API-Rollen unzugängliches Review-Installationsjournal mit Paket-/Kataloghash schreiben, erst dann Commit. Das Journal ist ein zusätzliches **vorgeschlagenes Testobjekt**, kein behauptetes Originalobjekt.
7. Wiederholung derselben Paketversion ist nur nach Postconditions und identischem Katalogfingerprint ein geprüfter **isolierter No-op**. Andere Checksum, unerwartete Journalzeilen oder Katalogdrift führen zum Abbruch, nicht zur blinden Reparatur. Fingerprint umfasst Tabellen/Spalten/Defaults/ACL/RLS/Policies, relevante Funktionsdefinitionen/Owner/EXECUTE, Trigger, Constraints, Rollen/Mitgliedschaften, Schema-/Default-ACL, Views, Indizes und Sequenzdefinitionen des isolierten Zustands. Er vergleicht denselben Testkatalog; er ist kein versions-/OID-unabhängiger Originalrestorebeweis und kein Datenbackup.

Scheitert Step4 etwa an einer vorhandenen kollidierenden RPC-Signatur, werden auch die bereits ausgeführten Step3-Änderungen vollständig zurückgerollt. Synthetische Zeilen in **allen zehn Tabellen** bleiben bei Installation, No-op und erzwungenem Fehler unverändert. Bei einem fehlgeschlagenen erstmaligen isolierten Installationsversuch existiert naturgemäß wieder die unsichere **Testbaseline**; das ist kein zulässiger Produktivrückfall. Im laufenden Originalsystem wird dieses Paket nie ausgeführt.

### Grenzen für Sessions und Parallelität

Tabellenlocks und Transaktionsgrenzen verhindern ein neu erzeugtes Berechtigungsfenster durch Teilcommits. PGlite beweist aber keine parallelen PostgreSQL-Verbindungen, laufenden API-Requests oder PostgREST-/Realtime-Schemacaches. Bereits vor einer späteren Migration laufende Transaktionen/alte Browser/Worker und vor Widerruf gelieferte Daten benötigen sicheren Schreibstopp, Wartungs-/Sessionplan, frische serverseitige Autorisierung und native Abnahme. Bei Locktimeout Abbruch; keine Erhöhung bis eine wartende Produktion „irgendwann durchgeht“. Bereits vorhandene Daten werden nicht nach Namen automatisch gebunden; unklare Altzeilen bleiben quarantiniert, bis D5 unabhängig gelöst ist.

## 5. Isolierte Tests und unveränderte frühere Nachweise

Neue `tests/security/step6-migration.test.js`:

- **Baseline separat:** zehn RLS-deaktivierte Tabellen und alle 20 bestätigten effektiven Tabellenrechtskombinationen im synthetischen Spiegel reproduzieren. Der Generator nimmt darüber hinaus breite Rechte auf andere Tabellen an; **diese zusätzlichen Annahmen werden durch Nachweis C nicht zu Originalfakten**.
- **Korrigierter Zustand:** alle zehn RLS-/Tabellen-/Spaltenpostconditions, Entfernung zusätzlicher permissiver Testpolicy/Spaltengrants, kein anonymer Direktzugriff, Privatschema gesperrt.
- Alle zehn synthetischen Tabellenbestände nach Installation und geprüfter Wiederholung unverändert; ungebundene Urlaubs-/Krankheits-/Tausch-/Ersatz-/Antragsdaten auch für Admin unsichtbar und raw nicht schreibbar. Das ist **nur Quarantäne-/Datenerhaltbeweis**, kein positiver vollständiger Fachworkflowtest.
- Tatsächliche Materialpayloads für Admin erlaubt, Mitarbeiter/AL/Scanner lesen; unberechtigtes Anlegen/Ändern sowie geschützte Materialspalten/physisches Löschen abgewiesen.
- AL eigene/fremde Abteilung, Mitarbeiter/Scanner/verwaiste Identität, explizite gültige/widerrufene/abgelaufene Vertretung, getrennte Veröffentlichung/Minimalprojektion und stale Version nach gemeinsamer Installation.
- Erzwungene geerbte UPDATE-/TRUNCATE-Rechte und kollidierende RPC-Definition: vollständiger Rollback der drei Schritte, alter Katalog/Zeilen unverändert. Unbekannter Vorzustand, Policy-/Spaltengrant-/Viewbodydrift und verändertes Journal werden abgewiesen.
- Neue rein lesende Administrator-Katalogvorlage wird nur gegen die synthetische DB ausgeführt und mit 31/308/127 und vollständigen lokalen Tabellen-/Spaltenrechtsausgaben geprüft; kein Produktionsabruf.
- Neue native Fortsetzung und unbekannter Runner werden auf Fail-Closed/Orchestrierung geprüft.

Alle früheren 32 Sicherheitsprüfungen bleiben bestehen, einschließlich eigenem Kontakt-Selbstservice nur mit synthetischem freigegebenem Feld, Rollen-/Auth-/Abteilungsmanipulation, private Lesegrants/Ablauf, Session-/Offline-/Antwort-Races und BA-/OP-/Tool-Fachtests. Kein geschützter Hash oder altes Testfixture wird geändert.

Separate neue native Pipeline `security-step6-isolated.yml` behält sämtliche ursprünglichen Step3-/Step4-/Baseline-/BA-OP-Prüfungen und führt erst danach vier additive echte JWT/REST-Material-/Anon-/Datenerhaltgruppen aus. Exakter reversibler Step6-Runneradapter auf gepinntem Step4-Runner, Quellenmanifest vor DB-Verbindung, unveränderte Target-Guards und always()-Cleanup. Vorherige Workflows bleiben bytegleich. Die vollständige atomische Neuinstallation/No-op-/Fehlersequenz ist hier zunächst eingebettet geprüft; die native Fortsetzung nach früheren Commits beweist sie **nicht automatisch**. Ein vollständiger nativer Installations-/Rollbacktest auf Originalrestore bleibt erforderlich.

Der erste neue Guard enthielt eine überzählige schließende Klammer. Die tatsächliche PGlite-Ausführung hat das aufgedeckt; beide neuen Guardstellen wurden korrigiert, die unveränderten #28/#29-Inputs blieben erhalten. Danach gezielte Tests erneut ausgeführt. Nach Erweiterung des Fingerprints um Sequenzdefinitionen zeigte der Gesamtlauf außerdem eine BigInt-Serialisierung im neuen Code; Sequenz-Metadaten werden jetzt verlustfrei als SQL-Text projiziert. Keine Sicherheitsassertion zur Fehlerumgehung angepasst.

## 6. Rein lesender Originalnachweis für Administrator

`supabase/review/security-step6-metadata-readonly.sql` ist eine **nicht produktiv ausgeführte Administratorvorlage**: ein REPEATABLE-READ-/READ-ONLY-Katalogsnapshot, keine Tabellenzeilen/Benutzer/Sessions, keine Sequenzstände oder Zeilenschätzungen. Liefert Schemas, Relationen/Spalten/Constraints/FKs, Funktionen/RPC-Signaturen/Bodies/Owner/ACL/Config, Views/Materialized Views, Trigger, Policies, Indexdefinitionen/Validität, Sequenzdefinitionen, Default ACL, Rollenflags/Mitgliedschaften, effektive Tabellen-/Spaltenrechte für API-Rollen und Abhängigkeiten in beide Richtungen für public/humbel_security.

**Nicht gleichbedeutend mit pg_dump:** vollständige Typ-/Domain-/Extensionobjekte, Eventtrigger, Foreign-Server und Plattformkonfiguration benötigen den im Schritt-5-Runbook exakt beschriebenen ergänzenden Schema-only-/Owner-/ACL-Export. Keine pg_authid-/auth.users-/auth.sessions-/vault-Secret-Abfragen. Vor Übergabe prüft der Administrator Funktionskörper/Defaults auf eingebettete Secrets; bereinigte Arbeitskopie und geschütztes Original/Hashmanifest getrennt halten. Ohne Exportlieferung bleiben alle Definitionen hier **UNVERIFIED**. Herkunft, Operator, Zielprojekt/Branch, konsistenter Zeitpunkt, Client-/Serverversion, Umfang/Ausnahmen/Exitcodes und SHA-256 sind beizulegen.

Die neue Vorlage wird weder mit execute_sql noch per SQL Editor durch den Agenten ausgeführt. Der dokumentierte Connectorfehler 28P01 wird nicht durch Passwortwechsel/Umgehung „repariert“. Die manuell bestätigte Dashboardverbindung beweist die Funktionsfähigkeit dieses Zugangs, nicht die Connectorcredentials.

## 7. Produktiveinsatz: eindeutige Entscheidung und nächste Schritte

**Nein. Die entwickelte Sicherheitsmigration ist anhand der tatsächlich verifizierten Originalinformationen noch nicht vollständig für einen kontrollierten Produktiveinsatz geeignet.** Der bestätigte P0-Ausgangsbefund ist belastbarer geworden. Vollständige strukturelle und fachliche Kompatibilität, Owner-/ACL-/Bypasspfade, gesamte sichere Anwendung und Betriebsabnahme bleiben offen.

Verbindliche vollständige Gate-/Verantwortungs-/Handlungsliste: **security-step6-release-checklist.md**. Alle nicht nachgewiesenen Gates bleiben OFFEN/BLOCKED_NOT_TESTED. Insbesondere:

1. **Technischer Originalnachweis:** DBA liefert aktuelle Definitionen/Rechteherkunft/Default-Grants/Owner/Abhängigkeiten und API/Auth-Konfiguration nach Vorlage + Schema-only-Runbook; unabhängig im genehmigten isolierten Stack restaurieren. Keine Repository-Substitute als Original ausgeben.
2. **D1–D6 Fach/Identity:** Leser/Felder/Publish, Urlaub/Krankheit/Tasks/Tausch/Ersatz/Anträge/Fristen, erlaubte Kontakte, unabhängige Personen-/Bereichs-/Altzeilenbindungen und Issuer-/Widerrufsprozess verbindlich bestimmen. Danach passende sichere atomare Schreib-/Leseadapter umsetzen, nicht ungeklärte Regeln erfinden.
3. **D7/D8 Betrieb/Fertigung/Werkzeuge:** Alt-/Neu-BA-Wahrheitsquelle und Scanner-/Buchungs-/Journal-/Undo-/Inventurmatrix bestätigen; volle positive Fachkompatibilität erhalten und testen.
4. **Gesamte App:** sichere Boot-/Projektion-/Mehrtabellen-/Privacy-/Cache-/Log-/Session-/Realtimeintegration vor Aktivierung fertigstellen. Heute wäre die direkte Alt-App nicht vollständig kompatibel.
5. **Abnahme:** aktuelle vollständige Browser-/Visual-/native JWT/REST-/Paralleltests sowie echte Originalrestore-, Neuinstallations-/Wiederholungs-/Fehler-/Recoverytests mit Daten-/Audit-/Historienerhalt bestehen lassen. Baselinebefund und korrigierten Zustand weiter trennen.
6. **Betriebliche Freigabe:** Betreiber bestimmt RPO/RTO, Wartungsfenster/Schreibruhe, Verantwortliche/Stop-Kriterien, verschlüsselte vollständige Backups/Schlüsselaufbewahrung/PITR und unabhängige isolierte Restoreprobe. Rückfall nie durch Reaktivierung unsicherer Altpolicies. #25-Pages-/Deployment-/Branchschutz-Nachweise schließen.
7. Erst anschließend eigenes aus dem **tatsächlichen** Originaldiff abgenommenes Produktivpaket und separate ausdrückliche Entscheidungen über Migration, Merge, Deployment und Aktivierung. Guards entfernen oder Reviewdateien kopieren genügt nicht.

Keine Backup-/Restoreprobe auf Produktion, keine Geschäftsdaten aus dem Original, keine produktive SQL-/RLS-/Auth-/Rollen-/Featureänderung. Kein Merge und kein Deployment. Nach Erstellung des neuen Draft-PR endet der Auftrag ohne Actions-Abfrage/Warteschleife/Dispatch.

Technische Referenzen (12.11.2026): https://www.postgresql.org/docs/17/ddl-priv.html ; https://www.postgresql.org/docs/17/ddl-rowsecurity.html ; https://www.postgresql.org/docs/17/transaction-iso.html ; https://supabase.com/docs/guides/database/postgres/row-level-security . Keine Paket-/CLI-/Engineversion angehoben.

## 8. Tatsächliche lokale Ergebnisse des Schritt-6-Standes

Maschinenlesbar: `security-step6-validation.json` mit tatsächlichen Exitcodes/Testzahlen/Loghashes und offengelegtem vorherigem neuen Serialisierungsfehler. Kein historischer CI-Erfolg wird auf diesen Stand übertragen.

| Prüfung | Ergebnis |
|---|---|
| npm ci --ignore-scripts, check:syntax einschließlich beider Pflicht-Nodechecks; Syntax der vier neuen JavaScriptdateien | PASS |
| test:ci | PASS **2/2** |
| test:qa | PASS **20/20** |
| qa:check | PASS; 251 Legacy-Identitäten, Geschäftslogik und 33 PNG-Referenzen erhalten |
| test:integration | PASS **5/5** |
| test:security | PASS **45/45**, keine Skips: bisherige 32 unverändert + 13 neue Step6-Prüfungen |
| Neun bestehende Fachmodulsuiten | PASS **168/168**, keine Skips |
| qa:agent-server | PASS |
| Exakte Quellenmontage, Clock-/Provenienz-/Step3-/Step4-/Step6-Adapter, Runner-/Seedsyntax und Manifest | PASS: **54** verwaltete Eingaben, kein nativer DB-Lauf |
| Ursprüngliche native Target-Guards im gepinnten Harness | PASS **2/2**, ausschließlich Zielschutz, keine native DB-Ausführung |
| test:e2e | **BLOCKED_NOT_TESTED**, Exit 1: alle 773 Fälle aufgerufen, 773 fehlende Executables vor Seitenaufruf, 0 bestanden. Fortschrittsreport complete=true, runnerStatus=failed, 773 Versuche |
| test:agents | **BLOCKED_NOT_TESTED**, Exit 1: Seed vor Seitenaufruf am fehlenden Browser blockiert |
| test:visual | **BLOCKED_NOT_TESTED**, Exit 2: vorgeschriebener Playwright-1.62.1-noble-Container fehlt; Referenzen unverändert |
| Native Step6-/Sicherheits-/BA-OP-/Parallel-/JWT-/REST-Abnahme | **BLOCKED_NOT_TESTED**: kein Docker/Podman/psql oder zugelassener Stack; DB-Preflight Exit 2 |
| Originalschema-/Upgrade-/Recoveryprobe und vollständige gehärtete Anwendung | **BLOCKED_NOT_TESTED**: Originalexport, genehmigtes Restoreziel und vollständige sichere Fachintegration fehlen |
| git diff --check | PASS |

Die READ-ONLY-Katalogvorlage wurde nur in der synthetischen eingebetteten Datenbank ausgeführt. Bestätigte Originalbefunde stammen ausschließlich aus dem manuellen Auftraggebernachweis. Kein produktiver DB-Zugriff und keine Produktivfreigabe.
