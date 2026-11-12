# Schritt 7 – tatsächlicher Originalkatalog und kritischer Mitarbeiterfix

**NEIN: keine sichere Produktivmigration freigegeben.** Neuer gestapelter Draft auf PR #31, verbindlicher Parent `1238bf0c9467f1d6bf5e83dd8e0aa8520ef4b7c7`, Branch `security/original-audit-step7`, APP_VERSION **0.16.4**, VERSION_LOG **12.11.2026 13:44 Europe/Berlin**. Refs #17, #25–#31, #15/#16. Keine Änderung bestehender Geschäftsprozesse/Startseite oder produktiver Schalter. Haupt-App ausschließlich Version/Log. Vorherige Berichte bleiben als historische Nachweise unverändert.

## 1. Automatisch durchgeführte Originalprüfung

Der vorhandene Supabase-SQL-Connector funktioniert in diesem Auftrag. Die erste lesende Verbindung meldet **postgres / postgres / PostgreSQL 17.11 auf aarch64**, 31 öffentliche Tabellen. Zielref `itntnpzqzyjfdnjdshvm` gemäß Auftrag, Projekt Schichtplan/Produktionsbranch main. Keine Passwortänderung/Authentifizierungsumgehung und keine erneute manuelle SQL-Arbeit für den Auftraggeber. **Browserfallback war nicht erforderlich**, keine Browsersession oder Anmeldung geprüft.

Ausgeführt wurden ausschließlich:

1. READ-ONLY-Verbindungs-/Kataloganzahlabfrage.
2. Die **unveränderte** vorbereitete `security-step6-metadata-readonly.sql` in REPEATABLE READ READ ONLY: vollständige JSON-Aggregation als **eine Ergebniszeile**. Snapshot **12.11.2026 12:36:34.766069 UTC / 13:36:34 Europe/Berlin**.
3. Unabhängige READ-ONLY-Katalogzählung für öffentliche Tabellen/Spalten/Policies/Constraints/Benutzertrigger/Funktionen/Indizes; alle Abschnittszahlen stimmen überein. Das ist ein zusätzlicher Vollständigkeitscheck, kein Beweis, dass zwischen getrennten Transaktionen nie DDL stattfand.
4. Ergänzende READ-ONLY-Abfrage `security-step7-extra-metadata-readonly.sql`: einzelne effektive Tabellenrechte einschließlich Grant Options, Schema-USAGE/CREATE, öffentliche Funktions-EXECUTE-Rechte und ausschließlich allowlistbasierte technische Rolleneinstellungen. Keine geheimen JWT-/Authkonfigurationen ausgewählt.
5. Schlüsselloser HTTPS-**HEAD** auf `/rest/v1/employees?select=id&limit=0`, Antwort **401**, kein Antwortbody gelesen. Kein API-Key/Bearer/Benutzerkonto verwendet.

Kein SELECT aus Geschäfts-/Mitarbeiter-/Auth-Benutzer-/Session-/Werkzeug-/Fertigungsrelationen, kein Ausführen von Geschäfts-RPCs oder Wartungsfunktionen. Katalogfunktionen wie pg_get_functiondef deparsen Definitionen, sie führen deren Körper nicht aus. Kein INSERT/UPDATE/DELETE/TRUNCATE/DDL/GRANT/REVOKE im Original. Die Kontrollen liefen als postgres, **nicht als tatsächliche Appidentität**; daraus folgt keine erfolgreiche RLS-/API-Rollenabnahme.

### Gelieferter Originalumfang und Resultatlimit

| Metadatenabschnitt | Erfasster Umfang |
|---|---|
| Öffentliche Relationen | **31 Tabellen**, 21 RLS an / zehn aus; keine öffentliche View/Materialized View |
| Öffentliche Spalten/Constraints | **308 Spalten / 134 Constraints**, einschließlich CHECK/PK/Unique/FK/Validierung |
| Öffentliche Policies | **127**, alle beobachteten Policies PERMISSIVE und an authenticated adressiert |
| Öffentliche Funktionen/Indizes/Benutzertrigger | **6 / 76 / 11** |
| Gesamter erfasster technischer Katalog | 85 Relationen, 843 Spalten, 292 Constraints, 319 Trigger einschließlich interner, 237 Indizes, 114 Funktionen, 68 Views (davon 65 information_schema), zwei Sequenzdefinitionen, fünf Extensions |
| Rechte-/Rollenumfang | neun Schemas, 30 Rollen, 21 Mitgliedschaften, 24 Default-ACL-Einträge; 62 Tabellen-/616 Spaltenrechtsdatensätze der ersten Vorlage |
| Erweiterter effektiver Rechtetest | **496 Einzelrechte** = 31 Tabellen × zwei API-Rollen × acht Privilegien, plus 12 Funktionsrechte, Schema-USAGE/CREATE und technische Einstellungen |
| Abhängigkeiten | **955** gerichtete technische Abhängigkeitseinträge gemäß Vorlagenfilter, einschließlich Referenzen zu/aus öffentlichen Objekten |

Die JSON-Aggregation vermeidet ein SQL-Editor-Zeilenlimit; es wurde kein SQL Editor verwendet. Ein Ergebnisrow enthält sämtliche Arrays; JSON-Parsing, alle erwarteten öffentlichen Anzahlen und der zusätzliche Zählcheck waren erfolgreich. Keine abgeschnittene Antwort oder fehlende Zeile wird als vollständig ausgegeben.

Artefakte:

- `security-step7-original-catalog.json.gz`: gzip-komprimierter **JSON-Metadatenexport**, mit deterministischer Kompression und Roh-/Komprimathashes im Paketmanifest. Entpackt 1.248.803 Bytes, komprimiert 82.477 Bytes. Keine Anwendungsdaten und keine Sequenzstände/Zeilenschätzungen.
- `security-step7-effective-rights.json`: ergänzende Einzelrechte-/Konfigurationsbeobachtung.
- `security-step7-evidence.json`: Herkunft, Ziel, Abfragenhash, Snapshot, Abschnittsprüfung, HTTP-Beweisgrenze und offene Nachweise.
- `security-step7-comparison.json`: konkreter historischer Spalten-/Constraint-/Funktionsvergleich.

**Secretprüfung vor Veröffentlichung:** Alle sechs öffentlichen Funktionskörper separat angesehen, ohne eingebettete Zugangswerte. Statische Prüfung der übrigen Definitionen auf credentialartige Literale; eine storage.list_objects_with_delimiter-Definition wurde konservativ vollständig als `REDACTED_PENDING_ADMIN_REVIEW` zurückgehalten. Das ist **kein Nachweis eines gefundenen Secrets**, sondern eine vorsichtige Begrenzung der Weitergabe. Die anderen Metadaten wurden als technische Definitionen erfasst; es gab keine Passwort-/API-Key-/Tokenabfragen. Die bereinigte Arbeitskopie ist deshalb kein byteidentischer vollständiger Originaldump.

**Weiter offen:** autoritativer pg_dump --schema-only einschließlich vollständiger Typ-/Domain-/Extension-/Eventtrigger-/Foreign-Server-/Owner-/ACL-Recovery, sicherer Exportweg und unabhängiger Restore. Der Connector bietet keinen pg_dump-Export; es werden dafür keine Passwörter ausgelesen. Der bisherige detaillierte einmalige administrative Exportplan aus Schritt 5 bleibt der konkrete Lieferweg, nicht eine neue Serie manueller Einzelabfragen.

## 2. Direkt bestätigte Originalbefunde – nach Kritikalität

### P0: Mitarbeiterrollen und Identitätsfelder

Auf employees haben anon und authenticated alle beobachteten Tabellenrechte; die effektiven UPDATE-Rechte auf **role, auth_user_id, department_id, is_active und id** ergeben sich aus dem bestätigten Tabellen-UPDATE. Aktuell sind **keine separaten öffentlichen Spalten-ACLs** gesetzt; effektive Spaltenrechte dürfen daher nicht als ausdrücklich vorhandener Einzelspaltengrant missverstanden werden.

Authenticated besitzt permissive employees_insert/UPDATE/DELETE/SELECT-Policies mit **true**. Die zusätzlichen Adminpolicies werden permissiv **OR**-verknüpft und schließen diese Öffnung nicht. Damit können nicht berechtigte Authenticated-Identitäten serverseitige Autoritätsfelder beeinflussen und fremde Personalprofile lesen. Es wurde **kein produktiver Schreibversuch** vorgenommen. Der Zusammenhang wird nur in einer synthetischen isolierten Baseline reproduziert: eigener role-UPDATE zu admin macht anschließend den Originalhelper is_admin wahr.

Die anon-CRUD-Grants bedeuten bei employees **nicht automatisch eine anon-RLS-Freigabe**: Auf dieser RLS-aktivierten Tabelle richten sich alle Policies an authenticated; ohne passende Rollenvererbung gibt es keinen anon-Zeilenpfad. Tabellenrecht, Policyfreigabe und API-Erreichbarkeit werden getrennt gehalten.

### P0: Wartungs- und DDL-bezogene Rechte auf allen 31 Tabellen

Die neue Einzelprüfung bestätigt für **anon und authenticated auf jeder der 31 Tabellen** SELECT, INSERT, UPDATE, DELETE, **TRUNCATE, REFERENCES, TRIGGER, MAINTAIN**. Alle **496** abgefragten Rechte sind true; Grant Options false. Öffentliches Schema: beide Rollen USAGE=true, CREATE=false.

**TRUNCATE wird nicht von RLS geschützt.** Die breiten SQL-Rechte stellen deshalb auch bei aktivierter RLS ein eigenständiges Risiko dar. Das beweist keine über Standard-PostgREST verfügbare TRUNCATE-Operation; API- und SQL-Angriffsoberflächen bleiben getrennt. Keine produktive Lösch-/Wartungsaktion wurde ausgeführt.

Der unveränderte PR-31-Kandidat entzieht diese Rechte nur in seinen **15 geschützten Tabellen**. Auf den übrigen **16** bleiben sie in einer rechtegetreu erweiterten isolierten Fixture bestehen. Dieser konkrete zusätzliche Kandidatenmangel wird nun negativ und korrigiert getestet.

Für postgres und supabase_admin sind außerdem breite **öffentliche Default-Tabellenrechte** direkt erfasst. Diese könnten die Lücke auf zukünftigen Tabellen neu einführen. Default-CRUD-/Funktions-/Sequenzrechte sind ebenfalls erfasst und bleiben über das gezielte neue Tabellenrechtsdelta hinaus ein gesonderter Abnahmepunkt.

### P0: Zehn Tabellen ohne RLS

Die zehn manuell bestätigten Tabellen sind nun auch direkt erfasst: planner_absence_replacements, planner_absences, planner_assignments, planner_availability, planner_saturday_requests, planner_shift_cancellations, planner_sick_leaves, planner_swaps, planner_vacations und tool_materials. RLS aus, API-Rollen mit vollen SQL-CRUD-Rechten und Schema-USAGE. Das lässt Schicht-/Abwesenheits-/Materialdaten ohne RLS-Schutz. **Externer authentifizierter oder anon-JWT-API-Datenzugriff wurde nicht getestet.**

### P1/P0 im Zusammenhang: weitere permissive Policies

tools hat unter anderem authenticated UPDATE USING/WITH CHECK true; tool_journal erlaubt authenticated breit SELECT/INSERT/UPDATE/DELETE, parallel zu scheinbar begrenzenden Adminpolicies. planner_tasks, planner_special_days und planner_special_day_assignments enthalten breite Authenticated-Regeln. Weitere Fertigungs-/Verwaltungstabellen werden vollständig in der 31-Zeilen-Matrix ausgewertet.

Eine erforderliche Werkzeugbuchung ist keine automatische Freigabe beliebiger Stammdatenänderungen oder des Löschens/Manipulierens fremder Journalhistorie. Ebenso ist ein freigegebener Schichtplan keine Freigabe aller Urlaubs-/Krankheits-/Taskfelder. Die genaue zulässige Feld-/Workflowmatrix bleibt nötig; pauschales Admin-only ist kein geeigneter Ersatz. Auch scheinbar bereichsbezogene Regeln hängen teilweise von der manipulierbaren Mitarbeiter-Autorität ab.

### P1: Rollen-Constraint und Definer-Vertrauenswurzel

Der validierte employees_role_check enthält **nur admin/employee/tool_scanner**, obwohl Policies und can_access_production_department zusätzlich department_admin prüfen. Bei normalen durch den CHECK zulässigen Zeilen ist dieser AL-Zweig nicht erreichbar. Ein solcher Rollenwert scheitert im isolierten unveränderten Constraint mit 23514. Daraus wird keine Aussage über konkrete Mitarbeiterpersonen abgeleitet; es wurden keine Personenzeilen gelesen.

Die fünf ursprünglichen Definer-Helfer sind Eigentum postgres, search_path=public; alle sechs öffentlichen Funktionen sind PUBLIC/anon/authenticated ausführbar. Die fünf lesen aktuell qualifiziert public.employees und auth.uid; das primär bewiesene Risiko ist die **mutable Mitarbeiter-Vertrauenswurzel**. Öffentliches CREATE ist für API-Rollen aktuell nicht erlaubt. Ein praktisch ausnutzbarer Namensersetzungspfad wird deshalb **nicht allein aus search_path=public behauptet**. Im Kandidaten werden leere Suchpfade, qualifizierte Referenzen, unabhängige Authority und kontrolliertes EXECUTE separat geprüft.

## 3. Vollständige Risikomatrix und Vergleich zu PR #31

`security-step7-risk-matrix.json` enthält **alle 31 öffentlichen Tabellen**: Original-RLS/FORCE/Owner/ACL, jedes einzelne API-Rollenrecht, sämtliche Tabellenpolicies, syntaktisch eindeutig offene Policykommandos, Abteilungs-/Privacy-/Manipulationsrisiken, API-Nachweisgrenze, PR-31-Abdeckung und verbleibende Schritte. Das ist keine vollständige symbolische SQL-Beweisengine; komplexe Prädikate bleiben als Originaltexte nachvollziehbar.

Die bestehende 152-Aufruf-Matrix wird weiterverwendet. **149 Aufrufe** betreffen tatsächlich vorhandene Originaltabellen; **drei Inventuraufrufe** adressieren inventory_sessions/entries, die im aktuellen Originalschema nicht vorhanden sind. Kein Schluss auf einen aktiven Inventurbetrieb; bestehende Schalter bleiben aus. Umgekehrt gibt es Originaltabellen ohne direkten inventarisierten UI-Aufruf. SQL-/Trigger-/RPC-Abhängigkeiten bleiben deshalb mit im Umfang.

Der historische Spiegel stimmt aktuell in **308 Spalten einschließlich Typ/Default/Nullbarkeit und 134 Constraintdefinitionen** überein, Drift=0. Alle **sechs Funktionsdefinitionen sind bytegleich** zu den gepinnten Originaldefinitionen aus #15. Das stärkt diese Bestandteile; Schemaowner/Rollenkonfiguration, Plattformobjekte und ein unabhängiger vollständiger Restore werden dadurch nicht ersetzt.

Aktuell gibt es im Original **kein humbel_security-Schema und keine security_*-RPCs**; nur die sechs beschriebenen öffentlichen Funktionen. Neue Artikel-/BA-/OP-/Zeit-/Mengen-/Lifecycle-Objekte bleiben Repository-Kandidaten, nicht als installierte Originalobjekte ausgegeben. Die tatsächlichen späteren Migrationen und der Legacy-/Neu-Mischbetrieb müssen gegen den Originalrestore abgenommen werden.

## 4. Isolierter Korrekturkandidat Schritt 7

### Mitarbeiterfix: vorhandener Schutz bestätigt

Die unveränderte #28-SQL beseitigt alle alten Mitarbeiterpolicies und PUBLIC-/Anon-/Spaltenaltrechte, prüft administrative Änderungen über eine unabhängig provisionierte, befristete und mit aktueller Mitarbeiterbindung konsistente Authority. #29 ergänzt den sicheren Actor; #31 die überprüfte Transaktions-/Materialhärtung. Der kritische Mitarbeiterfehler benötigt **keine neu erfundene Personalverwaltung**: Er ist in diesem Schutzmodell geschlossen, jetzt zusätzlich für **alle fünf geschützten Felder**, eigene/fremde Personen, Mitarbeiter/AL/Vertretung/Scanner und unberechtigtes Insert/Delete geprüft.

Administratoren behalten echte notwendige Stammdaten-CRUD; eine Profiländerung erzeugt keine Authority. Eigene Kontakte bleiben ausschließlich über den vorhandenen eigenen RPC mit ausdrücklich konfigurierter Test-Allowlist, ohne fremde Ziel-ID. Welche realen Kontaktfelder produktiv erlaubt werden, bleibt D4; employees hat weiterhin keine ursprünglichen Kontaktspalten. Ungebundene/inaktive/verwaiste/widersprüchliche Identitäten bleiben fail-closed.

### Zusätzliche belegte Korrekturen

`supabase/review/security-step7-isolated.sql` ist **nur ein isoliertes Reviewartefakt**, nicht in migrations registriert oder produktiv aktiviert:

1. Für exakt die **31 nachgewiesenen Tabellen** TRUNCATE/REFERENCES/TRIGGER/MAINTAIN von PUBLIC/anon/authenticated entziehen; zusätzliche REFERENCES-Spaltengrants ebenfalls entfernen. Gewöhnliche CRUD hinter ihren vorhandenen Regeln wird außerhalb des bisherigen geschützten Bereichs nicht pauschal neu erfunden.
2. Dasselbe gezielt auf die bestätigten öffentlichen Default-Tabellengrants der beiden beobachteten Ersteller postgres und supabase_admin anwenden. Keine globale Plattformrechteänderung und keine neue Rolle im Produktivsystem.
3. employees_role_check ausschließlich dann von den beobachteten drei auf die **bereits fachlich vorgesehene Rolle department_admin** ergänzen, wenn seine Definition exakt dem Original entspricht; bekannte identische Phase-2-Zieldefinition wird akzeptiert, jede unbekannte Definition führt zum Abbruch. **Keine Mitarbeiterwerte/Identitätsbindungen ändern oder automatisch enrollen.** Autorisierte Zuweisung ist ein späterer gesonderter Identityprozess.
4. Direktes API-EXECUTE auf set_updated_at entziehen. Bestehende Triggerverwendung und gewöhnliche Adminänderungen werden isoliert positiv geprüft. Die fünf anderen ursprünglichen Helfer sind bereits im unveränderten #28-Kandidaten mit unabhängiger Autorität, leerem Suchpfad und ohne Anon-/PUBLIC-Execute abgesichert.

Die neue SQL ist kein vollständiger Fix sämtlicher CRUD-/Journal-/Privatdatenregeln auf den übrigen 16 Tabellen. Diese bleiben **ausdrücklich offen**. Maintenance-Rechteentzug allein macht diese Tabellen nicht produktionssicher.

### Gemeinsame Transaktion statt halber Umstellung

`security-step7-review.js` verwendet die **unveränderten, gehashten** Step3-/Step4-/Step6-Quellen plus die gehashte neue Step7-Datei. Nur äußere, exakt geprüfte BEGIN/COMMIT werden zusammengeführt. Vor Mutation Isolations-/Sentinelguard, Advisory-Lock, begrenzte Wartezeiten und Tabellenlocks. Alle vier Schritte, Rollen-Constraint, Rechte-/Defaultänderungen und das neue private Reviewjournal werden in **einer Transaktion** geprüft und committed.

Nachprüfungen umfassen die bisherigen zehn RLS-/Spalten-/Materialverträge, alle 31 Tabellen bezüglich der vier neuen Rechte und REFERENCES-Spalten, Mitarbeiter-RLS/Anon-Denial/Entfernung true-Policies, Default-Tabellenrechte und Anon-EXECUTE. Eine bekannte identische Wiederholung muss Postconditions und Katalogfingerprint bestehen, sonst Abbruch; unbekannter Vorzustand wird nicht implizit überschrieben oder heruntergestuft. Kollision oder unbekannter Mitarbeiter-CHECK rollt sämtliche Schritte und Daten zurück. Dies ist kein produktives Migrationsjournal und kein getesteter Originalrestore.

### Was bewusst unverändert bleibt

Schichtplanung, Personal-/Abteilungsverwaltung, Urlaub/Krankheit/Tausch/Samstag/Tasks, Werkzeug-/Journal-/QR-/Inventurabläufe, Fertigung/Maschinen, BA-/OP/Rüstzeit/Mengen sowie Startseite bleiben im Anwendungscode unverändert. Die Produkt-HTML lädt den sicheren Adapter weiterhin nicht. Keine neuen produktiven Flags oder Service-Role-Fallbacks.

Die heutige Alt-App ist weiterhin nicht vollständig mit dem fail-closed Planner-/Privatdatenmodell kompatibel. Anträge/Ersatz/Storno/Tausch/Genehmigungen benötigen sichere atomare Befehle und bestätigte Leser-/Feldregeln, ebenso Journal-/Scanner-/Inventurtransaktionen und Cache-/Log-/Session-/Realtimebehandlung. Diese dürfen nicht durch erneute breite Tabellenrechte oder versteckte Legacy-Fallbacks scheinbar repariert werden.

## 5. Tatsächliche isolierte Prüfung und native Vorbereitung

Die neue Threat-Fixture nutzt den strukturellen Spiegel mit Auth-Shims und synthetischen Personen. Sie übernimmt **belegte effektive öffentliche API-Tabellenrechte**, die ursprüngliche Drei-Rollen-Constraint, öffentliche Helfer-EXECUTE und zwei beobachtete Default-Ersteller. Das ist eine präzisere Bedrohungsfixture, **kein vollständiger originalgetreuer Supabase-Restore**: privilegierte Operator-/Ownerrollen, Plattform/Auth/REST und echte Sessions bleiben anders.

Getrennt geprüft werden Original-Baseline und korrigierter Zustand: role→admin→is_admin nur synthetisch reproduziert; ursprüngliche AL-Constraint scheitert; PR31 lässt weitere Wartungsrechte bestehen. Danach alle 31 neuen Rechte/Defaults gesperrt, Constraint gezielt erweitert ohne Personenänderung/Enrollment, fünf Mitarbeiterfelder gegen vier unberechtigte Rollen auf eigenen/fremden Zeilen geschützt, unberechtigtes Insert/Delete/Truncate, zulässige Admin-CRUD und eigener freigegebener Testkontakt, eigene/fremde AL-Planung/Vertretungsablauf/-widerruf, getrennte Veröffentlichung/Privacy, Temp-/Public-Namensspoofing, No-op und vollständiger Fehlerrollback. Ein repräsentativer bestehender tools.stock-/tool_journal-Payload bleibt ausführbar; das ist **kein vollständiger Browser-/Scanner-/Buchungsatomizitätsnachweis**.

Die exakte Quellen-/Clock-/Provenienz-/Step3-/Step4-/Step6-/Step7-Runner-Montage ist mit **60** Eingaben geprüft, einschließlich Runner-/Seedsyntax und ursprünglichen Target-Guards **2/2**. Das ist ausschließlich Montage-/Integritätsnachweis, kein nativer DB-Lauf.

Ein neuer unabhängiger nativer Workflow setzt **nach sämtlichen bisherigen Baseline-/BA-OP-/Step3-/Step4-/Step6-Prüfungen** fort: bestätigte zusätzliche Privilegien nur im synthetischen Stack reproduzieren, Step7-Delta installieren, alle31 Postconditions, Mitarbeiterdatenerhalt sowie echte REST-Negativ-/Adminfälle. Alter Runner wird nur nach exaktem Hash reversibel ergänzt; Quellenmanifest/Target-Guards/Cleanup bleiben erhalten. Die neuen nativen Tests sind ohne Laufzeit nicht als bestanden auszugeben; additive native Fortsetzung ersetzt nicht die vollständige atomare Neuinstallation/Recoveryprobe.

Der neue synthetische Werkzeugtest verwendete zunächst den ungültigen Halterwert HSK. Die aktuelle Originalconstraint erlaubt HSK 63/HSK 100; der neue Fixturewert wurde entsprechend korrigiert, ohne Erwartung oder Constraint zu schwächen. Keine vorherige Testidentität/Assertion/PNG/Schutzhash wurde geändert.

## 6. Kontrollierte spätere Migrations- und Freigabereihenfolge

1. Erfasste Originalmetadaten/Hashes unabhängig reviewen, zurückgehaltene Plattformdefinition kontrolliert prüfen. Autorisierten vollständigen Schema-only-/Owner-/ACL-/Plattformexport auf sicherem Weg erstellen und in benanntem isolierten Ziel restaurieren, ohne Originaldaten in Git/CI. Jetzt bekannte öffentliche Definitionen nicht mehr pauschal als UNVERIFIED darstellen; nur tatsächlich fehlende Teile offenhalten.
2. Alle 31 CRUD-/Spalten-/Default-/Definer-/API-Zugriffspfade fachlich und technisch schließen. Besonders employees, zehn RLS-aus-Tabellen, tools/tool_journal und private Planner-/Taskdaten; keine zusätzliche permissive Policy neben true. Default-Funktions-/Sequenzprofile, geerbte Rechte und sichere Ownerrollen explizit abnehmen.
3. D1–D6 (Schichtleser/Felder/Publish, private Abläufe/Kontakte, unabhängige Identitäts-/Altzeilenbindungen und Delegations-/Widerrufsverfahren) sowie D7/D8 (BA-Mischbetrieb/Werkzeug-/Scanner-/Journal-/Inventurmatrix) entscheiden. Danach vollständige sichere Appadapter implementieren und positiv/negativ abnehmen. Der neue Rollen-CHECK ersetzt keinen Issuer-/Provisionierungsprozess.
4. Gemeinsamen Migrationskandidaten am Originalrestore mit echten PostgreSQL/GoTrue/PostgREST/Realtime-/Parallel-/Sessiontests, allen Browserprojekten, unveränderten Visuals, Daten-/Audit-/Mengen-/Werkzeugintegrität, Retry/Idempotenz und Fehlerrollback prüfen. Neue Fachmigrationen behalten explizite minimale Grants; kein Vertrauen in alte Defaultvollrechte.
5. Betreiber legt numerische RPO/RTO, Verantwortliche, Wartungs-/Schreibstoppfenster und Stop-Kriterien fest. Separat genehmigte vollständige verschlüsselte Sicherung einschließlich Schema/Rechte/Authority/Audit/Produktionshistorie, Secret-/Storage-Recovery und nachgewiesener PITR-Verfügbarkeit; unabhängige **isolierte** Wiederherstellungsprobe messen. Nichts davon wurde auf Produktion ausgeführt.
Der spätere Migrationsoperator benötigt nachgewiesene Objekt-/DDL-/Grantbefugnisse auf den 31 Tabellen und Befugnis zur Änderung der Default Privileges für postgres und supabase_admin. Ein funktionierender SQL-Editor oder aktueller Metadatenzugriff beweist diese Befugnisse nicht. Fehlende Operatorrechte führen zu BLOCKED; keine automatische Rollenerhöhung.

6. Vor einem später separat genehmigten Einsatz exakten Originalfingerprint/Version, Backup-/Restorebeleg und Artifact-SHAs prüfen. Wartungszustand/alte Browser/Worker koordinieren; Truststrukturen/RPCs, geprüfte Rechte-/Policy-/CHECK-Deltas und Postconditions transaktional anwenden, soweit geeignet. Vertrauensbindungen nur unabhängig genehmigt provisionieren, App-/Schema-/Realtimecache-/Sessionumstellung getrennt testen. Kein offenes Zwischenfenster, keine Aktivierung halb umgestellter Workflows.
7. Vor Commit bei Fehler Rollback und Postconditions. Nach Commit sicherer Wartungszustand und geprüfte Vorwärtskorrektur bzw. **isolierte** Recovery mit sicheren Rechten und kontrolliertem Journal-/Historienabgleich. Niemals alte unsichere Policies/Grants automatisch reaktivieren oder neuere Produktions-/Auditdaten durch ein altes Backup überschreiben. Wiederanlauf erst nach echten Rollen-/Integritätsprüfungen und Go/No-Go-Verantwortlichem.
8. #25-Pages-/Branchschutz-/Betriebsnachweise und ausdrückliche getrennte Entscheidungen zu Migration, Merge, Deployment und Produktivaktivierung bleiben erforderlich.

## 7. Abschlussentscheidung und vollständige Restblocker

**NEIN.** Die Originalprüfung ist deutlich belastbarer und eigenständig erfolgt; die Humbel-App ist nach diesen Nachweisen **noch nicht vollständig für eine sichere Produktivmigration vorbereitet**.

| Restblocker | Konkreter nächster Handlungsschritt |
|---|---|
| Vollständiger originaler Recovery-/Schema-Nachweis | DBA/Betrieb liefert einmaligen sicheren Schema-only-/Rechte-/Plattformexport mit Herkunft/Umfang/Hashes; genehmigtes isoliertes Ziel und unabhängiger Restore/Version-/Extensionabgleich. Kein manuelles Kopieren weiterer Einzelqueries nötig |
| API/Auth-Konfiguration und reale Rollen | Exponierte API-Schemas/Views/GraphQL, Auth-/JWT-/Revocation-/Realtime-Einstellungen gesichert bereitstellen/prüfen. pgrst.db_schemas war null; HEAD401 ersetzt keinen tatsächlichen Rollentest. Keine Keys in Berichte |
| Verbleibende Rechte auf allen31 | CRUD-/Journal-/Werkzeug-/Fertigungs-/Privatpolicies außerhalb geschützter15, Default-CRUD-/Funktions-/Sequenzrechte und Definer-/Owner-/Vererbungspfade explizit nach Matrix korrigieren und abnehmen |
| Fachentscheidungen D1–D4 | Leser/Felder/Publish, Urlaub/Krankheit/Tausch/Ersatz/Storno/Anträge/Tasks/Fristen, medizinische Privatsicht/Aufbewahrung/Offline und tatsächliche Kontaktfelder verbindlich entscheiden; sichere atomare Verträge ergänzen |
| Vertrauensbindungen D5/D6 | Unabhängiges Enrollment, Auth-/Person-/Abteilungs-/Legacybindungen, Issuer und Befristungs-/Widerrufsprozess mit Verantwortlichem bereitstellen. department_admin im CHECK allein gewährt keine Authority |
| Fertigung/Werkzeuge D7/D8 | Legacy-/Neu-BA-Wahrheitsquelle und erlaubte Werkzeug-/Scanner-/Journal-/Undo-/Inventuraktionen/Atomizität festlegen und positiv testen; vorhandene Übergaben/Rüst-/Mengenlogik erhalten |
| Vollständige sichere App | Alte breite Boot-/SELECT*/Planner-/Cache-/Logpfade, medizinische Projektionen und Session-/Realtime-/Offline-Races vollständig umstellen; kein unsicherer Fallback |
| Technische Abnahme | Fehlende Browser-/Visual-/native echte Rollen-/Parallel-/Originalrestore-/Installations-/Recoverytests am aktuellen endgültigen Gesamtstand bestehen lassen |
| Betrieb/Backup | Verantwortliche, RPO/RTO, Wartungs-/Schreibstopp-/Abbruchplan, verschlüsseltes Backup/PITR und unabhängige Recoveryprobe mit Audit-/Datenintegrität nachweisen |
| Veröffentlichung/Freigabe | Pages/Branchschutz/ausgelieferten SHA nachweisen; eigenes abgenommenes Produktivpaket und separate Migration-/Merge-/Deployment-/Aktivierungsfreigabe einholen |

Keine produktive SQL-Änderung, kein Merge, kein Deployment und keine Aktivierung. Nach Erstellung des neuen Draft-PR endet der Auftrag; keine Actions-Warteschleife oder zusätzlicher Dispatch.

Referenzen: PostgreSQL 17 Privilegien/RLS und Default Privileges (https://www.postgresql.org/docs/17/ddl-priv.html ; https://www.postgresql.org/docs/17/ddl-rowsecurity.html ; https://www.postgresql.org/docs/17/sql-alterdefaultprivileges.html), Supabase RLS (https://supabase.com/docs/guides/database/postgres/row-level-security). Keine Paket-/Runtime-/CLI-Version angehoben.

## 8. Tatsächliche lokale Ergebnisse

Maschinenlesbar: `security-step7-validation.json`. Original-Metadatenprüfung, eingebettete korrigierte Tests und noch offene native/App-Abnahme bleiben getrennte Nachweisklassen.

| Prüfung | Ergebnis |
|---|---|
| npm ci --ignore-scripts; check:syntax inkl. beider Pflicht-Nodechecks; alle fünf neuen JS-Dateien | PASS |
| test:ci / test:qa | PASS **2/2 / 20/20** |
| qa:check | PASS: 251 Legacy-Identitäten, Geschäftscode, Schutzhashes und 33 PNG-Referenzen erhalten |
| test:integration | PASS **5/5** |
| test:security | PASS **55/55**, keine Skips: 45 frühere + zehn neue Step7-Prüfungen |
| Neun vorhandene Fachmodulsuiten | PASS **168/168**, keine Skips |
| qa:agent-server | PASS |
| test:e2e | **BLOCKED_NOT_TESTED**, Exit1: vollständige773 Fälle;773 fehlende Executables vor Seitenaufruf;0 bestanden. Fortschrittsreport complete=true / runnerStatus=failed |
| test:agents | **BLOCKED_NOT_TESTED**, Exit1: Seed am Browserstart blockiert |
| test:visual | **BLOCKED_NOT_TESTED**, Exit2: vorgeschriebener Playwright1.62.1-noble-Container fehlt, keine Referenzänderung |
| Native PostgreSQL/GoTrue/PostgREST-/BA-OP-/Step7-/Parallelabnahme | **BLOCKED_NOT_TESTED**: kein Docker/Podman/psql/zugelassener lokaler Stack; DB-Preflight Exit2 |
| Vollständige Originalrestore-/gehärtete App-/native Installations-/Recoveryabnahme | **BLOCKED_NOT_TESTED**, Voraussetzungen gemäß Abschnitt7 fehlen |
| git diff --check | PASS |

Die produktiven Sicherheitslücken bleiben durch diesen Auftrag unverändert. Positive isolierte Nachweise sind keine produktive Behebung oder Freigabe. Nach Erstellung des Draft-PR wird weder CI nachgeprüft noch ein Lauf gestartet.
