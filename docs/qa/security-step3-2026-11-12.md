# Sicherheitsfreigabe Schritt 3: isolierter Korrekturkandidat

**Keine Produktivfreigabe. Keine produktive Änderung.** Neuer gestapelter Draft auf PR #27 / `feat/home-dashboard-v0153`, unveränderliche Basis `efeb9f5769056d490737c6d4327bb4ae9bc13e73` (0.15.5). Neuer Branch `security/rls-permissions-step3`. APP_VERSION **0.16.0**, VERSION_LOG **2026-11-12 07:42 Europe/Berlin**, tatsächlicher Versionsänderungszeitpunkt. Refs #17, #15, #16, #25–#27.

## 1. Evidenz und Reichweite

Gelesen: AGENTS.md, DEVELOPMENT.md, Sicherheitsbericht #25 einschließlich offener Fachentscheidungen und Originalmetadatengrenzen; Integrationsbericht #26; #15-Diagnose mit tatsächlichen Originalhelfern und strukturellem Spiegel; unveränderter isolierter #16-Kandidat; Fachspezifikation, Artikel-/BA-/OP-/Routing-/Zeit-/Mengen-/Lebenszyklus-Module und vorhandene Berechtigungspfade des kumulativen Entwicklungsstands. Die Startseite wird nicht erneut entwickelt. Der vom Auftraggeber bestätigte erfolgreiche #27-CI-Stand wird als Ausgangsbasis verwendet, nicht als Ergebnis des neuen Kandidaten.

Belegte Schwachstellen: zehn RLS-deaktivierte Tabellen, historisch bestätigte Anon-DML-Grants und permissive Mitarbeiter-/Planerpolicies. Eine zusätzliche permissive Admin-Policy beseitigt bestehendes `true` nicht. Die Tabellen employees und departments bilden zugleich eine manipulierbare Quelle für alte Rollen-/Leiterprüfungen. Die neue Baselineprüfung reproduziert die fremde Rollenänderung **vor** der Korrektur; dieser Test ist ausdrücklich Diagnose, kein Fixnachweis.

Quellen: `docs/qa/schema-inventory.json` (05.11.; ausdrücklich nicht restorable), `production-catalog-2026-11-08.json` (31/308, zehn RLS aus), sechs Originalfunktionen und 25 eigenständige Indizes aus gepinntem #15 `25e9a6914e38fc791d9ecad13d92a45e30d19bb3`. `tests/security/catalog-baseline.js` und `source-catalog.js` sind bytegleich aus dessen `scripts/db-integration/` übernommen; Hash-Provenienz wird getestet. Sie sind Testinfrastruktur/Metadaten, keine neu erfundene Produktionsdefinition.

Keine Live-Supabase-Verbindung, keine Personal-/Fertigungszeilen oder Auth-Benutzer gelesen. Aktuelle vollständige Owner/ACL/Default-Grants, Views, Funktionen und originalgetreuer Restore fehlen weiterhin. Neue Strukturen unter `humbel_security` sind **vorgeschlagenes Kandidatenmodell**, kein angeblich vorgefundenes Originalschema. `employees` hat im nachgewiesenen Schema **keine Kontaktspalten**; eine konkrete produktive Kontaktfeld-Allowlist wird daher nicht erfunden.

## 2. Tatsächlich implementierte Korrekturen

Review-SQL: `supabase/review/security-step3-isolated.sql`, keine registrierte Migration. Ein früher Guard verlangt entweder die bekannte native Testdatenbank `humbel_isolated_test` oder eine tatsächlich als PGlite identifizierte eingebettete Engine, außerdem einen mindestens 32 Zeichen langen Test-Sentinel. Ohne diese Bedingungen Abbruch vor neuen Objekten. Der native Runner behält zusätzlich unverändert Loopback-/Port-/Zufallssentinel-/Container-Netzgrenzen. Die Datei bleibt absichtlich gegen produktive Anwendung gesperrt.

### Vertrauenswurzel und Mitarbeiter

`humbel_security.principals` wird **leer** angelegt. Eine externe, unabhängig geprüfte Zuordnung enthält Auth-ID, eindeutige Mitarbeiter-ID, Rolle, Abteilung, Aktivierung, Ablauf und Evidenz. Kein Import alter Profilrollen, kein `user_metadata`, keine automatische Aufnahme durch ein manipuliertes Profil. Bei Zugriff müssen aktuelle Mitarbeiterbindung, Rolle, Abteilung und Aktivität mit dieser Quelle übereinstimmen. Fehlende, abgelaufene, deaktivierte oder widersprüchliche Bindungen und anonyme Auth-Identitäten scheitern geschlossen.

Alle bestehenden Policies auf employees/departments werden im isolierten Kandidaten ersetzt; auch alte Spaltengrants werden ausdrücklich entzogen. Normale Mitarbeiter können keine Mitarbeiterzeilen schreiben, keine Rolle/Auth-ID/Abteilung/Aktivität ändern und keine anderen Mitarbeiterdaten verändern. Der unabhängig verifizierte Admin behält Mitarbeiterstammdatenpflege. Ein Admin-Update von employee.role verleiht dem Ziel **nicht automatisch** eine neue privilegierte Identität. Die getrennte, überprüfte Authority-Provisionierung muss folgen; bis dahin ist die widersprüchliche Identität gesperrt. Das ist ein absichtlicher sicherer Übergang, kein selbstheilendes Bootstrap.

Die fünf bekannten Original-Helfersignaturen `is_admin`, `current_employee_id`, `current_app_employee_id`, `current_app_employee_role`, `can_access_production_department` werden auf die überprüfte Identität umgestellt, mit leerem search_path und ohne anonymes EXECUTE. Die bisherige Produktions-Abteilungsregel bleibt Admin oder tatsächlicher aktiver Abteilungsleiter; **Planungsvertretung erzeugt keine Produktionsvollmacht**. Die bestehenden unabhängigen BA-/OP-Authority-Tabellen und Fachprozesse werden nicht umgebaut.

### RLS und Grants

Alle zehn bekannten Tabellen erhalten RLS und verlieren PUBLIC-/Anon-Tabellen- und Spaltenprivilegien. Es gibt keine pauschale Admin-only-Ersetzung des gesamten Betriebs. Zusätzlich werden die ebenfalls permissiven vorhandenen `planner_tasks`, `planner_special_days`, `planner_special_day_assignments` sowie employees/departments in diesen isolierten Schutz einbezogen. Für alle zwölf Planertabellen sind rohe Client-Schreibrechte gesperrt, da Namen/Zeilen-ID allein keine sichere Subjekt-/Bereichsbindung und keinen Auditpfad beweisen.

Einheitliche `record_bindings` ordnen ausschließlich ausdrücklich verifizierte Zeilen stabilen Personen und Bereichen zu. Unzugeordnete Altzeilen bleiben auch für einen Admin quarantiniert; es wird kein Namens-Backfill erraten. Für vertrauliche/ungeklärte Daten sind zusätzlich explizite, befristete `record_grants` erforderlich; diese starten leer. Deren `can_write` ist allein **kein** Schreibweg: Raw-DML bleibt entzogen, ein neuer privater Schreibworkflow ist nicht implementiert. Damit werden unentschiedene Kranken-/Urlaubs-/Tauschrechte nicht stillschweigend erteilt.

### Erlaubte Abteilungsplanung, Stellvertretung und Veröffentlichung

`security_assignment_command` erlaubt den nachweislichen aktiven AL seiner Abteilung oder eine ausdrücklich zugeordnete, zeitlich gültige, nicht widerrufene Vertretung des weiterhin aktiven verifizierten Leiters. Kein pauschales Adminrecht für AL, keine abteilungsübergreifende Ableitung. Verifizierte Admins haben administrative Planungssicht. Mitarbeiter/Scanner dürfen keine Schichten ändern.

Der Aufruf prüft Abteilung und aktive eindeutige Zielperson serverseitig, leitet den vorhandenen Namensschlüssel aus der geprüften Person ab, sperrt relevante Zeilen, serialisiert dieselbe Zuordnung per Advisory-Lock, prüft die erwartete Version und protokolliert die Änderung atomar. Falsche/fehlende Version, fremde Abteilung, mehrdeutiger Name und ungebundene Altzeile werden abgewiesen. Ein Auditfehler rollt Daten und Bindung zurück. Direkte Planänderungen umgehen diesen Pfad nicht.

Schreiben, Freigeben und Lesen sind getrennt: `schedule_grants` sind unabhängig, befristet, widerrufbar und an eine konkrete Abteilung gebunden. Ein Planer darf nicht automatisch veröffentlichen. `security_publish_assignment` bindet eine Veröffentlichung an die aktuelle Version. Änderungen löschen die Veröffentlichung. `security_published_schedule` liefert nur explizit freigegebene aktuelle Zuordnungen im erlaubten Bereich: assignment_id, shift_id, shift_date, employee_id, keine Krankheitsgründe, Rollen, Auth-IDs oder Personalprofile. Dies ist eine vorgeschlagene minimale Projektion; D1 bestimmt deren endgültige Reichweite und Darstellung.

### Selbstservice und Werkzeugreferenzen

Mangels nachgewiesener Kontaktspalten gibt es einen **vorgeschlagenen separaten** eigenen Kontaktstore und `security_own_contacts`, ohne fremden Zielparameter. Die Feld-Allowlist ist im Kandidaten leer. Erst explizit freigegebene Zeichenkettenfelder sind schreibbar; sicherheitskritische Schlüssel bleiben abgewiesen. Die Tests erlauben ausschließlich `synthetic_contact` als ausdrücklich künstlichen Testfall. Das ist **keine Entscheidung**, welche produktiven Telefonnummern-/E-Mail-Felder geführt werden sollen, und keine fertige Anbindung an die App.

`tool_materials` bleibt für verifizierte aktive Mitarbeiter einschließlich Scanner lesbar, für verifizierte Admins schreibbar. Die übrigen Werkzeugtabellen/-Buchungsfunktionen werden nicht verändert; deren vollständige Fach-/UI-/Realtime-Abnahme gegen den gehärteten Zustand fehlt weiterhin. Aus einem Materialkatalogtest wird kein vollständiger Werkzeugverwaltungsnachweis abgeleitet.

## 3. Geänderte SQL-Objekte

- Neues privates Schema, neun neue Tabellen: principals, delegations, schedule_grants, record_grants, record_bindings, publications, contact_fields, contacts, audit (neun Tabellen; kein öffentliches Geschäftstabellen-Backfill).
- Private Helfer: actor_id, actor_role, is_admin, can_plan, schedule_scope, record_scope, guard_assignment; Trigger `security_assignment_guard` auf planner_assignments.
- Vier neue öffentliche, explizit für authenticated freigegebene RPCs: security_assignment_command, security_publish_assignment, security_published_schedule, security_own_contacts. Interne Tabellen/Sequenzen sind für API-Rollen unzugänglich; interne Funktionen erhalten nur die für Policies benötigten EXECUTE-Rechte.
- Fünf bekannte öffentliche Helfer mit erhaltener Signatur angepasst; ursprünglicher `set_updated_at` nicht geändert.
- Policies/Grants auf 15 explizit genannten vorhandenen Tabellen ersetzt: employees, departments, zwölf Planertabellen und tool_materials. Die übrigen 16 ursprünglichen Tabellen sowie bestehende Artikel-/BA-/OP-SQL-Dateien bleiben unverändert. Bestehende unbekannte Definer-/View-/Default-ACL-Pfade sind **nicht vollständig abgenommen**.
- PUBLIC-/Anon-/Authenticated-CREATE auf public entzogen. Default-EXECUTE wird für neue Funktionen des ausführenden Owners im privaten Schema entzogen; das ist kein globaler Nachweis sämtlicher ursprünglicher Owner-/Default-Grants.

Hinweis: Der Testoperator/Owner ist in der Spiegelumgebung privilegiert. Endgültige nicht anmeldbare Definer-Ownerrollen, minimale Objektgrants und eventuelle BYPASS-/FORCE-RLS-Auswirkungen sind erst anhand des autoritativen Restores abzunehmen. Leerer search_path und unzugängliches privates Schema ersetzen diese Ownerprüfung nicht.

## 4. Testtrennung und native Pipeline

`npm run test:security` führt die tatsächliche unveränderte Review-SQL in einer wegwerfbaren PGlite-Engine gegen den aus #15 übernommenen strukturellen Spiegel aus. Auth-Funktionen und Rollen im eingebetteten Lauf sind ausdrücklich synthetische Shims; keine JWT-/PostgREST-/Mehrverbindungsbehauptung. Ein Test reproduziert die unsichere Baseline. Davon getrennt folgen Tests für den korrigierten Zustand; ein grüner Baseline-Test wird nie als Korrektur gewertet.

Neuer separater Workflow `Security Step 3 Isolated Candidate`: erst eingebettete Sicherheitsverträge, dann der unverändert gepinnte native #15-Stack mit allen vorhandenen Baseline- und BA-/OP-Phasen/Erstzuordnungstests aus #26/#27. **Die drei bisherigen Diagnose-/#16-Härtungs-/BA-OP-Jobs bleiben unverändert.** Danach wird in dieser eigenen Umgebung der neue Kandidat installiert und mit echten GoTrue-JWTs/PostgREST/SQL-Verbindungen geprüft. Der neue Runneradapter akzeptiert nur den exakten SHA des bereits durch Clock-/Provenienzadapter geprüften Runners und fügt eine rückführbare Fortsetzung ein; keine Baselineassertion wird ersetzt. SHA-Manifest des konkreten aktuellen Checkouts bleibt vor der ersten DB-Verbindung zwingend. Native Statuswerte des Kandidaten und offener Baselinebefund bleiben getrennt; authoritativeRestore bleibt false. Cleanup und Artefakte bleiben garantiert.

Acht neue native Fallgruppen sind vorbereitet: Anon-/RLS-/Privatschemazugriff, fremde/eigene Rollenmanipulation, eigene/fremde Planung, Vertretung/Widerruf/Ablauf, getrennte Veröffentlichung/Leseprojektion, Kontakt-Allowlist/Materialkatalog, echte konkurrierende Versionsänderungen und BA-Anlage/abteilungsübergreifende Maschinenbindung nach Härtung. Diese native Fortsetzung wurde hier **nicht ausgeführt**; ihre erfolgreiche Syntax-/Adaptermontage ist kein nativer Testnachweis.

## 5. Lokale Ergebnisse

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| `npm ci --ignore-scripts` | PASS |
| Pflicht-`node --check`, `check:syntax`, alle neuen JS-Dateien | PASS |
| `test:ci` / `test:qa` / `qa:check` | PASS, **2/2**, **18/18**; 251 Legacy-Identitäten, unveränderte Geschäftslogik-/Test-/Bild-Hashes und 33 Referenzen |
| `test:integration` | PASS, **5/5**, vorhandene Integrationsverträge unverändert |
| `test:security` | PASS, **17/17**, keine Skips: 13 SQL-Fallgruppen (eine Baseline, zwölf korrigierter Zustand) und vier Provenienz-/Orchestrierungs-/Isolationsprüfungen |
| Bestehende Artikel/OP/Readiness/Forecast/BA/Routing/Rüsten/Menge/Lebenszyklus-Suiten | PASS, **8+15+17+41+9+10+17+14+37 = 168/168**, unverändert, keine Skips |
| `qa:agent-server` | PASS |
| Exakte Montage des gepinnten nativen Runners mit Clock-, Provenienz- und neuer Sicherheitsfortsetzung | PASS einschließlich JavaScript-Parse; kein DB-Runtime-Nachweis |
| Vollständige Quellenmontage aus sauberem lokalem Implementierungscommit | PASS: 44 manifestierte SQL-/JS-/Erweiterungsdateien; Manifest erneut verifiziert; beide unveränderten nativen Target-Guards **2/2** bestanden |
| `test:e2e` | **BLOCKED_NOT_TESTED**: vollständige **758** Fälle aufgerufen, alle 758 vor Seitenaufruf wegen fehlender Browser-Executables blockiert; 0 bestanden, 0 Skips |
| `test:agents` | **BLOCKED_NOT_TESTED**: Seed aufgerufen, gleicher fehlender Browser |
| `test:visual` | **BLOCKED_NOT_TESTED**, Exit 2: vorgeschriebener Playwright-1.62.1-noble-Container fehlt, keine Referenzänderung |
| Native PostgreSQL/GoTrue/PostgREST/Edge-/Paralleltests; `test:db:preflight` | **BLOCKED_NOT_TESTED**, Preflight Exit 2: kein Docker/Podman/psql oder geprüfter Teststack. Die neue native Fortsetzung ist vorbereitet, nicht bestanden |
| `git diff --check` | PASS |

Zusätzliche explizite Privatdaten-Leserechte werden nur mit synthetischer Testfreigabe und Zeilenbindung positiv geprüft; Ablauf sperrt den nächsten Lesezugriff und Lesen erteilt keinen Schreibzugriff. Dies entscheidet keine produktiven medizinischen Leserrechte. Der erste Implementierungsversuch zeigte zwei Fehler im neuen Testfixture (nicht eindeutige synthetische shift_id, anschließend uneinheitlicher Parametertyp beim Verketten von UUID/Text). Die neue Fixture wurde anhand des echten Unique-Constraints und expliziter UUID-Konvertierung korrigiert; keine Bestandsassertion entfernt oder abgeschwächt. Finale 17/17 erfolgreich.

Ein erster Clock-Prepare-Aufruf wurde im Featurecheckout ohne native Harnessdateien gestartet und brach vor einem Test ab. Im korrekt zusammengestellten Harness wurden danach alle Adapter, Syntax, Quellenmanifest und beide unveränderten Target-Guards erfolgreich geprüft. Keine Sicherheitsassertion angepasst.

Kein neuer GitHub-CI-Erfolg wird vorweggenommen. Der Auftrag endet nach PR-Erstellung ohne Actions-Abfrage, Warteschleife oder Dispatch.

## 6. Offene Fachentscheidungen und harte Freigabesperren

Maschinenlesbare Matrix: `security-step3-access-matrix.json`. Die neue Aufgabenbeschreibung entscheidet AL-Planung der eigenen Abteilung, unabhängige Rollenverwaltung, zeitlich gültige Vertretung und eigenen erlaubten Kontakt-Selbstservice. Sie bestimmt nicht automatisch D1–D6 im Detail:

1. **Schichtansicht:** Personenkreis (eigenes Team/Abteilung/Betrieb), konkret sichtbare Namen/Felder und Veröffentlichungszuständigkeit bestätigen. Die Minimalprojektion und Scope-Grants sind ein prüfbarer Vorschlag.
2. **Private Planung:** Genehmiger, Sichtbarkeit und Bearbeitung für Urlaub/Krankheit, Ausfälle, Tausche, Verfügbarkeit, Samstag und Ersatzzuordnungen festlegen. Dazu fehlen sichere Schreib-RPCs und UI-Adapter; Raw-DML bleibt gesperrt. Nicht als vollständig erhaltene Alt-UI-Kompatibilität ausgeben.
3. **Kontakte:** tatsächliche eigenen editierbaren Kontaktfelder, Validierung, Speicherung, Löschung und UI-Integration freigeben. Leere Default-Allowlist bleibt bis dahin bestehen.
4. **Vertrauenswürdige Identität und Vertretung:** Verantwortlichen und Prüfprozess für Erstprovisionierung, Namensauflösung, Zu-/Entzug, Befristung und Audit benennen. Keine Bestätigung aus derzeit frei manipulierbaren Profilen ableiten. Ein DBA darf nicht durch einfaches Kopieren aller heutigen employee.role-Werte neue Principals erzeugen.
5. **Originalnachweise:** aktueller vollständiger, konsistenter schema-only-Export einschließlich Owner/ACL/Default-Grants, Funktionskörper, Abhängigkeiten, Views/Invoker-Verhalten, Trigger, Sequenzen, Extensions/Versionen und API-Konfiguration; Restore ohne Testsubstitute in disponibler Umgebung. Für personenbezogene Zuordnung gesondertes autorisiertes Verfahren außerhalb öffentlicher Testartefakte.
6. **Gemeinsame Abnahme:** neue native Tests, alle erlaubten UI-/Schicht-/Werkzeug-/Fertigungsabläufe unter gehärteter DB, session_id-/Logout-/Tokenwiderruf, Rechteänderungen bei laufenden Requests, Realtime, Datenschutz-/Cacheminimierung und Last/Sperrverhalten. Bisherige BA-/OP-RPCs enthalten teils eigene Employee-/Authority-Abfragen; diese sind unverändert und nicht durch den neuen Root allein vollständig formal abgesichert. Alle erlaubten Autorisierungswege müssen am Originalrestore nochmals gemeinsam geprüft werden.

Der Kandidat beweist keine vollständige Kompatibilität mit der heutigen direkten Planner-/Employees-SELECT*-Oberfläche. Deren bisheriger breiter Datenabruf darf nicht dadurch „erhalten“ werden, dass unsichere Rechte fortbestehen. Der Kandidat wird **nicht aktiviert**, bevor sichere Adapter und positive vollständige Regression vorliegen. Die offenen #25-Freigabesperren zu Schemafidelität, Rollenprovisionierung, Cache, Pilot, Pages, Betriebs- und Backupnachweisen bleiben bestehen.

## 7. Kontrollierter späterer Migrations-, Restore- und Rückfallplan

1. DBA/Betrieb erstellt separat autorisierten konsistenten Metadaten-/Owner-/ACL-Export. Prüfsumme und Zeitpunkt festhalten; keine Mitarbeiter-/Fertigungszeilen in PR-Artefakte übernehmen. Freigegebene Fachmatrix, Identitäts-/Bereichsabgleich und Privacy-Verantwortliche dokumentieren.
2. Autoritativen Schema-Stand in frischen isolierten Stack mit passender PostgreSQL-/GoTrue-/PostgREST-/Extension-Konfiguration restaurieren. Kandidat gegen tatsächliche DDL/Owner/ACL anpassen und Schema-/Grantdiff unabhängig reviewen. Änderungen nicht über neue Hashes „akzeptieren“, sondern fachlich und mit Tests erklären.
3. UI-Adapter und alle bestätigten Planner-/Kontakt-/Werkzeug-/BA-/OP-Positiv- und Negativfälle am gehärteten Restore abnehmen, einschließlich fremder und verwaister Identitäten, Widerruf/Expiry und echter Paralleltransaktionen. Offene Rechte bleiben gesperrt, nicht stillschweigend Admin-only oder global freigeschaltet.
4. Vor späterer ausdrücklich genehmigter Produktionsmaßnahme verschlüsseltes Backup/PITR und vollständige Schema-/ACL-/Function-/Authority-Sicherung erstellen, **Restoreprobe in separater Umgebung messen**, Verantwortliche, RPO/RTO, Wartungsfenster, Schreibstopp und Stop-Kriterien festlegen. Pages-/Branch-/Environment-Nachweise aus #25 beibringen.
5. Erst danach eigener freigegebener Produktionsmigrationsentwurf mit konkretem erwarteten Schemasnapshot, Voraussetzungen, Ownern, Timeouts und atomarer Transaktion. Die hier vorliegende Review-Datei bleibt unverändert produktionsgesperrt; ihr Guard darf nicht im laufenden Betrieb „kurz entfernt“ werden. Keine automatischen Principal-/Personen-Backfills, keine Feature-/Workeraktivierung nebenbei.
6. Vor Commit bei Fehlern gesamte Transaktion abbrechen. Nach Commit bei Sicherheits-/Kompatibilitätsabweichung betroffenen Zugriff kontrolliert stoppen, Daten/Audits erhalten und vorab geprüfte sichere Kompatibilitätsvariante nutzen. **Unsichere permissive alte Policies nicht automatisch wieder einschalten.** Keine Snapshot-/Audit-/Zählerdaten löschen. PITR/Restore nach neuen Writes nur als separate Incident-Entscheidung mit gemessenem Verlustfenster und gesicherter Aufarbeitung.
7. Erst nach unabhängiger Review, vollständigen Nachweisen und ausdrücklicher Eigentümerfreigabe getrennt über Code-Merge, DB-Migration, Deployment und Aktivierung entscheiden. Dieser Draft führt keinen dieser Schritte aus.

## 8. Eindeutige Einschätzung

Im **eingebetteten isolierten Kandidaten** nachgewiesene Abweisungen: Anonzugriff auf zehn bekannte Tabellen, fremde/eigene Mitarbeitereskalation und Leiteränderung, rohe Planänderung, fremde Abteilung/Zielperson, unbestätigte/abgelaufene/widerrufene Vertretung, ungeprüfte Principal-/Kontaktfelder, unfreigegebene oder fremde Schichtprojektion. Positive SQL-Nachweise: verifizierte Administration, eigene AL-/befristete Vertreterplanung, getrennte Veröffentlichung/Schichtlesen, konfigurierter eigener Testkontakt und Materialreferenzlesen für Mitarbeiter/Scanner; Auditrollback und Versionskonflikte geprüft.

**Nicht nachgewiesen:** produktive Schließung, vollständiger Originalrestore, neue native Ausführung, volle heutige Planner-/Werkzeug-UI-Kompatibilität oder vollständiger Session-/Realtime-/Lastnachweis. Endgültige Produktivfreigabe bleibt ausdrücklich gesperrt. Benötigt werden die konkreten Entscheidungen und Administratornachweise aus Abschnitt 6, keine pauschale Bestätigung „alles freigeben“.

Quellen für PostgreSQL-/Supabase-Regeln: https://supabase.com/docs/guides/database/postgres/row-level-security ; https://supabase.com/docs/guides/database/postgres/column-level-security ; https://www.postgresql.org/docs/current/sql-revoke.html . Changelog geprüft, keine Bibliotheks-/CLI-Version angehoben. Keine Aussage, dass Dokumentation den individuellen Datenbankzustand belegt.
