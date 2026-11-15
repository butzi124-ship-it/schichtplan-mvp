# Schritt 5 – gesperrter Migrations-, Backup- und Abnahmeplan

**Plan, keine Produktionsausführung.** Basis PR #29 und Schritt-5-Bericht. Alle Gates sind kumulativ; ein fehlender Nachweis verhindert den nächsten Schritt. Produktive Daten, Credentials und vollständige Datenbackups gehören nicht ins Repository oder in CI-Artefakte. Bis Gate G7 ist keine produktive Umstellung vorgesehen; G7/G8 verlangen einen neuen ausdrücklichen Auftrag.

## G0 – Zuständigkeiten und Originalnachweis beschaffen

Betreiber benennt Datenbankadministrator (Export/Backup), Security-Reviewer (Rollen/ACL), Fachverantwortliche Personal/Planung/Fertigung/Werkzeuge, Restore-Operator und Go/No-Go-Verantwortlichen. Namentliche Zuständigkeiten, Vieraugenfreigabe, Wartungsfenster, Abbruchfrist, RPO/RTO und Aufbewahrung sind **offene Betreiberentscheidungen**. D1–D6 unverändert aus `security-step3-access-matrix.json`; zusätzlich D7 Alt-/Neu-BA-Zuordnung, D8 Werkzeug-/Scanner-/Inventurmatrix, O1 RPO/RTO/Backup/Restorebetrieb. Kein implizites „Admin darf alles“ als Ersatz.

### Präzise Lieferanforderung an den Administrator

**Keine Ausführung durch diesen Auftrag.** Ein Administrator mit bereits genehmigter Exportberechtigung liefert einen **Schema-only**-Export. Keine employees-/planner-/tools-/production-/auth.users-/Kundenzeilen. Folgendes Paket ist erforderlich:

| Datei / Nachweis | Erforderlicher Inhalt |
|---|---|
| manifest.json | Projektref/Region, DB-/Engineversion, UTC Beginn/Ende, Operatorbestätigung, Client-/Serverversion, vollständige Befehle ohne Credentials, ausgewählte/ausgeschlossene Schemas samt Grund, Exitcodes/Warnungen, Snapshot-/DDL-Freeze-Nachweis, SHA-256 aller gelieferten Dateien |
| schema.sql | pg_dump **--schema-only**, Owner und ACL erhalten, keine --no-owner/--no-acl, nicht nur public. Tabellen/Spalten/Typen/Defaults/Identity/Constraints/FK/Indizes inkl. invalid/unready, Views/Matviews, Sequenzdefinitionen, Funktionen mit vollen Signaturen/Bodies/Volatilität/Security/search_path, Trigger/Eventtrigger, Policies/RLS/FORCE RLS, Grants und Default ACL |
| roles-sanitized.sql/json | Rollenflags, Mitgliedschaften inkl. Admin-/Inherit-/Set-Optionen, Objektowner, Login/BYPASSRLS/CREATE/CREATEROLE; **keine Passwörter oder Passwort-Hashes**. Native Rolleexportvorlage pg_dumpall --roles-only --no-role-passwords, nur durch autorisierten Administrator, nicht durch Connector-General-SQL |
| catalog.json | Objektidentitäten/OIDs nur als lokale Vergleichshilfe; pg_class/attribute/constraint/index/policy/proc/trigger/namespace/sequence/extension/default_acl und effektive table/column/schema/sequence/EXECUTE-Rechte für PUBLIC/anon/authenticated/Service-/Ownerrollen; auch implizite PUBLIC- und geerbte Rechte. Read-only konsistent erfasst, keine Anwendungsrelationen abfragen |
| platform-redacted.json | API-exponierte Schemas und Suchpfad, Realtime-Publications/Tabellenauswahl, Auth-/JWT-Rollen-/TTL-/Revocation-Konfiguration, Extensionversionen und Migrationsidentitäten; keine Signingkeys, Passwörter, Tokens, Providersecrets, Benutzer oder Sessions |
| completeness.md | Alle Datenbankschemas aufzählen, auth/storage/realtime/vault/extension-Schemas und kundeneigene Änderungen explizit klassifizieren; ausgelassene/verwaltete Objekte mit Recoveryquelle. Nicht aus „CLI dump erfolgreich“ Vollständigkeit ableiten |

Eine administrative Befehlsvorlage für den **später separat autorisierten** Schemaexport ist `pg_dump --schema-only --format=plain --file=schema.sql service=humbel_metadata_export`. Der Service verweist auf eine administrativ eingerichtete credentialgeschützte Verbindung; kein Passwort in Kommandozeile/Repo. Vorher bestätigte pg_dump-17-Version und Serverkompatibilität prüfen. Ohne Schemafilter umfasst pg_dump alle zugänglichen Schemas; fehlende Zugänglichkeit, Extensionobjekte, Plattformobjekte und Warnungen müssen trotzdem ausgewiesen werden. Rollenexport ist ein eigener Zugriff und nur bei passender administrativer Berechtigung zulässig. Kein beliebiger execute_sql-Fallback und keine Benutzer-/Datenabfragen zur Berechtigungsprobe.

Schema-/Rollenexporte können sensible Defaults/Funktionsliterale/FDW-Optionen enthalten. Der Administrator prüft sie **vor Übergabe**. Secrets nicht auslesen oder in Git schreiben; falls authentische Original-DDL Secrets enthält, Original verschlüsselt in kontrolliertem Betreiberarchiv halten, bereinigte Arbeitskopie mit dokumentierter Redaktionsstelle/Originalhash liefern. Eine redigierte Kopie ist kein byteidentisches Original. Vollständigkeitslücken bleiben UNVERIFIED und Restore nur mit autorisiertem Secret-Provisioning.

Ein Schema-only-Export ist kein Datenbackup. Die technische Herkunft muss sich von einem aus Repositorydateien erzeugten Spiegel unterscheiden. Ein konsistenter pg_dump-Snapshot und zusätzliche zeitlich getrennte Rollen-/Konfigurationskataloge sind nicht automatisch ein gemeinsamer Snapshot: DDL-Änderungen währenddessen ausschließen oder Abweichungen mit Wiederholungs-/Freeze-Nachweis administrativ behandeln. Nichts davon ist heute erfolgt.

## G1 – Original isoliert wiederherstellen, ohne Kandidat

Voraussetzung: ausdrücklich benanntes freigegebenes **wegwerfbares isoliertes Ziel**, keine produktiven URLs/Credentials, ausgehendes Netzwerk gesperrt, keine E-Mail/Webhooks/Realtime-Brücken zum Original. PostgreSQL 17.11.0.003 und passende Auth/PostgREST-/Extensionversionen oder ausdrücklich bewertete Abweichung. Vorhandenes Harness nutzt 17.6.1.136; das ist kein Versionsgleichheitsbeweis.

1. Manifest/Hashes und Vollständigkeit unabhängig prüfen, Export als untrusted SQL prüfen. Rollen/Owner/Extensions vor Schema in passender isolierter Plattform provisionieren. Plattforminterne Objekte nicht blind über vorhandene Supabase-Instanz spielen. Jede nötige Übersetzung protokollieren und vergleichen.
2. Schemaimport mit ON_ERROR_STOP ausführen, Exitcodes/Warnungen sichern. Keine fehlenden Originaldefinitionen aus Kandidatdateien ergänzen und dann „Originalrestore“ nennen.
3. Katalog erneut extrahieren, vollständigen erwarteten/ist Diff erstellen: strukturelle Objekte, Indizes, Constraints, alle Policyausdrücke, Owners, ACL/Default ACL/Mitgliedschaften, Funktionsoverloads/search_path/EXECUTE, Trigger/Views/Dependencies. Semantische Unterschiede nicht durch pauschales Sortieren/Entfernen von ACL- oder Ownerzeilen weg-normalisieren.
4. Synthetische, rollengetrennte Identitäten und minimal notwendige Daten nur auf diesem Ziel erzeugen; unabhängige Authority-Zuordnung mit synthetischen Belegen. Original-Mitarbeiterdaten werden nicht in Tests übernommen.
5. Erst **unveränderter Originalzustand** separat diagnostizieren; Exploit-Reproduktion als BASELINE_FINDING dokumentieren. Nicht mit korrigiertem Kandidatenstatus vermischen.

**G1 PASS erfordert** einen nachweislich vollständigen Originalschema-Restore mit protokollierter Plattform-/Versionsabbildung. Der bisherige 31-Tabellen-Spiegel erfüllt das nicht.

## G2 – Fachverträge und Appadapter schließen

`security-step5-access-plan.json` ist die Arbeitsliste mit stabilen A001–A152-Identitäten. Jede Zeile erhält erst nach Umsetzung einen realen Test-/Diffbeleg; `fullyIntegratedApplication:false` ist derzeit korrekt. Ergänzend alle im Bericht genannten RPC-/Edge-/Realtime-/Storage-/Auth-Pfade prüfen.

| Reihenfolge | Präzise Änderung | Freigabe-/Testbedingung |
|---|---|---|
| 1 Identität | Boot ohne Personenprobe; getUser + serverseitiger Actor, stabile IDs; Felddaten von Authority trennen | Admin/AL/Vertretung/Mitarbeiter/Scanner, fehlende/mehrdeutige/verwaiste/inaktive/widersprüchliche Bindungen, kein Lokalrollenfallback |
| 2 Lesedaten | Explizite Feldprojektionen für Personal, veröffentlichte Schichten, medizinische Privatsichten, Material, Produktion/Werkzeuge | D1/D3/D4, fremde Felder auch per REST/RPC nicht lesbar; kein SELECT*-Ersatz als Sicherheitsumgehung |
| 3 Schreiben | Sichere einfache assignment/publish/contact-Verträge zuerst; mehrteilige Planungs-/Werkzeugaktionen als atomarer Fachbefehl | D2/D4/D8; stabile Request-ID/Version, Rollen-/Abteilungsscope, Audit, Konflikt/Retry, keine Teilbuchung |
| 4 Alt-/Neu-Fertigung | Bestehende BA-/OP-Commands beibehalten, Legacy-Direktpfad nur nach Originallogik-/Scopeprüfung | D7; ein Auftrag hat klare Buchungsquelle; Maschinenübergaben zwischen Abteilungen, Rüsten/Menge/Nachfertigung/Abschluss unverändert erlaubt |
| 5 Datenlebenszyklus | Sensible Logs entfernen; nur freigegebene UI-Präferenzen persistieren; Altdaten vor Laden migrieren/löschen; Generationen für Requests/DOM/Realtime | D3/D6; Logout/Sessionwechsel/Offline/Tokenrefresh/Widerruf, kein fremder Cache; vor Löschung ungesicherte Änderungen prüfen und kontrolliert erhalten |
| 6 Gesamtaktivierung | Vollständige UI mit sicheren Verträgen in isoliertem Restore laden | Keine Teilumstellung und kein Fallback auf permissive Raw-Tabellen. Alte normale Anwendung bis vollständiger Abnahme unverändert lassen |

Kein aus Mitarbeiterrolle/JWT-user_metadata/Displayname/Schichtkürzel automatisch erzeugtes Enrollment. Administrative Prüfung der tatsächlichen Auth-/Personal-/Abteilungszuordnung und historischer Zeilen ist eine getrennte geschützte Betriebsaufgabe mit Freigabe/Audit, nicht Teil des Metadatenauftrags.

## G3 – versioniertes Migrationspaket erst nach G1/G2 erstellen

**Heute kein ausführbares Produktivpaket.** #28/#29-Dateien behalten ihre Isolationsguards. Sie beginnen Transaktionen, enthalten CREATE-Objekte und sind nicht als wiederanwendbare Upgrade-Migration konstruiert. Guard entfernen reicht nicht. Originalhelper/Policies können andere Overloads/Abhängigkeiten/Owner haben.

Das spätere Paket muss enthalten: exakte Ausgangsschema-/Policy-/ACL-Fingerprints und Serverversion, zulässige Zielref/Umgebungsprüfung, freigegebene SQL-Objektdeltas, Vorwärtsversions-ID+Checksum, Abhängigkeitsreihenfolge, Lock-/Statement-Limits, erwartete Katalogpostconditions, Abbruchkriterien und passende Appversion. Eine unbekannte Ausgangsversion oder Abweichung führt zum Abbruch **vor** DDL. Advisory Lock/Versionsjournal verhindert konkurrierende Migrationen; gleiche Version+gleicher Hash darf nach Prüfung der Postconditions kontrolliert No-op sein, gleicher Name+anderer Hash muss abbrechen. Keine IF NOT EXISTS-Kaskade zum Verdecken unbekannter Definitionen.

Transaktional geeignete Reihenfolge: Objekt-/Rollen-/ACL-Preconditions → private Strukturen/Owner/Minimalgrants → gesicherte Helfer/RPCs → genau überprüfte Legacy-Policy-/Spalten-/Tabellenrechte ersetzen → RLS/FORCE-RLS-Entscheidung je Owner/Definer prüfen → erforderliche vertrauenswürdige Bindungen kontrolliert provisionieren → SQL-Postconditions → Commit. Keine Zwischenphase mit breiter PUBLIC-Ausführung oder wieder offener Policy. PostgREST-Schemacache/Realtime/Token-/Clientsynchronisierung außerhalb der SQL-Transaktion explizit orchestrieren und testen. Nichttransaktionale Schritte wie gegebenenfalls CREATE INDEX CONCURRENTLY separat planen; ihre Zwischenzustände und Rücknahme dürfen nicht pauschal als atomar gelten.

Keine automatische Produktionseinspielung, keine Registrierung in produktiven Migrationsworkflows und kein Aktivierungsflag in Schritt 5.

## G4 – getrennte technische Abnahmematrix

| Nachweisklasse | Test / Bestehende Infrastruktur | Erforderlicher Beleg vor Produktion |
|---|---|---|
| Historische Baseline | #15/#25-Diagnose, eigener unveränderter Restorezustand | Erfolgreiche Lückenreproduktion bleibt Finding, niemals FIXED |
| Isolierter struktureller Kandidat | test:security/PGlite, native Step3/4 und BA-/OP-Harness mit gehashten Inputs | Positive/negative Rollen-, ID-, Scope-, Vertretungs-/Versionstests; explizit authoritativeRestore:false, bis echte Originalbasis vorliegt |
| Originalschema + Migration | G1-Restore, versioniertes Paket G3, GoTrue/PostgREST mit verschiedenen JWTs | Anon und unauthentifizierte Writes denied; keine Rollen/Auth-/Abteilungsmanipulation, kein fremdes Personal/medizinische Felder; AL eigene Abteilung, gültige/abgelaufene/widerrufene Stellvertretung, Veröffentlichung getrennt, eigener freigegebener Kontakt erlaubt |
| Vollständig integrierte Anwendung | Vollständige Browsermatrix und tatsächliche sichere Clientanbindung | Keine Boot-/SELECT*/Cache-/Sessionumgehung; Schichtlesen/Schreiben/Veröffentlichen, Werkzeug/Scan/Journal/Inventur und BA-/OP-Übergaben fachlich richtig; unveränderte Visuals |
| Transaktionen / Fehler | Native getrennte DB-Verbindungen, nicht PGlite-Shims | Parallele Planung/Scan/Mengenbuchung, Deadlock/Timeout/Netzabbruch vor und nach Commit, sichere Wiederholung: keine Doppelbuchung/Teilerfolge, Audit vollständig, Datenerhalt |
| Upgrade/Idempotenz/Rückfall | Wegwerfbare Kopien A/B, nicht Originalsystem | Erste Installation, zweite gleiche Version, falscher Hash/Schema, Fehler mitten im Paket: korrekter Abbruch/No-op; kein Daten-/Historienverlust; getestete sichere Vorwärtskorrektur bzw. geschlossener Rückfall |
| Sitzungen / Widerruf | GoTrue+REST+Realtime, Browser-Race-/Offlinefälle | Alte JWTs/abgelaufene Claims, Tokenrefresh, Logout, parallele Tabs, Rollenwiderruf während Anfrage, gesperrte/offline Caches, keine Antwort der alten Identität im neuen DOM |
| Unabhängige Wiederherstellung | G5 isolierte Recoveryprobe | Vollständiger Restore, Owner/ACL/Policies/Versions-/Integritätsabgleich, gemessene RPO/RTO und Start nur unter sicheren Rechten |

Jeder Nachweis bindet geprüften App-SHA, SQL-Checksummen, Exportmanifest, Runtime-/Extensionversionen und tatsächliche Ergebnisse. BLOCKED_NOT_TESTED/fehlende Artefakte blockieren; keine Skips, gelockerten Assertions, neuen Visualreferenzen oder Stand-ins als Originalbeweis. Tabellen-/Buchungsanzahlen und Integritätschecks dürfen nur im ausdrücklich autorisierten isolierten Ziel laufen; produktive Datenprüfungen verlangen gesonderte Berechtigung.

## G5 – Backup und unabhängige Restoreprobe vorbereiten

**Heute weder Backup noch Probe durchgeführt.** Betreiber bestimmt RPO/RTO, Zeitraum, Wartung/Schreibsperre, Datenschutz und genehmigtes Recoveryziel. Schema-only deckt keine Datenwiederherstellung ab. Später separat autorisierte vollständige Sicherung benötigt Daten **und** Schema, Rollen/Mitgliedschaften ohne versehentliche Secretsveröffentlichung, Owner/ACL/Default Grants, Extensions, Auth-/API-/Realtimekonfiguration und Recoveryinformationen für separat verwaltete Secrets/Schlüssel. Vault-/Storage-Abhängigkeiten und externe Worker/Queues berücksichtigen. Ein Datenbankbackup sichert nicht automatisch Storage-Objektdateien.

Betreiber prüft vorhandenen Plan/PITR tatsächlich: aktives Recoveryfenster, letzter erfolgreicher Backup-/WAL-Stand, Zielzeitpunkt, Aufbewahrung, Kosten und Zuständigkeit; Planverfügbarkeit nicht aus Supabase-Projektstatus ableiten. Verschlüsselte, zugriffsbeschränkte getrennte Aufbewahrung, kontrollierter Schlüsselzugriff mit Notfallvertretung, Integritäts-/Hashprüfung und Audit. Kein volles personenbezogenes Backup als GitHub-Artefakt.

Unabhängiger Operator stellt in freigegebener isolierter, abgeschotteter Umgebung wieder her. Produktive Benutzer-/Sessiondaten nur mit eigener Datenschutz-/Restoreberechtigung, sonst synthetischer Datenlauf und ausdrücklicher Hinweis, dass Datenrestore noch nicht bewiesen ist. E-Mail/Jobs/Webhooks deaktiviert, Secrets kontrolliert ersetzt. Wiederherstellung von Rollen/Ownern/ACL/Policies und historischen Audit-/Produktionsketten prüfen, nicht nur Tabellenzählung. Anfang/Ende messen, Recoverypunkt/RPO und Wiederanlaufzeit/RTO mit Betreiberziel vergleichen. Restorefehler machen das Gate rot.

## G6 – Go/No-Go vor späterem Wartungsfenster

Alle Objekt-/Rechteabweichungen fachlich geschlossen; kein UNVERIFIED bei sicherheitswirksamen Abhängigkeiten. Aktueller SHA inklusive aktueller vollständiger Browser-/Visual-/Native-/Originalrestore-Abnahme, geprüfter Migrations- und Rückfallhash, Backup/Restorebelege, Verantwortliche und Kommunikations-/Abbruchplan müssen vorliegen. Sicherstellen, dass keine zweite Appversion/Worker parallel unter alten Verträgen schreibt. Branchschutz und Pages-Quelle/Deploymentzustand aus #25 administrativ nachweisen. Betreiber entscheidet explizit über RPO/RTO-Restabweichungen; sie werden nicht durch eine Testbehauptung akzeptiert.

## G7 – nur separat genehmigte spätere Durchführung

Neuer ausdrücklicher Auftrag erforderlich. Vorfenster: Backup-/PITR-Nachweis prüfen; sichere Wartungsoberfläche und Schreibruhe; exakte Zielref/Versions-/Katalogpreconditions nochmals read-only prüfen. Nur das abgenommene Paket anwenden. Bei Abweichung/Locktimeout/fehlender Authority abbrechen. Vor erneuter Öffnung minimale authentifizierte Funktions-/Negativchecks und Integritäts-/Auditabgleich durchführen, Cache-/Token-/Workerumstellung koordinieren. Autorisierte Betreiber prüfen produktive Integrität ohne Datenexport in Chat/CI. Keine freigegebene Produktion allein aufgrund erfolgreicher DDL.

## G8 – kontrollierter Rückfall

Vor Commit: Transaktion abbrechen und Rollback/Postconditions prüfen. Nach Commit: Zugriffe/Schreibpfade sicher schließen und Incidentverantwortlichen einschalten; bevorzugt getestete vorwärts gerichtete Korrektur unter weiter restriktiven Rechten. **Kein automatisches Wiederherstellen der unsicheren true-/Anon-/breiten Authenticated-Policies, auch nicht aus altem Backup.** Die alte App mit offenen Rechten ist kein zulässiges Rückfallpaket.

Ist Recovery nötig, zuerst in isoliertes Ziel zum genehmigten Recoverypunkt wiederherstellen; sichere Policy-/Rechtehärtung vor Verbindung zu Clients anwenden und prüfen. Seit Backup entstandene Aufträge, Mengen, Werkzeugbewegungen und Auditereignisse über freigegebenes kontrolliertes Journalverfahren abgleichen; keine stille Löschung oder Doppelbuchung und kein pauschaler Restore über neuere Produktionshistorie. Unveränderliche Originalbackups und Incident-/Auditbelege aufbewahren. Wenn sicherer Rückfall nicht nachweislich möglich ist, bleibt die Anwendung im Wartungszustand, bis die verantwortliche Stelle entscheidet.

## Quellen und heutige Grenzen

PostgreSQL 17: https://www.postgresql.org/docs/17/app-pgdump.html und https://www.postgresql.org/docs/17/app-pg-dumpall.html . Supabase: https://supabase.com/docs/guides/platform/backups und https://supabase.com/docs/guides/platform/migrating-within-supabase/backup-restore . Abruf 12.11.2026. Versions-/Plattformunterschiede und verwaltete Schemas müssen beim echten Restore geprüft werden. Diese Vorlagen wurden **nicht** auf der Originaldatenbank ausgeführt. Dieser Plan ist überprüfbar und gateweise konkret; ein belastbares ausführbares Produktions-DDL bleibt mangels Originaldefinitionen/Fachfreigaben absichtlich gesperrt.
