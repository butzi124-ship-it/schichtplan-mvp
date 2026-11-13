# Schritt 8B Phase 1 – isolierte Identitäts- und Berechtigungsgrundlage

**Keine Produktivfreigabe, keine Aktivierung.** Basis PR #34 / `66795417d3a940bcb1273b46113f7f45b5b3a8e6`, Version 0.16.6; Sicherheitsbasis PR #32. Neuer Branch `security/identity-foundation-8b1`, Version **0.17.0**. Refs #33 und [verbindliches Konzept](../architecture/modular-permissions-8a.md). Kein bestehender PR wird umgeschrieben. Haupt-App ausschließlich Version/Versionslog, alle anderen Module/Flags/Tests/Bildreferenzen unverändert.

## Vorprüfung und vorhandene Sicherheitsarbeit

Am 13.11.2026 Remote-Refs neu abgerufen und Heads #28–#34 geprüft: #28 dd64a8e…, #29 2e46d3a…, #30 35a56bb…, #31 1238bf0…, #32 67ac80a…, #34 6679541…. Kein Rebase/Merge. Die vorherigen SQL-Pakete und ihre Hashmanifeste bleiben unverändert. Keine neue produktive Metadaten- oder Datenabfrage; Originalkatalog weiterhin datierter Nachweis vom12.11., kein vollständiger Originalrestore.

`humbel_security.actor_id()` aus #28–#32 ist bereits eine unabhängig registrierte Auth-/Mitarbeiterbindung mit Aktivitäts-/Legacyrollen-/Primärabteilungsprüfung. Diese wird **wiederverwendet**, nicht durch ein lokales Profil oder einen neuen offenen Login ersetzt. `js/secure-access.js` bleibt unverändert deaktiviert. Bestehende Personal-/Abteilungs-/Planer-/Werkzeug-/BA-/OP-Pfade werden nicht neu implementiert oder auf diesen Kandidaten umgestellt.

## Implementierter, nicht angebundener Umfang

SQL nur unter `supabase/review/permissions-8b1-isolated.sql`, **nicht** unter registrierten Migrationen. Zielguard: ausdrücklich benannte isolierte Datenbank bzw. PGlite mit Sentinel; zusätzlich vorhandenes geprüftes Step7-Paket. Der netzwerkfreie Reviewrunner bekommt eine bereits isolierte Datenbankinstanz, liest keine URLs/Passwörter und stellt keine Verbindung her. Keine Deployment-/App-/Workfloweinbindung.

Alle neuen Objekte liegen in **humbel_permissions**; keine vorhandene Definition im Schema public oder humbel_security wird ersetzt. Alle Tabellen haben RLS ohne Clientpolicies; PUBLIC/anon/authenticated/service_role erhalten keine Schema-, Tabellen-, Spalten- oder Funktionsrechte. Kein neuer PostgreSQL-Superuser/Service-Role-Client. Eine fehlende externe API-Anbindung ist eine beabsichtigte Sperre, kein umgangener Test.

| Vorbereitetes Objekt | Zweck / sichere Grenze |
|---|---|
| runtime | Ein Singleton, enabled=false; zusätzlich sind alle APIrechte geschlossen. Kein aktiver Produktionsschalter |
| roles | Exakt sieben nicht hierarchische Namen: executive (GL), production_manager (PL), department_head (AL), deputy (STV), hr, employee (Mitarbeiter), technical_admin (TA). **Keine Zuweisungen**; Scanner bleibt technischer Legacyfall |
| identities | Auth-ID primärer FK zur geprüften Registry, eindeutige Mitarbeiter-ID/FK; enabled=false, blocked, Bindungs-/Berechtigungs-/Primärversion, Gültigkeit/Evidenz. Keine automatische Übernahme eines alten admin |
| sessions | Leere unabhängige Sessionfreigabeliste: Session-ID, Userbindung, Bindungs-/Berechtigungsversion, Frist, Widerruf, Evidenz. Keine Tokens/Passwörter, kein behaupteter Restore von auth.sessions |
| memberships | Eindeutige Mitarbeiter-/Abteilungsrelation mit Gültigkeit/Aktivität/Widerruf. Mitarbeiter ohne Login können organisatorisch zugeordnet werden; keine automatischen Rechte |
| organization_roles | Mehrere ausdrücklich zugewiesene Rollen möglich, Zeitraum/Widerruf/Evidenz. AL/STV brauchen konkreten Department-FK; keine numerische Rangfolge |
| leaderships | Getrennte konkrete Leiterzuweisung mit zusammengesetztem FK auf Person/Rolle/Abteilung. Eine Person kann mehrere Bereiche gleichberechtigt leiten. Vorläufig ein aktueller Leiter je Bereich gemäß fortgeltender Eindeutigkeitsregel/O03 |
| capabilities | **Leerer** Aktionskatalog, keine beschlossenen Fachprofile vorgespiegelt. Reservierte Organisations-/Identity-/Berechtigungs-/Protokoll-/Aufbewahrungsnamespaces nicht als Modulrechte verfügbar. Scopevertrag muss ausdrücklich gewählt werden |
| module_grants | **Leere** explizite Person-/Rollen-/Modul-/Aktions-/Abteilungsfreigaben, disabled default; Primaryepoch, Frist auch unbefristet, Widerruf/Evidenz. Kein universelles Adminrecht |
| installation | Privates Reviewjournal für bekannte Version, Quellhash, unveränderte Basis und Definitionsfingerprint; kein produktives Migrationsjournal |

Kein Rollenbootstrap, keine GL-/PL-Ernennungs-/Widerrufs-RPC, keine Stellvertretungsfreigabe, keine Vergabeoberfläche, keine Protokollkorrektur oder Aufbewahrungs-/Löschfunktion. Die Regeln aus 8A bleiben verbindlich, werden aber nicht durch ungeklärte Enrollmentannahmen halb implementiert.

## Serverseitiger Prüfvertrag

Zwei **rein lesende private** Funktionen: `actor_context()` ohne Actorparameter und `can_perform(module,action,department)` mit lediglich Zielaktions-/Ressourcenscope. Keine übermittelbare Mitarbeiter-ID/Rolle als Authority. Eindeutige, aktive, freigegebene Identität **UND** unveränderte geprüfte Legacybindung **UND** passende aktive Sessionfreigabe/-Version **UND** Runtimegate sind nötig. Fehlende, gesperrte, abgelaufene, widersprüchliche oder anonyme Identität liefert keinen Kontext. Ungültige Session-/Anonymitätsformate verweigern; sonstige technische Fehler werden nicht als Allow kaschiert.

Zusätzlich für eine Aktion: aktueller exakter Modul-/Aktions-/Abteilungsgrant, aktiver Bereich, Gültigkeit/Widerruf/Primaryepoch und passende aktive Organisationszuweisung. Mehrdeutige gleichzeitig passende Rollen-/Grantbelege verweigern. Bei AL muss die konkrete Leitungszuweisung stimmen. STV-Ausübung bleibt bewusst **false**, bis Abwesenheitsquelle und Freigabevertrag O04/O05 umgesetzt/abgenommen sind. HR und technischer Admin erhalten keine eingebauten Fach-/Privatdatenrechte. Rollen allein oder Mitgliedschaft allein reichen nie.

`scope_rule=membership` verlangt die explizite aktuelle Membership; `explicit_department` verlangt stattdessen den expliziten Ressourcengrant. Dies ist **kein allgemeiner Membershipfilter über bestehende BA-/OP-Berechtigungen**. Der Katalog ist leer; welcher Fachfall welche Variante verwendet, bleibt O06/O08/O15 und eigener späterer Review. Aktuelle Fachmodule benutzen keine dieser Funktionen und behalten ihre bisherigen Zustands-/Teilnehmer-/Mengen-/Reviewerregeln.

Zwei eng begrenzte SECURITY-DEFINER-Lesefunktionen sind notwendig, damit bei späterer geprüfter Exposition keine Authoritytabellen an Clients freigegeben werden müssten. Sie sind im isolierten Paket operator-owned, mit leerem search_path, qualifizierten Tabellen-/Auth-/Helperreferenzen, ohne dynamisches clientabhängiges SQL und ohne API-EXECUTE. **Vor jeder Exposition** sind dedizierter minimaler Owner, RLS-/Owner-/EXECUTE-/API-/GraphQL-/Sessionprüfung erforderlich. Kein Anspruch auf Vertraulichkeit gegen einen tatsächlichen DBA/Superuser. Tests erweitern Schema-/EXECUTE-Rechte ausschließlich temporär in synthetischen Fixtures; diese GRANTs stehen nicht im Kandidaten.

## Datenintegrität, Wiederholung und Transaktionen

Keine Kopie/Änderung von employees.department_id oder departments.leader_employee_id; Stammabteilung bleibt einzige organisatorische Primärquelle. Keine produktiven Personenwerte, Namen oder Login-/Rollenzuweisungen eingelesen/backgefüllt. Keine kaskadierenden Löschungen von Personal/Fachhistorie.

Neue SQL läuft transaktional. Reviewrunner prüft zuvor die unveränderte Step7-Basis samt vorhandenen Schutzpostconditions und Fingerprint; nach Installation nochmals bestehende Definitionen/ACLs, neue Defaultsperren und leere Zuweisungstabellen. Nur bekannte identische Wiederholung mit gleichem Quellen-/Versions-/Definitionszustand ist No-op. Drift, unbekannte Teilinstallation oder falscher Zielguard bricht ab; keine automatische Reparatur/Überschreibung oder Aktivierung. Quelle ist im **neuen** Manifest gepinnt; keine bisherigen Schutzhashes aktualisiert.

Die Lesefunktion ist **keine** Autorisierung zukünftiger Schreibtransaktionen: sie darf nicht einmal vor einem separaten ungesicherten Write geprüft werden. Mutations-/Lifecyclekommandos müssen später Actor-/Authority-/Zielversionen sperren, nach Wartezeit neu prüfen und Fachaktion/Audit atomar ausführen. Mehrverbindungstests vor Aktivierung zwingend.

## Bewusste Freigabeblocker und Phase 2

1. **O01/O02/O11/O12:** unabhängiges Enrollment/GL-Bootstrap, Organisationsscope, Mehrfachrollen-/Konfliktregeln; keine Self-Service-Registry oder Rollenvergabe vorhanden. Bindungsänderungen nur im vertrauenswürdigen Operatorpfad, noch kein geprüfter Produktivoperator.
2. **Session-Lifecycle:** Liste ist unabhängig/leergestartet, kein GoTrue-/auth.sessions-/JWT-Signatur-/Refresh-/Logout-/Realtime-Synchronisierungsnachweis. Claims sind erst hinter einem tatsächlich validierenden Gateway vertrauenswürdig. Freigabe, Widerruf, Epochanhebung und Reaktivierung nach O09 fehlen als sichere Kommandos. Bereits gelieferte Daten sind nicht rückrufbar.
3. **Stammwechsel/Handover:** Snapshot-/Primaryepoch verweigert veraltete Kontexte, aber atomare Suspendierung **aller** Alt-/Sondergrants, Schutz vor Wechsel-zurück-Reaktivierung und Reviewhistorie müssen vor erster Aktivierung gebaut werden. Keine Ablösung bestehender strenger Scalar-Guards in dieser Phase.
4. **O03–O08/O10/O13–O18:** Co-Leitung, Abwesenheit/STV-Freigabe, Modul-/Feld-/Scope-/Privacyprofile, Review/Widerruf, PL-Nachricht und Protokoll-/Fristfragen bleiben entsprechend 8A offen. Mehrfachzuordnung wird gespeichert, nicht als vollständiger Organisations-/Delegationsworkflow ausgegeben.
5. **Bestehende Risiken:** 31-Tabellen-/Default-/CRUD-/Definer-/Privat-/Altcache-/Log-/Sessionprobleme aus #32 sind nicht durch diese zusätzliche Grundlage geschlossen. Keine unsicheren Altzugriffe als neues Allowfallback verwenden.
6. **Native und Betrieb:** vollständiger Originalrestore, echte PostgreSQL/GoTrue/PostgREST-/RLS-/API-/Parallel-/Sessiontests sowie Backup/PITR/RPO/RTO/Recovery, Owner-/API-Minimierung und separate Aktivierungsfreigabe erforderlich. Ein PGlite-Spiegel ist kein Originalrestore.

Empfehlung Phase 2: zuerst sichere, auditierte und idempotente Identitäts-/Session-Sperr-/Versionskommandos mit echter isolierter Auth-/REST-Laufzeit und Transaktionsrennen nachweisen. Danach bestätigte Membership-/Primärwechselkommandos samt atomarer Sonderrollensuspendierung. Ernennungen/Modulprofile nur nach ihren noch offenen fachlichen Voraussetzungen. Kein Appcutover, bevor ein vollständiger Modulvertrag und dessen Positivregression abgenommen sind.

## Nachweise

Aktuelle Ergebnisse und blockierte Prüfungen werden in [permissions-8b1-validation.json](permissions-8b1-validation.json) dokumentiert. Die neuen SQL-/Rollenprüfungen verwenden ausschließlich synthetische Daten im wegwerfbaren PGlite, nicht Supabase-Token oder Produktivdaten. REST-äquivalente direkte SQL-Angriffe sind **keine tatsächlich ausgeführten HTTP-/GoTrue-Tests**. Keine Produktivaktivierung aus isoliert grünem Ergebnis ableiten.

Technische Referenzen, am13.11.2026 geprüft: https://supabase.com/docs/guides/database/postgres/row-level-security und https://www.postgresql.org/docs/17/sql-createfunction.html. Supabase-Changelog geprüft; kein CLI-/Paket-/Engineupgrade in diesem Paket.

## Zusätzliche Integrationsgrenzen des vorbereiteten Speichers

Memberships/Leitungszeiger enthalten in dieser Phase nur den aktuellen vorbereiteten Zustand; ein vollständiger auditierter Historien-/Ernennungsworkflow wird nicht behauptet. Mehrdeutige überlappende Rollen-/Aktionsbelege werden bei der Prüfung abgewiesen, nicht automatisch vereinigt. Vor einer späteren Vergabe-API sind zeitliche Überlappungsconstraints, explizite Rollen-/Profilkombinationen und unveränderliche Evidenz-/Änderungshistorie abgenommen umzusetzen. Neue FKs verhindern das harte Löschen referenzierter Mitarbeiter; der spätere Personal-Lifecycle muss deshalb Deaktivierung/Datenschutz-/Aufbewahrung nach O09/O10 koordinieren, bevor ein Appcutover erlaubt ist. Im ausgelieferten System ist nichts davon installiert.

Die neun additiven Testgruppen decken Defaults/ACL/No-op, kein implizites Admin-/Rollenrecht, Auth-/Session-/Identitätsmanipulation, Deaktivierung/Epochen/Widerruf, gleichberechtigte Mehrfachleitung, getrennte zeitliche Rollen-/Grant-/Membershipprüfung, direkte SQL-/Namensspoofingangriffe, Datenerhalt/Batchrollback sowie atomare Installation/Fault-Injection/Drift ab. Die Baseline-Reproduktion aus früheren PRs bleibt getrennt und wird nicht als neuer Fix ausgegeben.

Während der Entwicklung wurden ein Syntaxfehler im **neuen** Zielguard und der fehlende Connection-Sentinel nach dem Klonen einer **synthetischen** Testdatenbank korrigiert. Keine Guards oder Assertions gelockert; neue Negativtests bleiben erhalten, vorhandene Tests/Hashmanifeste unverändert. Der neue Paketmanifesthash beschreibt ausschließlich die überprüfte neue SQL-Datei. Frühere fehlgeschlagene Entwicklungsversuche sind keine grünen Läufe.

## Tatsächlicher lokaler Abschluss

Bestanden: npm ci --ignore-scripts; Pflicht-/neue JS-Syntax; test:ci 2/2, test:qa 20/20, qa:check, test:integration 5/5, **test:security 64/64** (55 bestehende + 9 neue). Danach wurde die zusätzliche Step7-Pakethashprüfung ergänzt und die **finale neue Suite 9/9** vollständig erneut bestanden; die 55 Bestandsfälle/Quellen blieben unverändert. Neun Fachmodulsuiten **168/168**, qa:agent-server, Dokument-/Manifest-/Änderungsscopeprüfung und git diff --check ebenfalls bestanden. Keine Skips/Assertion-/Referenzlockerung.

**BLOCKED_NOT_TESTED:** E2E vollständig gestartet, nach **94 Browserstartfehlern** und Nachweis fehlender Chromium-/Firefox-/WebKit-Binaries beendet (Exit130); Restfälle nicht ausgeführt, keine 773 bestandenen oder vollständig abgearbeiteten Fälle behauptet. Agentenseed am fehlenden Browser blockiert; Visual-Guard Exit2, DB-Preflight Exit2. Neue echte native Rollen-/JWT-/API-/Mehrverbindungs-/Restoreabnahme nicht verfügbar. Aktuelle Einzelwerte/Loghashes und Quellenhashes im Prüfprotokoll. **NEIN: keine Freigabe für Produktivmigration oder Aktivierung.**

## Fortsetzung: Schritt 8B Phase 2

Der historische Phase1-Nachweis bleibt erhalten. Der [Phase2-Kandidat](permissions-8b2.md) ergänzt ab Version0.17.1 ausschließlich isolierte, deaktivierte Verwaltungsoperationen und technischen Audit. Echte Auth-/JWT-/API-/Parallelabnahme und produktive Aktivierung werden dadurch nicht als erledigt ausgegeben.
