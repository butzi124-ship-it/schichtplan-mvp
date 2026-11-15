# Schritt 8B Phase 2 – isolierte Identitäts- und Sitzungsverwaltung

**Version 0.17.1. Keine Produktivfreigabe oder Aktivierung.** Basis PR #35, `security/identity-foundation-8b1`, Commit `c5384669a17b11d7d97e42f3bd9dafd27df5fe16`, Version0.17.0. Neuer Branch `security/identity-sessions-8b2`. Refs #33/#34 und Sicherheitsbasis #32. Die vollständigen [8A-Konzepte](../architecture/modular-permissions-8a.md) und der [Phase-1-Bericht](permissions-8b1.md) bleiben maßgeblich. Keine bestehenden PRs überschrieben/zusammengeführt, keine produktive Datenbankverbindung oder Geschäftsprozessänderung.

## 1. Geprüfte Basis und Abgrenzung

PR #35/#34/#32 und Remote-Refs am13.11.2026 geprüft: c5384669… /66795417… /67ac80ac…. Unveränderter Phase1-Installer und seine gepinnten vorherigen Kandidaten werden vor Installation verifiziert. Kein neuer Originalschema-/Authkonfigurationsnachweis behauptet; #32-Katalog bleibt datierter Metadatenbefund vom12.11., kein vollständiger Originalrestore. Keine neue produktive Supabase-Abfrage oder Nutzung von Produktionsdaten.

Phase1 wird **nicht erneut implementiert**: identities, sessions, Rollen, Memberships, Leitungen und Grants bleiben dieselben Objekte. Die neuen Kommandos ändern ausschließlich dafür vorgesehene private Zustände. Kein App-/HTML-/Frontend-/Featureflag-/Workflowcutover; Haupt-App nur APP_VERSION/VERSION_LOG. Mitarbeiter-/Auth-ID-Bindungen, public-Rollen/Abteilungen, Schicht-/Werkzeug-/BA-/OP-Funktionen bleiben unverändert.

## 2. Vertrauensgrenze statt ungeprüfter JWT-Autorität

**In diesem Kandidaten gibt es bewusst keine öffentlich nutzbare Supabase-Auth-Integration.** Die neue Verwaltungsautorität stammt aus einer explizit unabhängig freigegebenen **SQL-Verbindungsidentität**: session_user → OID/Name des tatsächlichen DB-Logins → vorbereitete aktive Mitarbeiteridentität → zeitlich und versionsgebundene Ziel-/Aktionsfreigabe. Kein frei übermittelter Actor, keine Rollenentscheidung aus employees.role, auth.uid(), auth.jwt(), user_metadata, app_metadata oder einem übergebenen Benutzer-/Session-JSON.

Die SQL-Logins müssen LOGIN, aber NOSUPERUSER/NOBYPASSRLS/NOCREATEROLE/NOCREATEDB haben. OID und Name müssen zur Freigabe passen; Namenswiederverwendung nach Rollen-Neuanlage reicht nicht. Shared-API-/Plattformrollen wie authenticator, anon/authenticated/service_role/postgres/supabase_admin sind ausgeschlossen. SET ROLE wird nicht als Verwaltungsidentität akzeptiert. Die Rolle des Funktionsowners (current_user im SECURITY-DEFINER-Kontext) ist ebenfalls kein Actorbeweis.

**O09 und O01 sind damit nicht entschieden.** Wer reale Verwaltungsfreigaben/erste privilegierte Organisationsrollen erteilen darf, bleibt offen. Der Kandidat enthält **null Operatorbindungen, null Scopes, null Sessionnachweise und keinerlei Enrollment-/Ernennungsfunktion**. Ein technischer App-Administrator, vorhandener Legacy-admin, HR, GL oder PL bekommt daraus keine automatische Verwaltungs-/Fachvollmacht. Test-Operatoren und deren genau begrenzte Freigaben sind ausschließlich synthetische Fixtures.

Freie Ziel- und Session-UUIDs sind nur **Ressourcenparameter**, nie Autorität. Der tatsächliche SQL-Actor braucht für jede konkrete Zielidentität die ausdrücklich passende Aktion. Eigenfreigabe ist ausgeschlossen. Der SQL-Verbindungsweg erfordert später unabhängige native Login-/Pool-/Ownerprüfung; die PGlite-Simulation ist kein Netzwerk-Anmeldenachweis und keine Sicherheitsgarantie gegen einen tatsächlich privilegierten DBA.

## 3. Neue Reviewobjekte

Alle Objekte ausschließlich im neuen privaten Schema **humbel_identity_admin**, SQL unter `supabase/review/permissions-8b2-isolated.sql`, nicht unter registrierten Migrationen. APIrollen/PUBLIC erhalten weder Schema-/Tabellen-/Spalten-/Funktionsrechte, noch gibt es Client-RLS-Policies. RLS auf allen neun Tabellen einschließlich Installerjournal.

| Objekt | Funktion / Schutz |
|---|---|
| runtime | Eigenes Gate, **enabled=false**. Management zusätzlich verweigert, falls der bisherige Phase1-Clientkontext eingeschaltet wäre; kein Legacy-Allow-Fallback |
| operators | **Leere** unabhängige DB-Loginbindung an Actoridentität, OID/Name, enabled=false, Zeit/Widerruf, Bindungs-/Berechtigungsversion und Evidenz-UUID |
| scopes | **Leere** explizite Operator-/Ziel-/Aktionsmatrix, kein NULL-/Wildcardbereich, enabled=false, Zeit/Widerruf/Evidenz |
| session_attestations | **Leerer** Nachweisspeicher eines später separat geprüften Auth-Verifiers: Session-/User-ID, Bindungs-/Berechtigungsversion, endliche Gültigkeit, Evidenz-UUID. Kein Token, Passwort, JWT-Payload oder medizinischer Inhalt |
| identity_versions | Optimistische Verwaltungsrevision pro vorbereiteter Identität; initial0 ohne vorhandene Zeile, jede angewandte Änderung +1 |
| requests | Unveränderliche Request-ID, tatsächliche Autorität, kanonische strukturierte Parameter und Ergebnis; keine allgemeinen Freitextpayloads |
| session_tombstones | Unveränderliche Widerrufs-ID je Session mit Ziel, Zeitpunkt und deferrierbarem FK auf angewandten Request; verhindert Wiederzulassung derselben ID |
| audit | Unveränderlicher technischer Änderungsnachweis mit Actor, Ziel, Aktion, Session-ID falls relevant, Serverzeit, Versionen/Epochen, Autoritäts-/Scope-Evidenz und Ergebnis |
| installation | Geprüfte Paket-/Basis-/Definitionshashes, kein produktives Migrationsjournal |

Sechs begrenzte Funktionen: **command**, **server_context**, interne **authority**, **binding_ok**, **live_session** sowie Triggerfunktion **immutable_event**. Rein private, parameterarme Implementierung; keine allgemeinen SQL-/Tabellenpatch-/Rollenvergabe-Endpunkte. SEARCH_PATH leer, qualifizierte Tabellenreferenzen, Lockwartezeit3s für die sperrenden Funktionen. Owner ist im isolierten Paket weiterhin der geprüfte Installeroperator; **minimaler dedizierter Owner und dessen echte Berechtigungsabnahme sind vor Exposition zwingend offen**. Kein neuer PostgreSQL-Superuser oder Service-Role-Client.

## 4. Implementierte Kommandos und Sitzungsschutz

| Aktion | Atomare Wirkung |
|---|---|
| identity_activate | Verifizierte bestehende Auth-/Employeebindung und aktive Grunddaten nötig; private Identität enabled/entsperrt, Berechtigungsepoche erhöht. Alle bisherigen Sessions widerrufen/tombstoned, bisherige Grants deaktiviert/widerrufen und Organisationszuweisungen deaktiviert. **Keine Wiederbelebung oder Vergabe von Rollen** |
| identity_deactivate | Identität gesperrt/deaktiviert, Epoche erhöht; Sessions/alte Grants gesperrt, Audit und Receipt im selben Commit |
| epoch_bump | Aktuellen Aktivitätszustand erhalten, Epoche erhöhen und bisherige Sessions/Grants ungültig machen. Kein neues Recht |
| session_allow | Nur aktive verifizierte Zielidentität und bereits **unabhängig gespeicherter, unveränderlicher** Sessionnachweis mit aktuellen Bindungs-/Berechtigungsversionen, passendem User und nicht abgelaufenem Prüfzeitraum. Freigabe höchstens bis zum nachgewiesenen endlichen Ende. Keine Wiederverwendung einer existierenden oder tombstoned ID |
| session_revoke | Nur bekannte eigene Session der freigegebenen Zielidentität, dauerhaftes Tombstone und revoked=true; Audit/Version atomar |
| server_context | Interne serverseitige erneute Ziel-/Session-/Versions-/Widerrufsprüfung unter Sperren. Rückgabe nur IDs/Versionen, kein Fachrecht, Personalprofil oder Token. Kein Browser-/HTTP-Login-Endpunkt |

Der Sessionnachweis kann **nicht** durch command aus Claims erzeugt werden; auch der Test-Operator hat keine direkten Tabellen-Schreibrechte. Neue Ops verwenden überhaupt keine JWT-Claims als Verwaltungsbeweis. Verfälschte Benutzer-/Rollen-/Sessionclaims ersetzen weder SQL-Loginbindung noch Scope oder Nachweis. Bereits vorhandene Phase1-Sessionzeilen erhalten nicht automatisch eine Attestation.

Widerruf/Epochenwechsel macht sowohl freigegebene Sessions als auch **noch ungenutzte alte Sessionnachweise** unwirksam. Eine Aktivierung führt nicht zu erneuter Freigabe alter Rollen oder Session-IDs. Tokenrefresh/erneute Freigabe einer weiterhin bestehenden tatsächlichen GoTrue-Session ist noch kein implementierter Produktivvertrag; native Lebenszyklus-/Issuerentscheidung bleibt erforderlich.

**Keine vorgetäuschte JWT-Prüfung:** Der Attestationsspeicher ist ein geschützter Vertrag für einen späteren vertrauenswürdigen Verifikationspfad, kein implementierter kryptographischer GoTrue-/JWT-Verifier. Es gibt keinen produktiven Schreiber/Adapter dafür. Ohne ihn gibt es keinen echten Sessionfreigabenachweis. Eine gehashte oder behauptete JWT-Payload würde diesen Nachweis nicht ersetzen. Der bisherige Phase1-JWT-Kontext bleibt gesperrt und wird nicht als Umgehungsweg aktiviert.

## 5. Transaktionen, Sperren und Wiederholungen

Schreiboperationen prüfen Authority und Gates **innerhalb** der geschützten Transaktion, sperren Operator/Scope und betroffene Actor-/Zielidentitäten in fester UUID-Reihenfolge und prüfen die Autorisierung nach Wartezeiten erneut. Vor Aktivierung werden Grundbindung/Rootregistry/Employeeprofil gesperrt geprüft, vor Sessionfreigabe der Nachweis. Sessionwiderruf sperrt die konkrete Session. Alle Gültigkeiten werden mit aktueller Serverzeit geprüft, nicht aus einem vor dem Warten berechneten Clientkontext.

Unterstützter Vertrag ist **READ COMMITTED**, andere Isolationsebenen werden ausdrücklich verweigert statt alte Snapshots als Autorisierung zu verwenden. Request-ID/identischer Payload erlauben eine geprüfte Wiederholung ohne doppelte Änderung/Auditzeile. Fremder Actor oder geänderter Payload bei derselben ID wird abgewiesen. Falsche Verwaltungsrevision führt zu Versionskonflikt. Wiederholung eines früheren session_allow nach Widerruf/Ablauf verweigert; alte Successreceipts sind keine wiederverwendbaren Rechte. Historische Wiederholungsantworten enthalten auch die aktuelle Revision.

Jede Mutation, Epoche/Revision, Rollen-/Sessionentwertung, Receipt und Audit bilden **eine** Transaktion. Ein injizierter Auditfehler rollt alles zurück. Kein freier Callback oder späteres ungesichertes Fachschreiben. Die neue Verwaltungsautorisierung darf nicht einmalig vor einer anderen Transaktion geprüft und dann als Berechtigungstoken weitergegeben werden.

Präzise Grenze: Ein vorher berechtigter, gesperrt laufender Kontext darf vor dem Widerrufscommit enden; neue Kontexte nach dem Commit müssen verweigern. **Echte Parallelverbindungen sind hier nicht nachgewiesen.** Alle späteren Schreiber/Verifikationsadapter müssen dieselbe Sperr-/Versionsordnung einhalten; tatsächliche Auth-Revocation, lange Lesetransaktionen, Ablauf während Sperrwartezeit und Idle-/Statementlimits gehören in die native Abnahme. Locktimeout allein begrenzt nicht eine vom Aufrufer offen gehaltene Transaktion.

Bestehende App-/Legacy-Aktivitäts-/Abteilungsänderungen werden nicht heimlich umgeleitet. Vor Produktivaktivierung müssen alle relevanten Lebenszyklus-Schreibwege koordiniert oder gleichwertig geschützt sein; andernfalls bleibt der Cutover gesperrt.

## 6. Audit und Datenschutz

Technischer Audit ist **kein fachliches Vertretungsprotokollkorrekturrecht** aus8A. Weder technische Administration noch Operatorlogin erhält dadurch HR-/Produktions-/Werkzeug-/Gesundheitsleserechte. Keine entsprechende View oder Audit-Gesamtauskunft veröffentlicht.

Angewandte Änderungen protokollieren unveränderlich nachgewiesenen Actor/Autoritätsnachweis, Ziel, Aktion, Serverzeit, Vorher-/Nachherrevision und Epoche, Ergebnis applied sowie gegebenenfalls die Session-ID. Receipts/Sessionnachweise/Tombstones/Audit lehnen UPDATE, DELETE und TRUNCATE durch Trigger ab; API-ACLs sind zusätzlich geschlossen. Korrekturen/Entfernung/Aufbewahrung sind keine hier implementierte Administrationsfunktion. Ein unbeschränkter DBA kann DDL/Trigger ändern; dessen Betriebskontrolle bleibt unabhängig erforderlich.

Es werden **keine Passwörter, Tokens, JWT-Rohdaten, medizinischen Inhalte oder freie Client-Gründe** gespeichert. Payloads bestehen aus zulässigem Aktionscode, UUIDs, Revision und numerischem Ablaufzeitpunkt. Evidenz ist UUID-Referenz, nicht der Beleginhalt.

Bei einem SQL-Fehler wird auch ein innerhalb derselben Transaktion geschriebener Audit zurückgerollt. Es gibt deshalb **keinen behaupteten dauerhaften SQL-Audit sämtlicher abgewiesener Versuche**. Ein separater vertrauenswürdiger, datensparsamer Ablehnungs-/Betriebs-Auditsink ist vor entsprechender Produktivanforderung noch zu entwickeln. Jede tatsächlich angewandte administrative Änderung muss ihren unveränderlichen Audit im selben Commit haben.

## 7. Reproduzierbarkeit, Regression und feste Laufzeitgrenzen

`permissions-8b2-review.js` validiert den unveränderten Phase1-Installer und alle vorherigen gepinnten Sicherheitsquellen, wendet nur das neue additive Reviewpaket an und prüft unveränderte Basisdefinitionen/ACLs, neue Sperren/leere Trusttabellen sowie Versions-/Definitionsfingerprint. Ein vorhandener Phase1-Installationsnachweis ist zwingend; der Runner darf Phase1 nicht implizit neu installieren. Nur bekannte identische Wiederholung ist No-op. Drift/Teilinstallation/unerwartete Herkunft bricht ab. Bisherige SQL-Dateien und Schutzmanifeste bleiben unverändert; neues Manifest beschreibt ausschließlich neue geprüfte SQL.

Der eigenständige lokale Linux-/POSIX-Supervisor `tests/db-integration/permissions-8b2-verify.py` liest die tatsächlichen package-Skripte, protokolliert Prozess-ID/Deadline/Exitcode und setzt je Prozess feste Grenzen (Syntax30s, QA60–120s, Integration/Fachmodule180s, Security900s), insgesamt1800s. Bei Timeout beendet er die gesamte Prozessgruppe mit TERM und nötigenfalls KILL, ohne automatische Wiederholung; nicht fertig geprüfte Fälle werden BLOCKED_NOT_TESTED. Neue einzelne Node-Testgruppen haben zusätzlich300s, gezielter Gesamtlauf maximal300s. Browser-/Container-/DB-Voraussetzungen werden **vor** dem Start geprüft; fehlende Voraussetzungen führen zu BLOCKED_NOT_TESTED ohne aussichtslose Wiederholungsschleifen. Keine Skipmarker/Testfilter/Assertion-/Referenzlockerung.

Alle Testdaten sind synthetisch. SQL-Sessionwechsel im PGlite-Test erfolgt durch den isolierten Superuser, nicht durch echte Netzwerk-Anmeldung. PGlite stellt beim RESET SESSION AUTHORIZATION hier nicht wie erwartet zurück; die Fixture stellt explizit den ursprünglichen Testowner wieder her. Das ist eine Testumgebungsgrenze, keine gelockerte Produktivprüfung. Nach Versionsbindung wurde ein veralteter INSERT im neuen Negativtest auf die gültige Spaltenform korrigiert, damit der unveränderte erwartete Berechtigungsfehler tatsächlich geprüft wird. Frühere rote Läufe bleiben im Prüfprotokoll sichtbar.

## 8. Noch bewusst gesperrt / Empfehlung Phase3

1. **O09/O01/O02/O11/O12:** fachlich genehmigte Operator-/Lifecycle-/Enrollmentbefugnisse, erste privilegierte Organisationsrollen, Scope und Konfliktregeln. Keine Quelle darf aus alten admin-Profilen automatisch GL/PL/TA ableiten. Operatorprovisionierung/-änderungen und ihre unabhängige Auditierung sind nicht als freigegebener Produktionsprozess implementiert.
2. **Verifikationsadapter:** echte Token-/Signatur-/Issuer-/Audience-/Expiry-/Session-/Widerrufsprüfung und vertrauenswürdiges Erzeugen/Widerrufen von zeitlich und versionsgebundenen Attestationen. Keine Exposition als allgemeines PostgREST-RPC unter gemeinsamer authenticator-/Service-Role-Verbindung. Aktuelle Quelle lehnt diesen Weg ab.
3. **Native Abnahme:** PostgreSQL-Mehrverbindungen mit echten nichtprivilegierten Logins, GoTrue/PostgREST/RLS/JWT/Refresh/Logout/Realtime und Widerrufsraces. Restore mit tatsächlichen Ownern/ACLs/Funktionen/Rollenvererbung; minimale Owner und realer Privilegienumfang noch nicht abgenommen. Kein vollständiger Originalrestore vorliegend.
4. **Historische Risiken:** übrige31-Tabellen-/CRUD-/Default-/Definer-/Altcache-/Privat-/Log-/API-Fragen aus #32 sowie vollständige sichere Appintegration bleiben offen. Alle8A-Entscheidungen außerhalb dieses begrenzten Kandidaten gelten weiter; kein Fachmodul wurde umgestellt.
5. **Betrieb:** Backup/PITR/RPO/RTO, geprüfter Restore, Auth-/Authority-Secrets/DB-Loginverwaltung, Wartung/Stopkriterien, unabhängiger Auditpfad und getrennte Migration-/Merge-/Deployment-/Aktivierungsfreigaben erforderlich.

**Phase3 zuerst:** Autorisierungs-/Verifikationsvertrag fachlich bestätigen, minimalen vertrauenswürdigen Serverpfad für echte Auth-Sessionattestationen bauen und im genehmigten nativen Stack einschließlich Sperr-/Widerrufsraces abnehmen. Danach Betreiber-/Issuer-/Authority-Provisionierungs- und Auditverfahren sowie Lifecycle-/Primärwechseladapter gezielt integrieren. Kein neues Design, keine erneute BA-/OP-Implementierung und kein Produktivcutover aus PGlite-Ergebnissen.

Technische Referenzen geprüft am13.11.2026: https://supabase.com/docs/guides/auth/sessions ; https://www.postgresql.org/docs/17/sql-set-session-authorization.html ; https://www.postgresql.org/docs/17/explicit-locking.html . Changelog geprüft, keine Paket-/CLI-/Engineupdates.

## 9. Ergebnis und Freigabeentscheidung

Die tatsächlichen Läufe, ihre festen Limits, die korrigierte Sicherheitsprüfung und blockierte Umgebungen sind im [Prüfprotokoll](permissions-8b2-validation.json) getrennt dokumentiert. Ein synthetischer SQL-/RLS-/Rollbacknachweis ist kein echter JWT-/API-/Auth-Session-/Mehrverbindungsnachweis.

**NEIN – keine Produktivfreigabe.** Keine produktiven Daten gelesen/geändert, keine produktive Migration, Auth-/Rollenänderung oder Berechtigung aktiviert. Kein Merge, Deployment oder zusätzlicher Dispatch. Der Auftrag endet nach Erstellung des gestapelten Draft-PR auf #35.

Lokale Ergebnisse: **Security69/69**, davon64 erhaltene Fälle und5 neue Testgruppen, keine Skips; **Fachmodule168/168**, CI-Berichtstests2/2, QA20/20, Integration5/5, qa:check und qa:agent-server bestanden. Pflicht-/neue JS-Syntax, tatsächliche Ausführung der Kandidaten-SQL in PGlite, Quellen-/Änderungsscope-/Dokumentprüfung und git diff --check bestanden. Die vollständige Security-Suite wurde nach der oben beschriebenen Negativtestkorrektur einmalig begrenzt wiederholt; die übrigen geprüften Quellen/Tests blieben unverändert.

**BLOCKED_NOT_TESTED:** Browser-/Agentenmatrix vor Start wegen fehlender Chromium-/Firefox-/WebKit-Binaries; Visuals wegen fehlendem gepinnten Container; native Datenbankabnahme wegen fehlender zugelassener Laufzeit/Auth-/REST-/Restorebasis (DB-Preflight Exit2). Keine dieser Prüfungen als bestanden ausgegeben. Kein Testtimeout ist in den abgeschlossenen Läufen eingetreten; alle gestarteten Testprozesse hatten feste Limits.
