# Schritt 8A – Bestandsaufnahme und Konzept für modulare Berechtigungen

**Status: Analyse-/Reviewentwurf, keine Implementierung oder Aktivierungsfreigabe.** Auftrag vom 13.11.2026, Refs [Issue #33](https://github.com/butzi124-ship-it/schichtplan-mvp/issues/33), #17 und PR #15–#32. Basis **PR #32 / 0.16.4 / security/original-audit-step7 / 67ac80ac1b649fbdf975aeb2d005f827cda8a8a1**. Dokumentationsbranch `docs/modular-permissions-step8a`, APP_VERSION **0.16.5**, VERSION_LOG **13.11.2026 09:20 Europe/Berlin** entsprechend AGENTS.md. Haupt-App ausschließlich Version/Log.

Dieses Arbeitspaket legt keine neuen Rollen in Supabase an, führt kein SQL aus und verändert weder Accounts noch Geschäftsprozesse. Die neuen Rollen-/Tabellennamen in den Konzeptdateien sind **VORSCHLÄGE**, keine vorhandenen Produktionsobjekte. Die neuen fachlichen Vorgaben ersetzen widersprechende alte Konzeptannahmen, aber ändern nicht rückwirkend alte Prüfberichte oder Testidentitäten.

## Dokumente und Verbindlichkeit

| Dokument | Inhalt |
|---|---|
| [Rollen- und Rechtematrix](permissions-8a-matrix.md) | Sieben Organisationsrollen, Scannerabgrenzung, Modulaktionen, Zuständigkeit und Datenschutz |
| [Datenmodell und Autorisierung](permissions-8a-data-model.md) | Mehrfachabteilungen, Leitungs-/Vertretungsfunktionen, dauerhafte/befristete Sonderrollen, Lifecycle und sichere Durchsetzung |
| [Migration, Phasen und Testplan](permissions-8a-rollout-test-plan.md) | Kleine isoliert abnehmbare Implementierungsphasen, Modulbrücken, Regression, spätere Betriebsfreigabe |
| [Offene Entscheidungen](permissions-8a-decisions.md) | Konkrete Fragen, sichere Zwischenposition und betroffene Phase; bereits entschiedene Regeln nicht erneut öffnen |
| [GitHub-/Quellenstand](permissions-8a-source-state.json) | Tatsächliche PR-Heads, Branches/Ancestry, CI-Abfrage, Quellblob-IDs und Aussagegrenzen |
| [Lokale Prüfergebnisse](permissions-8a-validation.json) | Tatsächlich ausgeführte Prüfungen dieses Dokumentationsstands |

**B** = durch diesen Auftrag/fortgeltende Fachspezifikation bestätigt. **I** = aktuell im Code beobachtet, keine neue fachliche Autorisierung. **P** = technischer Vorschlag, später zu reviewen/implementieren. **O** = offene Entscheidung, keine stillschweigende Freigabe. Historische DB-Metadaten bleiben mit Aufnahmedatum gekennzeichnet.

## 1. Tatsächlich überprüfter Entwicklungsstand

Am 13.11.2026 Issue #33 vollständig inklusive Kommentarabfrage gelesen (keine Kommentare), aktuelle offene PRs repositoryweit gesucht, PR-Metadaten und Remote-Refs neu abgerufen. **17 offene Draft-PRs: #15, #16, #18–#32; kein neuerer offener Implementierungs-PR gefunden.** 26 Remote-Refs einschließlich origin/HEAD inventarisiert. `main` bleibt **e764cf2e93382e847627d8bc7c802b02ff70c375**. Abweichender Branchname security/rls-migration-step7 aus einer alten Kurzfassung ist **nicht** PR-32-Head. Der richtige Ref ist security/original-audit-step7.

| Entwicklungsbereich | Verifizierter Stand / Bedeutung |
|---|---|
| Artikel/Operations/Maschinenübersicht/Forecast | PR #11–#14 bereits gemergt; vorhandene Module nicht neu bauen. Gemergter Code ist kein Beweis installierter Produktions-DB-Objekte oder aktiver Featureflags |
| #18 Spezifikation → #19–#24 BA-/OP-Phasen | Offener kumulativer Stapel mit BA-Anlage, Maschinenübergaben, Rüstzeit, bestehendem Zähler, Mengen-/Lifecycle- und Hauptverantwortungsregel; bereits implementiert, erhalten |
| #25/#26 | Sicherheitsbericht, AFTER-STATEMENT-Forecastkorrektur, getrennte native Modi und Quellenprovenienz sind im Stapel enthalten |
| #27 | Startseiten-Design und Navigation abgeschlossen; keine neue Designphase |
| #28/#29 | Isolierter Vertrauens-/RLS-Kandidat und deaktivierter sicherer Clientadapter vorhanden; keine normale Produktivintegration behauptet |
| #30/#31/#32 | Originalnachweis, atomare Reviewpakete, Mitarbeiter-/Material-/31-Tabellen-Rechteprüfungen; keine ausgeführte Produktivmigration |
| #15/#16 | Separate DB-/Härteprüfungsbranches sind nicht pauschal Vorfahren von #32. Übernommene native Harness-/Quelladapter und bewusst isolierte Admin-only-Diagnose separat behalten; diese Diagnose wird keine zukünftige Rollenmatrix |

Die vollständigen Heads/Parents stehen im Quellen-JSON. #18–#32 sind kumulativ in #32 enthalten; überholte Branches nicht erneut zusammenkopieren. Neuer Draft auf #32 erhält diese History. Kein Merge/Rebase abgeschlossener Arbeit erforderlich.

### CI des Ausgangspunkts, neu abgefragt

Für Commit 67ac80ac… meldet der GitHub-Connector sechs PR-Workflows **completed/success**: Security Step7 **43224389227**, Step6 **43224390230**, BA/OP **43224389142**, Step4 **43224389151**, Step3 **43224389238**, QA and Release **43224389154**. Dessen Jobs QA Integrity, Functional Browsers, Visual Regression, Quality Gate und e2e sind success; **Deploy Tested Main ist skipped**. Step7-Vertrags- und native Fortsetzungsjobs sind success.

Das sind neu gelesene **Statusmetadaten**, kein neuer Audit aller ZIPs/Einzelfallzahlen. Frühere lokale BLOCKED_NOT_TESTED-Angaben bleiben historische lokale Angaben und werden nicht überschrieben. Der erfolgreiche strukturelle native CI-Lauf bleibt **kein vollständiger Originalrestore, keine neue modulare Rollenabnahme und keine Produktivfreigabe**. Kein Run wurde nachgestartet; keine Warteschleife.

## 2. Datenbank-Istzustand: vorhandene Strukturen und Beweisgrenzen

Verwendet wird der Originalkatalog aus PR #32, Snapshot **12.11.2026 12:36:34 UTC**, nicht als Live-Abnahme vom 13.11. ausgegeben. In 8A keine neue produktive DB-Verbindung und keine Daten-/Auth-Abfrage. Dieser Katalog belegt 31 öffentliche Tabellen/308 Spalten/134 Constraints/127 Policies/6 öffentliche Funktionen/76 Indizes/11 Benutzertrigger. Vollständiger unabhängiger pg_dump-Restore und API-/Authkonfiguration bleiben offen.

| Objekt / Quelle | Tatsächlich vorhandene oder implementierte Eigenschaft | Grenze für das neue Zielmodell |
|---|---|---|
| Original public.employees | UUID id; nullable/unique auth_user_id mit Auth-FK; einzelnes department_id und zusätzlich Freitext department; role, is_active und Personalfelder. Keine ursprünglichen Kontaktspalten | Eine Stammabteilung/Profilrolle ist keine N:M-Mitgliedschaft oder modulare Authority. Freitext nicht für Rechte verwenden. Leere/unklare Bindungen nicht erraten |
| Original employees_role_check | admin/employee/tool_scanner. Frontend akzeptiert zusätzlich department_admin | Widerspruch belegt; isolierter Step7 erweitert gezielt auf vier Werte, nicht auf sieben neue Organisationsrollen |
| Original public.departments | id, name/code, active, ein leader_employee_id als FK | Ein Mitarbeiter kann schon von mehreren Abteilungszeilen als Leiter referenziert werden. Je Abteilung aber nur ein solcher Zeiger; keine expliziten zeitlichen Leitungs-/Stellvertreterzuweisungen |
| Original SQL/RLS | Zehn Tabellen ohne RLS; breite Rollen- und Wartungsrechte auf allen31, permissive employees/tools/journal/Planner-Regeln | Kritische Befunde werden durch dieses Konzept nicht behoben. Neuen Modellbau nicht als Grund verwenden, den kritischen Fix beliebig aufzuschieben |
| Original fünf Definer-Helfer | is_admin/current_employee_id/current_app_employee_id/current_app_employee_role/can_access_production_department; postgres/search_path=public, mutable Employee-Quelle | Keine sichere neue Autorität aus derzeitigen Profilwerten ableiten. PUBLIC-/Anon-EXECUTE, Owner, Grants und sichere Namensauflösung beim echten Cutover weiter prüfen |
| Review humbel_security.principals (#28) | Unabhängige Auth-/Mitarbeiterbindung, enabled/valid_until/evidence, aber noch **eine role/eine department_id** und Gleichheitsprüfungen mit Employeeprofil | Sichere Grundidee erhalten; Struktur reicht für Organisations-/Mehrfachrollen nicht aus. Keine beliebigen Rollenstrings/NULL-Abteilung als Globalvollmacht |
| Review delegations/schedule_grants/record_grants/bindings | Bereich, Befristung/Widerruf, getrenntes Lesen/Planen/Veröffentlichen, stabile Bindungen; Privatsichten bewusst nicht breit freigegeben | Operative Abwesenheitsvertretung/PL-Fallback und dauerhafte Sonderrollen/Handoververhalten sind nicht vollständig umgesetzt |
| Review security_actor_context / js/secure-access.js (#29) | Parameterlose Serveridentität; getUser, aktuelle SQL-Prüfung, eigene Kontakte, minimale Schicht-/Materialprojektion; Memorycache mit Session-/Offline-Invalidierung; standardmäßig deaktiviert, nicht in Produkt-HTML | Antwort und Client kennen vier Rollen, eine Primärabteilung und Planungsbereiche. Keine sieben Rollen, HR-Feldmatrix oder allgemeine modulare Vergabe-API vorhanden |
| Review Step6/7-Pakete | Atomare Guards/Locks/Postconditions/No-op/Rollback; Materialspaltenrechte, alle31 Maintenance-Rechte und bekannter CHECK-Übergang | Nicht als fertige Organisationsrechteinstallation registrieren. Quelle/Assertions/Provenienz bleiben erhalten |

Im Originalsnapshot fehlen humbel_security, security_*-RPCs sowie die vorgeschlagenen neuen BA-/OP-/Artikelstrukturen. **Inventur inventory_sessions/entries fehlt dort ebenfalls**. Das bedeutet nicht, dass deren Code neu entwickelt werden soll: es existieren Code/Review-SQL, aber kein belegter produktiver Installationsstand. Die 152 vorhandenen direkten Zugriffe bleiben das Quellinventar; 149 betreffen aktuelle Originaltabellen, drei die vorgesehenen Inventurobjekte.

## 3. Frontend, Personal- und Abteilungsverwaltung

Quellreferenzen gelten am Basis-SHA; spätere Versionszeile verschiebt Zeilen geringfügig. Exakte Blob-IDs im Quellenstand.

| Pfad | Aktuelles Verhalten | Geplante Konsequenz ohne Umsetzung in 8A |
|---|---|---|
| VALID_EMPLOYEE_ROLES (schichtplan_mvp.js), Navigation, canAccessProduction/canManageProduction/canManagePersonnel | Harte Prüfung admin/employee/tool_scanner/department_admin; Personal-/globale Abteilungsadministration admin, Produktion admin/dept_admin, Mitarbeiter Werkzeugtab | Capabilities pro Aktion/Scope statt neuer globaler Stringhierarchie. HR/Technik dürfen nicht über Alias admin alle Altfunktionen sehen/laden |
| getManagedDepartmentIds/getVisibleProduction*/canEditProduction* | Berechnet bereits **alle** departments mit leader_employee_id = aktuelle Mitarbeiter-ID; Admin sieht alles | Diesen Mehrfachleitungsansatz erhalten. Keine Abteilungsermittlung über erste passende Zeile oder Stammabteilung als alleinigen Scope |
| createPersonnelEmployee/savePersonnelEmployee/autosavePersonnelEmployeeDepartment | Direktes employees INSERT/UPDATE; ein department_id/Freitext und UpdatedAt. Stammwechsel derzeit kein transaktionaler Sonderrollenentzug | Später semantisches Stammwechselkommando mit atomarer Suspendierung aller bisherigen Sonderrollen und Audit. Kein unverändertes Autosave als sichere Migration ausgeben |
| createProductionDepartment/saveProductionDepartment | Direktes departments INSERT/UPDATE, ein Leiterfeld; UI verlangt Admin | Abteilungsstammpflege, Leitungsbestellung durch PL und Mitgliedschaft trennen; UI darf nicht allein Authority herstellen |
| getCurrentEmployeeRecord/syncSupabaseSessionToApp/bootSupabase | Eigenes Profil SELECT*, Rolle aus Profil; breite nachfolgende Datenladepfade | Kein sicherer Mehrfachrollen-Boot. Neu versionierter serverseitiger Kontext und minimale Projektionen nötig |
| persist()/Login-/Tasklogs | Planner-/Abwesenheits-/Taskdaten und breite Logs verbleiben als offene Sicherheitsbefunde | Privatspeicher/DOM/Subscriptions beim Rechteentzug bereinigen. RLS schützt keine bereits gespeicherten Browserdaten |
| Aktuelle UI-Funktionen und Module | Viele direkte Rollenchecks und geschäftliche Zustandsprüfungen | Je Modul gezielt umstellen und Positivregression erhalten, kein neuer pauschaler Admin-only-Wrapper |

Weder das Ausblenden einer Schaltfläche noch eine lokale Rolle ist serverseitige Autorisierung. Mehrere Orgrollen dürfen nicht durch einen einzigen neuen role-Wert oder numerisches Rangsystem simuliert werden. Eine Person kann HR und eine ausdrücklich zusätzliche Werkzeugrolle besitzen; HR oder technischer Admin **allein** erzeugen diese Modulrechte nicht.

## 4. Bereits vorhandene modulare Autorität – wiederverwenden

| Bestehendes Review-/Modulobjekt | Vorhandene Trennung | Konflikt/Integrationsbedarf |
|---|---|---|
| article_master_scopes und article_master_assignments | Mehrere explizite Abteilungsscopes pro Auth-ID; individuelle Artikel-/Erst-OP-Zuweisung, expires_at/revoked; getrennte Reviewerregeln | Kein neuer Artikelstamm. Orgrolle allein darf Erst-OP-/Artikelbindung nicht erfinden; Versions-/Vieraugenregeln behalten |
| production_readiness_grants | user_id/machine_id, can_confirm **getrennt von** can_work, expires_at/revoked | Maschinen-/OP-Autorität nicht pauschal durch Abteilungsmitgliedschaft ersetzen oder aus technischer Administration ableiten |
| production_routing_authorities | user/department, leader/deputy/admin, Principal, Zeit/Widerruf; Weitergabe und Maschinenübergabe vorhanden | routing_manager bindet Leiterpointer/role; routing_operator vergleicht noch employees.department_id mit Maschinenabteilung. Für Mehrfachabteilungen gezielte neue Scopebrücke nötig |
| BA-Einstieg, Setup, Menge, Lifecycle | Sichere Fach-RPCs, eigene Arbeit/Teilnahme, Audit, Version/Request-ID und Geschäftsstatus | Bestehende Auftrags-/Operations-/Maschinen-/Zeit-/Mengenlogik behalten. Capability erlaubt keine Umgehung von Teilnehmer-, Mengen- oder Abschlussregeln |
| production_lifecycle_orders und automatische Hauptverantwortung (#24) | BA-Hauptverantwortung separat von operativer OP-Verantwortung; genau ein gültiger Leiter der Erst-OP-Abteilung, Mehrdeutigkeit bleibt offen | Neue Mitgliedschaften/Orgrollen ändern gespeicherte BA-Verantwortung nicht automatisch. Co-Leitung wäre Fachänderung, keine implizite Erweiterung |
| Routing-/Lifecycle-Notifications | Bereits auditable Benachrichtigungs-/Request-ID-Muster vorhanden | PL-Übernahmen später an dieses Outbox-/Idempotenzmuster anbinden; kein neues Mitteilungszentrum oder tatsächlicher Versand in 8A |
| Werkzeuge/Journal/Scan und Inventur | Bestehende Lager-/QR-/Buchungsabläufe und Rollenfilter | Lesen, Buchen, Stammdaten, Undo und Inventur getrennt autorisieren. Geometrische Suche ist nicht Teil dieses Auftrags |

Die alten Module haben keine einheitliche neue Organisationsrolle. Bestehende unabhängige Scopes sind wertvolle Autoritätsbelege, aber nicht automatisch vollständig kompatibel: etwa Pflicht-Enddaten vs. künftig unbefristete Sonderrollen, skalare Primärabteilungen und derzeitige vier Rollen. Quelltabellen und Serverprüfungen müssen pro Adapter vollständig betrachtet werden; kein zusätzlicher alternativer Grantspeicher ohne Migrationszuständigkeit.

## 5. Verbindliche Anforderungsabdeckung

| ID | Bestätigte Vorgabe | Konzept-/Testanker |
|---|---|---|
| R01 | Sieben getrennte Organisationsrollen; Technik ohne automatische vertrauliche Personal-/Gesundheitsrechte | Matrix M1, Modell Rollen-/Datenprojektionen, Tests T01/T07 |
| R02 | Nur GL ernennt PL; PL ernennt AL | Matrix M2, semantische Ernennungskommandos, T02 |
| R03 | Mehrfachabteilungen bei erhaltener Stammabteilung | Mitgliedschaften/Primärfeld getrennt, T03 |
| R04 | AL kann mehrere Abteilungen gleichberechtigt leiten | Leitungszuweisungen statt Primärfilter, T04 |
| R05 | Zusätzliche Mitgliedschaft verlangt ausdrückliche Rechteauswahl | Explizit Keine zusätzlichen Rechte oder konkrete Grants; Mitgliedschaft kein Grant, T03 |
| R06 | Modulare Sonderrechte für mehrere ausgewählte Abteilungen, befristet/unbefristet | Versioniertes Template + pro Scope Grantzeile + atomarer Vergabebatch, T05 |
| R07 | AL gibt nur eigene Rechte im zuständigen Bereich weiter; keine höhere Selbstvergabe | Delegationsobergrenzen/Serverkontext, T06 |
| R08 | Stammwechsel deaktiviert alle bisherigen Sonderrollen bis erneuter Freigabe | Atomarer Primary-Change/Epoch/Suspendierung, T08 |
| R09 | Leiterwechsel erhält gültige Sonderrollen, neue Leitung prüft später | Beständige Abteilungsfreigabe + nicht sperrende Reviewqueue, T09 |
| R10 | Ausscheiden/Deaktivierung beendet Zugriff sofort | DB-Sperr-/Epoch-/Sessionvertrag, T10 |
| R11 | Eigene Stellvertretung pro Abteilung, operative Übernahme bei Leiterabwesenheit | Bereichsgebundene Vertreterzuweisung + vertrauenswürdige operative Verfügbarkeit, T11 |
| R12 | Stellvertreter braucht Freigabe für Sonderrollenvergabe | Separater genehmigter Vergabeumfang, T12 |
| R13 | Sind beide abwesend, übernimmt PL und wird benachrichtigt | Gesicherte Fallbackepisode + transaktionale Outbox, T13 |
| R14 | Vertretungshandlungen protokollieren | Actor/Vertretener/Bereich/Quelle/Aktion, T11–T14 |
| R15 | Alle bestehenden Fachprozesse/Identitäten/Tests erhalten | Modulweise Integrationsphasen, T15–T22; kein Implementierungsdiff in 8A |

## 6. Konflikte mit den offenen Sicherheits-PRs

- **Legacy admin ist keine automatische GL, PL, HR oder technische Administration.** Jede bisherige Identität benötigt unabhängige Bestätigung ihrer künftigen Teilrechte. Testpersona admin bleibt als alte Identität erhalten und wird später nur mit explizit dokumentiertem Kompatibilitätsrechtepaket abgebildet; sie wird nicht still zu technical_admin umbenannt.
- **principals.role/department_id und actor_id-Gleichheiten** aus #28 sind bewusst streng, aber skalar. Ihre Ablösung erfordert gleichwertige vertrauenswürdige Bindungs-/Aktivitäts-/Versionsprüfungen; nicht bloß Gleichheiten entfernen, um Multi-Scope zu ermöglichen.
- **Verknüpfte Leiteridentität in delegations/routing** darf nicht unbemerkt jede beständige Sonderrolle beim Leiterwechsel ungültig machen. Operative Stellvertretung und beständige Sonderrolle brauchen getrennte Lebenszyklen.
- **Unbefristet** passt nicht in überall verpflichtendes valid_until/expires_at. Keine Jahreszahl 9999 als verdeckter Workaround; eigene kontrollierte Zielsemantik und ein versionierter Adapter sind erforderlich.
- **RLS mit true-/breiten Grants** bleibt der kritische Ausgangsbefund. Das neue Modell darf diese Pfade nicht als Legacy-Fallback offenlassen. Neue Scopes schützen keine bereits separat zugängliche Tabelle/RPC.
- **CHECK-Erweiterung auf department_admin** aus #32 löst keinen sieben Rollen umfassenden Katalog. Weder bestehende Daten noch alte Constraints jetzt ändern; neue Rollen zunächst in separaten vertrauenswürdigen Zuweisungen vorsehen.
- **Mehrfachleitung/Erst-OP-Hauptverantwortung**: mehrere von einer Person geleitete Abteilungen sind bereits vereinbar; mehrere gleichzeitige Leiter in derselben Abteilung wären eine bislang offene neue Kardinalität, nicht still mit implementieren.

## 7. Ergebnis, Priorität und Abnahmegrenze

Empfohlene Reihenfolge: bestätigte Matrix/Privatsichten und unabhängige Identitygrundlage → sichere isolierte Mehrfachmitgliedschaften → Organisationsrollen/Ernennungen → modulare Sonderrollen/Lifecycle → operative Vertretung/PL-Outbox → modulweise Adapter unter Erhalt bestehender Fachregeln → vollständiger Originalrestore/Endabnahme → separat genehmigte Betriebsmaßnahme. Kritische Mindesthärtung aus #28–#32 parallel isoliert fortführen, ohne sie vor vollständiger Appkompatibilität produktiv zu aktivieren oder durch einen unvollständigen Gesamtneubau zu ersetzen.

Die Dokumentation bereitet die Implementierung vor; **keine der neuen Modelltabellen, sieben Rollen, Freigabe-RPCs, Benachrichtigungen oder Testszenarien ist damit implementiert**. Sicherheitsfreigabe weiterhin **NEIN**. Vollständiger Originalrestore, verbleibende31-Tabellenrechte, API-/Sessionnachweis, offene Fachentscheidungen, vollständige gehärtete App und Backup-/Betriebs-/Pagesfreigaben bleiben erforderlich.

Technische Referenzen, geprüft am 13.11.2026: https://supabase.com/docs/guides/database/postgres/row-level-security ; https://www.postgresql.org/docs/17/ddl-rowsecurity.html . Kein Paket-/CLI-/Engineupdate, kein Deployment und kein produktiver Zugriff in 8A.

## 8. Tatsächliche lokale Prüfung dieses Dokumentationsstands

`npm ci --ignore-scripts`, beide einzelnen Syntaxprüfungen, `check:syntax`, `test:ci` (2), `test:qa` (20), `qa:check`, `test:integration` (5), `test:security` (55) und `qa:agent-server` sind bestanden. Links, Quellen-/Anforderungsabdeckung und die auf Version/Log begrenzte Appänderung wurden geprüft; `git diff --check` ist sauber.

**BLOCKED_NOT_TESTED:** E2E meldet 773 fehlgeschlagene Browserstarts, Agentenseed einen; installierte Playwright-Browser fehlen. Zusätzlich traten ENOENT-Fehler beim Progress-Reporter auf; deren Ursache ist nicht unabhängig nachgewiesen und wird hier nicht als Anwendungsfehler oder behoben ausgegeben. Visualprüfung blockiert am vorgeschriebenen Container, native Originalrestore-/GoTrue-/PostgREST-Abnahme an fehlender genehmigter Runtime/Restore-/Fixturebasis. Keine Browserassertion und kein neuer Rollenfall wird als bestanden gezählt. Ein voriger unvollständiger Browserversuch ist kein grüner Lauf. Genaue Exitcodes und Log-Hashes stehen im [Prüfprotokoll](permissions-8a-validation.json). Für T01–T22 ist die fehlende Implementierung die bewusste Abnahmegrenze dieses Analyseauftrags.
