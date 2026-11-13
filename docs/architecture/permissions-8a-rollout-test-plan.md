# Schritt 8A – Migration, Integrationsphasen und Regression

**Konzept, keine Umsetzung.** [Quellen/Iststand](modular-permissions-8a.md), [Matrix](permissions-8a-matrix.md), [Zielmodell](permissions-8a-data-model.md), [Entscheidungen](permissions-8a-decisions.md). Sämtliche folgenden Implementierungen verlangen eigene kleine Review-/Testaufträge. In diesem PR keine SQL-Datei, Migration, Adapter- oder Geschäftslogikänderung.

## P1. Migrationsgrundsätze und bestehende Identitäten

1. Den tatsächlichen vollständigen Originalstand sicher exportieren/restaurieren, Owner/ACL/Default Grants/Definer/API/Auth/Session/Extensions prüfen. #32-Katalog ist ein datierter Metadatennachweis, kein vollständiger Restore und keine Personenzuordnung. Kein produktiver Zugriff erforderlich, um 8A zu dokumentieren.
2. Vorhandene Mitarbeiter-/Abteilungs-/Auth-IDs, Personal-/Fertigungshistorie und BA-/OP-Snapshots erhalten. Keine Umnummerierung, keine Namens-/Slotzuordnung als Authority und keine automatische Authkontoanlage. Mitarbeiter ohne Login dürfen weiter planbar sein.
3. Pro bestehendem Account unabhängig prüfen: reale Person oder technisches Scannerkonto, aktuelle Zuständigkeit, Primär-/Zusatzmitgliedschaft, Leitung, Vertretung, vorhandene Artikel-/Maschinen-/OP-/Modulfreigaben, Zeit/Gültigkeit und zulässige Felder. In isolierten Tests nur synthetische Daten; personenbezogene echte Zuordnung ist später ein eigener geschützter Betreiberprozess.
4. **Kein** Mapping admin→GL/PL/HR/TA oder department_admin→Leitung aller Bereiche. employees.department_id ist zunächst ein organisatorischer Altbeleg, departments.leader_employee_id ein zu bestätigender Leitungsbeleg; alte Profilrechte sind wegen der nachgewiesenen Manipulierbarkeit keine alleinige Vertrauenswurzel.
5. Neue Mitgliedschaft braucht ausdrückliche Berechtigungsentscheidung, auch keine zusätzlichen Rechte. Bestehende erlaubte Mitarbeiter-/Scannerbuchungen anhand Fachvertrag bestätigen und explizit in neue Profile übertragen, statt sie durch einen neuen Minimal-Mitarbeiterdefault versehentlich abzuschalten.
6. Unbekannte/mehrdeutige Bindungen quarantinieren und kontrolliert klären. Keine scheinbar sichere Admin-only-Pauschalsperre berechtigter Produktionsprozesse. Read/Write/Confirm/Publish/Admin unabhängig abnehmen.
7. Zunächst additive, standardmäßig deaktivierte Strukturen/Brücken im **isolierten** Ziel. Bestehende geprüfte Autorität nicht gleichzeitig mit einer zweiten offenen Wahrheit betreiben. Datenerhalt, Auth-/Berechtigungsversionen, Invarianten und Cutoverzuständigkeit vor erstem produktiven Paket festlegen.

## P2. Konkrete Konflikt- und Adapterliste

| Bestehender Pfad | Sichere geplante Anpassung | Vorbedingung / Regression |
|---|---|---|
| employees.role / department_id, #28 principals.role/department_id und actor_id | Vertrauenswürdige Identität getrennt von pluralen Rollen/Memberships; Primärfeld erhalten, direkte Sicherheitsfeldschreibwege schließen | Keine Gleichheitsprüfungen ersatzlos entfernen. Unbekannte/alte Accounts und primärbezogene Sonderrollensuspendierung testen |
| departments.leader_employee_id / getManagedDepartmentIds | Geprüfte Leitungszuweisungen liefern alle gleichberechtigten Bereiche; alter Pointer ggf. kontrollierte **einseitige** Kompatibilitätsprojektion | Kein Zwei-Wege-Sync, keine spontan ausgewählte erste Abteilung. Keine Co-Leitung pro Abteilung ohne O03/BA-Hauptregel |
| security_actor_context + vier Rollen in js/secure-access.js | Neuer versionierter minimaler Fähigkeitskontext, bewusst migrierte UI und Serverprüfungen | Kein stiller Vertragsbruch, keine neuen Rollen als admin/employee-Fallback, kein unbekannter Kontext als scheinbarer Login-Erfolg |
| Personnel-/Department-Direktupdates/Autosave | Semantische getrennte Personal-, Membership-, Primärwechsel-, Ernennungs- und Deaktivierungskommandos | HR-Feldmatrix, PL-/GL-Ernennung, Suspendierung/Audit/Version/Outbox atomar, Rollen-/Auth-ID-Änderung nicht als allgemeiner Profilpatch |
| schedule.assign/publish/read und private Mehrtabellenaktionen | Bereits sichere einfache RPCs weiterverwenden; bestätigte fehlende private/storno/tausch/ersatz-Kommandos gezielt ergänzen | Keine halbe RPC-Umstellung plus alte Raw-DML. Publikations-/Privatdatentrennung und alte Planner-Positivfälle |
| article_master_scopes/assignments/settings | Explizite neue Capabilityquelle oder kontrollierte Brücke zu genau diesen bestehenden Regeln | Artikel-/Erst-OP-/Autoren-/Reviewer-/Revision-/Snapshotregeln unverändert; keine lossy Rollenübersetzung |
| production_readiness_grants und routing_manager/operator | Bestätigte Mehrfach-/Ressourcenscopes statt pauschalem Primärfilter; aktuelle unabhängige Authority weiter prüfen | Confirm/Work/Plan getrennt, Fremdabteilungsübergaben und Maschinenbindungen erhalten, Mitgliedschaft nicht unzulässig als neuer Universalfilter |
| Lifecycle-/Setup-/Mengen-/Hauptverantwortung | Gezielte Autorisierungsbrücke am vorhandenen Fachkommando | Kein neuer Zähler, keine neue Rüstzeit-/BA-Logik. Exakt eine Erst-OP-Hauptzuständigkeit und gespeicherte Verantwortung nicht durch normalen Rollenwechsel umschreiben |
| Werkzeug/Journal/QR/Inventur | Aktions-/Feld-/Bereichsvertrag, atomare bestätigte Buchungs-/Journalbefehle und getrennte Scannerprofile | Keine geometrische Suche; Material-/Stamm-/Bestand-/Journal-/Undo-/Inventurrechte getrennt. Nicht im Original vorhandene Inventurstruktur zuerst mit sicherer Migration bereitstellen |
| persist()/Datenlogs/Session/Realtime | Nur benötigte Datenprojektionen, keine vertraulichen Offlinekopien; Generation/Epoch an Antworten/Anzeige | Keine Daten des früheren Benutzers/Scopes nach Loginwechsel oder Widerruf. Unbestätigte lokale Änderungen vor kontrollierter Bereinigung erhalten |

Ein pro Modul aktivierter neuer Vertrag muss **alle alternativen Tabellen-/RPC-/View-/GraphQL-/Workerpfade** schließen oder gleichwertig prüfen. Neue UI-Prüfung vor alter offen erreichbarer Tabelle ist kein Sicherheitsfix. Kein `legacy_allow OR new_allow`, um Regression zu kaschieren. Bestehende Fachzustandsprüfungen bleiben zusätzlich zur Capability nötig.

Im Shadowmodus darf nur im isolierten Ziel verglichen werden: Welche bestätigten alten Positivfälle werden vom neuen Modell erlaubt/abgewiesen? Keine produktiven Personaldaten spiegeln und keine ungesicherte Produktionstelemetrie einführen. Durchsetzung später je vollständig abgenommenem Modul auf **eine** Autoritätsquelle schalten; bei Fehler sicher deaktiviert/gewartet statt zurück auf permissive Legacygrants.

## P3. Kleine unabhängig testbare Implementierungsphasen

Jede Phase: eigener konkreter Scope/Draft, passende isolierte Tests, Originalassertionen/Identitäten erhalten, keine automatische Produktivaktivierung. Vorgänger sind Voraussetzungen, **keine** Erlaubnis zur gleichzeitigen Aktivierung unvollständiger Folgephasen.

| Phase | Abgrenzbares Ergebnis / ausdrücklich nicht enthalten | Vorbedingungen | Unabhängige Abnahme / Exit |
|---|---|---|---|
| 8A (dieser PR) | Bestandsaufnahme, Matrix, Modell, Entscheidungen, Test-/Migrationsplan. **Keine Geschäftslogik** | #32/#33 gelesen | Quell-/Link-/Anforderungskonsistenz, Version/QA, unveränderte SQL/Module |
| 8B – Freigabeverträge/Identity | Zuerst isolierter Actor-/Identitätsvertrag (P6); für folgende Rollen-/Datenadapter bestätigte O01/O02/O10/O11. Kein Rollenbootstrap aus Profilen | Originalschema-/Runtimegrenzen und kritische Härtung verstanden | Inaktive/verwaiste/uneindeutige/stale Identitäten verweigert, alte zulässige sichere Fälle erhalten; T01/T07/T10 |
| 8C – Membership-/Primärmodell | Additives N:M-Modell und gesicherte Stamm-/Zusatzzuordnung, explizite Rechtewahl; noch keine neue Org-Vollmacht | O03/O08/O14, 8B | Mehrfachmitgliedschaft ohne Grant, Primärwechselentzug vorbereitet/integriert sobald Grants existieren, History/FK/Parallelität; T03/T04/T08 |
| 8D – Organisation/Ernennung | GL→PL (Ernennung/Abberufung), GL/PL→AL, PL/zuständiger AL→STV-Verträge, explizite Bereichsleitung, HR/TA-Trennung, ein kontrollierter Legacy-Leiteradapter | 8B/8C und Bootstrap/O12 bestätigt | Nur GL ernennt/abberuft PL; GL/PL bestellen/abberufen AL; PL ernennt und AL benennt STV bereichsgebunden; Multi-Leitung gleich; keine Selbstaufwertung/Techvollmacht; T01/T02/T04/T07 |
| 8E – modulare Sonderrollen | Versionierte Rechteprofile, Scopebatches, befristet/unbefristet, eigene Delegationsobergrenzen, Handoverreview | 8D, O05/O06/O07 | Erteilen/Widerrufen/Reissue/Suspendierung, Primärwechsel **aller** Specials, Leiterwechselerhalt; T05/T06/T08/T09/T14 |
| 8F – operative Vertretung/PL-Fallback | Gesicherte Availability, STV-Kontext, freigegebene STV-Vergabe, deduplizierte PL-Übernahme/Outbox | 8D/8E, O04/O13 | Eigener Bereich, Abwesenheit/Unknown/Expiry, beide fehlen→PL, Nachricht/Audit/Retry; T11–T14; Protokollrechte T23/T24 |
| 8F1 – Vertretungsprotokoll/Korrektur | Append-only-Protokoll mit Pflichtfeldern, HR-Gesamtsicht/Leitungsscope und GL/PL-Korrektur, keine Fachbuchungsänderung | 8B/8D/8F und geklärte historische Lesescopes O18 | T23/T24/T26, direkte CRUD-Denials, keine Selbstfreigabe oder TA-Korrektur |
| 8F2 – Aufbewahrungskonfiguration | Versionierte GL-/TA-Parameterverwaltung und Audit; **kein Löschjob** | 8F1, entschiedenes Konfigurationsschema | T25/T26, korrekter Default, keine Datenlese-/Korrektur-/Löschfreigabe aus TA-Konfiguration |
| Späteres separates Löschpaket | Erst nach Frist-/Gesetz-/Freigeber-/Recoveryfreigabe echter manueller/automatischer Ablauf | O17 vollständig entschieden, genehmigte isolierte Tests | T25/T26, kein Verlust geschützter Originale/Ketten/Auditnachweise; keine automatische Aktivierung |
| 8G1 – Personal/Abteilung | Vorhandene Verwaltung auf genehmigte semantische Aktionen/minimale Projektion umstellen | 8B–8F soweit benötigt, O09/O10 | HR/Tech/GL-Datentrennung, aktive/inaktive Ziele, Lifecycle und keine Sicherheitsfeldpatches; T07/T10/T15 |
| 8G2 – Planner/Privatsichten | Nur bestehende Schicht-/Freigabe-/Antrags-/Privatworkflows sicher anbinden | D1–D4 und passende Kommandoverträge | Lesen/Planen/Publish/Bestätigen getrennt, medizinische Daten geschützt; T11/T16/T19 |
| 8G3 – Werkzeuge/Scanner/Inventur | Bestehende betriebliche Aktionen und sichere atomare Buchungs-/Journalgrenzen; keine neue Suche | O15/D8, sichere Originalstruktur | Bestände/QR/Journal/Undo/Parallelbuchung, HR/TA keine Automatismen; T17/T20 |
| 8G4 – Fertigung/Artikel/BA/OP | Vorhandene Authority-Helfer gezielt anbinden, kein neuer Geschäftsprozess | O08/O15/historische Sicherheitsentscheidung D7, G1/G2 soweit Querschnitt | Fremdabteilungsübergaben, bestehende Zähler/Rüst-/Mengen-/Lifecycle-/Erstverantwortungsregeln; T18/T20 |
| 8H – gemeinsame Abnahme | Vollständiger Originalrestore, echte Rollen/API/Parallelität, UI und Betriebsprobe | Alle relevanten Modulphasen und O16 | T01–T26 tatsächlich ausführen; kein offener sicherheitskritischer Bypass/ungeprüfter Positivfall |
| Später separat – Betrieb | Reviewtes konkretes Produktionspaket und kontrollierter Pilot | Unabhängiger Restore/Backup/PITR/RPO/RTO, Wartung, Versions-/Pagesnachweis | Ausdrückliche getrennte Freigaben für Migration/Merge/Deployment/Aktivierung |

Membershipdaten dürfen früher isoliert erfasst werden, aber Rollen-/Vergabe-/RLSdurchsetzung erst mit der zugehörigen vollständigen Authority. Primärwechsel-Suspendierung muss spätestens **vor erster Sonderrollenaktivierung** verbindlich atomar sein; kein Zeitraum mit schon aktiven Grants und noch wirkungslosem Lifecyclekommando.

Kritische Sicherheitsfixes #28–#32 werden nicht neu implementiert. Bei ihrer isolierten Fortführung alte globale admin-Annahmen prüfen, bevor das neue HR-/TA-Modell aktiviert wird. Das fachlich noch unvollständige neue Gesamtmodell wird nicht als Vorwand zur Produktivaktivierung eines ebenso unvollständigen alten Kandidaten verwendet.

## P4. Sicherheits- und Regressionstestplan

**Die folgenden Tests sind konzipiert, nicht neu implementiert oder bestanden. Für das neue Rollenmodell: BLOCKED_NOT_TESTED bis Implementierung/isolierter Laufzeit.** Die jetzt erneut ausführbaren Bestandstests prüfen ausschließlich unveränderte bisherige Kandidaten. Basisdiagnose, korrigierter Kandidat, integrierte UI und echter Originalrestore bleiben separate Ergebnisse.

| ID | Zwingende positive und negative Fälle | Nachweisebene |
|---|---|---|
| T01 | Alle sieben Rollen einzeln, ausdrücklich kombinierte Rollen, keine Rang-/Employeevererbung; HR/TA ohne Automatik für Werkzeuge/Produktion/Medizin; reservierte Orgaktionen nicht im Modulprofil | SQL/RPC + echte Rollen/JWT + UI |
| T02 | Nur GL ernennt/abberuft PL; GL und PL bestellen/abberufen AL im erlaubten Scope; PL ernennt Abteilungs-STV, AL benennt nur eigene STV; andere/Selbsternennung/Selbstfreigabe/indirekte Ketten/Role-Payload/Template-Umgehung verweigert; unbekannter Bootstrap gesperrt | DB-Constraints/Kommandos/Audit, direkte REST-Angriffe |
| T03 | Mehrere aktive Mitgliedschaften, eindeutige Primärbeziehung; explizite Auswahl Keine Rechte oder konkretes Paket; Mitgliedschaft allein gibt nichts, fehlende Auswahl/Scopebatch abgewiesen | DB-/RPC-/Formularintegration |
| T04 | Person führt A/B gleichberechtigt bei Primär A oder C; C ohne Zuweisung gesperrt; keine erste-Abteilung-Abkürzung, zeitliche Überlappung und unentschiedene Co-Leitung | SQL/Scopekontext + UI + BA-Hauptregel |
| T05 | Modulprofil nur gewählte Aktionen/Felder/Bereiche; A/B vs C; vor Beginn/am Beginn/am Ende, unbefristet, Widerruf, DST; kein 9999-/NULL-Legacyfehler | PostgreSQL + DB-Uhr + Adapter |
| T06 | AL kann nur eigene delegierbare Teilrechte im Scope vergeben; keine fremden Module/Felder/Meta-Rechte, kein höheres Selbstrecht; A+B-Batch atomar denied statt halber Vergabe | RPC/REST + Parallelität + Audit |
| T07 | HR-Core/Privatdaten, eigene Kontakte, Gesundheits-/Abwesenheitsdetails und technischer Sanitized-Kontext getrennt; direkte Basistabellen/Views/Exporte/GraphQL keine Umgehung | RLS/Spalten-/View-/RPC-/APItests |
| T08 | Stammwechsel suspendiert **alle** bisherigen Sonderrollen einschließlich anderer Abteilungen/unbefristeter; keine automatische Reaktivierung; parallele Vergabe/Wechsel serialisiert; Fehler erhält Primär-/Grant-/Auditkonsistenz | Echte Mehrverbindungstransaktionen |
| T09 | Normaler Leiterwechsel/Alter-Aussteller-Austritt erhält gültige Grants anderer Begünstigter; Reviewowner wechselt, kein stiller Ablauf. Expliziter Sicherheitswiderruf separat nach entschiedenem Kaskadenvertrag | SQL/Audit/Reviewtests |
| T10 | Deaktivierung/Austritt: alte JWTs/Requests/Subscriptions/Idempotenzantworten verweigert oder vor Commit beendet; Reaktivierung ohne unprüfte Privilegien; lange Snapshots/Lockwarten/Racegrenzen | GoTrue/REST/DB + Browser/Realtime |
| T11 | Je Abteilung eigene STV, Leiter da/abwesend, STV gültig/expired/revoked, unknown/stale Verfügbarkeit; Fremdabteilung verweigert, keine Diagnosefelder im Kontext | SQL/RPC/UI/Availabilityadapter |
| T12 | STV operativ erlaubt, Sondergrant ohne Freigabe denied, passende Freigabe nur erlaubtes Modul/Scope/Zeit/Ziel; keine eigene höhere Genehmigung | SQL/RPC/Audit |
| T13 | Beide abwesend→zuständiger aktiver PL; fehlender/mehrdeutiger PL fail-closed; Übernahme/Benachrichtigung auch ohne sofortige Buchung, deduplizierte Episode/Retry, Deliveryfehler/Outboxrollback | Native Worker-/Outbox-/RESTtests |
| T14 | Parallele Ernennung/Handover/Sperre/Grant/Primarychange, Versions-/Payloadkonflikte, Retry vor/nach Commit, Audit-/Outboxfehler, falsche Lockreihenfolge; keine Doppelwirkung/kein Datenverlust | Echte PostgreSQL-Verbindungen + Fehlerinjektion |
| T15 | Bestehende Personal-/Abteilungs-CRUD mit zugelassenen Feldern; HR kann zulässige inaktive Stammsicht pflegen, inaktive Person hat keinen eigenen Zugriff; Auth-/ID-/Rollenpatches nicht aus allgemeiner Verwaltung | UI + SQL/REST, bestehende Identitäten |
| T16 | Schichten/Urlaub/Krankheit/Ersatz/Tausch/Samstag/Verfügbarkeit/Tasks/Publish, erlaubte MA-/AL-/STV-/PL-Positivfälle und private Fremdsicht-Denials | Ganze bestätigte Planneroberfläche + DB |
| T17 | Werkzeugmaterialien/Bestände/Stamm/Journal/QR/Entnahme/Auffüllen/Undo/Inventur, gleichzeitige Buchungen und Fehler, Scannerbereiche, HR-/TA-Denial | Bestehende Werkzeugfälle + Native/API/Scan |
| T18 | Artikel-/Erst-OP-/Reviewerrechte, BA-Anlage, Maschinenwechsel zwischen Abteilungen, Rüsten/Teilnahme/Mengen/Nachfertigung/Klärfall/Abschluss, bestehender Zähler und unveränderliche Historie | Vorhandene neun Fachsuiten + native BA/OP |
| T19 | Mehrere Tabs, Logout/Loginwechsel/Tokenrefresh/Epoch, verspätete Antworten/Offline-/Altcaches; keine alte Person/Abteilung im neuen DOM oder Snapshot | Playwright + echte Session-/Realtimetests |
| T20 | Anon, authentifiziert ohne Authority, fremde/fehlende/uneindeutige Bindung, Tabellen-/Spalten-/Schema-/Default-/Functiongrants, SECURITY DEFINER und Namensspoofing; keine Legacy-Allow-OR-Umgehung | Baseline und korrigierte DB getrennt, API |
| T21 | Vollständiger Originalschema-/Owner-/ACLrestore, wiederholte Migration/Version-/Hashdrift, vorher/nachher FK-/Daten-/Audit-/Zeithistorie, sicherer Rollback/Recovery und zulässige Downgradeverweigerung | Genehmigter isolierter Originalrestore, unabhängige Probe |
| T22 | Bestehende Dashboard-/Unterbereichsnavigation, zulässige Aktionen und ruhige Denials für neue Rollen; alle Geräte/Touch/Accessibility, unveränderte Bilder | Vollständige Playwrightmatrix/Visualcontainer |

Alle 251 bisherigen Legacy-Chromium-Testfälle mit ihren Testidentitäten, Assertions, Diagnose- und positiven Fachtests bleiben erhalten. Keine Skips, Referenzupdates, größeren Toleranzen oder Umbenennung alter Identitäten zum Grünerwerden. Neue Rollencases additiv. Wenn eine alte synthetische Adminpersona bislang fachlich zulässige Verwaltung prüft, erhält sie später eine ausdrücklich dokumentierte entsprechende **Business-Authority** im isolierten Setup; sie wird nicht mit der engeren TA-Rolle gleichgesetzt. Existiert ein echter Widerspruch zwischen altem Testvertrag und neuer genehmigter Fachregel, offen reviewen statt Assertion verschleiern oder abschwächen.

### Zusätzliche Testgruppen für den Beschluss vom 13.11.2026

Auch diese Fälle sind **BLOCKED_NOT_TESTED**, nur geplanter Vertrag; kein neuer automatisierter Sicherheitstest oder Berechtigungsfix wird behauptet.

| ID | Positive und negative Abnahme | Ebene |
|---|---|---|
| T23 | Jede Vertretung mit Person/Funktion/Abteilung/Aktion/Zeitpunkt; Leitung nur Zuständigkeit, HR sämtliche Logs/Korrekturen lesend; keine HR-Operationen, TA-/MA-/Fremdscope-Denials, keine Diagnose-/Rohpayloadlecks | SQL/RPC/API/UI, aktuelle und historische Scopewechsel |
| T24 | GL/PL veranlassen berechtigte Korrektur; AL/STV/HR/MA/TA und delegierte Modulprofile denied; kurze Begründung zwingend, auch Leerraum ablehnen; Original unverändert, referenzierte Zusatz-/Folgekorrektur und vollständiger Korrekturaudit; keine rückwirkende Fachbuchung/Authority | DB-/API-Direktangriffe, UI und Auditintegrität |
| T25 | Default fünf Jahre/letzte Änderung/manuell; wählbare Dauer, Ereignisdatum/letzte Änderung, Auto/manuell; nur GL/TA konfigurieren, übrige Rollen denied; jede Änderung/Freigabe auditiert; Konfiguration gibt kein Lesen/Korrigieren/Löschfreigaberecht. Nach O17: Ablaufgrenzen/gesetzliche Sperren/Versionen/Backupnachbehandlung und tatsächliche autorisierte Löschung getrennt testen | Isolierte Clock-/Policy-/Rollen-/APItests, später separate native Löschabnahme |
| T26 | Korrektur vs. Fristablauf/Policywechsel/Löschfreigabe, Abberufung vs. Aktion, gleichzeitige Korrekturen/Retry, Auditfehlerrollback; keine verlorene Kette, verwaiste Verweise, Fristumgehung oder geschützter Datenverlust; Löschtests ausschließlich später isoliert | Echte PostgreSQL-Mehrverbindungstransaktionen, Fehlerinjektion/Recovery |

## P5. Späterer Betrieb und Rückfall

An G0–G8 des vorhandenen [Step5-Runbooks](../qa/security-step5-migration-runbook.md) und die [Step6-Freigabecheckliste](../qa/security-step6-release-checklist.md) anschließen, durch aktuelle #32-Nachweise präzisieren. Originalmetadaten vom 12.11. nicht erneut als generell unbekannt darstellen; vollständige Recovery-/API-/Betriebsnachweise sind dennoch offen.

Voraussetzungen: genehmigtes isoliertes Ziel, echte passende Runtime/Owner/Defaults, unabhängiger Restore, vertrauenswürdiger Person-/Bereichsabgleich ohne Produktionsdaten in CI, numerische RPO/RTO, verschlüsselte vollständige Backups/Schlüssel-/Storage-Recovery, PITR-Nachweis, Wartung/Schreibruhe, Verantwortliche/Stop-Kriterien, tatsächlicher App-/SQL-SHA und Pages-/Branch-/Deploymentzustand.

Späterer Ablauf: Originalfingerprint/Version prüfen → Backup/Restorebeleg → sichere Wartung/alte Schreiber stoppen → additive Identity-/Modellstruktur und unabhängige Enrollmententscheidungen → abgenommene modulare RPC-/RLS-/ACL-/Kompatibilitätsdeltas transaktional → Postconditions unter echten Rollen → Cache-/Session-/Worker-/Clientkoordination → kontrollierter Pilot. Nichttransaktionale Schritte separat planen. Kein automatisches Flag und keine parallele freizügige Altquelle.

Vor Commit: bei Fehler Transaktion zurückrollen und Daten/Rechte prüfen. Nach Commit: sichere Wartung/Write-Sperre, getestete Vorwärtskorrektur oder isolierte Recovery plus kontrollierter Journal-/Historienabgleich; kein unsicheres Restore alter true-Policies/Grants und kein Überschreiben neuerer Auftrags-/Werkzeug-/Auditdaten. Organisations-/Delegations-/Ernennungshistorie bleibt erhalten, Berechtigungsversionen/Sessionentzug müssen Wiederanlauf berücksichtigen. Erst danach eigene Freigabe zu Betrieb, Migration, Merge und Deployment. **Nichts davon wird in 8A durchgeführt.**

## P6. Erster sicher abgrenzbarer Teil von 8B

Empfehlung: **isolierter Identity-/Actor-Vertrag auf synthetischen Fixtures**. Unveränderte #28–#32-Identitätswurzel und aktuelle Actor-/Bindungs-/Sessionversion prüfen, Denials für deaktiviert/verwaist/mehrdeutig sowie Regression zulässiger bestehender Fälle. Zuerst einen kleinen reviewbaren Vertrag mit Positiv-/Negativtests festlegen; nur danach eigener Implementierungsauftrag. Kein GL-Bootstrap, keine Ernennung, kein Protokollkorrektur- oder Löschpfad und keine neue Berechtigung bereits in diesem ersten Teil. O01/O02/O10/O11 bleiben Voraussetzungen für die späteren Rollen-/Datenadapter; keine Offenheit durch angenommene Produktionszuordnung ersetzen.

Die Ergänzung löst den früher offenen GL→AL-Vertrag und präzisiert Abberufung, STV-Benennung, HR-Lesen und TA-Konfiguration. Sie widerspricht nicht der bisherigen Audit-Unveränderbarkeit: fachliche Korrektur ist append-only, physische Bereinigung erst nach geltender Frist/gesetzlichem Vorrang und separater Abnahme. Vorhandene allgemeine Admin-/Audit-/Journalrechte aus Altkandidaten reichen nicht als neue GL-/PL-/HR-/TA-Freigabe. Keine Änderung bestehender Werkzeug-/BA-/OP-Korrekturen oder deren Historienaufbewahrung.
