# Gamification und adaptive Theoriefragen – Entwicklungsstand 0.18.0

Stand: 14.11.2026, Europe/Berlin. Auftrag: [Issue #39](https://github.com/butzi124-ship-it/schichtplan-mvp/issues/39), einschließlich Kommentare (zum Abrufzeitpunkt keine). Der nachfolgende Chat-Auftrag autorisiert die Implementierung über den ursprünglichen Konzeptumfang des Issues hinaus.

## Ausgangsbasis und Abnahmegrenze

Branch `feature/adaptive-learning-issue39`, gestapelt auf PR #38 / `docs/digital-quality-specification`, Commit `e57035398870b74d128ea56c080d7c2769fced04` (0.17.3). Die bestehende Historie von #32, #34, Identitäts-/Sitzungskandidaten 8B sowie Dokumenten-/Qualitätskonzepten bleibt erhalten. Verbindliche Regeln: `AGENTS.md`, `docs/DEVELOPMENT.md`, Architekturkonzept 8A und Sicherheitsberichte 8B1/8B2. Keine vorherige Phase erneut entwickelt.

**Ergebnis: funktionsfähiger, RPC-gestützter Entwicklungsstand mit ausführbarem isoliertem SQL und integrierten Oberflächen; keine vollständige produktive Integration oder Produktivfreigabe.** Alle neuen Datenbankobjekte sind Vorschläge, keine verifizierten Originalobjekte. Produktive Supabase-Daten, Auth, Rollen, RLS, Feature-Flags und Deployment wurden nicht angesprochen oder verändert.

Die unveränderte normale Anwendung hat weiterhin ihre bisherigen Funktionen. `HUMBEL_LEARNING_UI_ENABLED=false` blendet die beiden neuen Bereiche aus. Zusätzlich sind die unabhängigen serverseitigen Einstellungen `enabled=false` und `safety_enabled=false`; Identitäten, Befugnisse und Arbeitskontexte werden nicht automatisch übernommen. Ein im Browser manipulierter UI-Schalter gewährt keine Datenbankbefugnisse. Eine fehlende RPC-/Identitätsintegration zeigt einen ausdrücklichen Nichtverfügbarkeitszustand statt erfundener Punkte.

## Tatsächlich implementierte Funktionen

| Anforderung | Umgesetzt und isoliert nachgewiesen | Grenze / noch notwendig |
|---|---|---|
| Dashboard | Drei getrennte Konten, Level aus historischer Erfahrung, aktueller Fachstand, fünf SVG-Auszeichnungsstufen, Lern-/Auszeichnungshistorie, Wiederholungsbedarf, nächste Frage | Browserdarstellung nicht ausgeführt; Level-Schritt 100 und Teilnahmebonus +1 sind Kandidatenparameter, keine beschlossenen Produktionswerte |
| Fragenpool | Neun initiale Fachbereichsmetadaten, erweiterbarer Katalog; Thema, Konzept, Schwierigkeit, 2–8 Antworten, exakte Mehrfachauswahl, Erklärung, Quelle, Entwurf/Freigabe/Deaktivierung, unveränderliche Revisionen | Keine ungeprüften Fachfragen veröffentlicht; fachlich geprüfter Startkatalog und Varianten fehlen |
| Gemeinsamer Pool | Abteilungsübergreifende Auswahl nur aus freigegebenen globalen Fragen; globale Veröffentlichung zusätzlich über exakte unabhängig bestätigte Nutzlast und befristete Befugnis geschützt | Vergabeverfahren für globale Redaktionsbefugnis separat entscheiden; keine automatische Ernennung von Admin/GL |
| Adaptivität | Fällige Wiederholungen priorisiert, Lernstand/Schwierigkeit berücksichtigt, Wiederholung anderer freigegebener Frage desselben Konzepts; korrekt 14 Tage, falsch 1 Tag; nächstes Angebot zufällig 24–72 Stunden nach Abschluss | Keine automatische Erzeugung fachlich veränderter Fragen; kein Hintergrunddienst, keine garantierte Zustellung innerhalb 30 Tagen und keine Monatsquote von 50 Fragen |
| Punkte | Korrekt +5/+10/+20; erstmals falsch 0; falsch bei Wiederholung bereits gelernten Wissens −3/−6/−12; Teilnahme und historische Erfahrung getrennt | Kandidatenintervall und weitere Betriebsparameter vor Aktivierung bestätigen |
| Abzeichen | Bronze 100, Silber 500, Gold 1000, Platin 2500, Diamant 5000 pro Fach; Auf-/Abstieg, Historie bleibt; Fräser, Platte, Scheibe, Zahnrad, Messschieber | Schwellen konfigurierbar und sichtbar vorläufig; finale Freigabe fehlt |
| Verschieben | Erstes kostenlos, zweites Erinnerung, drittes und jedes weitere −5 Teilnahme; Abschluss setzt Zähler zurück; immutable Ereignisse und eindeutige Request-IDs | Arbeits-/Abwesenheits-/Maschinensicherheitsadapter noch nicht angebunden; unbekannter Kontext führt zu keinem Angebot/keinem Abzug |
| Verschiebetermine | 30 Minuten, Kandidat „später heute“ (+3 Stunden), nächstes bestätigtes zukünftiges Arbeitsfenster | „Heute“ kann am Tagesrand überschritten werden; Arbeitsfenster ist noch kein fachlich bestätigter nächster Arbeitstag. Kalender-/Schichtgrenzen vor Aktivierung präzisieren |
| Privatsphäre | Privat als Standard, freiwillige widerrufbare Abteilungsrangliste, getrennte Lernendengruppe, nur Name/Stand/Level/Abzeichen der Freigegebenen; keine Fehlerantworten an Kollegen | Zulässige Aufbewahrungs-/Auskunftsfristen sowie Datenschutzprüfung offen |
| Leitung | Gesonderte eigene Abteilungsübersicht: Aktivierung, Teilnahme, Verschiebungen/Abzüge und Unterstützungsthemen; Frageneditor und Versionsverwaltung; Theorieunterweisungen zuweisen/bestätigen | Befugnis verlangt aktive unabhängige Zuweisung UND tatsächliche Abteilungsleitung; Stellvertretung/weitere Organisationsrollen werden nicht automatisch interpretiert |
| Lernende | Explizit bestätigtes Curriculum und freigegebene Themen mit Ablaufdatum; allgemeine Fragen ebenfalls nur innerhalb bestätigter Themen; keine praktischen Lernaufgaben | Vollständiger Bildungsplankatalog und individuelle Freigaben fehlen; Ausbildungsbeginn allein erteilt kein Recht |
| Unterweisungen | Eigener Navigationsbereich, eigene Aktivierung und Ereignisse; Gründe Neueintritt/Maschine/Tätigkeit/Vorfall/periodisch; genaue Dokumentversion, Theorie mit Erklärung/Fehlversuchen, Unterstützungsmarker, getrennte tatsächliche Bestätigung mit Nachweis; Folgeauftrag standardmäßig nach 6 Monaten | Keine vollständige Unterweisungsplattform: automatische Ereignisanbindung, konkrete gesetzliche/riskobasierte Fristen, qualifizierte Unterweiserrechte und Tätigkeitsfreigaben offen. Keine automatisierte praktische Zertifizierung |
| Qualität als Lernquelle | Private bestätigte Themenreferenzen beeinflussen Auswahl, ausschließlich mit Nachweis und Gültigkeit | Noch kein Anschluss an das bisher nur spezifizierte Qualitäts-/Problemmodul; keine automatische Schuldzuordnung |

Die Bereiche „Gamification“ und „Sicherheitsunterweisungen“ befinden sich in der vorhandenen Navigation, nutzen helle blaue ERP-Karten, fokussierbare 44-px-Bedienelemente und responsive Raster. Mitarbeiter können Fragen beantworten, verschieben, ihre Sichtbarkeit ändern und ihre eigene Historie ansehen. Berechtigte Leitungen erhalten einen realen RPC-basierten Editor statt Demo-Daten. Das optionale Ranglisten-Opt-in beeinflusst ihre berechtigte administrative Übersicht nicht. Es gibt keine Popups oder Zeitgeber, die Maschinenarbeiten unterbrechen, und keine Speicherung der Lern-/Antwortdaten in localStorage oder IndexedDB. Logout, Benutzerwechsel und Offline-Ereignis leeren den Modulzustand; verspätete Antworten benötigen passende Generation UND serverseitig zurückgegebene Benutzeridentität.

## Datenbank- und Autorisierungsarchitektur

Datei `supabase/review/learning-issue39.sql`: **Review-/Testartefakt, nicht in Deployment-/Migrationserkennung registriert.** Transaktionale Erstinstallation mit isoliertem Datenbanknamen/PGlite-Kennung plus explizitem Sentinel (mindestens 32 Zeichen), bestehender `humbel_security.actor_id()` als Voraussetzung, begrenzter Lock-/Statement-Laufzeit. Bestehendes gleichnamiges Schema führt zum Abbruch statt stillschweigender Änderung. Die Sperre ist eine Fehlbedienungssicherung, kein Ersatz für überprüfte Testzugangsdaten und Zielkontrolle.

Neues privates Schema `humbel_learning`, 17 Tabellen:

- `settings`, `people`, `department_settings`, `scopes`: getrennte Rollout-, Zuordnungs- und befristete Befugnisebene.
- `context_windows`, `education_topics`: ausschließlich vertrauenswürdige Eligibility-/Ausbildungsnachweise; kein Mitarbeiterendpunkt zum Erfinden von Arbeitszeiten oder Fehlzeiten.
- `domains`, `questions`, `global_approvals`: Inhalte, unveränderliche Revisionen und unabhängige Freigaben für den gemeinsamen Pool.
- `units`, `mastery`: persönliche Einheiten, Wiederholungs- und Lernstand.
- `events`, `badge_history`, `receipts`: additive Buchungen/Erfolge, Idempotenz und Konflikterkennung.
- `instruction_assignments`, `instruction_events`: eigenständige Unterweisungsnachweise und Korrekturhistorie.
- `confirmed_topics`: minimierte bestätigte Problem-/Themenreferenzen, keine offenen Qualitätsberichte.

Alle Tabellen haben RLS und keine direkten Rechte für `PUBLIC`, `anon`, `authenticated` oder `service_role`. Drei öffentliche **SECURITY INVOKER**-RPCs (`learning_view`, `learning_admin_view`, `learning_command`) delegieren auf genau drei private **SECURITY DEFINER**-Einstiege; sämtliche übrigen Hilfsfunktionen sind nicht öffentlich ausführbar. Private Einstiege revalidieren den unabhängigen Sicherheitsprincipal, Mitarbeiteraktivität und Übereinstimmung von Benutzer/Mitarbeiter/Stammabteilung. Scanner erhalten keinen Lernzugriff. Leere `search_path`, vollständig qualifizierte Objekte und explizite EXECUTE-Rechte schützen die Namensauflösung. Keine Clientrolle, Clientpunktzahl oder frei übermittelte Mitarbeiter-ID ist Autoritätsnachweis. Zielpersonen-IDs bei Verwaltungsbefehlen werden auf die nachgewiesene Zuständigkeit geprüft.

SQL-Kandidaten und synthetische Root-Fixtures sind keine gehärtete Originaldatenbank. Die Funktionen besitzen zunächst den isolierten Installationsowner. Ein dedizierter minimal privilegierter NOLOGIN-Funktionsowner, Abhängigkeitsrechte und sichere Auth-/PostgREST-Konfiguration sind vor nativer Freigabe verbindlich auszulegen und zu testen. Die offenen Legacy-`employees`-Policies dürfen nicht als vertrauenswürdiger Weg zur Aufnahme in `people`/`principals` verwendet werden. Ein manipulierter Rollen-/Abteilungswert widerspricht dem unabhängigen Principal und wird verweigert. Es gibt keinen automatischen Enrollment-/Bootstrap-Endpunkt.

### Transaktionen und Audit

Ein gemeinsames transaktionales Advisory-Lock serialisiert neue Modulschreibvorgänge; Actor-/Scope-/Mitarbeiterzeilen und die konkrete Einheit werden im selben Ablauf gesperrt und erneut geprüft. Request-ID plus vollständige kanonische Nutzlast verhindern geänderte Wiederholungsanfragen. Identischer Replay erzeugt keine weitere Buchung und liefert ausschließlich die aktuelle persönliche Projektion. Auditfehler rollen Ergebnis, Score, Status und Receipt gemeinsam zurück. Originale Ereignisse, Frageninhalte und Abzeichenhistorie dürfen nicht überschrieben, gelöscht oder geleert werden. Datenbankowner sind technisch nicht durch diese Trigger absolut bindbar: Betreiberzugriff, externe Sicherung/Audit und Rechtehärtung bleiben erforderlich.

Angebot/Verschieben verwenden zusätzlich einen SHARE-Lock auf Kontextfenstern, damit während einer Abzugsentscheidung keine gegenteilige Ausnahme eingefügt wird. Ein späterer Kontext-/Widerrufsadapter muss dieselbe konsistente Sperrreihenfolge respektieren. Lang laufende Fremdtransaktionen, echte parallele Sessions und Widerrufe wurden hier **nicht** nativ geprüft; dafür sind Mehrverbindungstests Pflicht. Der globale Modul-Lock ist ein bewusst konservativer Kandidat mit möglichem Durchsatzlimit, kein Skalierungsnachweis.

## Schweizer Ausbildungsgrundlagen – Quellenstand 14.11.2026

Offizielle Berufsseite: https://www.becc.admin.ch/becc/public/bvz/beruf/show/45708 (Polymechaniker/in EFZ ab Lehrbeginn 2026, Verordnung vom 29.08.2025, Inkrafttreten 01.01.2026).

Verordnung, Art. 25: https://futuremem.swiss/assets/inhalt/Downloads/Bildungsverordnungen/Def-Stand-September-2025/SR-412.101.220.88-01012026-DE.pdf ; offizieller Erlass https://www.fedlex.admin.ch/eli/cc/2025/572/de . Der PDF-Text wurde geprüft; Fedlex lieferte hier nur die JavaScript-Hülle. Neues Qualifikationsverfahren erstmals 2030, Teilprüfung erstmals 2028. Vor 2026 begonnene Ausbildungen werden bei Abschluss vor 31.12.2031 nach bisherigem Recht abgeschlossen; für verkürzte neuere Ausbildungen und Wiederholungen gelten zusätzliche Übergangsbestimmungen. Deshalb keine pauschale Zuordnung allein aus einem Datum.

Der verlinkte Bildungsplan https://www.becc.admin.ch/becc/public/bvz/beruf/download/17128 war über den Abrufweg nicht vollständig lesbar; daraus wird keine vollständige Themen-/Ausbildungsfreigabe behauptet. Zuständige Ausbildungsfachpersonen müssen Curriculum und bereits vermittelte Themen ausdrücklich bestätigen. Keine automatisch publizierten KI-Fragen, keine praktischen Aufgaben und kein Ausbildungsheft.

## Tests und tatsächliche Nachweise

Maschinenlesbares Protokoll: [issue39-validation.json](issue39-validation.json). Sämtliche lokalen Testprozesse hatten feste Timeouts (30–300 Sekunden mit kontrolliertem Kill-Nachlauf); keine unbegrenzten Wiederholungen. Die vollständigen bestehenden Node-Suites bestanden. Nach den letzten SQL-/UI-/Sessionänderungen bestand `test:learning` mit **16/16**, nach Browserregistrierung und Workflow-Ergänzung Syntax, **22/22 QA** und QA Integrity erneut.

| Prüfung | Ergebnis |
|---|---|
| `npm ci --ignore-scripts`, beide direkten JS-Syntaxprüfungen, `check:syntax` | PASS |
| `test:ci` / `test:qa` / `test:integration` / `test:security` | PASS: 2 / 22 / 5 / 69 |
| Artikel, OP, Bereitschaft, Forecast, BA-Anlage, Routing, Rüstzeit, Mengen, Abschluss | PASS: 168 vorhandene Fachtests zusammen |
| `test:learning` | PASS: 11 SQL-/Transaktionsgruppen und 5 UI-/Sessiongruppen; zusammen 16 |
| `qa:check`, `qa:agent-server` | PASS; 251 Legacy-Identitäten und 33 Bildreferenzen erhalten |
| Playwright `--list` | PASS: 3 neue Fälle × 5 Geräte-/Browserprofile = 15 registriert; **keine Ausführung** |
| `test:e2e`, `test:agents` | **BLOCKED_NOT_TESTED**: Chromium/Firefox/WebKit-Binaries fehlen |
| `test:visual` | **BLOCKED_NOT_TESTED**: vorgeschriebener Playwright-1.62.1-noble-Container fehlt; keine Referenzen verändert |
| Native PostgreSQL/GoTrue/JWT/PostgREST/RLS-/Session-/Parallelitätsabnahme | **BLOCKED_NOT_TESTED**: keine zugelassene lokale vollständige Laufzeit/Identitäten/Originalrestore; DB-Preflight meldet korrekt blockiert |
| `git diff --check` und JSON-/Workflow-Konsistenz | PASS, abschließend vor Commit geprüft |

**282 bestandene Node-Testfälle** über die aufgeführten Suites, keine ausgelassenen oder abgeschwächten Assertions. PGlite führt die tatsächliche neue SQL-Datei mit synthetischen Tabellen/Auth-Shims aus; es ist weder ein Supabase-Auth-/JWT-Sicherheitsnachweis noch ein Mehrverbindungs- oder Originalschema-Restoretest. Die UI-Node-Tests prüfen Rendering/Handler, nicht reale Pixel-/Touchbedienung.

Während der Entwicklung fanden die neuen Tests einen SQL-Guard-Klammerfehler und einen mehrdeutigen PL/pgSQL-Bezeichner; beide wurden im Anwendungskandidaten korrigiert. Ein bestehender QA-Test entdeckte die Auftrennung des geschützten Dashboard-HTML-Blocks; die zusätzlichen Assets wurden davor verschoben, der Originalblock blieb unverändert. Neue eng begrenzte, exakt einmalige Integrationsadapter dokumentieren ausschließlich neue App-/Assethaken und die additive Playwright-Registrierung. Kein geschützter erwarteter Hash wurde geändert. Die zusätzliche Learning-Suite läuft im bestehenden QA-Integrity-Job mit fünf Minuten Timeout; keine neue konkurrierende Browserpipeline und kein Dispatch.

## Kontrollierter Weg zu einer separat freizugebenden Testversion

1. Review von SQL, UI und Berechtigungsentwurf; Issue #39 und Stack #28–#38 auf aktuellen Integrationsstand prüfen. Offene Originalschema-/Legacy-`employees`-/Auth-Freigabeblocker bleiben eigenständige Blocker.
2. Autorisierte wegwerfbare PostgreSQL-/GoTrue-/PostgREST-Umgebung mit synthetischen Benutzern, überprüften Schemaabhängigkeiten und lokaler Zielprüfung bereitstellen. Vorlagen-/Root-Fixtures nicht mit verifiziertem Originalrestore verwechseln. Backup und unabhängigen Restore in isolierter Umgebung, RPO/RTO/Retention mit Betreiber festlegen.
3. Minimalen Funktionsowner, unabhängige Aufnahme-/Widerrufsautorität und globale Inhaltsfreigabe klären; keinen bestehenden Admin automatisch umwandeln. SQL-Datei zuerst transaktional im isolierten Ziel prüfen; keine Registrierung als produktive Migration. Ausgangsdaten sichern und Integritätszahlen vergleichen. Erstinstallationsabbruch/idempotente Request-Verarbeitung sind verschieden: die Installationsdatei ist absichtlich keine wiederholt reparierende Migration.
4. Vertrauenswürdige Arbeitszeit-/Urlaub-/Krankheit-/Maschinensicherheitsadapter und Kalendergrenzen anbinden; keine Rohdiagnosen in Lernkontexte übernehmen. Angebotserzeugung mit mindestens einer sicheren Gelegenheit innerhalb 30 Tagen, Urlaubsausnahmen, aufgelaufenen Angeboten und Monatsziel separat fachlich auslegen/implementieren. „Später heute“ und „nächster Arbeitstag“ korrekt in lokalen Schicht-/Feiertagsgrenzen testen.
5. Fachredaktion liefert geprüften Fragenpool einschließlich Wiederholungsvarianten; Ausbildungsfachperson bestätigt vollständigen aktuellen Bildungsplan, Curriculum und Themen. Schwellen, Level, Teilnahmebonus, Wiederholungsfristen und Datenschutz-/Aufbewahrungsparameter ausdrücklich freigeben.
6. Pflichtunterweisungen: risikobasierte Fristen, qualifizierte Durchführende, Eskalation/Benachrichtigungen, praktische Nachweise und Tätigkeitsfreigaben separat abnehmen. Deaktivierte Gamification darf diese nie deaktivieren. Die neue Theorieoberfläche ersetzt bestehende betriebliche Nachweise nicht.
7. Native Negativtests: anon, unenrolled, Scanner, verändertes Profil, abgelaufene/widerrufene Scope-/Sessionrechte, fremde Abteilungen, Lösungsexfiltration, Replay und Selbstfreigabe. Mehrverbindung: Antwort/Antwort, Abzug/Ausnahme, Verwaltung/Widerruf, Consent/Leaderboard, Auditfehler/Rollback; keine Doppelbuchung und keine Freigabe veralteter Kontexte. Backend-Ressourcenlimits/PostgREST-Berechtigungen prüfen.
8. Vollständige vorhandene Browsermatrix plus 15 neue Fälle, Desktop/Tablet/Telefon, Fokus/Touch/Scroll/Screenreader, Legacy-Workflows und unveränderte visuelle Referenzen im gepinnten Container abnehmen. Bisherige Node-Pässe ersetzen diese Abnahmen nicht.
9. Separat freigegebene Testinstanz: unabhängige synthetische Zuordnungen und Scopes bewusst einrichten, serverseitige Testgates und nur dort den UI-Schalter aktivieren; keine Demoidentitäten oder Testnachweise produktiv übernehmen. Verfügbarkeit, Widerruf und Unterweisungstrennung nachweisen.
10. Späterer Produktivplan benötigt separate Freigabe, geprüften Originalschemaexport/Backup/Restore, Wartungsfenster und Verantwortlichkeiten. Bei Fehlern zunächst neue Gates schließen und schreibende Zugriffe stoppen, Audit-/Lernhistorie erhalten; keine alten unsicheren Policies reaktivieren. Destruktive Rückmigration ist hier nicht freigegeben.

## Verbleibende Entscheidungen und ehrliche Fertigstellungsgrenze

**Keine Produktionsbereitschaft.** Native Identität/Sessions und Ownerrechte, vollständige Browser-/Visuellabnahme, vertrauenswürdige Eligibility, Kalender-/30-Tage-Hintergrunddienst, überprüfter Inhalts-/Ausbildungsstand, konkrete Unterweisungsautorität und Datenschutz-/Betriebsparameter fehlen. Qualitätsanbindung, externe Meldungen, weitere modulare Rollen und Mehrmandantentrennung sind nicht als fertig implementiert auszugeben. Insbesondere stammt eine öffentliche Lernrangliste nicht aus einem privaten Fehlerprofil.

Nächster sinnvoller Schritt ist die native isolierte Auth-/RLS-/Parallelitätsabnahme der vorhandenen RPCs, anschließend der vertrauenswürdige Kalender-/Sicherheitskontextadapter mit tatsächlichem Angebotsdienst. Weitere UI-Funktionen sollten diese Sicherheitsabhängigkeiten nicht vorwegnehmen. Keine produktiven Änderungen, kein Merge und kein Deployment in diesem Auftrag.
