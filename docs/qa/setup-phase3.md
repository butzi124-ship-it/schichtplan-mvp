# Phase 3 – Rüstzeiterfassung (Reviewstand 0.12.4)

Gestapelter Draft auf PR #20 `feature/ba-op-phase2`, geprüft gegen dessen tatsächlichen Head `e276270adc704fc9033e304ff5ad81db56593ecd`. Ref #17, verbindliche Spezifikation PR #18. Neuer Branch `feature/ba-op-phase3`. Kein Merge, Deployment, produktiver Datenbankzugriff oder Feature-Aktivierung.

## Implementierung

- `js/setup-times.js`: expliziter Start mit intern/extern, beliebige Moduswechsel, persönliche Pause/Ende und gemeinsamer Abschluss; Teilnehmer dürfen auch nach persönlichem Ende den gemeinsamen Vorgang schließen. Serverzeit statt Gerätezeit.
- Ein offener gemeinsamer Vorgang je Maschinenabschnitt der OP, persönliche Segmente über die vorhandene Phase-2-Schnittstelle. Interne persönliche Zeiten belegen die Maschine, externe ändern keinen Maschinen-Fertigungsstatus. Interne Rüstung blockiert parallelen Produktionsstart; laufende Produktion blockiert Wechsel zu intern. Externe Rüstung lässt vorhandene Produktion weiterlaufen.
- Allgemeine Mitarbeiterleiste außerhalb der Rüstmaske: Pause unterbricht sämtliche laufenden persönlichen Produktions-/Rüstsegmente. Gemeinsame Rüstung und Kollegen bleiben unberührt. Rückkehr fragt Ja/Nein; Ja verlangt erneute Moduswahl, Nein startet nichts. Bereits geschlossene Vorgänge werden nicht angeboten.
- Tatsächliche Abmeldung schließt persönliche Segmente erst nach Serverbestätigung; bei Fehler bleibt die Session bedienbar. Ausdrückliche Arbeitsplatzübernahme schließt ausschließlich die gewählte Vorgängerbuchung dieses Maschinenabschnitts und bietet offene Rüstung zur Fortsetzung an. Andere OPs/Personen bleiben erhalten. Schichtplan liefert nur Erinnerung, niemals einen Stopp.
- Maschinenwechsel/Rückgabe benutzt Phase 2: einheitlicher Zeitabschluss, konfigurierter Abbruchgrund auch bei pausiertem offenen Rüstvorgang, unveränderliche alte Maschinenabschnitte und neuer expliziter Rüststart. Kein zweiter Zähler und keine Änderung alter Ist-Minuten.
- Korrektur geschlossener persönlicher Segmente mit Pflichtbegründung, vollständigem effektivem Alt-/Neu-JSON, optimistischem Vergleich, Überlappungsschutz und unveränderlichem Original. Effektive Auswertung über Korrektur-Overlay. Eigene Korrekturen sowie unabhängig provisionierte zuständige AL/Vertreter/Admin serverseitig geprüft.
- Getrennte Kennzahlen: Summe persönlicher Minuten, Kalenderdauer zwischen gemeinsamem Start/Ende einschließlich Pausen, Vereinigungsmenge aktiver Kalenderintervalle, Vereinigungsmenge interner Maschinenstillstandsintervalle. Mehrere gleichzeitig rüstende Mitarbeiter verdoppeln keine Maschinenzeit. Artikel-Soll bleibt Gesamtvorgabe.
- Kein Anwesenheitskonto, kein MOBATIME-Aufruf, keine personenbezogenen Produktionsdaten in localStorage. Alle Schalter bleiben false; Reviewtabellen/RPC nur in synthetischen Tests aktiviert.

## Datenbank-/Nebenläufigkeitsentwurf

`supabase/review/setup-phase3.sql` ist unregistriert und nicht produktiv angewandt. Fünf neue Tabellen plus `setup_run_id` am vorhandenen Segment. Neue Tabellen RLS-aktiv, ohne Raw-Rechte für anon/authenticated; öffentliche RPC mit Invoker und eng begrenzten privaten Definer-Funktionen (leerer search_path, Auth-/Authority-Prüfung). Unveränderliche Events/Korrekturen. Auditfehler rollen Geschäftstransaktion zurück.

Konservativer gemeinsamer transaktionaler Advisory-Lock vor BA-/Abschnittssperren serialisiert Rüst-, Routing- und Readiness-Kommandos. Dadurch verhindern auch Pause über mehrere OPs, paralleler Maschinenwechsel und Teilnehmerabschluss doppelte Abschlüsse. Idempotente Requests sind an Actor, Aktion und vollständige Payload gebunden. Produktiver Durchsatz und feinere Sperren bleiben vor Freigabe zu messen; nicht als bereits skalierbare Produktivlösung ausgeben.

## Tests und Nachweise

Lokaler Stand vor CI:

| Prüfung | Ergebnis |
|---|---|
| npm ci --ignore-scripts | bestanden |
| Hauptdatei/Smoke node --check + check:syntax | bestanden |
| Phase 3 SQL/Unit | 17/17 bestanden, eingebettetes PGlite, keine native Parallelität |
| Artikel / OP / Readiness / Forecast / BA / Routing | 8 + 15 + 17 + 30 + 9 + 10 = 89/89 bestanden |
| test:ci / test:qa | 2/2 und 9/9 bestanden |
| qa:check | 251 Legacy-Identitäten, alter Business-Hash nach expliziter Hook-Umkehr und 33 Bildreferenzen erhalten |
| qa:agent-server | bestanden |
| Lokale test:agents / test:e2e | blockiert: Playwright-Browser fehlen; Download liefert ungültiges ZIP. Kein bestandener Browsertest behauptet |
| Lokale Visuals | blockiert: vorgeschriebener gepinnter Container fehlt |
| Allgemeiner DB-Preflight | BLOCKED_NOT_TESTED; kein autoritativer Restore/isolierter nativer Stack lokal |

40 neue Phase-3-Browserausführungen (8 Fälle × Desktop/Smartphone/Tablet/Firefox/WebKit) sind der vollständigen unveränderten bisherigen Matrix hinzugefügt. Insgesamt 683 Browserfälle. Echte Geräte/Scannerhardware nicht getestet; Gerätebezeichnungen meinen Browser-/Viewportprofile.

Native CI-Extension ergänzt die unveränderten 14 Baseline-, 6 Phase-1- und 7 Phase-2-Fälle um Phase-3-Fälle, inklusive realer JWT-/Raw-Denials, unabhängiger PostgreSQL-Verbindungen, parallelem Start/Abschluss/Pause/Maschinenwechsel, HTTP-409, Audit-Rollback und bestehender Stückzahlen. Pinned Infrastruktur aus #15: `25e9a6914e38fc791d9ecad13d92a45e30d19bb3`. Ergebnisse/CI-Links und tatsächlicher finaler SHA werden im Draft-PR nach dem Lauf ergänzt. Keine grüne Datenbankabnahme vor tatsächlicher Ausführung.

## Geänderte Dateien

Frontend: `js/setup-times.js`, `js/machine-readiness.js`, `schichtplan_mvp.html`, `schichtplan_mvp.js` (gezielte Hook-Datei beschreibt jede Abweichung vom geschützten Legacy-Code). SQL: `supabase/review/setup-phase3.sql`. Tests: `tests/helpers/setup-db.js`, `tests/setup/database.test.js`, `tests/setup/unit.test.js`, `tests/setup.spec.js`, `tests/db-integration/setup-extension.js`. QA/CI: `package.json`, `playwright.config.js`, `scripts/qa-check.js`, `tests/qa/setup-integration-patches.json`, `tests/qa/protected-files.json`, beide bestehenden Workflows. Dieser Bericht.

## Offene Freigaben, Risiken und Phase 4

- Produktivschema, Ownership, ACLs/Default-Grants und tatsächliche Abteilungs-/Mitarbeiterzuordnungen müssen autoritativ abgeglichen werden. #15/#16 bleiben offene Sicherheitsbefunde; dieser PR behebt keine Legacy-RLS-Schwachstellen und ersetzt keinen Restore.
- Vertrauen in unabhängig administrativ provisionierte Routing-Authorities, echte Leiter/Vertreter, Maschinenberechtigungen und Abbruchgründe bleibt zwingend. Mutable Mitarbeiterrolle/JWT-user_metadata alleine genügt nicht.
- Keine produktive SQL-Anwendung, RLS-/Auth-Änderung, Schalterfreigabe, Integration von #15/#16 oder Deployment autorisiert. Schalter erst nach separater DB-/Sicherheits-/Fachabnahme aktivieren.
- Legacy-Spannungs-Ist-Minuten sind manuelle Werte, keine laufenden persönlichen Timer. Sie bleiben unverändert. Allgemeine Pause unterbricht bestehende laufende Segmentbuchungen; ein neuer persönlicher Produktionsstart wird hier nicht erfunden.
- Geschlossene Browser, verlorene Netzwerkverbindungen oder abgelaufene Auth-Sessions sind keine bestätigte Abmeldung. Keine automatische Zeitabschaltung nach Browser-/Schichtende; spätere berechtigte Korrektur/Audit bleibt nötig.
- Historische interne Maschinenintervalle beruhen auf Rüstmodus, nicht auf Telemetrie; reale Maschinenstillstandsursachen und alte Produktionsintervalle können ohne spätere Phase-4-Zuordnung nicht rekonstruiert werden.
- Phase 4 integriert bestehende Spannungs-/Gut-/Ausschussereignisse exakt in dieselbe OP und ihre Maschinenabschnitte. Keine künstliche historische Mengenaufteilung, keine zweite Zähllogik. Mengenklärfälle/Nachfertigung/OP- und BA-Abschlüsse folgen gesondert.

## Zuordnung der zusätzlichen Abnahmekriterien

| Nutzerfall | Automatisierter Nachweis |
|---|---|
| 1 Intern/extern wechseln | SQL `explicit internal/external switches`; Browser `explicit start then mode switch` |
| 2 Mehrere Mitarbeiter | SQL `two employees share run`; native `Native mode switches` |
| 3 Persönliches Ende | dieselben Tests prüfen explizit weiterlaufenden Kollegen |
| 4 Gemeinsames Ende | identischer Abschlusszeitpunkt, keine offenen Teilnehmersegmente |
| 5 Pause/Fortsetzung | SQL `general pause`; zwei Browserfälle mit Nein/Ja/Modus und inzwischen geschlossenem Vorgang |
| 6 Schicht/Übernahme | SQL `logout and explicit workstation takeover`, `stale takeover`; Unit `planned shift ending` |
| 7 Parallele Maschinenfertigung | SQL/native `external setup`; zusätzlicher Produktionsstart-/Resume-Sperrtest für intern |
| 8 Maschinenwechsel | SQL `paused shared setup`; Browser `machine switch`; echte native Parallelverbindung gegen Logout |
| 9 Korrektur/Audit | SQL und Browser für Begründung, Alt-/Neu-Overlay, unverändertes Original, Stale und Überlappung |
| 10 Konkurrierende Ereignisse | native Doppelstart-Idempotenz, Pause gegen Teilnehmerabschluss, Routing gegen Logout |
| 11 Rollen/Abteilungen/Vertreter | server flag/scanner/raw, fremder Mitarbeiter, mutable admin, übertragene Abteilung und abgelaufene/aktive Vertretung; Phase-2-JWT-Regression unverändert |
| 12 Datenintegrität | unveränderte Altzähler/Ausschuss, unveränderte Snapshots/Ist-Minuten in bestehenden Tests; 251 geschützte Testidentitäten und exakter Legacy-Business-Hash |

Erster nativer Vorabstand `34f3729a0713d68f4d4c037e006991cccc36f484` (0.12.0): [35/35 native Tests und 2/2 Zielschutztests bestanden](https://github.com/butzi124-ship-it/schichtplan-mvp/actions/runs/42393464728). Strikte Visuals dieses Vorabstands bestanden zweimal 33/33. Der nachfolgende Stand 0.12.1 ergänzt insbesondere die veraltete Übernahme und die interne Belegung bei Produktionsfortsetzung; hierfür werden neue vollständige CI-Läufe verlangt. Finale CI-Nachweise und Commit-SHA stehen im PR-Bericht und sind nicht durch diesen historischen Vorablauf ersetzt.

0.12.2 ergänzt den historischen Eigenzugriff nach Abteilungswechsel sowie AL-Korrekturrechte nach aktueller ausführender Zuständigkeit, mit SQL-, nativen JWT- und Browserregressionen.

0.12.3 fragt auch beim OP-Öffnen bzw. beim Einstieg eines nachfolgenden Mitarbeiters in eine offene Rüstung ausdrücklich nach Fortsetzung (Ja/Nein, dann Modus). Kein automatischer Zeitstart.

0.12.4 verknüpft das Korrektur-Moduslabel eindeutig und verwendet im Pausen-Browsertest den tatsächlichen Navigationsnamen `Start / Dashboard`. Keine Assertion gelockert.
