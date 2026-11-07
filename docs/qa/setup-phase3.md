# Phase 3 – Rüstzeiterfassung (Reviewstand 0.12.0)

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
| Phase 3 SQL/Unit | 14/14 bestanden, eingebettetes PGlite, keine native Parallelität |
| Artikel / OP / Readiness / Forecast / BA / Routing | 8 + 15 + 17 + 30 + 9 + 10 = 89/89 bestanden |
| test:ci / test:qa | 2/2 und 9/9 bestanden |
| qa:check | 251 Legacy-Identitäten, alter Business-Hash nach expliziter Hook-Umkehr und 33 Bildreferenzen erhalten |
| qa:agent-server | bestanden |
| Lokale test:agents / test:e2e | blockiert: Playwright-Browser fehlen; Download liefert ungültiges ZIP. Kein bestandener Browsertest behauptet |
| Lokale Visuals | blockiert: vorgeschriebener gepinnter Container fehlt |
| Allgemeiner DB-Preflight | BLOCKED_NOT_TESTED; kein autoritativer Restore/isolierter nativer Stack lokal |

30 neue Phase-3-Browserausführungen (6 Fälle × Desktop/Smartphone/Tablet/Firefox/WebKit) sind der vollständigen unveränderten bisherigen Matrix hinzugefügt. Insgesamt 673 Browserfälle. Echte Geräte/Scannerhardware nicht getestet; Gerätebezeichnungen meinen Browser-/Viewportprofile.

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
