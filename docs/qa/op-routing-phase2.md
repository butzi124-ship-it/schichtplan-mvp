# Phase 2 – OP-Maschinenzuordnung und operative Verantwortung

## Basis und Umfang

Basis: Phase 1, PR #19, Commit `8348fadc7bd1f9042b4147bd47568cb484547e3a` (0.10.1). Spezifikation PR #18 / Issue #17. Eigener gestapelter Branch `feature/ba-op-phase2`; kein Merge, kein Deployment, keine produktive SQL-/RLS-/Auth-Änderung und keine Aktivierung. Phase 1 wird nicht neu implementiert.

Bestehende Grundlagen: `production_order_operations` und Artikel-/BA-Snapshots sind unveränderlich; `production_operation_execution` hält aktuelle Maschine, Status und kumulative Ausführungswerte. Die bestehende Readiness-API bleibt für physische Übergabe, Start, Unterbrechung und Abschluss erhalten. Bestehende Spannungs-/Mitarbeiterzähler, Ausschuss-/6M-Buchungen, Ist-Zeiten und Historien werden nicht umgeschrieben.

## Implementierte Abläufe

- OP-Barcode bzw. Erst-OP-Anlage öffnet bei aktivem Routing direkt dieselbe ausgewählte OP ohne zweite OP-Auswahl. Beim Öffnen einer OP: konkrete Maschine aus dem eingefrorenen Artikel-OP-Snapshot vorauswählen; nach erfolgter Zuordnung aktuelle BA-OP-Maschine anzeigen. Eine fehlende historische Standardmaschine wird nicht erfunden.
- Mitarbeiter dürfen geeignete, aktive Maschinen ihrer ausführenden Abteilung wählen. Eignung einer Alternative benötigt eine ausdrücklich provisionierte Maschinen-/Maschinengruppen-Zuordnung und einen gültigen Maschinen-Arbeitsgrant. Die bereits freigegebene konkrete Standardmaschine gilt als dokumentierte Vorgabe, nicht als automatisch zwingende Bindung.
- Aktiver zuständiger AL, aktive legitimierte Stellvertretung bzw. unabhängig berechtigter Admin können eine BA-OP verbindlich festlegen, auch in eine andere Abteilung. Normale Mitarbeiter können diese Bindung nicht aufheben oder ändern. Eine technische zwingende Maschinenbindung ist vorrangig und wird auch bei Rückgabe geprüft.
- Die ausführende Abteilung folgt der tatsächlich zugeordneten Maschine; die ursprüngliche Abteilung bleibt unveränderlich in der OP-Zuordnung vermerkt. Ursprüngliche AL behalten Übersicht, der ausführende AL übernimmt operative Rechte. BA-Dept, Artikelrevision, OP-ID und BA-Hauptverantwortung werden nicht verändert.
- Übernahme oder direkte Rückgabe erzeugen atomar ein unveränderliches Audit und dauerhafte Benachrichtigungen für verifizierte aktive AL/Vertreter der empfangenden Abteilung. Keine Zustimmung des Empfängers erforderlich. Fehlt ein verifizierter Empfänger, wird die Transaktion abgewiesen, statt eine Verantwortung zu erraten.
- Benachrichtigungen innerhalb der App; bei aktivem angemeldetem Modul spätestens beim nächsten 30-Sekunden-Abruf oder manuellen Neuladen sichtbar. Kein externer E-Mail-/Push-Dienst.
- Bei Rückgabe wird eine zulässige Maschine der ursprünglichen Abteilung ausdrücklich gewählt. Es wird keine bislang unbeschlossene Rückgabe-Standardregel geraten.
- Jeder tatsächliche Maschinenwechsel schließt den aktuellen Maschinenabschnitt und öffnet einen neuen Abschnitt derselben OP. Vorhandene Mengen und der erste Arbeitsbeginn bleiben bestehen; physischer Empfang an der neuen Maschine wird nicht erfunden. Aktuelle Ausführung pausiert bzw. wartet auf ausdrückliche Übergabe/Arbeitsaufnahme.
- Alle offenen persönlichen Segmente dieser OP werden am selben Serverzeitpunkt beendet. Bei aktiver interner/externer Rüstung ist ein aktiver administrativ konfigurierter Abbruchgrund zwingend; freie Gründe werden nicht automatisch angelegt. Andere OPs und deren Teilnehmer bleiben unberührt. Kein automatischer neuer Zeit-/Rüststart.
- Auftrag → Ausführung → Routing/Segmente werden konsistent gesperrt. Versionsprüfung verhindert verlorene Änderungen; Request-ID plus unveränderter Payload macht Wiederholung idempotent. Audit-/Benachrichtigungsfehler rollen auch Abschnittswechsel und persönliche Unterbrechungen zurück.
- Ein Trigger verhindert, dass die bisherige Readiness-Assign-API eine getrackte Maschine ohne den geprüften Routing-Ablauf ändert. Bestehende Arbeits-/Übergabe- und Planungsfunktionen bleiben daneben verfügbar; aktuelle Abteilungsprüfungen beachten die neue ausführende Abteilung.

## Datenmodell als nicht angewandter Reviewentwurf

`supabase/review/op-routing-phase2.sql` ist **keine registrierte/ausgeführte Produktionsmigration**. Voraussetzung sind die abgenommenen Strukturen #7/#8/#9 und Phase 1. Zusätzlich enthält der Entwurf die explizite Erweiterung des bekannten Legacy-Constraints employees_role_check um die bereits in App und Originalfunktion verwendete Rolle department_admin. Der ursprüngliche Katalog erlaubte nur admin/employee/tool_scanner; der erste native Lauf stoppte folgerichtig mit SQLSTATE 23514 vor den neuen AL-Fällen. Die Korrektur ist **Bestandteil des Review-SQL**, keine abgeschwächte Testfixture: unbekannte/missing Constraintdefinitionen blockieren die Anwendung. Keine bestehenden Rollenwerte oder Mitarbeiter werden konvertiert, keine Authority automatisch vergeben. Dieser zusätzliche Schema-/Berechtigungsaspekt benötigt gesonderte produktive Sicherheitsabnahme.

Zehn neue Tabellen:

| Tabelle | Zweck |
|---|---|
| production_routing_settings | Serverfreigabe, standardmäßig false |
| production_routing_authorities | Unabhängig provisionierte AL-/Admin-/Vertretungsrechte, Zeitfenster und Widerruf |
| production_routing_capabilities | Bestätigte technische Eignung von Maschine und Maschinengruppe |
| production_routing_abort_reasons | Administrativ provisionierte Abbruchgründe; keine erfundene produktive Liste |
| production_operation_machine_constraints | Bestätigte zwingende BA-OP-Maschinenbindung mit Herkunftsnachweis |
| production_operation_routing | Aktuelle Maschine, ursprüngliche/ausführende Abteilung, BA-OP-Bindung, Version |
| production_operation_machine_sections | Unveränderlich erhaltene abgeschlossene Maschinenabschnitte mit Zustandsreferenzen |
| production_operation_activity_segments | Minimaler Phase-3-Anschluss für atomare persönliche Unterbrechungen |
| production_operation_routing_events | Unveränderlicher Alt-/Neu-/Actor-/Grund-/Request-Audit |
| production_operation_routing_notifications | Persistente App-Benachrichtigungen pro Wechsel und Empfänger |

Alle Tabellen: RLS aktiv, keine anonymen oder authenticated Raw-Table-Rechte. Zwei geprüfte öffentliche invoker-RPCs; privilegierte Implementierungen im bestehenden privaten Schema. Neue private Helfer haben keine direkten Client-EXECUTE-Rechte. Keine Änderung an produktiven oder alten Legacy-RLS-Policies.

AL-Rechte benötigen sowohl eine unabhängig provisionierte Authority als auch aktive reale Mitarbeiter-/Leiterzuordnung `departments.leader_employee_id`. Stellvertretung benötigt ein gültiges Zeitfenster und weiterhin gültigen aktiven Leiter als Principal. Ein alleiniger, nach früherem Legacy-Befund manipulierbarer `employees.role`-Wert oder JWT-user_metadata verleiht keine AL-/Admin-Rechte. Authorities, relevante Mitarbeiter, Departments und Grants werden während des Wechsels gesperrt. Echte Provisionierung wird erst nach gesonderter Sicherheitsabnahme festgelegt.

Native Abnahme 0.11.1: 26/27 bestanden; der erwartete Stale-Konflikt lief in einen HTTP-Timeout. Ursache: Der neue Code verwendete SQLSTATE `40001` für deterministische fachliche Konflikte; PostgREST 14 kann diesen Fehler unbegrenzt wiederholen ([Supabase-Hinweis](https://supabase.com/kb/troubleshooting/high-cpu-and-infinite-transaction-retries-when-using-custom-error-codes-in-rpc-functions-77326b)). Neue Routing-Konflikte verwenden nun `PT409` (HTTP 409); Versionsvergleich, Ablehnung, Rollback und Assertions bleiben bestehen. Ein nativer HTTP-Regressionstest verlangt Code und HTTP-Status. Die anfängliche JIT-Vermutung wurde verworfen; keine JIT-/Timeout-Änderung. Ältere `40001`-Pfade in Readiness/Forecast bleiben ein separat zu prüfendes Integrationsrisiko; sie werden in Phase 2 nicht pauschal umgeschrieben.

Die bestehende `readiness_allowed`-Prüfung wird für getrackte OPs erweitert, für nicht getrackte OPs unverändert delegiert. Drei konkrete Abteilungsprädikate in Readiness-Command/View werden auf die ausführende Abteilung angepasst; der Reviewentwurf prüft exakte bekannte Voraussetzungstexte und bricht bei Abweichung ab. Kein pauschales Abschalten von Berechtigungen.

## Bewusste Integrationsgrenzen

1. **Kein zweiter Stückzahlzähler.** Maschinenabschnitte speichern Grenzen und referenzierte Ausführungszustände. Die endgültige Zuordnung einzelner bestehender Zählerereignisse/Spannungen zu OP und Abschnitt gehört in Phase 4 und ist noch nicht aktiv. Eine exakte historische Stückzahlaufteilung auf frühere Maschinen wird nicht rückwirkend erfunden. Kein Counter-Reset oder Kopieren von Gut-/Ausschuss-/Abklärungsmengen.
2. Bereits vorhandene Ist-Zeitfelder und historische Buchungen bleiben unverändert. Der minimale neue Zeitsegment-Anschluss hat **keine öffentliche Start-/Korrektur-API** und keine Anwesenheits- oder MOBATIME-Funktion. Aktive Segmente dürfen nicht nachträglich in einen bereits geschlossenen Abschnitt gestartet werden; der Insert-Guard nutzt dieselbe BA-Sperre wie ein Maschinenwechsel. Vollständige Bedienung entsteht in Phase 3.
3. Interner/externer Wechsel, gemeinsamer Rüstvorgang, persönlicher Start/Pause/Fortsetzung/Abschluss, Schicht-/Arbeitsplatzwechsel und Korrektur-Audit sind noch Phase 3. Diese PR startet keine Rüstung automatisch und behauptet keine Anbindung unbekannter alter Timer.
4. Erfassung/Provisionierung zwingender technischer Artikelbindungen im Stammdateneditor und der Eignungs-/Vertretungskonfiguration ist noch separat zu integrieren. Servermodell und Sperrprüfung sind vorhanden; normale optionale historische Artikelmaschinen werden nicht ohne Freigabe zu Pflichtmaschinen konvertiert. BA-spezifische verbindliche AL-Festlegung ist in der App implementiert.
5. Regeln für Teilmengen, Mengenklärfälle, Nachfertigung und neue Abschlüsse bleiben Phase 4/5. Bestehender Readiness-Gesamtmengenpfad bleibt bis zu deren kontrollierter Weiterentwicklung erhalten. Neue Personal-/Maschinenprognosen bleiben Phase 6; alte Prognosetests werden weiter ausgeführt.
6. BA-Hauptverantwortlichen-Regel ist wie in Phase 1 offen. Eine OP-Verlagerung ordnet keinen neuen BA-AL zu.

## Tests und Freigabegrenzen

Neue Vertragstests: Standard/Alternative, fachliche Eignung, Bindung, abteilungsübergreifender AL, aktiver/abgelaufener/widerrufener Stellvertreter, falscher Principal, Empfänger, Rückgabe, unveränderte BA-/OP-/Altzähler-/Ist-Zeitdaten, gemeinsamer Segmentabschluss und unberührte fremde OP, Replay/Stale, Audit-/Benachrichtigungsrollback, Umgehungsversuch über Legacy-API, weiterhin nutzbare bestehende Work-APIs, Flag-off und Rolleneskalation.

Neue Browserfälle auf Desktop, Smartphone, Tablet, Firefox und WebKit ergänzen sämtliche bestehenden Tests. Keine geschützte Assertion oder Bildreferenz wird entfernt/gelockert. Pins werden ausschließlich für konkret veränderte Module und neue Tests aktualisiert; zentraler Legacy-Business-Hash und Phase-1-Tests bleiben erhalten. Die neue Oberfläche ergänzt die bisherige Readiness-Oberfläche, statt Arbeitsfunktionen zu entfernen.

Native CI kombiniert die bisherigen 14 Baseline- und 6 Phase-1-Fälle mit Phase 2 über die unveränderte isolierte Infrastruktur aus #15 auf SHA `25e9a6914e38fc791d9ecad13d92a45e30d19bb3`. Aktuelle Review-SQL und Testextensions werden in separatem Checkout überlagert. Echte GoTrue-JWT-/REST-, RLS-/Raw-Denial-, Mehrverbindungs-, Audit- und Rückgabeprüfungen ausschließlich im internen Wegwerfstack. **Kein autoritativer Produktionsrestore** und keine Behebung der offenen Legacy-Befunde aus #15/#16 behauptet. Deren Originalschema-/Ownership-/ACL-Abgleich und Sicherheitsfreigabe bleiben erforderlich. Die dortige Forecast-Lock-Korrektur ist Teil der gepinnten isolierten Testbasis und vor späterer Integration gesondert zu berücksichtigen.

Lokale npm-/Syntax-/Unit-/Vertrags-/Integritätsprüfungen und tatsächliche CI-Ergebnisse mit SHA/Links stehen im PR. Lokale Browser-/Visual-Prüfung bleibt mangels Browserpaketen/gepinntem CI-Container nicht durchgeführt; deren vollständige Ausführung erfolgt in GitHub Actions. Keine Tests auf der Live-App oder mit produktiven Datensätzen.

`HUMBEL_OP_ROUTING_ENABLED` und serverseitiges Routing-Setting bleiben false; vorhandene Modulschalter ebenfalls unverändert deaktiviert. Vor Aktivierung: vollständige Tests, originale Schema-/ACL-Prüfung, vertrauenswürdige Authority-/Eignungs-/Abbruchgrund-Provisionierung, ergänzende Phasenintegrationen und gesonderte Freigabe. Abgeschlossene Sections/Audits dürfen beim Rückbau nicht gelöscht werden. Kein Merge, Deployment, Worker oder produktiver SQL-Lauf durch diese Phase.

Native Abnahme 0.11.2: 27/27 bestanden (PR-Lauf 42373347656). Browserartefakte aus 0.11.1 zeigten ungenaue implizite Feldlabels (Labeltext enthielt Select-Optionen); 0.11.3 verknüpft die drei Labels explizit über eindeutige IDs, ohne Testlocators oder Assertions zu ändern. Vollständiger Browserlauf folgt für diesen Stand.

Browserabnahme 0.11.3: 638/643 bestanden; alle fünf AL-Varianten warteten im bestehenden Abteilungs-Starttab auf die noch nicht geöffnete Routing-Maske. 0.11.4 ergänzt den regulären Klick auf „OP-Maschinenübersicht“ im Login-Helfer (wie im bestehenden Readiness-Test). Kein künstlicher Direktaufruf, keine Produktnavigation geändert, keine Assertion entfernt oder gelockert. Vollständiger Wiederholungslauf für den Folgecommit erforderlich.
