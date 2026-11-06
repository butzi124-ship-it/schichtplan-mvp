# Legacy-RLS- und Rollen-Härtung – getrennter, gesperrter Review-Entwurf

Status: **nur auf isoliertem strukturellem Testschema**, nicht produktionsreif. Grundlage ist Draft-PR #15 mit echten GoTrue-/PostgREST-/PostgreSQL-Tests. Das Quellschema ist **kein vollständiger autoritativer Restore**; echte Funktionskörper, ACLs, Indizes und Default-Grants fehlen. Keine produktive Änderung, keine Freigabe oder Veröffentlichung.

## Nachgewiesene Baseline-Schwachstellen

- 10 `public`-Tabellen mit deaktivierter RLS und rohen Anon-DML-Rechten: neun Planner-Tabellen sowie `tool_materials`.
- Die 9 Planner-Tabellen haben bereits permissive authenticated-Richtlinien mit `true`. **RLS einfach einschalten reicht daher nicht aus**, auch angemeldete Benutzer können bislang fremde Abwesenheiten sehen und ändern.
- `employees` besitzt `employees_insert`, `employees_update` und `employees_delete` mit `true`. Diese überstimmen die restriktiven Admin-Richtlinien durch permissive OR-Kombination; echte JWT-Testanfragen konnten fremde Mitarbeiter zu Admins machen.

## Testkandidat

Die Datei `supabase/review/legacy-security-hardening-isolated.sql` ist **durch Datenbankname und isolierten Sentinel technisch auf den Teststack begrenzt**. Sie ersetzt dort die permissiven Planner-Policies durch vorläufige Admin-only-Policies, sperrt Anon vollständig, reduziert Werkzeugmaterialien auf aktive angemeldete Mitarbeiter (Lesen) und Admins (Schreiben) und entfernt drei unbeschränkte Mitarbeiter-Schreibpolicies. Echte negative JWT-/REST-Tests prüfen diese Effekte und den unveränderten Fertigungszähler.

**Keine automatische Übernahme der Policy-Matrix in die Produktion:** Admin-only kann die normale Schichtplanung für Nicht-Admins blockieren; deshalb ist diese Zwischenlösung nur für den isolierten Nachweis vorgesehen. Vor einer produktiven Variante müssen verbindliche Leserechte für Einsatzpläne und Abwesenheitsdaten, Person-zu-Planer-Bindung, Stellvertretung, Schichttausch und Mitarbeiter-Selbstservice fachlich geklärt und auf vollständigem Restore regressionsgetestet werden. Beispielsweise sind reine `user_name`-Vergleiche nicht vertrauenswürdig, solange ihre Verbindung mit `auth.uid()` nicht eindeutig und serverseitig verifiziert ist.

## Gesonderte Freigabeschranken

1. Originales, gesondert genehmigtes **Schema-only**-Abbild ohne Personen- oder Fertigungsdatensätze beschaffen; die unvollständigen Hilfsfunktions-Substitute aus #15 ersetzen und tatsächliche Grants analysieren.
2. Eine vollständige Abteilungs-/Rollen-/Datenklassifikationsmatrix mit der Schichtplan-Fachseite abstimmen; insbesondere Mitarbeiter-Selbstservice und Schichtplan-Lesezugriff erhalten.
3. Im separaten, isolierten Supabase-Stack dieselben Auth-, Negativ-, Parallel-, Datenintegritäts-, Browser- und visuellen Regressionen mit der finalen Rollenmatrix ausführen. Test der korrekten Wiederherstellung alter Policy-Definitionen als Rückfallplan.
4. Erst danach ein produktionsfähiges, separat genehmigtes Migrationsskript mit Backup, Rollback, Wartungsfenster und schrittweiser Aktivierung vorbereiten.

**Kein Merge, Deploy, Produktions-RLS-Eingriff oder Datenmigration ohne ausdrückliche Freigabe.**

## Präzisierter Prüfstand (20:17 Europe/Berlin)

Die unterlagerte Testinfrastruktur #15 verwendet nun sechs tatsächlich per lesender Katalogabfrage verifizierte Originalfunktionen und 25 eigenständige Indexdefinitionen. Die ältere permissive Testsubstitution für Abteilungszugriff ist nicht mehr vorhanden; die finale Sicherheitsmatrix muss dennoch fachlich abgestimmt und mit einem vollständigen autoritativen Restore geprüft werden.
