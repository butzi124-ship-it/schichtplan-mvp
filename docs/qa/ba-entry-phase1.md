# Phase 1: Drei-Felder-BA-Einstieg (Issue #17, Spezifikation PR #18)

## Ausgangspunkt und Bestand

Basis: PR #18, `91d13e4bed23174625a808213f6f0517ea046891`, auf main `e764cf2e93382e847627d8bc7c802b02ff70c375`. #15 `25e9a6914e38fc791d9ecad13d92a45e30d19bb3` und #16 `bf14f4ed32fd8dc8e4e8374a5faa0c51a29416c7` bleiben unverändert und ungemergt. Dieser Implementierungs-PR wird auf #18 gestapelt, damit sein Diff ausschließlich Phase 1 enthält.

Vorhanden, aber deaktiviert: Artikelrevisionen/Freigabe, OP-Pläne, unveränderliche BA-/OP-Snapshots, Code128 und Parser für bestehendes `BA.OP` (Referenz `812150.100`), Maschinenbereitschaft und Prognose. Die alte BA-aus-Artikel-API verlangt Termin/Maschine und Verantwortlichen-Scope. Sie wird für bestehende Aufrufer unverändert erhalten. Der neue Einstieg verwendet denselben Parser, dieselben OP-/Artikel-Snapshottabellen und dieselbe Snapshot-/Barcode-Anzeige.

Bestehender Zähler: `production_station_counts.good_qty` pro `station_id`/`order_employee_id`; Ausschuss und Abklärung liegen an `production_order_stations`, Ursachen in der 6M-/QA-Erfassung und Ereignishistorie. `getProductionOrderCompletedGoodTotal` nutzt die letzte Spannung. Es wird kein zweiter Zähler eingeführt. Ein Literal `1AA` bzw. ein eindeutig dokumentierter Decoder hierfür ist im geprüften Hauptcode nicht vorhanden; seine Bedeutung wird nicht erfunden. Bestehende Codes und Zählerpfade werden nicht verändert. Hardware-Scan nicht getestet; Barcodeformat ist durch Repository-Referenz, Parser und vorhandenen unabhängigen Code128-Decoder belegt.

## Tatsächlich implementiert

- Neuer, standardmäßig deaktivierter Adapter am bestehenden Operationsmodul; kein pauschales Öffnen von `canAccessProduction()` für Mitarbeiter.
- BA-Anlage mit genau drei manuellen Pflichtfeldern: Humbel-BA-Nummer, Artikelnummer, positive ganze Stückzahl. Führende BA-Nullen bleiben erhalten. Kein ERP-Abgleich oder Nummerngenerator.
- Server wählt die höchste freigegebene Revision des exakt passenden und autorisierten Artikels. Mehrdeutigkeit über mehrere Gültigkeitsbereiche wird abgewiesen; keine zufällige Bereichs-/AL-Zuordnung.
- Berechtigung: aktiver Mitarbeiter, aktive Abteilung und unabhängig provisionierter Artikel-Scope oder nicht abgelaufener/nicht widerrufener Artikelauftrag mit exakt passender erster OP des Plans. Ein beliebiger Mitarbeiter mit dieser Erst-OP-Zulassung darf anlegen; weder Artikelersteller noch Freigeber muss mit ihm identisch sein. Die erste OP wird aus der Planreihenfolge bestimmt, nicht als OP-Nummer 1 angenommen.
- Mitarbeiterrollen/JWT-user_metadata allein erzeugen keine Berechtigung. Die aus #15/#16 bekannte manipulierbare Mitarbeiterrolle reicht ohne unabhängigen Scope/Grant nicht aus. Scanner und inaktive Mitarbeiter werden zusätzlich ausgeschlossen. Bestehende produktive Rechte werden nicht geändert.
- Atomare Anlage von BA, unveränderlichem Artikel-/OP-Snapshot, Revisionslink und unveränderlichem Anlageaudit; OP-Reihenfolge, Standardmaschinen und Sollzeiten werden unverändert kopiert. Keine Stations-, Zähler-, Zeit-, Ausschuss- oder Checklistenzeilen werden übernommen/erzeugt.
- Serialisierung mit derselben `operations-ba:`-Sperre und Parent/Revision-Reihenfolge wie die bestehende API. Ein vorhandener eindeutiger BA wird geöffnet, nicht überschrieben. Doppelte Legacy-Nummern werden nicht bereinigt, sondern als mehrdeutig abgewiesen. Bestehender Trigger schützt neue Snapshot-BA auch gegen konkurrierende Legacy-Schreibwege.
- Direkter Barcode-Aufruf oder manuelle BA-/OP-Suche selektiert dieselbe OP ohne Mengenbuchung. Transportfehler bewahren Eingaben; wiederholte Klicks werden im UI verriegelt und serverseitig idempotent behandelt. Sitzungswechsel verwerfen sensible Anzeigezustände.
- Bestehende Legacy-BA werden für autorisierte Verantwortlichen-Scopes als Bestand geöffnet, ohne nachträglichen OP-Snapshot. Für Mitarbeiter wird ihre bislang nicht belegte erste Legacy-OP-Zulassung nicht aus Spannungsnummern erraten.
- Hauptverantwortlicher AL bleibt ausdrücklich `pending_rule`; keine unbestätigte Zuordnung implementiert und kein viertes Pflichtfeld.

## SQL-Entwurf, Schalter und Grenzen

`supabase/review/ba-entry-phase1.sql` ist ein nicht registrierter Reviewentwurf, keine auf einer gehosteten Datenbank ausgeführte Migration. Er setzt die noch nicht produktiv abgenommenen #7/#8-Strukturen voraus. Neu: zwei RLS-geschützte Tabellen für serverseitige Aktivierung (Standard false) und unveränderliches Anlageaudit, ein geprüfter RPC mit privaten Implementierungen. Keine Client-DML-Rechte auf neue Tabellen, kein anonymer RPC, keine Änderung bestehender Legacy-Policies oder Authentifizierung. Snapshot-`due_date` wird im Entwurf nullable; vorhandene Werte bleiben bestehen und die alte API verlangt weiter einen Termin.

`HUMBEL_BA_ENTRY_ENABLED`, Artikel- und Operationsschalter bleiben false. Ohne Aktivierung im Server-Setting ist auch der direkte neue RPC gesperrt. Kein Merge, Deployment, produktive Migration, Workerstart oder Feature-Aktivierung.

Der vorhandene Legacy-BA-Datensatz verlangt `status` und `started_at`; Phase 1 übernimmt dessen vorhandene Anlagekonvention. Diese Felder sind kein nachgewiesener tatsächlicher OP-Arbeitsbeginn und kein neu gestarteter persönlicher Zeitzähler. Planung/OP-Ausführung folgen in späteren Phasen.

Noch nicht umgesetzt bzw. vor Freigabe offen:

- Fachliche AL-Hauptverantwortungsregel und Vertretung; sichere Provisionierung der realen Erst-OP-Grants.
- Vollständiger Originalschema-/Ownership-/ACL-/Default-Privilege-Abgleich und autoritativer Restore. Native Tests verwenden ausschließlich #15s strukturellen Katalogspiegel. Kritische Legacy-RLS-Befunde bleiben offen.
- Explizite Spannung-zu-OP-Vorlagen und neue Pflichtmaschinenbindungen fehlen im bisherigen Planmodell; keine automatische Legacy-Konvertierung. Erweiterung in den zugehörigen Folgephasen.
- OP-Öffnen zeigt den bestehenden Snapshot; kein Rüst-/Fertigungsstart und keine zweite Produktionszählung. Maschinenzuweisung/Abschnitte, Rüstung, Spannungsintegration, Teilmengen, Mengenklärfälle, Nachfertigung, Abschlüsse und neue Planung aus Phase 2–6 sind nicht Teil dieses PR.
- Späterer gezielter Artikelabgleich ist nicht implementiert; automatische Änderungen bestehender BA bleiben ausgeschlossen.
- Nicht eingeplante neue BA haben keinen erfundenen Endtermin oder BA-Maschinenwert. Produktionsfreigabe und Prognose benötigen die späteren Planungs-/Sicherheitsnachweise.

## Validierung

Lokale Pflichtprüfungen und Vertragstests werden im PR mit tatsächlichen Ergebnissen angegeben. PGlite-Vertragstests sind keine echte Supabase-/JWT-/Parallelabnahme. Lokal fehlt ein funktionsfähiges Playwright-Browserpaket; Installationsversuch scheiterte an ungültigem Downloadarchiv, der Agent-Seed konnte daher nicht starten. Der MCP-Handshake selbst funktioniert. CI führt die unveränderten bestehenden Browser-/Visual-Suites plus die neue Phase-1-Suite aus; keine Snapshotaktualisierung.

`BA Entry Isolated Acceptance` verwendet #15s Infrastruktur auf festem SHA in einem separaten Checkout und überlagert ausschließlich diesen SQL-Entwurf und die Phase-1-Testextension. Baseline-14 bleiben aktiv; zusätzliche echte JWT-/REST-, Raw-Table-Denial-, Rollback- und Mehrverbindungsfälle laufen nur im internen Wegwerfstack. Keine Verbindung zur produktiven Supabase-Instanz, keine produktiven Daten oder Secrets. Ergebnisstatus und getestete SHAs sind in CI/PR dokumentiert. Dieser zusätzliche DB-Check ist weiterhin separat vom bestehenden e2e-Ruleset zu bewerten.

Die Integritätsprüfung reversiert nur die drei expliziten neuen App-Hooks vor dem unveränderten Legacy-Business-Hash. Bestehende 251 Legacy-Testidentitäten und 33 Bildreferenzen bleiben geschützt. Der alte Artikel-Footer-Test erhält die bereits in #15 geprüfte exakte Kopplung an APP_VERSION und VERSION_LOG statt einer historischen festen Versionsnummer; keine Geschäftsassertion entfällt.

## Rücknahme und Freigabe

Vor Veröffentlichung reichen deaktivierte Client-/Serverschalter; bestehender Pfad bleibt verfügbar. Bei späterer DB-Einführung sind Schema-/ACL-Abnahme und Wiederherstellungsplan gesondert erforderlich. Nach echten neuen BA-Buchungen dürfen neue Snapshots/Audits nicht gelöscht und nullable Termine nicht blind wieder auf NOT NULL gesetzt werden. Abschalten bewahrt die Daten. Jeder Merge, jede produktive SQL-Ausführung und jede Aktivierung benötigt eigene ausdrückliche Freigabe.
