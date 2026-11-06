# Artikelstamm Phase 1 – Entwurf für Review (#7)

## Freigabestand

Die Produktionsseite setzt `HUMBEL_ARTICLE_MASTER_ENABLED = false`. Es gibt keine Artikel-RPCs und keinen neuen Menüpunkt, solange dieser Schalter aus ist. Kein Schema-/Policy-/Auth-/Daten-Write auf einem bestehenden Supabase-Projekt wurde ausgeführt. Die Datei unter supabase/migrations ist ausschließlich ein Vorschlag. Auch CI führt sie nur in einer wegwerfbaren PGlite-Instanz im Arbeitsspeicher aus. Es gibt keinen automatischen Migrations- oder Aktivierungsschritt.

Dieser Feature-Branch baut auf Draft-PR #2 (`codex/qa-0.5.17`, 6b8deb5) auf. Dessen Browser-/Visual-/Integritätsstrecke wird verwendet; Regeln/Vorlagen und Berichtstests aus #1 sind zusätzlich übernommen. Vor Integration nach main muss #2 mit dem aktuellen main synchronisiert und separat freigegeben werden. Der erforderliche e2e-Name aus #1 bleibt als Fail-closed-Kompatibilitätsjob hinter Quality Gate erhalten. Die bereits in main/PR #3 enthaltene Reload-Navigationskorrektur wird in den QA-Test übernommen. Ein exakter reversibler Testpatch prüft weiter den ursprünglichen geschützten Testhash; keine Geschäftsassertion oder Testidentität ändert sich. Keine zweite Browserpipeline, keine automatische Zusammenführung. Keine Änderungen der produktiven Pages-Einstellung.

## Vor Coding gelesener Ist-Zustand (06.11.2026)

Nur Katalog-SELECTs im Projekt Schichtplan: vorhandene Tabellen-/Spaltennamen, Policies und Funktionsdefinitionen; keine Anwendungszeilen und keine Auth-Benutzer wurden exportiert. `employees` besitzt `auth_user_id`, `role`, `is_active`, `department_id`. `departments` besitzt `leader_employee_id`, `active`. Die vorhandene Funktion `can_access_production_department` erlaubt aktive Admins und zugewiesene Abteilungsleiter. Keine Artikelstammtabellen vorhanden. `production_order_stations` ist ein Spannungsmodell und kein verlässlicher Erst-OP-Nachweis.

Wichtiger Befund: Neben restriktiveren employee-Policies bestehen permissive INSERT/UPDATE/DELETE-Policies für authenticated mit true. Daher können neue sensible Freigaben nicht sicher aus employees.role, einem lokalen Demo-Login oder user_metadata abgeleitet werden. Dieser PR verändert die Alt-Policies nicht. Ihre Prüfung/Korrektur ist ein eigener Freigabepunkt, kein behaupteter Sicherheitsfix dieses Features.

## Eigenständiger Berechtigungsnachweis

| Akteur | Voraussetzung | Rechte |
| --- | --- | --- |
| Artikel-/Produktionsverantwortlicher (Admin oder Abteilungsleitung) | Serverseitig provisionierter Eintrag in article_master_scopes für auth.uid + Abteilung | Lesen, Anlegen, Entwurf ändern, neue Revision, Freigabe, Archivierung innerhalb dieses Bereichs |
| Mitarbeiter erste OP | Serverseitiger article_master_assignments-Eintrag für auth.uid, Abteilung, genaue Artikelnummer, unabhängig geprüfte erste OP mit Beleg, Ablauf und Widerruf | Initialer Entwurf und eigene Inhalte dieses Artikels; keine Freigabe, Revision oder Archivierung |
| Sonstiger employee / tool_scanner / unzugewiesener Admin | Kein entsprechender separater Nachweis | Keine Artikeldaten oder Aktionen; vorhandene Produktionsrechte unverändert |
| Nicht angemeldet / anon | Kein gültiger Auth-Kontext | Keine technische Information oder API-Ausführung |

Scope-/Assignment-/Vier-Augen-Einstellungen werden ausschließlich über einen separat geprüften Server-Administrationsprozess mit serverseitigen Zugangsdaten provisioniert. Es gibt keinen Browser-Endpunkt zum Selbstzuweisen. Niemals den service_role-Schlüssel im Browser verwenden. Zuweisung nur nach unabhängiger Identitäts-/Rollen-/Erst-OP-Prüfung; kein automatisches Übernehmen aus dem unsicheren employees-Feld. Scanner niemals provisionieren. Ein lokaler Demo-Login ist keine Authentifizierung. Das UI verlangt eine Supabase-User-ID plus RPC-Capabilities, die SQL-API prüft auth.uid bei jedem Befehl erneut.

Neue Tabellen sind RLS-aktiv, öffentliche/anon/authenticated Tabellenrechte werden widerrufen. Tabellenzugriff wird nicht als API angeboten: nur explizit freigegebene RPCs. Public-Wrapper sind SECURITY INVOKER; notwendige privilegierte Implementierungen liegen in einem nicht exponierten Schema mit leerem search_path, expliziten EXECUTE-Grants und eigenen Auth-/Scope-Prüfungen. Keine user_metadata-basierten Entscheidungen. Auch direkt manipulierte RPC-Payloads werden geprüft. Eine Capability-Anzeige im Browser ist niemals der Sicherheitsmechanismus.

## Modell und Unveränderlichkeit

- Identität: Abteilung + getrimmte, großgeschriebene Artikelnummer (eindeutig). Nummer/Scope sind nach Anlage unveränderlich; Schreibweise wird für die Anzeige bewahrt. Prefix-Suche behandelt `%`, `_` und Backslash als Literale.
- Revisionen: separate UUID, aufsteigende Nummer pro Artikel, Entwurf/Freigegeben/Archiviert, Inhalt, row_version und Actor-/Zeitfelder. Genau ein offener Entwurf pro Artikel. Zeichnungsrevision ist ein Inhaltsfeld und nicht die interne Artikelrevision.
- Daten: feste Allowlist für Bezeichnung, Kunde, Zeichnung/-revision, Material, Abmessungen, Spezifikation, technische Hinweise, Programme, Referenzen, Prüfhinweise, interne Bemerkungen. Jeder Wert Text, höchstens 12000 Zeichen. Keine strukturierte BA-Nummer, Menge oder Auftragstermine. Freitext bleibt fachlich zu prüfen; das System kann solche Angaben nicht semantisch aus beliebigem Freitext entfernen.
- Freigabe friert Inhalte ein. Bearbeiten erfolgt nur auf einer neuen Entwurfsrevision mit kopierten Inhalten. Auch ein verknüpfter Entwurf kann nicht inhaltlich verändert werden. Alle Revisionen bleiben erhalten; keine Delete-API. Archivierung erhält Inhalte und Referenzen.
- BA-Verknüpfungen sind in einer neuen, in Phase 1 noch nicht beschreibbaren Link-Tabelle vorgesehen (FK mit RESTRICT). Bestehende production_orders/stations/counts und reale BA werden weder erweitert noch umgeschrieben. Anlegen einer BA aus Artikel gehört zu einer Folgephase.
- Mutationen laufen in einer SQL-Funktion innerhalb einer Transaktion, sperren Elternartikel und Revision in fester Reihenfolge und prüfen row_version. Unique Constraints verhindern gleiche Identitäten/Revisionen und mehrere offene Entwürfe. Audit wird nur beim erfolgreichen Commit angelegt. Keine automatische Wiederholung fehlgeschlagener Writes.
- Optional verlangt die Abteilung einen anderen Freigeber als den Ersteller. Standard aus, explizite serverseitige Einstellung. Creator, Reviewer und Archivierungsverantwortliche werden vom Auth-Kontext gesetzt, niemals aus Browserfeldern übernommen.

## Browser und sensible Daten

Artikel werden ausschließlich im Arbeitsspeicher des separaten Moduls gehalten. Kein localStorage-/IndexedDB-/Export-/Download-Pfad und keine Daten im allgemeinen persist()-State. Der Benutzerwechsel verwirft den Cache; ausstehende Antworten werden an eine Session-Generation gebunden. Nach Widerruf verweigert die DB weitere Zugriffe. Bereits sichtbare Informationen lassen sich nachträglich nicht vom Bildschirm eines Benutzers zurückholen; erneutes Laden entfernt nicht mehr lesbare Zeilen. DOM-Ausgabe wird escaped, Formularfelder haben zugängliche Labels und große mobile Eingabeflächen. Laufende Writes sperren erneute UI-Übermittlung. Fehler erhalten ungespeicherte Eingaben; Konflikte verlangen eine bewusste Aktualisierung statt stiller Überschreibung.

## Nachgewiesen und noch nicht nachgewiesen

Die PGlite-Tests wenden genau die vorgeschlagene SQL-Datei auf eine isolierte PostgreSQL-Engine an. auth.uid ist dabei ein ausdrücklich synthetischer Test-Shim; auth.users, Abteilungen und bestehende BA sind minimale künstliche Vertragsfixtures. Rollen-/EXECUTE-/DML-Denials, RPC-Scope/Assignments, Felder, Revisionen, Audit und Constraints werden tatsächlich von SQL ausgeführt, nicht nur von JavaScript nachgebildet. Browserfälle verwenden denselben SQL-Vertrag über eine lokale Node-Bridge; sämtlicher externer Seitenverkehr ist gesperrt.

**Das ist kein vollständiger Supabase-Integrationstest.** Kein GoTrue/PostgREST/JWT-/Session-Refresh-/Realtime-Nachweis, kein vollständiges Produktionsschema-Restore und kein Test unabhängiger paralleler DB-Verbindungen (PGlite hat eine exklusive Verbindung). Stale-Version- und Unique-Konflikte sind geprüft; echte konkurrierende Transaktionen bleiben separat erforderlich. Die bisherige allgemeine DB-Preflight bleibt BLOCKED, kein grüner Supabase-RLS-Claim.

Vor Aktivierung erforderlich: SQL/Security-Review; Prüfung der bestehenden employee-Policies; geprüfte separate Supabase-Testumgebung mit Schema-only-Restore; reale Auth-JWTs jeder Rolle und Widerruf/Expiry; direkte REST/RPC-BOLA-Versuche; konkurrierende Anlage/Revision/Release/Edit mit mehreren Connections; Advisors; Backup/Rollback-Plan; gesonderte Live-Migrations- und Feature-Freigabe. Anwendung und Rollback niemals automatisch aus diesem PR starten.

Referenzfixture: nur sicher erkannte Werte zu 32505, Spannflansch, ITEM A (Switzerland) Ltd., 90MnCrV8. Herkunft Seite 7/10, kein vollständiger Operationsplan. BA 812150 und Menge 48 existieren ausschließlich als separate synthetische Auftragsfixture, nie im Artikelinhalt.

Quellen: https://supabase.com/docs/guides/database/postgres/row-level-security ; https://supabase.com/docs/guides/database/functions ; https://pglite.dev/docs/ ; https://pglite.dev/docs/api . CLI-Changelog vom 04.11.2026 gelesen: v3 ist lokal als Standard; hier ausschließlich migration new verwendet, weder link noch db push/up aufgerufen.

## Stabile Bildvergleichs-Metadaten

Der erste CI-Lauf 42153076757 meldete sechs Bildunterschiede. Actual/Expected wurden heruntergeladen und die Differenzrechtecke geprüft: ausschließlich der feste Versionsknopf (v0.6.0 statt v0.5.17), keine anderen Pixel. Der Visual-Fixture prüft nun zuerst den echten Knopftext gegen APP_VERSION und setzt nur diesen Build-Metadatenwert für den Bildvergleich auf den bekannten Referenzwert, analog zur festgesetzten Uhrzeit. Ein separater funktionaler Fall prüft den echten Text v0.6.0. Kein Referenz-PNG, Screenshot-Ziel, Toleranzwert oder Geschäftsassert wurde geändert. Ein exakter reversibler Fixture-Patch schützt weiter den ursprünglichen Regressionstesthash.
