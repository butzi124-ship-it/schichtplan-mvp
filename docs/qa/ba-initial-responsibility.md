# Automatische BA-Hauptverantwortung – Version 0.15.0

Gestapelt auf PR #23 / `feature/ba-op-phase5`, verifizierter Parent `05b57de2f1dfcedf5575a4064d8ff79acafa0d1b`, Version 0.14.2. Neuer Branch `feature/ba-initial-responsibility`. Refs #17 und verbindliche Spezifikation `docs/requirements/ba-op-shopfloor-2026-11-07.md`; die neue ausdrückliche Entscheidung vom 08.11.2026 ersetzt die bisher offene Erstzuordnungsregel. APP_VERSION **0.15.0**, VERSION_LOG **2026-11-08 07:13 Europe/Berlin**. Commit/PR-Link stehen in der PR-Beschreibung.

## Verhalten

Der bestehende `ba_entry_command` behält genau seine drei Anlageangaben. Ein separater Reviewentwurf `supabase/review/ba-initial-responsibility.sql` ergänzt ausschließlich den bisherigen geprüften Einstieg; die SQL der Phasen 1–5 bleibt unverändert. Auch der bisherige Fünf-Felder-Einstieg wird nicht verändert. Die neue Regel gilt ausschließlich für neue BA am Phase-1-/OP-1-Einstieg.

Die erste OP wird nach gespeicherter Position aus dem freigegebenen Artikel-/OP-Snapshot ermittelt, nicht aus einer angenommenen OP-Nummer 1 oder aus einer Bedienereingabe. Das vorhandene OP-Modell erbt die zuständige Abteilung vom Artikel-Snapshot; es hat kein eigenständiges frei eingebbares OP-Abteilungsfeld. `routing_department` liefert bei der Anlage diese Abteilung. Keine Maschinenwahl, Routingzeile oder operative Zuweisung wird durch die Erstzuordnung erzeugt.

Genau eine aktuell gültige, nicht widerrufene unabhängige `leader`-Authority mit aktivem AL-Profil muss vorhanden sein und zur vertrauenswürdigen `departments.leader_employee_id` passen. Eine veränderbare Profilrolle allein verleiht keine Rechte. Administratoren und Stellvertreter werden nicht als automatische Ersatz-AL gewählt. Mehrere widersprüchliche aktive Leader-Authorities führen auch dann zu offenem Ergebnis, wenn der Abteilungszeiger einen bevorzugt. Ungültige/abgelaufene/zukünftige Authorities, inaktive Mitarbeiter, Scanner und fehlende Abteilungszuordnung werden nicht zugeordnet.

Ohne gültigen eindeutigen AL wird der BA erfolgreich angelegt und **„Hauptverantwortung offen“** angezeigt. Die bereits vorhandene geprüfte administrative Erstzuordnung mit Kataloggrund bleibt verfügbar. Spätere Änderungen erfolgen über die vorhandene Übertragung mit Pflichtgrund und Audit; normale Mitarbeiter erhalten keine Verwaltungsrechte.

## Atomizität, Historie und Zugriff

Die Zuordnung speichert `principal_user_id`/`principal_department_id` im vorhandenen Lifecycle-Datensatz. Eine erfolgreiche oder offene Entscheidung erzeugt genau ein unveränderliches `initial_responsibility`-Ereignis im vorhandenen Audit: Actor, Serverzeit, vorheriger/nachheriger Stand, Quellrevision, erste OP, Abteilung, Kandidatenzahl, Regelversion, Ergebnisgrund und verwendete Authority. Erfolgreiche Zuordnung benachrichtigt intern über die bestehende Lifecycle-Funktion. Der unveränderliche Phase-1-Anlagebeleg bleibt als ursprünglicher Anlagebeleg erhalten; die neue Entscheidung wird ergänzend protokolliert und nicht durch Umschreiben alter Belege verborgen.

Bestehender globaler Lifecycle-Lock, geordnete Authority-/Abteilungs-/Mitarbeitersperren und erneute Schalterprüfung serialisieren Anlage/Übertragung. Kurze SHARE-Tabellensperren auf den drei Zuordnungsquellen verhindern zusätzlich konkurrierende neu eingefügte Profile/Authorities (Phantome), nicht nur Widerruf bereits vorhandener Zeilen. Unter hoher Last ist diese bewusst konservative Sperrstrategie separat zu messen; sie wird ausschließlich im neuen aktivierten Anlagepfad verwendet.

Die ursprüngliche Phase-1-Eindeutigkeitsprüfung und BA-Sperre bleiben erhalten. Wiederholte oder parallele Anlage derselben BA liefert den bestehenden BA und erzeugt keine zweite Entscheidung. Öffnen, später geänderte Leitungsdaten, ein neues Login oder Retry nach einer verlorenen Antwort weisen niemals erneut zu. Insbesondere wird eine spätere geprüfte Übertragung nicht überschrieben. Bereits bestehende BA werden nicht nachträglich automatisch zugeordnet.

Audit-, Fremdschlüssel- oder Benachrichtigungsfehler rollen die gesamte Anlage samt Snapshot und Verantwortung zurück. Die Aufteilung in neue private Wrapper und unverändert erhaltene geprüfte Originalfunktionen ist kein zusätzlicher Clientzugang: umbenannte Originale/Leser/Helfer sind für `PUBLIC`, `anon` und `authenticated` gesperrt; nur der überprüfte bestehende Command bleibt aufrufbar. Der öffentliche Invoker wird mit unveränderter Signatur erneuert, damit gepoolte Sessions das neue Ziel auflösen. Keine neue öffentliche Definer-Funktion.

Alle bisherigen Fertigungsdaten, Spannungszähler, Rüstzeiten und OP-Verantwortlichkeiten bleiben unverändert. Ein neuer Test erzeugt reale synthetische Gutteilbuchungen und aktive Rüstzeit, legt einen weiteren BA an und vergleicht alle bisherigen Zähler-/Zeit-/Abschnitts-/Snapshotdaten exakt.

## Deaktivierter Reviewstand

Neue Einstellung `production_initial_responsibility_settings.enabled=false`, RLS aktiviert und alle direkten Clientrechte entzogen. Zusätzlich müssen die vorhandenen Phase-5-/Vorphasenschalter aktiv sein. Bestehende Frontendschalter und alle Produktiveinstellungen bleiben unverändert aus. Ohne neuen Serverschalter bleibt das bisherige Verhalten einschließlich bestehender Test-Fixtures erhalten. Die Anzeige verarbeitet neue bestätigte Serverzustände `assigned`/`open`, mit unverändertem Legacy-Fallback für noch nicht erweiterte Server.

SQL bleibt **unregistrierter, nicht angewandter Reviewentwurf**. Keine Migration registriert, keine produktive Verbindung, keine Aktivierung, kein Merge oder Deployment. Offene Befunde zu Ownership, Originalrestore, RLS/ACL und vertrauenswürdiger Berechtigungsprovisionierung aus PR #15/#16 bleiben unverändert offen. Erst fachlich/technisch geprüfte Provisionierung darf später aktiviert werden.

## Lokale Prüfungen

| Prüfung | Ergebnis |
|---|---|
| `npm ci --ignore-scripts`, beide Pflicht-`node --check`, `check:syntax` einschließlich neuer Tests | bestanden |
| Lifecycle-SQL-/Berechtigungs-/Adapter-/Installationssuite | **37/37**, davon 10 neue Erstzuordnungs-/Integritätsfälle |
| Unverengte bestehende Artikel-/OP-/Readiness-/BA-/Routing-/Rüst-/Mengen-/Prognoseverträge | **131/131** |
| CI-Bericht / QA-Gates | **2/2**, **9/9** |
| `qa:check`, `qa:agent-server`, `git diff --check` | bestanden |
| Vollständige Browsermatrix | **733 Fälle aufgerufen**: 718 bisherige unverändert + 3 neue Fälle in allen 5 Profilen; Browserstart wegen fehlender Binaries blockiert |
| Agent-Seed | Browserstart blockiert |
| Gepinnte Visuals | BLOCKED_NOT_TESTED, vorgeschriebener Playwright-1.62.1-noble-Container fehlt; keine Referenzänderung |
| DB-Preflight / native JWT-/Mehrverbindungsabnahme | BLOCKED_NOT_TESTED, isolierte lokale Dockerlaufzeit fehlt |

Der QA-Schutz behält alle bisherigen Hashes, Assertions, 251 Legacy-Identitäten und 33 Bildreferenzen. Eine zusätzliche exakt begrenzte Integrationsliste erlaubt ausschließlich die beauftragte Verantwortungsanzeige in `js/ba-entry.js`; jede andere Änderung bleibt durch den ursprünglichen Hash geschützt. Keine bestehenden Tests geändert, entfernt oder übersprungen, keine Filter verengt, keine Assertions oder Bildtoleranzen gelockert.

## Vorbereitete native CI und offene Abnahme

Die native Erweiterung führt die vorhandene Phase-5-Erweiterung einschließlich aller bisherigen **52** Fälle unverändert zuerst aus und installiert danach den neuen Reviewentwurf. **Zehn zusätzliche native Fälle, 62 insgesamt**: Schalter/Raw-/Anon-Schutz, reale Drei-Felder-JWT-Anlage, abgelaufene/widerrufene/zukünftige Authorities, Inaktivität/Scanner/fehlende Zuordnung, Mehrdeutigkeit, administrative Nachzuordnung mit Pflichtgrund, zwei echte gleichzeitige Anlageverbindungen, tatsächlich beobachtete Sperre bei konkurrierendem Widerruf, Auditrollback und unveränderte Verantwortung bei Retries/geänderter Leitung sowie verbotene Client-Zielvorgabe.

Die drei neuen Browserfälle prüfen Drei-Felder-Anlage mit automatischem AL, offenes Ergebnis bei widerrufener Berechtigung sowie idempotentes Öffnen. Sie laufen in Desktop Chromium, Mobile Chromium, Tablet Chromium, Firefox und WebKit. Neue echte native und Browserergebnisse werden erst durch die anschließende GitHub-CI geliefert; PGlite ist kein Ersatz für echte JWT-/HTTP-/Mehrverbindungsabnahme. Keine vollständige Freigabe behauptet.

GitHub-Workflows **QA and Release** und **BA and OP Isolated Acceptance** sind für Push/PR konfiguriert. Nach Upload des gestapelten Draft-PR keine Statusabfragen, kein Warten und kein zusätzlicher Dispatch. Die abschließende GitHub-CI-Prüfung erfolgt separat im Humbel-App-Chat.
