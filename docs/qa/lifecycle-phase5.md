# Korrektur 0.14.1 – Installation auf befülltem Phase-4-Bestand

Ausgangspunkt: PR #23, Branch `feature/ba-op-phase5`, Commit `13acc5dcbe2df70c397ac5dcdc0bc6771448178e`. APP_VERSION **0.14.1**, VERSION_LOG mit tatsächlicher Europe/Berlin-Zeit. Ausschließlich Installationsreihenfolge, zusätzlicher Regressionstest, Version und dieser Bericht geändert; Fachfunktionen der Phasen 1–5 unverändert.

## Nachgewiesene Ursache und Korrektur

Der isolierte [native Lauf 42446965606](https://github.com/butzi124-ship-it/schichtplan-mvp/actions/runs/42446965606) beendet alle bisherigen **43/43** Tests erfolgreich. Sein Ergebnisartefakt meldet danach beim Einrichten von Phase 5: `55006: cannot ALTER TABLE "production_lifecycle_cases" because it has pending trigger events`. Der Fehler wurde lokal mit der unveränderten SQL gegen PostgreSQL/PGlite reproduziert. Der PL/pgSQL-Kontext identifiziert `alter table public.production_lifecycle_cases enable row level security` im abschließenden Rechteblock.

Die vorherige Abnahme hat bereits echte Phase-4-Mengenklärfälle erzeugt. Deren Übernahme in `production_lifecycle_cases` löst aufgeschobene Fremdschlüsselprüfungen aus (`DEFERRABLE INITIALLY DEFERRED`). Das spätere `ALTER TABLE` innerhalb derselben Transaktion ist deshalb unzulässig. Die bisherigen lokalen Lifecycle-Fixtures installierten Phase 5 vor dem Erzeugen solcher Fälle und erfassten diesen Migrationspfad nicht.

Die unveränderte Übernahme steht jetzt **nach sämtlichen DDL-, Trigger-, RLS- und Rechteanweisungen, unmittelbar vor COMMIT**. Die Installation bleibt eine atomare Transaktion; alle Fremdschlüssel und Rechteprüfungen bleiben aktiv. Keine Triggerabschaltung, Constraint-Umgehung, Statusregeländerung oder automatische Mengenkorrektur. Die neun nativen Phase-5-Tests und die Workflowkette **43 + 9 = 52** bleiben unverändert und werden nach erfolgreicher Installation erreicht.

Der neue Test `tests/lifecycle/installation.test.js` erzeugt über die vorhandenen RPCs einen realen Phase-4-Spannungswiderspruch, installiert die echte SQL und prüft erhaltene Mengen-, Benachrichtigungs- und Zeitdaten, eindeutige Übernahme samt Ursprungsreferenzen, unveränderten deaktivierten Schalter, alle zehn RLS-Tabellen, direkte Zugriffssperren, unveränderliche Historie und alle vier aufgeschobenen Fremdschlüssel. Ein absichtlich ungültiger Fremdschlüssel wird weiterhin am Commit mit `23503` zurückgewiesen. Anschließend ist der übernommene Fall über das berechtigte RPC im rein synthetischen Test lesbar. Keine bestehenden Tests oder Assertions geändert.

## Lokale Ergebnisse dieses Fixes

| Prüfung | Ergebnis |
|---|---|
| Installation, Pflicht-Syntaxprüfungen einschließlich neuer Testdatei | bestanden |
| Lifecycle-SQL-/Adapter-/Installationsregression | **26/26** |
| Unverengte bestehende Modulverträge einschließlich Phase 4 und Prognose | **131/131** |
| CI-Bericht / QA-Gates | **2/2**, **9/9** |
| `qa:check`, `qa:agent-server`, `git diff --check` | bestanden |
| Vollständiges `test:e2e`, alle fünf Browserprofile | 718 aufgerufen; Browserstart wegen fehlender Binaries blockiert, keine fachliche Browserabnahme |
| `test:agents` | Browserstart wegen fehlendem Binary blockiert |
| Gepinnte Visuals | BLOCKED_NOT_TESTED: erforderlicher Playwright-1.62.1-noble-Container fehlt |
| `test:db:preflight`, native 52-Fälle-/JWT-/Mehrverbindungsabnahme | BLOCKED_NOT_TESTED: lokale isolierte Dockerlaufzeit fehlt; PGlite ersetzt diese Abnahme nicht |

Die bereits erfolgreiche GitHub-QA des vorherigen Standes (718 Browserfälle, 66 visuelle Prüfungen und Quality Gate/e2e) gilt nicht als Ausführung dieses neuen Commits. Die native Gesamtabnahme der Korrektur bleibt GitHub vorbehalten. Keine produktive Datenbankverbindung oder Aktivierung; keine Berechtigungsänderung. Bestehende offene fachliche und technische Punkte aus dem ursprünglichen Bericht bleiben bestehen.

Push aktualisiert den bestehenden Draft-PR; die konfigurierten Workflows **QA and Release** sowie **BA and OP Isolated Acceptance** werden anschließend durch GitHub ausgelöst. Kein zusätzlicher Dispatch und keine Statusabfrage nach dem Push. Kein Merge oder Deployment.

---

Der folgende ursprüngliche Bericht dokumentiert den historischen Entwicklungsstand **0.14.0**.

# Phase 5 – BA-Mengen, Nachfertigung, Mengenklärfälle und Abschlüsse

## Basis und Umfang

Gestapelt auf PR #22, Commit `72034961668e32b7503a7980f3382c5deb4a84e6`, Version 0.13.2, ursprünglicher Tree `83a4aec54d3dad53ec655970703a63e120629fea`. Vor Beginn einmalig bestätigt: [QA and Release](https://github.com/butzi124-ship-it/schichtplan-mvp/actions/runs/42425098874) und [native BA-/OP-Abnahme](https://github.com/butzi124-ship-it/schichtplan-mvp/actions/runs/42425098875) jeweils erfolgreich. Keine erneute Implementierung der Phasen 1–4.

Bezug: [Issue #17](https://github.com/butzi124-ship-it/schichtplan-mvp/issues/17), verbindliche Spezifikation `docs/requirements/ba-op-shopfloor-2026-11-07.md` / PR #18. Separater Branch `feature/ba-op-phase5`; APP_VERSION **0.14.0**, tatsächliche Europe/Berlin-Zeit im VERSION_LOG. PR-/Commit-Zuordnung siehe Draft-PR-Beschreibung.

Neue Dateien: `js/ba-lifecycle.js`, Reviewentwurf `supabase/review/lifecycle-phase5.sql`, isolierter DB-Adapter, SQL-/Adapter-/Browser-/native Tests und exakte QA-Hook-Liste. SQL **nicht registriert oder produktiv angewandt**. Frontend `HUMBEL_BA_LIFECYCLE_ENABLED=false` und serverseitige Einstellung `production_lifecycle_settings.enabled=false`. Alle bestehenden Produktivschalter bleiben unverändert aus. Keine produktive Verbindung, kein Merge oder Deployment.

## Implementierung und Datenquellen

| Anforderung | Umsetzung und Beleg |
|---|---|
| Sollmenge erhöhen | Geprüftes BA-Kommando ändert ausschließlich aktuelle `ba_quantity`/`target_quantity`; ursprünglicher Artikelsnapshot bleibt unverändert. Ist-Zähler und Zeitsegmente bleiben erhalten. |
| Nachfertigung nach OP-/BA-Abschluss | Betroffene abgeschlossene OPs unter neuer Zielmenge erhalten neue Maschinenabschnitte; alte Abschnitte und Abschlüsse bleiben geschlossen/historisch. Nur wenn tatsächlich weitere OP-Arbeit erforderlich ist, wird ein abgeschlossener BA wieder freigegeben. Kein automatischer Rüst-/Produktionsstart. Maschinenübersicht zeigt bei aktivem Phase 5 den aktuellen BA-Zielwert und die zusätzliche tatsächlich verfügbare Übergabemenge des aktuellen Abschnitts. Eine ausdrücklich bestätigte positive Teilübergabe bleibt von der gesamten Sollmenge getrennt. |
| Reale Folgemengen | Vorhandene Phase-4-Mengenquelle bleibt maßgeblich. Neue physische Übergaben im Nachfertigungsabschnitt und nach weiterem Maschinenwechsel berücksichtigen den zum Abschnittsbeginn belegten Gutteilbestand; alte Vorgängergutteile können nicht nochmals als zusätzliche Teile übergeben werden. |
| Sollmenge reduzieren | Auch unter Gutmenge erlaubt, mit Pflichtgrund und ausdrücklicher Warnbestätigung. Überproduktion separat; keine Gutteile werden gelöscht oder zu Ausschuss umgewandelt. Folge-OPs dürfen sämtliche tatsächlichen Vorgängergutteile bearbeiten. |
| Mengenklärfälle | Ausschließlich vorhandene Phase-4-Erkennung. Neue Episoden referenzieren denselben bestehenden Konflikt; kein zweiter Stückzahlzähler/Detektor. Kommentare, `Geklärt` und begründetes `Geklärt mit verbleibender Abweichung`; Restdifferenz und alle vorherigen Einträge bleiben sichtbar. |
| Zuständige Klärberechtigung | Mitarbeiter erklären/korrigieren berechtigte Buchungen; nur aktuell unabhängig berechtigter operativer AL, aktive Stellvertretung oder zuständiger Administrator schließen Fälle. Weder automatische Klärung noch Mengenanpassung. |
| OP ausdrücklich abschließen | Reine Soll-erreicht-Anzeige, niemals automatischer Abschluss. Aktuell berechtigter tatsächlicher Teilnehmer oder zuständige Leitung darf abschließen. Alle zugeordneten offenen Fälle sperren atomar. Unter Soll nur mit Pflichtgrund und interner Leitungsbenachrichtigung. |
| Aktive Zeitbuchungen | Oberfläche zeigt persönliche Buchungen und offene/pausierte gemeinsame Rüstvorgänge. Abschluss verlangt ausdrückliche Stop-Bestätigung; sämtliche zugehörigen offenen Segmente/Rüstvorgänge enden an einem einheitlichen Serverzeitpunkt. Bereits geschlossene Intervalle bleiben unverändert. |
| Gewöhnliche Buchungen nach Abschluss | Bestehende Terminalprüfungen bleiben wirksam. Nur ausdrücklich begründete berechtigte Zählerkorrekturen dürfen auch auf geschlossene Abschnitte zugreifen; positive Korrekturen erhalten ebenfalls einen Korrektur-Ereignistyp. Bestehende eigene Rüstzeitkorrektur bleibt erhalten. |
| BA ausdrücklich abschließen | Alle OPs müssen explizit abgeschlossen und alle zum Zeitpunkt offenen Fälle geklärt sein. Begründete Restabweichungen zählen als geklärt. Nur gespeicherter hauptverantwortlicher AL, seine aktive Stellvertretung oder unabhängig berechtigter Administrator. |
| Später entstehende Widersprüche | Neue Fälle erhalten einen beobachteten Nach-Abschluss-Hinweis; historische OP-/BA-Abschlüsse bleiben erhalten. Erst ausdrücklich veranlasste Mengenerhöhung kann Nachfertigung freigeben. |
| BA-Hauptverantwortung übertragen | Vorheriger Haupt-AL, aktive Stellvertretung oder Administrator; verifizierter neuer AL und aktiver Kataloggrund zwingend. Bei mehreren Leitungsabteilungen muss die verantwortliche Abteilung ausdrücklich gewählt werden; keine erste beliebige Zuordnung. Grundauswahl samt damaliger Katalogdefinition und Freitext wird eingefroren. Neuer AL erhält interne Meldung. Operative OP-Zuordnungen ändern sich nicht. |
| Administrativer Grundkatalog | Geprüfte Admin-Aktion mit Alt-/Neu-Audit und eigenem optimistischem Katalogvergleich; Aktivierung/Deaktivierung und Freitextpflicht pflegbar. Kein vorgegebener erfundener fachlicher Katalog. |

Die Erstzuordnungsregel für den Haupt-AL bleibt **offen**. Neue BA behalten ihren Drei-Felder-Einstieg; kein erratener AL, kein viertes Pflichtfeld. Die Verwaltung zeigt `pending` und erlaubt ausschließlich eine ausdrückliche, begründete administrative Erstzuordnung über denselben geprüften Verantwortlichkeitsprozess. Danach gelten ausschließlich die gespeicherte Verantwortung und die aktuell gültigen Authorities.

Die BA-Sollmenge ist ein Planwert, keine nachgewiesene physische Materialmenge. Bei aktivem Phase 5 wird Überproduktion an der ersten Spannung der ersten OP daher nicht als erfundener Materialwiderspruch gegen den gesenkten Sollwert behandelt. Die vorhandene Erkennung zwischen realen Vorgänger-/Spannungsmengen bleibt erhalten. Vorhandene Phase-4-Fälle werden trotzdem übernommen und niemals automatisch gelöscht/geklärt. Fehlende historische Zuordnung zu einem Abschluss wird beim Import als **unbekannt** gespeichert.

## Transaktionen, Rechte und Historie

Zehn neue Tabellen mit RLS und entzogenen direkten Clientrechten: Einstellungen, aktuelle BA-Verantwortung/Version, Grundkatalog, Ereignisse, Abschlusssnapshots, Nachfertigungsbezug, Fall-Episoden, Kläreinträge, interne Meldungen und private Transaktionsfähigkeiten. Die sechs Historien-/Meldungstabellen sind gegen Änderung/Löschung geschützt. Abschlusssnapshots enthalten einen flachen damaligen Mengen-/Zeit-/Fallstand, keine rekursiv vervielfachten früheren Snapshots.

`ba_lifecycle_command` ist ein schmaler Invoker-Einstieg in private, serverseitig geprüfte Transaktionen. Derselbe globale Lock wie Phasen 2–4, BA-Zeilensperre und Versionsvergleich serialisieren Mengenänderung, Abschluss, Routing und Zählung. Unabhängig provisionierte Authority-, Mitarbeiter-, Abteilungs- und Gate-Zeilen werden während der Prüfung gesperrt. Request-ID ist an Actor, Objekt, Aktion, Version und vollständige Daten gebunden. Gleiches Retry liefert bei unveränderten Rechten den gespeicherten Zustand; nach einem späteren Scope-Entzug ausschließlich den aktuell noch lesbaren Zustand; verändertes Payload oder veraltete konkurrierende Aktion wird zurückgewiesen. Audit-/Benachrichtigungs-/FK-Fehler rollen die gesamte Geschäftstransaktion zurück.

Die wenigen erforderlichen Phase-4-/Snapshot-/Übergabeadapter sind im neuen Reviewentwurf enthalten, an den serverseitigen Schalter gebunden und verweigern unerwartete oder mehrdeutige Voraussetzungen. Identitäten und eingefrorene Artikelsnapshots werden nicht geöffnet. Ausnahmefähigkeiten werden ausschließlich in einer privaten Transaktionstabelle erzeugt und wieder entfernt, **nicht** über vom Client setzbare Sessionvariablen. Ein GUC-Spoofing-Test und direkte Raw-Zugriffsversuche sind enthalten.

Die Oberfläche hält nur flüchtigen Sessionzustand, schützt doppelte Aktionen und behält nach unklarer Transportantwort dieselbe Request-ID. Kein localStorage, kein zweiter Zählspeicher. 6M-Auswahl verwendet die vorhandenen Ursachen. Loginwechsel, Scanner und Abmeldung löschen den Modulzustand. Neue Zugriffe auf die BA-Verwaltung erweitern keine Legacy-Produktionsrechte. Mitarbeiter erhalten nur zugängliche OP-/Fall-/Abschlusssichten; fremde OPs werden auch nicht über historische Gesamtsnapshots oder alte Request-Replays offengelegt. Die berechtigte BA-Hauptleitung behält den Gesamtüberblick.

Alle ursprünglichen geschützten Hashes bleiben unverändert. QA reversiert ausschließlich die exakt dokumentierten neuen Hooks/Präsentationsadapter aus `tests/qa/lifecycle-integration-patches.json`, danach sämtliche vorhandenen Adapter. **251** Legacy-Identitäten, sämtliche bisherigen Assertionen und **33** Bildreferenzen bleiben geschützt.

## Lokale Prüfung

| Prüfung | Ergebnis |
|---|---|
| `npm ci --ignore-scripts` | bestanden |
| Beide vorgeschriebenen `node --check`, `npm run check:syntax`, neue Modul-/Test-/native Syntax | bestanden |
| Neue Phase-5-SQL-/Berechtigungs-/Adapter-Verträge | **25/25** (22 SQL-/Integritätsfälle, 3 Adapterfälle) |
| Unverengte vorhandene Artikel-/OP-/Readiness-/BA-/Routing-/Rüst-/Mengen-/Prognoseverträge | **131/131**, einschließlich Phase 4 **14/14** und Prognose **41/41** |
| CI-Bericht / QA-Gates | **2/2**, **9/9** |
| `qa:check` / `qa:agent-server` / `git diff --check` | bestanden |
| Ursprüngliche gepinnte native Ziel-/Runtime-Guards | **2/2** |
| Vollständiges `test:e2e` | **718 Fälle aufgerufen**, Browserstart wegen fehlender Binaries blockiert; keine fachliche Browserabnahme |
| `test:agents` | aufgerufen, Browserstart blockiert |
| `test:visual` | **BLOCKED_NOT_TESTED**, vorgeschriebener Playwright-1.62.1-noble-Container lokal nicht verfügbar; keine Referenzänderung |
| `test:db:preflight` / echte native JWT-/PostgREST-/Mehrverbindungsabnahme | **BLOCKED_NOT_TESTED**, isolierte Dockerlaufzeit lokal fehlt; kein Produktionsfallback |

Die SQL-Verträge prüfen Erhöhung nach BA-Abschluss, Reduzierung unter Ist, Restabweichung, OP-Abschluss unter/über Soll, BA-Sperren bei neuen Fällen trotz geschlossener OPs, spätere Fälle nach BA-Abschluss, positive/negative historische Korrekturen, identische Retries, konkurrierende Versionen, Auditrollback, alte Zeit-/Zähler-/Abschlusssnapshots, aktive Zeitenden, verifizierte Rolle/Scanner/Inaktivität, abgelaufene/entzogene Stellvertretung, Verantwortungsübertragung und Katalogpflege. Der PGlite-Parallelfall ist **kein** Ersatz für unabhängige native Verbindungen.

## Vorbereitete unverengte CI

Drei neue Browserfälle laufen in allen fünf unveränderten Profilen: Desktop Chromium, Mobile Chromium, Tablet Chromium, Firefox, WebKit. Damit **703 + 15 = 718**: Flag-off ohne RPC, ausdrücklicher Teilnehmerabschluss ohne erfundenen Ausschuss/ohne BA-Abschlussrecht und verlorene Antwort mit identischem Retry. Alle bisherigen Browserfälle bleiben erhalten.

Der native Workflow verwendet weiterhin exakt die gepinnte Infrastruktur `25e9a6914e38fc791d9ecad13d92a45e30d19bb3` und die geprüfte Testuhr aus PR #22. Unveränderte Phasen 1–4 laufen vollständig zuerst; anschließend **neun** zusätzliche echte PostgreSQL/GoTrue/PostgREST-Fälle: Flags/Raw-/Anon-Denials, Reduzierung, ausdrückliche OP-/BA-Abschlüsse, historische Korrektur/Restklärung/Nachfertigung, zwei echte Verbindungen mit konkurrierendem Abschluss/Mengenänderung, doppelte Abschlussabsicht, Auditrollback, aktiver Zeitabschluss und echte JWT bei AL-Übertragung/Vertretung. Erwarteter Umfang **43 + 9 = 52**, vorbereitet und lokal **nicht ausgeführt**. Frühere Phase-4-only-Limitierung im stufenweisen Ergebnis wird ausdrücklich durch Phase-5-Metadaten ergänzt; kritische Legacy-Sicherheitsbefunde bleiben weiterhin offen.

## Offene technische und fachliche Punkte

- Automatische AL-Erstzuordnungsregel und dauerhaftes fachliches Übertragungsgrundsortiment sind weiterhin festzulegen; kein Automatismus oder Katalog wurde erfunden. Manuelle geprüfte Zuordnung und administrative Katalogpflege sind implementiert.
- Originalschema, tatsächliche Function-/Table-Ownership, Trigger, ACL/Default-Grants, RLS und vertrauenswürdige Authority-Provisionierung müssen entsprechend PR #15/#16 separat produktiv abgeglichen werden. Keine Behebung der kritischen Legacy-Befunde behauptet.
- Echte neue native JWT-/Mehrverbindungsergebnisse, alle Browserprofile/Visuals, Hardware sowie Last-/Durchsatzabnahme stehen noch aus. Der konservative gemeinsame Lock bleibt erhalten. Kein grüner Gesamtfreigabestatus aus lokalen Ersatztests.
- Die vorhandene, weiterhin deaktivierte Prognose nutzt ihren ursprünglichen Artikelsnapshot. Ihre Umstellung auf dynamische BA-Zielmengen gehört zum gesonderten Planungs-/Phase-6-Abgleich und ist vor gemeinsamer produktiver Aktivierung zu prüfen.
- Meldungen sind interne nachvollziehbare App-Meldungen wie in den bestehenden Modulen; kein unbestellter externer E-Mail-/Pushversand.

Branch und Draft-PR werden nach diesen lokalen Prüfungen hochgeladen. Durch Push/PR sind **QA and Release** und **BA and OP Isolated Acceptance** konfiguriert. Keine anschließende Statusabfrage, Warteschleife oder zusätzlicher Dispatch. Die abschließende CI-Prüfung übernimmt der Humbel-App-Chat.

**Kein Merge, Deployment, produktiver Datenbankeingriff oder Aktivierung.**
