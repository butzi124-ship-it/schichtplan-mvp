# Phase 4 – bestehender Stückzahlzähler / v0.13.0

Refs #17, Spezifikation #18; gestapelter Draft auf #21 (`feature/ba-op-phase3`, `6e7bd8a85c06b4cf73bb2f94114200b4e3d8dbde`). Eigener Branch `feature/ba-op-phase4`. Keine Wiederimplementierung der Phasen 1–3. Kein Merge, Deployment, produktiver DB-Lauf oder Feature-Aktivierung. Alle produktiven Schalter bleiben aus.

## Tatsächlicher Bestand und Schnitt

Kataloggrundlage: `docs/qa/schema-inventory.json` (SELECT-Katalogspiegel, kein Restore). Bestehende Funktionen in `schichtplan_mvp.js`: `getProductionStationGoodTotal`, `getProductionOrderCompletedGoodTotal`, `adjustProductionStationGoodQty`, `adjustProductionStationAmount`, `openProductionQaCauseModal`.

| Bestandsobjekt | Unveränderte Bedeutung / Verwendung |
|---|---|
| `production_order_stations` | Spannung, freie `op_number`-Kennzeichnung, aktuelle Scrap-/Clarify-Summen und Lebenszeitsummen |
| `production_order_employees` | Mitarbeiterzuordnung; Unique `(order_id,employee_id)` |
| `production_station_counts` | Einziger Gutteilzähler je `(station_id,order_employee_id)`, `good_qty` |
| `production_station_events` | Vorhandene Ereignisse `good`, `good_correction`, `scrap`, `scrap_correction`, `clarify`, `clarify_correction` mit Mitarbeiter, Ursache und Notiz |
| `production_qa_causes` | Bestehende 6M-Ursachen unverändert |
| `production_order_operations` | Unveränderlicher BA-OP-Snapshot, **keine** Spannung |
| `production_operation_machine_sections` | Vorhandene Maschinenabschnitte aus Phase 2 |

`1AA` wird nicht als dreistellige OP interpretiert oder umgeschrieben. Die Zuordnung geschieht ausschließlich über OP-UUID und eine separate eindeutige Spannungsbindung. Der Test mit `op_number='1AA'` und historischen 17 Gutteilen/2 Ausschuss bestätigt den Erhalt. Keine Umwandlung oder automatische Gutbuchung von Ausschuss/„In Abklärung“.

Die neuen Tabellen sind Zuordnungen, unveränderliches Audit sowie minimale Widerspruchserkennung/Leitungsmeldungen. **Kein neuer Gut-/Ausschusszähler.** Abschnittssummen sind Abfragen der vorhandenen Ereignisse, nicht eine weitere Mengenquelle. Bestehende Renderer/Plus-/Minus-Buttons und 6M-Dialog werden wiederverwendet; nur verknüpfte Spannungen verwenden den neuen atomaren RPC.

## Implementierte Anforderungen

| Anforderung | Umsetzung / gezielte Prüfung |
|---|---|
| Genau eine OP pro Spannung, mehrere Spannungen pro OP | Eindeutiger Spannungs-PK, geordnete vollständige Bestätigung, nach Bestätigung unveränderlich. Neue Spannungen bleiben normale Bestandsstationen. Keine erratene historische OP-Nummer. |
| 100/100/98 = 98 OP-Gutteile | Nur die letzte erforderliche Spannung trägt OP-Gutteile bei. Personensummen aller Spannungen bleiben separat erhalten. |
| Neue OP 0; Fortsetzung erhält Stand | Neue Stationen haben keine übernommenen Counts; erneutes Öffnen/Join/Pause setzt keine Menge zurück. |
| Mitarbeiter und Ausschuss erhalten | Eigene aktive Zuordnung oder unabhängig berechtigter aktueller AL/Vertreter/Admin; ursprüngliche Ereignistypen, Ursachen und Lebenszeitsummen bleiben bestehen. Eigener Einstieg verwendet dieselbe Mitarbeiterzuordnung und kann inaktive Zuordnung ohne Mengenreset reaktivieren. |
| Maschinenwechsel | OP-UUID und Zähler bleiben; jedes neue Ereignis referenziert den geöffneten Abschnitt. Veraltete Abschnitts-ID wird mit HTTP409 abgewiesen. |
| Teilmengen vor Vorgängerabschluss | Rechnerische Verfügbarkeit aus fertigen Vorgänger-Gutteilen. Eng begrenzter Adapter der alten Vollmengen-Sperre ermöglicht ausdrücklich bestätigte physische Teilmengenübergabe und Arbeitsstart. Die Berechnung bestätigt selbst keine physische Übergabe. |
| Keine Doppelzählung | Transaktionaler gemeinsamer Lock in derselben Reihenfolge wie Phase 3/Routing, atomare Deltas, Request-ID inklusive Actor/OP/Payload-Abgleich; Ereignis-Attribution eindeutig. |
| Mengenwiderspruch behält Buchung | Gesamter BA wird nach jeder Buchung neu bewertet, auch nach rückwirkender Vorgängerkorrektur. Ein offener Fall je betroffener Spannung, Erstwerte/Request/Zeitpunkt und deduplizierte Benachrichtigung an unabhängig berechtigte aktuelle Leitung. Aktuelle Restdifferenz bleibt sichtbar. |
| Historische Mengen | Nur explizite begründete AL-Adoption; bereits gespeicherte Mengen bleiben erhalten und werden keinem erfundenen Maschinenabschnitt zugeteilt. Widersprüchliche historische Fremd-BA/Mitarbeiterbeziehungen werden vor Adoption zur gesonderten Prüfung abgewiesen. |
| Keine Phase 5 | Kein Klärfallabschluss, keine BA-Mengenänderung, Nachfertigungs- oder Abschluss-API. Alte Abschluss-/Mengenwege für verknüpfte BA/OP werden abgewiesen, damit sie Phase 5 nicht umgehen. Eine rechnerisch wieder ausgeglichene Differenz schließt ihren Fall nicht automatisch. |

Frontend benötigt zusätzlich `HUMBEL_OP_QUANTITIES_ENABLED`; Standard **false**. Server benötigt zusätzlich `production_quantity_settings.enabled`; Standard **false**. Die vorbereitete SQL-Datei liegt ausschließlich unter `supabase/review/quantity-phase4.sql`, nicht als registrierte Produktivmigration. Vorhandene unverbundene Altaufträge behalten ihre Regeln.

## Lokale Prüfungen

| Prüfung | Tatsächliches Ergebnis |
|---|---|
| `npm ci --ignore-scripts` | Erfolgreich |
| Beide vorgeschriebenen `node --check`, `check:syntax`, `git diff --check` | Erfolgreich |
| Phase-4-Suite `test:quantity` | **14/14**: elf PGlite-Datenintegritäts-/Berechtigungsfälle und drei Adapter-Verträge |
| Artikel / OP / Readiness | **8/8**, **15/15**, **17/17** |
| BA-Einstieg / Routing / Rüsten | **9/9**, **10/10**, **17/17** |
| CI-Bericht / QA-Gates | **2/2**, **9/9** |
| Forecast | **28/30**; zwei unveränderte Bestandsassertionen schlagen zur späten Tageszeit fehl, unten belegt |
| Summe Unit-/DB-Verträge einschließlich Phase 4 | **129/131**, zwei bestehende Forecast-Fehler |
| `qa:check` | Erfolgreich; 251 geschützte Legacy-Identitäten, bestehender Business-Hash nach exakter Hook-Umkehr und 33 unveränderte Bildreferenzen erhalten |
| `qa:agent-server` | Erfolgreich; keine App-/DB-Schreibaktion |
| `test:agents` | Aufgerufen; Browserstart blockiert durch fehlendes Chromium-Binary, keine fachliche Ausführung |
| Vollständiges `test:e2e` | Unverengt aufgerufen: **703 Browserfälle**. Lokal am fehlenden Browser-Binary gescheitert; **kein** bestandener Browser-/Gerätenachweis |
| `test:visual` | **BLOCKED_NOT_TESTED**: vorgeschriebener `playwright:v1.62.1-noble`-Container lokal nicht vorhanden; zweimal 33 strikte Vergleiche unverändert in CI vorbereitet |
| `test:db:preflight` | **BLOCKED_NOT_TESTED**, explizite isolierte native Laufzeit fehlt; keine produktive Verbindung |

Die Browserinstallation wurde versucht: CDN lieferte ein defektes/trunkiertes ZIP (`End of central directory record signature not found`). Kein alternativer Browser oder gelockerter Vergleich wurde als Abnahme verwendet. Desktop/Smartphone/Tablet/Firefox/WebKit bleiben lokal **nicht getestet**.

**Bestehende Forecast-Fehler:** `tests/forecast/database.test.js`, Fälle „service-only publication atomically caches results and announces without physical release or counter writes“ und „publication failure rolls back run, pointer and announcements atomically“. Der Fixture-Auftrag braucht mit drei OPs mehr als die verbleibende heutige Schicht; dadurch entsteht kein heute erwarteter Arrival-Datensatz. Der erste Test erhält `planned` statt `expected`, der zweite erreicht seinen absichtlich fehlschlagenden Arrival-Insert nicht. Dieselben beiden Fehler wurden separat auf dem **unveränderten Phase-3-Head `6e7bd8a…`** reproduziert (28/30). Weder Quelle noch Fixture noch Assertion dieser Suite geändert. Kein Grün behauptet; Zeitdeterminismus bleibt ein gesonderter Prüfblocker.

## Neue Browser- und native CI-Fälle

Vier neue vollständige Browserfälle laufen in allen fünf bestehenden Profilen, also 20 zusätzliche Ausführungen ohne Filterverengung: Originalbuttons/letzte Spannung/mobile Breite, verlorene Antwort und identischer Retry, Maschinenwechsel mit zwei Attributionsabschnitten, Original-6M-Dialog/Flag-off. Keine vorhandene Testassertion geändert.

Der bestehende isolierte Workflow verwendet weiterhin die gepinnte #15-Infrastruktur `25e9a6914e38fc791d9ecad13d92a45e30d19bb3`. Sämtliche Phase-1/2/3-Tests werden unverändert vor die sechs neuen Phase-4-Prüfungen geschaltet: ausgeschaltete Gates und echte JWT/Raw-Denials; vorhandener Zähler und letzte Spannung; physischer Teilmengenstart; zwei unabhängige Verbindungen (identische und verschiedene gleichzeitige Requests); HTTP409 beim veralteten Maschinenabschnitt; Audit-Rollback und Fremdzugriff. Diese nativen Prüfungen sind **vorbereitet, lokal nicht ausgeführt**. Der Upload startet die bestehenden Push-/PR-Workflows. CI-Ergebnisse werden anschließend separat geprüft; keine Warteschleife, kein zusätzlicher Dispatch zum Warten.

## Offene Risiken und Freigabegrenzen

- SQL, native JWT/PostgREST-/Mehrverbindungstests und vollständige Browser-/Visualabnahme stehen noch zur Prüfung. Zwei reproduzierte Bestandsfehler verhindern eine pauschale grüne Gesamtabnahme.
- Katalogspiegel und synthetische PGlite-Tabellen sind kein autoritativer Produktionsrestore. Originalschema, Ownership, bestehende Trigger, ACL/Default-Grants und Alt-RLS müssen aus #15/#16 separat abgeglichen werden. Neue Raw-Write-Guards prüfen sowohl BA- als auch Spannungsbezug gegen gefälschte `order_id`; ihre Owner-Übereinstimmung mit den Security-Definer-Kommandos muss bei Aktivierung verifiziert werden. Keine Behebung aller kritischen Legacy-Befunde behauptet.
- Historische Spannungspläne werden nicht aus freiem Kennzeichnungstext geraten. Erforderliche Reihenfolge muss ausdrücklich bestätigt werden. Der alte Artikelsnapshot enthält keine verifizierte maschinenlesbare Spannungsplanung; eine automatische Stammdatenübernahme wird daher nicht erfunden. Bestätigte Pläne sind in Phase 4 absichtlich unveränderlich.
- Historische Gutteile vor Adoption sind in der OP-Menge, bleiben aber ohne belegte Herkunft außerhalb der Maschinensummen. Neue Korrekturen können negative Netto-Buchungen eines aktuellen Abschnitts zeigen; sie schreiben die historischen Produktionsereignisse nicht um.
- Der konservative gemeinsame Lock serialisiert Zähl-, Rüst- und Routing-Kommandos. Produktive Last-/Durchsatzabnahme fehlt. Ein Ereignis pro Buttonabsicht; physisch doppelt gezählte Werkstücke verschiedener Personen sind ohne Teileidentität nicht automatisch erkennbar.
- Leitungsmeldungen entstehen intern aus unabhängig provisionierten aktuellen Authorities. Kein E-Mail-/Pushversand und keine automatische Fallklärung. Bestehende Versorgung mit gültigen Leiter-/Vertreterrechten bleibt Voraussetzung.
- Phase-5-Fallbearbeitung, BA-Mengenänderung, Nachfertigung und Abschlüsse sind bewusst nicht implementiert. Verknüpfte Fälle können erst nach der gesonderten Phase-5-Implementierung abgeschlossen werden.
- Physische Scanner/Terminals/Smartphones/Tablets, tatsächliche Materialübergaben, Produktionsrestore und produktive Skalierung nicht getestet.

**Keine Freigabe. Kein Merge, Deployment, produktiver Supabase-Lauf oder Aktivierung.**
