# Digitale Qualitätsprüfung – technischer Entwurf

Zum [verbindlichen Fachvertrag](../requirements/digital-quality-2026-11-14.md). **Nur P: keine neuen Datenbankobjekte oder Endpunkte implementiert.** Ergänzt das [Dokumentenmodell](documents-inspections-model.md), ersetzt weder bestehende Tabellen noch die Identitätskandidaten aus [8B1](../security/permissions-8b1.md)/[8B2](../security/permissions-8b2.md).

## 1. Bestand und Integrationsschnitt

Am Basiscommit #37 vorhanden: `js/op-quantities.js` bindet die bisherigen Stationsbedienelemente über `op_quantity_view` und `op_quantity_command` an. `adjust` nutzt station/person/kind/delta plus Maschinenabschnitt und Request-ID. `supabase/review/quantity-phase4.sql` schützt bestehende Gut-/Ausschuss-/Abklärungsdaten, `lifecycle-phase5.sql` ergänzt Klärfälle/Nachfertigung/Abschluss. Diese Aggregatbuchungen enthalten **noch keinen vollständig geprüften individuellen Teile-/Merkmal-/Mandantenvertrag**. Neue Einzelteilidentität darf nicht aus bestehenden Personensummen rekonstruiert werden.

Artikelrevision/OP-Plan/BA-Snapshot existieren im Code/Review-SQL; produktive Installation weiter nicht nachgewiesen. `production_order_stations` ist ein Spannungsmodell, keine physische Seriennummer. Ein station_event oder dieselbe sichtbare OP-Laufnummer in zwei OPs beweist keinen identischen Werkstückkörper. Bestehende Stückzähl-/QS-Ursachen-/Produktionsprotokolle bleiben erhalten. Neuer Adapter muss sämtliche alternativen Alt-/RPC-/Tabellen-/UI-Pfade pro Pilotoperation schließen oder gleichwertig prüfen, ohne die übrigen Prozesse umzubauen. Kein zweiter aggregierter Qualitätszähler als Ersatz.

## 2. Prüfintervalle und Fälligkeit (P zu F02/F03)

Je Merkmal eine versionierte Regel mit Firma, Artikelrevision, OP, Bereich, Zyklustyp, Parametern/Phase, gültigem Zeit-/Versionsbereich, ausdrücklich bestätigter Priorität sowie Erst-/Letzt-/Nachprüfbedingungen. 100 % entspricht jedem relevanten Teil, jedes zweite/fünfte/zehnte einem expliziten N-Zyklus; freie Zyklen als begrenzter geprüfter Katalog, keine frei ausführbaren Formeln/Skripte. Prozentsatz als exaktes rationales Verhältnis, nicht gerundete binäre Floatquote.

**Nicht beschlossen ist ein universeller Prozentalgorithmus.** Vorschlag zur Entscheidung O01: deterministische kumulative Quote (fällig, wenn der ganzzahlige Quotenstand für Teil n gegenüber n−1 wächst); bei 10 % an 10/20/30…, bei 30 % an 4/7/10… . Erststück/Nachprüfung können zusätzliche Messungen auslösen, ersetzen aber keinen geplanten Pflichttermin. Alternative losbezogene Stichprobe nur mit ausdrücklich definierter Losgröße/Auswahlregel und reproduzierbarer Revision. Nicht einfach 100/p auf ganze Teile runden: das verfälscht z.B. 30 %. „10 % alle vier“ ist nicht zulässig.

Fälligkeit je Merkmal = regelmäßiger Termin vereinigt mit erzwungenen Erst-/Letzt-/Nachprüfungen und expliziten gültigen Bereichsregeln. Gleiche Merkmal-/Teilpflicht einmal, alle Gründe als Evidenz erhalten. Zwei Merkmale mit N=2 und N=5: bei Teil 2 nur erstes, bei Teil 5 nur zweites, bei Teil 10 beide; Sonderprüfungen zusätzlich. Ohne Fälligkeit kein Messdialog. Allein eine vorhandene Merkmalsliste bedeutet nicht, dass jedes Merkmal bei jedem Teil fällig ist.

Zählbasis (alle Bearbeitungsversuche vs. Gutteile), Startphase, Reset bei neuer OP/Nachfertigung, Erstteil je Schicht, Maschinen-/Werkzeugwechsel und Konfigurationspriorität sind O01/O02/O04. Zur Vermeidung „Ausschuss verschiebt den Termin unerkannt“ wird eine unveränderliche technische Bearbeitungsfolge unabhängig vom Gutteilaggregat vorgeschlagen. Sie ist kein zweiter Mengenzähler und braucht Abgleich zum Spannungsmodell. Schicht-/Personen-/Maschinenwechsel setzt offene Nachprüfpflichten nicht zurück. Keine neue Regel wird aktiviert, bevor diese Details und die Bestandsadaption entschieden sind.

Auftrag/OP/Teil erhält bei Buchungsbeginn genaues Prüfmanifest einschließlich Regel-/Vorlagenrevision und konkreter fälliger Merkmale. Server berechnet es erneut unter Sperre, prüft konfigurierte Änderungen/Rückzüge und erlaubt nicht, dass Client eine Pflicht weglässt. Neue Revisionen verändern alte Messungen nicht rückwirkend. Regelwechsel mit offenen Teilen benötigt expliziten Übergang, niemals automatische Befreiung alter Pflichten. Fehlende/mehrdeutige Regeln → Klärung, kein manueller Gut-Fallback.

## 3. Teileidentität, Nummern und Produzent

OP-Übernahme erfasst ausdrücklich Ja/Nein zur bereits vorhandenen individuellen Nummerierung. Mehrere Mitarbeitende derselben OP dürfen keine widersprüchlichen privaten Nummerierungsmodi erzeugen; gültige Operationseinstellung/Übernahmeversion serverseitig vergleichen, erneute Frage ersetzt nicht ungeprüft laufende Zuordnungen. Wechsel auf „Ja“ bei späterer Gravur in dieser/folgenden OP wird als neue belegte Zuordnung dokumentiert (O07).

**Nein (B):** sichtbare Folge ab 1 innerhalb der OP, gemeinsam über alle Mitarbeiter/Schichten, beim Wiederaufnehmen nicht neu beginnen. Eindeutige technische ID separat; Auftrag+OP+Mitarbeiter+Erfassungszeit bleiben erhalten. Zwei OPs können beide „Teil 1“ haben, ohne dass dies dasselbe physische Teil ist. Reservierte/vergebene Nummern nicht nach Fehler einfach einer anderen Person wiederverwenden; Lücken/Abbruchstatus sichtbar, endgültige Nummernvergabe O04.

**Ja (B):** vorhandene Nummer bei jeder Teilbuchung abfragen und prüfen. Technischer Vorschlag: eindeutiger fachlicher Vorgang je Firma+BA+OP+normalisierter vorhandener Teilenummer; Retry referenziert denselben Vorgang. Derselbe physische markierte Gegenstand darf in Folge-OP wieder erkannt werden, ohne fälschlich als globales Duplikat abgewiesen zu werden. Ob Nummer firmenweit, artikel-/los-/BA-bezogen eindeutig ist, zulässige Zeichen/führende Nullen und Baugruppenbeziehungen bleiben O07. Kein ungeprüftes Normalisieren, das unterschiedliche gültige Nummern verschmilzt. Reine Nachprüfung ist **kein** erneutes Zählen des bereits vorhandenen Teils.

Spätere Gravur: neue Nummer an belegte gegenwärtige technische Identität binden; frühere unmarkierte Teile nicht per Reihenfolge, Zeitnähe oder gleicher Zählernummer rückwirkend automatisch verbinden. Verknüpfung braucht konkreten Beleg/Actor/Zeit und Vertrauensstatus. „Nicht nachweisbar“ bleibt sichtbar. Datenmodell trennt Vorgangs-ID je OP, sichtbare Kennung und optional bestätigte physische Teileidentität. Produzent nicht bei Entscheidung auf AL/QS umschreiben; Prüfer/Korrekturausführender/Entscheider gesondert.

## 4. Logisches additives Datenmodell (P)

Alle Relationen firmengeschützt; konkrete Namen und SQL erst in späterem Implementierungsauftrag, bestehende Originaldefinitionen **nicht** als hierfür verifiziert darstellen.

| Entwurfsobjekt | Zentrale Felder/Beziehungen | Zwingende Invariante |
|---|---|---|
| companies / company_memberships | Firmen-ID, aktive unabhängige Identitätszuweisung, Zeitraum/Version | Kein Tenant aus frei änderbarem Profil/Browserparameter; O10 Provisionierung/Mehrfirmenwechsel |
| quality_profiles / profile_versions | Firma, Abteilung/Artikelrevision/OP, Regelstand, Status, genehmigte Priorität | Keine globale Humbel-Defaultvererbung, unklare Konflikte verweigern |
| characteristic_rules | Vorlagenmerkmal, eigene Intervallregel, Toleranzgrenzen, zwei Eingriffsgrenzen, Regeltyp/Einheit | Untere/obere Grenze konsistent, keine erfundenen einseitigen Endpunkte |
| operation_quality_contexts | Firma, BA/OP, Regelmanifest, Nummerierungsmodus/Version, nächste Folge, offene Sperrgründe | Ein kohärenter OP-Zustand für alle Mitarbeiter, keine Reset-/Rennumgehung |
| operation_piece_records | Technische ID, Firma/BA/OP, lokale Nummer oder vorhandene sichtbare Nummer, Produzent, Abschnitt/Spannungsbezug, Server-/Erfassungszeit, Version | Unikat nach entschiedenem Namespace; stabile Produzentzuordnung; keine firmengrenzübergreifenden Links |
| physical_identity_links | Vorgangs-ID, optionale belegte physische ID/Gravur, Evidenz/Actor/Zeit/Vertrauensstatus | Unbelegte OP-übergreifende Identität ausdrücklich nicht behaupten |
| due_checks / measurements | Stück/Merkmal/Manifest, Fälligkeitsgründe, Erst-/Letzt-/Nachprüfung, typed value, Originalevent/Prüfer/Zeit/Version | Eine Pflicht je konkreter Kombination, mehrere Messversuche erhalten; fehlend/N/A nicht bestanden |
| quality_cases / decisions | Stück/Ergebnissatz, Gelb/Rot, aktueller Handlungsbedarf, Nacharbeit/Sonderfreigabe/Ausschuss, Authority/Entscheider/Grund/Zeit | Entscheidung exakt an aktuellen Mess-/Fallstand gebunden, ursprünglicher Rotwert bleibt |
| corrective_actions / operation_holds | OP/Fall/Merkmale, tatsächliche Maßnahme, Werkzeugreferenz/alter-neuer dokumentierter Korrekturwert, Ausführender/Zeit, einzelne Sperrgründe | Kein CNC-Schreibbefehl; Korrekturdokumentation beendet nur zugehörige Sperre |
| mandatory_rechecks | Auslöser, OP, betroffene Merkmale, nächstes Teil, Status/Ergebnis | Korrektur hebt Nachprüfpflicht nicht auf, kein Verlust bei Schichtwechsel |
| acknowledgements / notification_outbox | Fall, zuständiger AL, Kenntnisnahme/Serverzeit; deduplizierter Zustellauftrag/Status | Kenntnisnahme ist keine Freigabe. Rot zunächst nur AL; Sonderfreigabe kein QS-Push |
| count_transition_links / receipts / audit | Konkreter Teilevorgang, bestehende Zähl-/Ereignisreferenzen, alter/neuer Status, Request-ID/Payloadfingerprint/Actor/Version | Einmalige Anwendung auf bestehenden Zähler, atomare Umbuchung, kein Löschen der Ursprungsereignisse |
| company_layouts / template_scope | Firma, Hoch/Quer/Format, Dokument-/Merkmal-/Benachrichtigungsrevision | Humbel A4 quer; fremde Firmen nur eigene/neutrale Inhalte |

Alle PK-/FK-/Unique-/Linkprüfungen im Ziel nehmen Firma mit auf oder erzwingen sie gleichwertig unvermeidbar. UUID allein ist keine Mandantengrenze. Dateien/Derivate, PDFjobs, Audits, Anwendungs-/Worker-/Realtime-/Such-/Export-/Restorepfade tragen dieselbe Bindung. Private Buckets ohne breite authenticated-/Service-Key-Clientrechte; Firmenpfad ist nur Struktur, nicht Sicherheit. Bestehendes #37-Storage-/Recoverykonzept fortführen, kein öffentlicher Download.

## 5. Getrennte Zustandsübergänge

| Ereignis / Vorzustand | Zielzustand / Mengenwirkung | Gate / Historie |
|---|---|---|
| Unkritisch, bewusst ohne Prüfpflicht | Manuell Gut/Ausschuss/Abklärung über bestehende Regeln | Nicht mit fehlendem Profil verwechseln |
| Teil identifiziert, Pflicht fällig und offen | Prüfwartend; **kein endgültiges Gutplus** | Wartend vs. gezählte Abklärung O04 |
| Alle nötigen Merkmale unkritisch bestanden | Gut-fähig, genau einmal bestehende Gutbuchung | Weiterhin letzte erforderliche Spannung/Teilnehmer-/Verfügbarkeitsregeln |
| Gelb / Eingriff erreicht | Fall + OP-Korrektursperre + Meldung; noch kein Umgehen offener Pflicht | Tatsächliche Maßnahme/deren Dokumentation erforderlich, kein AL-Freigabegate |
| Gelbmaßnahme dokumentiert | Nur zugehörige Korrektursperre weg; Nachprüfung nächstes Teil bleibt | Ggf. aktuelles Teil gutfähig, Messgelb unverändert; offene Kenntnisnahme bleibt sichtbar |
| Rot / qualitativ negativ | In Abklärung / nicht konform; kein Gut | AL oder gültiger berechtigter STV entscheidet |
| Abklärung → Ausschuss | Abklärung −1, Ausschuss +1 **für dasselbe Teil atomar** | Produzent unverändert, Requestreplay keine zweite Umbuchung |
| Abklärung → Nacharbeit | Nacharbeit/Nachprüfung erforderlich, nicht sofort Gut | Keine neue unabhängige Zählung desselben Körpers; Nacharbeits-/Mengendetail O04/O05 |
| Abklärung → zulässige Sonderfreigabe | Explizit akzeptiertes Teil, einmalige berechtigte Mengenüberleitung | Originalrot bleibt, QS-Auftragsanzeige Name/Zeit/Grund ohne Push |
| Falscher Messwert / Neumessung | Zusätzliches referenziertes Ereignis | Kein UPDATE des Originals, alte Entscheidung nicht auf neuen Stand kopieren |
| Später Fehler nach Abschluss | Neuer nachvollziehbarer Qualitätsfall | Kein historisches automatisches Wiederöffnen abgeschlossener BA oder stilles Löschen von Mengen |

Mehrere Merkmale: ein Rot verhindert Gut unabhängig von anderen Grünwerten; fehlende Pflichtwerte verhindern Gut; Gelb löst die Maßnahme auch bei sonst Grün aus. Mehrere Gelbereignisse können mehrere unabhängige Sperrgründe haben. Eine AL-Kenntnisnahme entfernt keine Sperre, eine Korrektur keine fremde Sperre. Aggregierter Cockpit-Handlungsbedarf ≠ historische Messfarbe ≠ Mengenzustand.

## 6. Transaktions-/Sicherheitsvertrag

Gemeinsame feste Sperrreihenfolge über aktuelle Firmen-/Actor-/Authoritybindung, OP/Maschinen-/Spannungskontext, Teilevorgang, Pflicht-/Mess-/Fallstand und Zählreferenzen; nach Wartezeit Auth-, Versions-, Status- und Fachregeln erneut prüfen. Fälligkeit, Messnachweis, Gate, Counttransition, Audit und verpflichtende Outbox in **einem geschützten Fachkommando** committen. Vorprüfung im Browser vor altem ungesicherten Count-RPC genügt nicht. Gegenläufige Ausschuss-/Freigabe-/Widerrufs-/Korrekturanfragen werden serialisiert, eine veraltete zweite Entscheidung scheitert.

Request-ID + exakter Payloadfingerprint + erwartete Revisionen; identischer Retry nur bei aktueller Leseberechtigung, anderer Payload/Nummer/Actor ist Konflikt. Einzigartige Übergangsreferenz verhindert doppelte Gut-/Ausschusswirkung auch bei neu gewählter Request-ID. Netzwerkantwort unklar → vorhandenen Vorgangsstatus abfragen, nicht neues Teil erzeugen. Fehler am Audit/Outbox/Zähler rollt alle zugehörigen DB-Änderungen zurück, Fotos/Dateien bleiben bei nichtatomarem Upload in Quarantäne nach #37.

Mandantentrennung ist **Voraussetzung vor Mehrfirmenbetrieb**, nicht allein bei neuen QS-Tabellen: alte employees/departments/production_orders und sämtliche Helfer/Policies/Grants/Exporte/Worker sind bislang kein belegtes mandantenfähiges Gesamtsystem. Keine produktive Tenant-Backfillmigration, keine pauschale GL-/TA-Globalvollmacht, keine Authentifizierung aus raw Claims. Plattformbetrieb/Backupschlüssel separat minimal und auditiert; TA erhält keine fachliche Gesamtsicht. Neutrale Seeds werden eigens erstellt, nicht durch Umbenennen von Humbel-Unterlagen.

Runtime-Prüfungen mit echten unabhängigen PostgreSQL-/Auth-/API-/Storageverbindungen und gleichzeitigen Transaktionen bleiben vor späterem Einsatz Pflicht. Dieser Entwurf und grüne Bestandstests sind keine native Abnahme oder Produktionsfreigabe.
