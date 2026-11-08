# Humbel-App: fachliche Spezifikation BA / OP / Fertigungsrückmeldung (Stand 07.11.2026)

> **Status: beschlossene fachliche Zielregeln, NICHT implementiert oder produktiv freigegeben.**
> Dokumentations-PR zu [Issue #17](https://github.com/butzi124-ship-it/schichtplan-mvp/issues/17), im Kontext von Epic #6 sowie #7, #8, #9, #4. Vor Entwicklung vorhandene App-/DB-Logik prüfen.
> Die jüngeren expliziten Entscheidungen in diesem Dokument haben Vorrang vor Widersprüchen in älteren Planungs-Issues. Keine Datenbankmigration, kein RLS-/Auth-Eingriff, kein Deployment oder Feature-Schalter durch diesen PR.

## A. Systemgrenze, Nummerierung und BA-Anlage

1. **Eigenständige Humbel-App**: keine Abhängigkeit vom bisherigen ERP, kein automatischer ERP-Import/Abgleich. Langfristig eigenständiges Humbel-ERP. Bis dahin existiert die BA- und Artikelverwaltung vollständig innerhalb der App.
2. Bestehende **Humbel-BA-Nummern** und feste **Humbel-Artikelnummern** weiterverwenden. Bei BA-Anlage manuell eingegebene BA-Nummer auf Eindeutigkeit prüfen; vorhandenen BA öffnen statt duplizieren. Später optional administrativ konfigurierbare automatische BA-Nummernvergabe, ohne alte IDs zu verändern. Artikelnummern vorerst nicht automatisch vergeben.
3. **Jeder für die erste OP zugelassene Mitarbeiter** darf einen BA direkt an der ersten OP anlegen. Die **einzigen verpflichtenden manuellen Angaben für diesen Einstieg**: BA-Nummer, Artikelnummer und Stückzahl. Kein Pflichtfeld für Endtermin oder Maschinenwahl beim Erstaufruf; weitere Planungsdaten können im BA später gesondert verwaltet werden. Hauptverantwortlichen AL automatisch aus der zuständigen Abteilung der ersten OP im gültigen Artikel-/OP-Snapshot zuordnen: genau ein aktiver, unabhängig serverseitig berechtigter und dort eindeutig hinterlegter AL. Abgelaufene, widerrufene, inaktive oder mehrdeutige Zuordnungen sind ungültig. Ohne eindeutigen gültigen AL BA trotzdem mit **„Hauptverantwortung offen“** anlegen; Administrator darf die Erstzuordnung über die geprüfte Funktion nachholen. Kein viertes Pflichtfeld. Automatische Zuordnung beziehungsweise offenes Ergebnis nachvollziehbar auditieren; OP-Zuständigkeiten bleiben unabhängig. Spätere Änderungen ausschließlich über die bestehende begründete Verantwortungsübertragung. Beschluss ergänzt am 08.11.2026.
4. Aus freigegebener Artikelrevision werden OPs mit Reihenfolge, Spannungszuordnungen, Standardmaschinen/-bindungen und Sollzeiten als **BA-Snapshot** übernommen. Keine alten Ist-Buchungen, Zählerstände, Abklärungsfälle oder Fremd-BA-Nummern kopieren. Nach Anlage alle OPs unter derselben BA-ID.
5. Vorhandene OP-Barcodes verwenden (z. B. `812150.100`; tatsächlichen Aufbau an realen Codes prüfen). OP direkt per Scan oder manuell per BA-/OP-Suche öffnen. Kein selbst erfundenes neues Barcodeformat. OP-Barcode öffnet die entsprechende OP ohne zusätzliche Auswahl.

## B. Artikelstamm und BA-Snapshots

1. Artikelstamm mit freigegebenen Revisionen; Änderungen betreffen **neu angelegte** BA, nie automatisch bereits bestehende.
2. Berechtigte Benutzer sehen Unterschiede **pro OP und Feld**; gezielte Auswahl einzelner neuer Artikelvorgaben für den bestehenden BA. Nicht ausgewählte Werte bleiben gleich.
3. BA-Artikel-Abgleich: verantwortlicher AL, aktive Stellvertretung und Administrator dürfen übernehmen; normale Mitarbeiter sehen zugängliche Unterschiede, können jedoch nicht übernehmen. Serverprüfungen zwingend.
4. Änderungen auch bei **bereits begonnenen OPs** zulässig, wenn historische Vorgaben/Ist-Werte geschützt und Maschinenwechselregeln eingehalten werden. Neue Vorgabezeiten nur für weitere Planung; bereits erfasste Ist-Daten und historische Vergleichsbasis unverändert.
5. **Vollständig abgeschlossene BA** dürfen nicht nachträglich über den Artikelstamm-Abgleich geändert werden. Wird nach einer BA-Mengenerhöhung regelkonform zur Nachfertigung wieder freigegeben, ist ein gezielter Abgleich für die weitere Bearbeitung wieder möglich.

## C. Maschine pro OP, Ausführung und Verantwortlichkeit

1. Jede Artikel-OP hat eine **Standardmaschine**. Beim OP-Öffnen vorauswählen. Ohne spezielle Sperre darf ein Mitarbeiter nur aktive/zugelassene Maschinen der für diese OP zuständigen Abteilung auswählen.
2. Beim Anlegen von Artikel **oder BA** ist die Option **zwingende konkrete Maschine** möglich. Diese Bindung darf nicht durch einen normalen Maschinenwechsel übergangen werden.
3. Ein zuständiger AL darf eine andere technisch geeignete Maschine **einmalig für eine konkrete Auftrags-OP** festlegen, auch abteilungsübergreifend. Diese OP-Maschine ist danach für normale Mitarbeiter **gesperrt**; der Artikelstamm wird nicht geändert. Dauerhafte neue Standardmaschine nur über Artikelrevision/Stammdaten.
4. Abteilungsübergreifende Einmalzuweisung braucht **keine vorherige Zustimmung** des ausführenden AL; dieser wird automatisch informiert. Der ausführende AL bzw. aktive Vertretung übernimmt operative Verantwortung (Maschine, Rüsten, Fortschritt). Der ursprüngliche AL behält Überblick und BA-Gesamtverantwortung bleibt unverändert.
5. Ausführender AL bzw. aktive Vertretung kann die OP **direkt** an die ursprüngliche Abteilung zurückgeben. Der ursprüngliche AL wird informiert und übernimmt operative Verantwortung. Alle bisher gebuchten Daten bleiben bestehen.
6. Rückgabe auch bei aktiven Rüst-/Fertigungsbuchungen möglich: aktuelle persönliche Buchungen zur Rückgabezeit eindeutig und atomar unterbrechen/schließen, Grund protokollieren; **nicht** als fertige OP werten.
7. **Jeder Maschinenwechsel während aktiver Rüstung**, auch innerhalb derselben Abteilung, bricht den bisherigen Rüstvorgang mit **verpflichtendem, administrativ konfigurierbarem Abbruchgrund** ab. Freitext optional, bei „Sonstiges“ aussagekräftige Erläuterung; alle Ist-Zeiten erhalten. Auf neuer Maschine neuer Rüstvorgang, Start nur explizit.
8. Maschinenwechsel nach Fertigungsbeginn: **dieselbe Auftrags-OP behalten**, aber getrennte Maschinen-/Fertigungsabschnitte mit eigenen Zeiten, Stückzahlen und Rüstvorgängen. Keine neue künstliche Teil-OP und keine Mehrfachzählung.
9. Maschinenkapazität der tatsächlich ausgeführten Maschine/Abteilung zuordnen. Zwingende Maschinenbindungen und aktuell gültige Sperren vor jeder Änderung serverseitig prüfen.

## D. Rüstzeiterfassung und allgemeine Pause

1. Budgetierte **Gesamtrüstzeit** aus dem Artikel/BA bleibt Sollwert; keine feste interne/externe Aufteilung verlangen. Ist-Rüstung: Mitarbeiter betätigt `Rüsten beginnen`, wählt **intern/extern**, kann beliebig zwischen Modi wechseln, `Rüsten pausieren`, `Eigene Rüstzeit beenden` oder `Rüsten beenden`.
2. Jeder Wechsel schreibt abgegrenztes Zeitsegment. **Intern** belegt persönliche Zeit und die konkrete Maschine, **extern** persönliche Zeit, jedoch nicht automatisch Maschinenstillstand. Maschine kann parallel einen anderen Auftrag fertigen. Maschinenbelegung/Stillstand eigenständig ermitteln, nicht personenbezogene Minuten addieren.
3. Gemeinsamer Rüstvorgang: **jeder beteiligte Mitarbeiter erfasst seine eigene Zeit**; alle Buchungen gehören zu einem gemeinsamen Rüstvorgang einer OP. Persönlicher Aufwand ist Summe über Personen, Kalenderdauer und Maschinenstillstand sind separate Kennzahlen.
4. `Eigene Rüstzeit beenden` schließt nur diese Buchung. **`Rüsten beenden` durch einen beteiligten Mitarbeiter** (oder verantwortlichen AL/Vertreter/Admin) schließt **gemeinsamen Rüstvorgang für alle**, auch noch laufende persönliche Buchungen mit einheitlichem Zeitpunkt und dokumentiertem automatischen Abschluss.
5. **Allgemeiner**, von der Rüstmaske unabhängiger Button `Zur Pause abmelden` in der Mitarbeiteroberfläche. Stoppt sämtliche laufenden persönlichen Fertigungszeitbuchungen (einschließlich Rüsten) ohne Schließen des gemeinsamen Vorgangs. `Aus Pause zurückmelden` startet keine Zeitmessung; Rückkehrfrage, ob zuletzt unterbrochene Rüstung fortsetzen? Bei „Ja“ Modus auswählen und neue persönliche Buchung beginnen; bei „Nein“ offen lassen, jederzeit später fortsetzbar.
6. Beim geplanten Schichtende **nur erinnern**; kein automatischer Stopp allein wegen Soll-Schichtende. Bei tatsächlicher Abmeldung oder nachfolgender Arbeitsplatzübernahme wird die persönliche Buchung des Vorgängers einmalig unterbrochen. Nachfolger wird auf offene Rüstung hingewiesen und kann sofort „Ja“ oder „Nein, später“ wählen.
7. Andere weiterhin aktive Mitarbeiterbuchungen bleiben bei Pause, Abmeldung oder Übergabe unberührt. Bei gemeinsamem `Rüsten beenden` endet hingegen jeder aktive Teilnehmer.
8. Mitarbeiter dürfen eigene fehlerhafte Rüstbuchungen korrigieren; zuständiger AL/aktive Vertretung ebenso die ihrer Abteilung. **Begründung und nicht manipulierbarer Alt-/Neu-Audit** für Korrekturen. Überlappungen und konkurrierende Abschlüsse/Übergaben verhindern.
9. Keine MOBATIME-Anbindung jetzt, kein zweites Anwesenheits-/gesetzliches Arbeitszeitkonto. Spätere Integration ausschließlich als **deaktivierter Platzhalter**. Fertigungs-Ist-Zeiten sind keine MOBATIME-Logins.

## E. Spannungen, bestehender Stückzahlzähler und Menge

1. **Einziger vorhandener Humbel-App-Stückzahlzähler** bleibt bestehen; keine zweite Zähllogik entwickeln, keine direkte Heidenhain-Schnittstelle in diesem Auftrag. Bestehende Gut-/Ausschuss-/„In Abklärung“-Logik, Spannungs- und Mitarbeiterzuordnung, Kennzeichnungen z. B. `1AA`, Ausschuss-Historie und 6M-Erfassung erhalten; konkrete Codes anhand vorhandener Implementierung verifizieren, nicht neu interpretieren.
2. Jede Spannung hat eine explizite Zuordnung zu **genau einer OP**; eine OP kann **mehrere Spannungen** enthalten. Standardfall: eine Spannung pro OP. Der Artikel/BA trägt Reihenfolge und Zuordnung.
3. **Erst die letzte erforderliche Spannung** darf fertige OP-Gutteile beitragen. Frühere Spannungen melden bearbeitete Stücke für ihren Mitarbeiter, aber **noch keine OP-Gutteile**. Bei verschiedenen Mitarbeitern je Spannung Personenzählungen separat, OP-Gutteile nur **einmal**. Ausschuss/ungeklärte Teile keinesfalls automatisch als Gutteile verbuchen.
4. Neu gestartete **Auftrags-OP** erhält einen eigenen, bei **0 beginnenden Zähler**. Dieselbe OP nach Pause/Schichtwechsel **nicht zurücksetzen**. Neue Maschinen-/Fertigungsabschnitte dürfen je Abschnitt bei 0 starten, während der **OP-Gesamtstand erhalten bleibt**. Keine alten Daten löschen; Idempotenz und parallele Zählerbuchungen prüfen.
5. Maschinenabschnitt, Mitarbeiter, Spannung und OP an jeder Zählung nachvollziehbar zuordnen, ohne vorhandenes Stations-/Spannungsmodell unkontrolliert zu konvertieren.
6. **Parallele Folge-OPs/Teilmengen sind erlaubt**, bevor Vorgänger-OP formal abgeschlossen ist. Bereits fertiggestellte Gutteile der Vorgänger-OP erhöhen laufend die aktuell verfügbare Bearbeitungsmenge der Folge-OP; die ursprüngliche/aktuelle BA-Sollmenge bleibt separat sichtbar. Rein rechnerische Verfügbarkeit nicht mit physisch bestätigtem Materialstandort/Materialübergabe verwechseln.
7. Mehr gezählte Stücke in der Folge-OP als bisher als Gutteile aus der Vorgänger-OP: **Buchung behalten** (nicht blockieren), Mengenwiderspruch als offenen Klärfall kennzeichnen, zuständigen AL benachrichtigen. Weder automatisch Stücke erfinden noch Mengen ändern.

## F. BA-Menge, Nachfertigung und Klärfälle

1. **BA-Sollmenge ausschließlich in der eigenständigen Humbel-BA-Verwaltung** ändern (keine ERP-Kopplung). Alte Menge, neue Menge, Änderer, Zeitpunkt und Grund dort, wo vorgeschrieben, auditieren.
2. **Mengenerhöhung**: betroffene abgeschlossene OPs automatisch für **Nachfertigung** freigeben; zusätzliche Fertigungs-/Rüstabschnitte neu erfassen, alte Daten unverändert. BA-/OP-Historie nicht ersetzen, OP-Gesamtzähler bei Fortsetzung nicht zurücksetzen. Folge-OPs erhalten zusätzliche Mengen erst nach tatsächlich erfassten Gutteilen der Vorgänger-OP.
3. **Mengenreduzierung** sogar unter bereits gefertigte Gutteile zulässig – mit Warnung, **Pflichtbegründung**, separater Überproduktionsanzeige; tatsächliche Daten nicht verändern. Alle bereits verfügbaren Gutteile dürfen nachfolgende OPs weiterhin bearbeiten; die BA-Sollmenge ist nur der separate Zielwert.
4. Mengenwiderspruch darf erst durch zuständigen AL, aktive Vertretung oder Administrator **als geklärt** markiert werden. Mitarbeiter dürfen Erklärungen und ihnen zustehende Buchungskorrekturen vornehmen, aber keinen Klärfall abschließen.
5. Mit Pflichtbegründung ist **`Geklärt mit verbleibender Abweichung`** zulässig; Restdifferenz weiterhin sichtbar. Fallabschluss verändert keine Mengen.
6. **OP-Abschluss** ist nur möglich, wenn alle **zum Abschlusszeitpunkt offenen zugeordneten Mengenklärfälle** geschlossen sind; „Geklärt mit verbleibender Abweichung“ erfüllt diese Bedingung. Frühere regelkonforme OP-Abschlüsse bleiben auch bei später entstehenden neuen Klärfällen historisch bestehen. Gleiches gilt für vollständig abgeschlossene BA.
7. Bei einer erst **nach BA-Abschluss** entdeckten Differenz BA nicht automatisch öffnen; separaten Klärfall eröffnen, deutlich kennzeichnen, AL informieren. Nur Mengenerhöhung zur Nachfertigung bewirkt die vereinbarte Wiederfreigabe.

## G. Abschlüsse und Verantwortlichkeiten

1. **Sollstückzahl erreicht** zeigt Status/Hinweis, schließt OP **nicht automatisch**. OP nur ausdrücklich durch an der OP berechtigten beteiligten Mitarbeiter, zuständigen AL/aktive Vertretung oder Administrator abschließen. Mitarbeiter darf OP auch unter Sollmenge mit **Pflichtgrund** abschließen, AL benachrichtigen. Fehlmenge nicht automatisch als Ausschuss buchen.
2. Vor OP-Abschluss Ist-Mengen, Ausschuss, offene Klärfälle, aktive Tätigkeitsbuchungen anzeigen. Offene Mengenklärfälle **sperren** den Abschluss (serverseitig atomar prüfen). Nach Abschluss keine normalen neuen Zähl-/Zeitbuchungen, nur berechtigte Korrekturen beziehungsweise echte Nachfertigungsfreigabe.
3. **BA-Abschluss** erst nach ausdrücklichem Abschluss **aller OPs** und aller **offenen** Mengenklärfälle. Ebenfalls **ausdrücklich** (nie automatisch) und ausschließlich durch den beim BA geführten **hauptverantwortlichen AL**, dessen aktive Vertretung oder Administrator.
4. Bei BA-Anlage Hauptverantwortlichen AL hinterlegen; ausführende AL können bei OPs wechseln, **BA-Hauptverantwortung bleibt** bis zur berechtigten Übertragung bestehen.
5. Der bisherige hauptverantwortliche AL, seine aktive Stellvertretung oder Administrator dürfen BA-Hauptverantwortung direkt an neuen AL übertragen, **Pflichtgrund** aus administrativ pflegbarer Liste und optionalem Freitext; Benachrichtigung an neuen AL, vollständige alte/neue Historie.

## H. Planung, Berechtigungen und Sicherheit

- **Kapazitätsplanung ausschließlich PLAN**: geplante Schichten, Sollarbeitszeiten, budgetierte Rüst-/Bearbeitungszeiten, Urlaub/Krankheit und Maschinenbelegung; Personalaufwand und Maschinenlaufzeit **getrennt**. Mannlose CNC-Laufzeit belastet Maschine, nicht automatisch Personal. Fertigungs-Ist (Rüst-/Stückzeiten) für nachträgliche Soll-Ist-Auswertung, **nicht** als personalrechtliche Anwesenheitszeit; keine aktive MOBATIME-Verbindung.
- Alle fachlichen Standardwerte/Schwellen/Berechtigungszuordnungen unter Admin-Konfiguration im zulässigen Umfang. Keine Konfiguration darf zwingende Integritäts-, Datenschutz-, Rollen- oder Audit-Regeln abschalten.
- Reale Rollen / Abteilungszuordnung / aktive automatische oder befristete Vertretung **serverseitig** prüfen, auch bei direktem RPC. Normale Mitarbeiter keine unberechtigten AL-Funktionen. Bei abteilungsübergreifenden OPs beachten: ursprünglich verantwortlicher AL, ausführender AL und BA-hauptverantwortlicher AL können verschieden sein.
- Wichtige Vorgänge transaktional und konsistent: Rüstsegmentwechsel, gemeinsames Beenden, Maschinenwechsel, Stückzählung, Mengenklärfall, OP-/BA-Abschluss, Mengen-/Verantwortungsänderung. Audit: Actor, Zeitpunkt, Bereich, Objekt, Alt-/Neuwert, Grund, Autorisierung; Fehlschlag des Audits blockiert Geschäftstransaktion.
- Keine produktive DB-/RLS-/Auth-Anpassung vor Autorisierung, Originalschema-/ACL-Abgleich und echten isolierten JWT-/RLS-/PostgREST-Tests; alte Schwachstellen aus Draft #15/#16 **nicht als behoben ausgeben**.
- Mobile Debugkonsole/temporäre Scanner-Overlays aus fertiger App entfernen; nur produktive Logs behalten.

## I. Zu verifizierende Grundlagen / keine stillschweigenden Annahmen

1. Bestand des vorhandenen Stückzahlzählers: exakte Bedeutung von `1AA`, Spannungszähler, Ausschussbuchungen, 6M, Idempotenz und Speicherpfad **im Code/DB belegen**.
2. Exakte OP-Barcodecodierung, Barcode-Scanner/Hardware, vorhandene OP-/BA-Tabellen, Feature-Flags und Rollengrenzen vor Änderung prüfen.
3. Neue BA-Erstmaske hat nur drei manuelle Pflichtfelder; **bestehendes #8 sieht derzeit auch Termin und Maschine vor**. Vor Codeänderung kontrollierten Übergang und erforderliche optionale Planfelder designen; keine alten Daten ausradieren.
4. Hauptverantwortlichen AL beim automatischen BA-Anlegen eindeutig zuordnen: Regel offen, darf den Drei-Felder-Einstieg nicht verkomplizieren; keine unautorisierte Rollenvergabe.
5. Physische Materialübergabe vs. rechnerische Teilmengenverfügbarkeit separat definieren. Alte #4/#6 „Teilmenge nur Ausnahme / Vorgänger vollständig“ **durch jüngere Freigabe zur parallelen Bearbeitung ersetzt**.
6. Künftiges Berechtigungs-/RLS-Konzept und Legacyschema sind noch nicht produktionsreif (PR #15/#16 Draft). Diese Spezifikation ersetzt keine technische DB-Abnahme.

## J. Geplanter technischer Schnitt / Abnahmematrix

**Getrennte zukünftige Implementierungs-PRs; dieser PR selbst ist nur Dokumentation.**

1. Artikel-/BA-Snapshot und einfache BA-Anlage an erster OP, Rollen, Barcode-Aufruf (ohne produktive Datenmigration).
2. Maschinenfestlegung, Sperren, Rückgabe, Maschinenabschnitte und Audit.
3. Persönliche/gemeinsame Rüstzeitsegmente und allgemeine Pause, auch Schichtwechsel, ohne MOBATIME.
4. Bestehenden Spannungszähler **integrieren statt ersetzen**, OP-Mengenflüsse und parallele Folge-OPs.
5. Mengenänderungen, Klärfälle, OP-/BA-Abschlüsse, Nachfertigung und Abteilungsverantwortlichkeiten.
6. Planbasierte Maschinen-/Personalprognose und separate Ist-Auswertung.

**Pflichtprüfungen pro Code-PR:** `npm ci --ignore-scripts`, `node --check schichtplan_mvp.js`, `node --check tests/smoke.spec.js`, `npm run check:syntax`, `npm run test:ci`, `npm run test:qa`, `npm run qa:check`, `npm run test:agents`, `npm run qa:agent-server`, `npm run test:e2e`, `git diff --check`; weitere vorhandene SQL-/RLS-/Browser-/Visual-/Datenintegritätstests nicht abschwächen. Echte DB-Tests nur in isolierter freigegebener Testumgebung. Version und VERSION_LOG mit tatsächlicher Berlin-Zeit je Änderung; PR-Head-SHA, CI-Link, Testergebnisse und Risiken ausweisen. **Kein automatischer Merge oder Deployment.**

### Mindest-Akzeptanzfälle für spätere Implementierung

- OP-Barcode erkennt exakt die richtige OP; BA direkt von OP 1 mit genau drei Pflichtangaben anlegen, identische BA nicht duplizieren.
- Eine OP hat 3 Spannungen; zwei verschiedene Mitarbeiter zählen 100/100/98: OP-Gutteile = 98, nicht 298; Ausschuss `1AA` bleibt nach Legacy-Regeln.
- Gemeinsame Rüstung mit intern/extern-Wechsel und zwei Mitarbeitern; Pause, Mitarbeiterwechsel, Schichtende-Erinnerung, gemeinsames Beenden ohne doppelten Zeitabschluss.
- Abteilungsübergreifende OP-Maschinenfestlegung, spätere direkte Rückgabe, Pflichtgrund, historischer Maschinenabschnitt, neuer Rüstvorgang.
- Nach 40 Gutteilen startet Folge-OP vor Abschluss des Vorgängers. Nach 41. Stück bei nur 40 verfügbaren Stück: Buchung bleibt, ein offener Mengenklärfall entsteht und blockiert den OP-Abschluss.
- AL schließt mit Restabweichung + Pflichtgrund → OP-Abschluss erlaubt; nach Abschluss auftauchender Konflikt erzeugt Klärfall ohne historische Wiederöffnung.
- Mengenerhöhung BA: Nachfertigung freigegeben, vorhandene Zeiten/Gutteile unberührt; Mengenreduzierung auf unter Gutmenge: Überproduktion, Warnung und Grund, nachfolgende Bearbeitung aller vorhandenen Gutteile erlaubt.
- Artikel-Änderung betrifft alte BA nicht automatisch; gezielter Feldabgleich, bei aktiver OP mit Sicherheitsregeln; bei vollständig abgeschlossenem BA gesperrt.
- Mitarbeiter kann OP, nicht BA abschließen; BA-Abschluss nur hauptverantwortlicher AL/Vertretung/Admin, nachdem alle OPs und Klärfälle abgeschlossen.
- Keine unautorisierte Rolleneskalation, keine doppelte Buchung bei konkurrierenden Klicks, vollständiges unveränderliches Audit; kein Live-DB-Schreibtest.
