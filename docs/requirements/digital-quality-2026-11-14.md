# Digitale Qualitätsprüfung – verbindliche Zielregeln

**Stand 14.11.2026, Dokumentationsversion 0.17.3. Nicht implementiert oder produktiv freigegeben.** Auftrag aus dem Chat, Übergabebezug [Issue #33](https://github.com/butzi124-ship-it/schichtplan-mvp/issues/33). Basis PR #37, `a5c08a4f1cf72c544a12e59da4528bf30c9dd154`, Version 0.17.2. Neuer Branch `docs/digital-quality-specification`. PR #37 ist weiterhin offen/Draft; am 14.11. Remote-Refs und vollständige Liste offener PRs geprüft (21, kein neuerer QS-Konzept-PR). Bestehende Implementierungen/History unverändert.

**B** = mit diesem Auftrag verbindlich beschlossen; **P** = technischer Vorschlag; **O** = noch offene Detailentscheidung. Alle beschriebenen Automationen sind Zielverhalten, keine aktivierten Funktionen. Keine neuen SQL-Dateien, Rollen, Buckets, Frontendmodule oder Geschäftslogik. APP_VERSION/VERSION_LOG ausschließlich nach AGENTS.md.

## 1. Lesepfad und Vorrang

| Dokument | Zweck |
|---|---|
| Dieses Dokument | Fachspezifikation und neue Beschlüsse F01–F12 |
| [Technik/Datenmodell](../architecture/digital-quality-model.md) | Bestehender Zähler, Transaktionen, Fälligkeit und Teilenummern, Mandantentrennung |
| [UI und Referenzen](../architecture/digital-quality-ui.md) | Verbindliches Bildlayout, Toleranzbalken, Druck, Arbeitsanweisung und Unsicherheiten |
| [Entscheidungen/Testplan/Änderungen](../architecture/digital-quality-acceptance.md) | Auflösung bisheriger D-Fragen, verbliebene O-Punkte, Abnahmekriterien und Phasen |
| [Aktuelle Prüfergebnisse](../architecture/digital-quality-validation.json) | Tatsächliche Dokumentations-/Bestandsprüfungen; neue Funktionstests getrennt |
| [Dokumentenbasis aus #37](../architecture/documents-inspections.md) | Dokumenttypen, sichere Dateiablage, Revisionen, Vorlagen/Originale und Korrekturketten |

Jüngere explizite Beschlüsse dieses Auftrags ersetzen widersprechende **Zielannahmen** aus #37/älteren Spezifikationen nur im angegebenen QS-Umfang. Historische Fotos, Sicherheitsbefunde und Prüfnachweise werden nicht umgedeutet. Die Varianten A (digital) und B (Scan) bleiben für die allgemeine Dokumentenverwaltung offen. Für die hier beschlossene dynamische Gutteilprüfung ist ein serverseitig auswertbares Ergebnis nötig; ein bloßer Scan oder OCR-Text ist kein automatischer Gutteilnachweis. Ein alternatives kontrolliertes Transkriptions-/Papierverfahren bleibt gesondert abzunehmen.

## 2. Anforderungsregister (B)

| ID | Beschluss | Abgrenzung / Testanker |
|---|---|---|
| F01 | Blau-weißes ERP-Layout nach beigefügtem Messprotokoll: Tabelle, Toleranzbalken, Verlauf, Werkzeugkorrekturen, deutliche Statusanzeigen, Touch | Fachregeln vor Bildzahlen. Humbel-Druck DIN A4 quer; andere Firmen Hoch/Quer konfigurierbar; A01/A02 |
| F02 | Je Merkmal eigenes Intervall; bei Teilzählung nur aktuell fällige Merkmale in Messmaske; ohne Fälligkeit keine Aufforderung | Erst-/Letztstück und zwingende Nachprüfung zusätzlich; Gutteil erst nach erfolgreichem Pflichtnachweis; A03–A05 |
| F03 | 100 %, jedes 2./5./10. Teil, individuelle Prozentsätze, freie Zyklen, Abteilungs-/Artikel-/Merkmalsvorgaben | Dreherei-Anweisung nur Referenz, keine globale Vorgabe. Prozentalgorithmus/Scopekonflikte präzisieren; A03 |
| F04 | Balkenmitte = Toleranzfeldmitte, Auslenkung links/rechts nach Messabweichung. Grün innerhalb Eingriffsgrenzen, Gelb Grenze erreicht, Rot außerhalb Toleranz | Untere/obere Eingriffsgrenze unabhängig je Merkmal. Gut/Nicht-gut ohne Zahl, negativ voller roter Balken; A02/A06 |
| F05 | Gelb: tatsächliche Korrektur durchführen, dokumentieren, nächstes Teil zwingend prüfen; OP bis dokumentierter Korrektur gesperrt | AL erhält Meldung und bestätigt nur Kenntnisnahme; **keine zusätzliche Freigabe**. Fehlende Kenntnisnahme sperrt OP nicht weiter; A07 |
| F06 | Rot: Teil nicht konform / In Abklärung; zuständiger AL oder berechtigter STV entscheidet Ausschuss, Nacharbeit oder zulässige Sonderfreigabe | Keine Entscheidungsmacht allein aus PL-/QS-/TA-Titel. Originalmessung bleibt unverändert; A08 |
| F07 | Unkritisch ohne hinterlegte Prüfpflicht: manuell Gut/Ausschuss/In Abklärung. Prüfpflichtig: Qualitätsstatus automatisch nach fälligen Nachweisen/Entscheidungen | Keine Umgehung durch manuelles Gut, keinen Doppelzähler. Abklärung→Ausschuss atomar umbuchen; Produzent bleibt erhalten; A09 |
| F08 | Bei OP-Übernahme Ja/Nein zur bestehenden individuellen Nummerierung | Nein: OP-lokal ab 1; Ja: bei jeder Teilbuchung vorhandene Nummer und Eindeutigkeitsprüfung. Technische ID getrennt; A10 |
| F09 | Später gravierte Nummer möglich; keine behauptete physische Identität zwischen früher unmarkierten Teilen verschiedener OPs | Auftrag/OP/Produzent/Erfassungszeit als Nachweis; nachträgliche Links nur mit belegter Zuordnung, niemals aus gleicher Laufnummer; A10 |
| F10 | Eigenständiges Qualitätscockpit: Rot zuerst, Gelb danach, Grün/abgeschlossene Vorgänge zuletzt; je Kategorie nach Aktualität | PL sämtliche Produktionsabteilungen, AL zugeordnete, QS alle Qualitätsbereiche der Produktion, Mitarbeiter arbeitsnotwendig; A11 |
| F11 | Rote Abweichung benachrichtigt zunächst nur zuständigen AL. Sonderfreigabe im betreffenden Auftrag für QS mit Entscheidername/Zeit/Grund sichtbar, **kein QS-Push** | Gelbmeldung plus Kenntnisnahme, keine versteckte Freigabe. Zustellkanal/fehlender AL offen; A11/A12 |
| F12 | Allgemeine Software und Firmeneinstellungen strikt trennen. Humbel-Vorlagen exklusiv, andere Firmen neutrale Vorlagen | Intervalle/Grenzen/Merkmale/Layout/Benachrichtigungen mandantenspezifisch; keine Kopie interner Anweisungen in generische Seeds; A13 |

## 3. Dynamischer Prüfablauf

Eine Teilbuchung identifiziert zunächst Firma, BA, OP, notwendigen Spannungs-/Maschinenabschnitt, produzierenden Mitarbeiter und technische Teilbuchungs-ID. Danach wird unter dem gültigen eingefrorenen Regelstand die Vereinigungsmenge fälliger Merkmale bestimmt: regelmäßiger Zyklus **oder** Erststück **oder** Letztstück **oder** verpflichtende Nachprüfung. Mehrere Gründe für dasselbe Merkmal erzeugen nur eine konkrete Fälligkeit mit mehreren Ursachen, keine mehrfachen Gutbuchungen. Andere Merkmale bleiben in der Eingabemaske verborgen, ihre historischen Ergebnisse dürfen in berechtigter Historie sichtbar bleiben.

Keine Fälligkeit: keine Messmaske. Eine ausdrücklich hinterlegte Prüfpflicht mit gerade nicht fälligem Merkmal wird nicht zur manuellen Qualitätswahl umgedeutet: der Regelstand bestimmt den zulässigen Status. Ein fehlendes/defektes Profil ist nicht dasselbe wie ausdrücklich ohne Prüfpflicht. Bei unbestimmter Pflicht keine automatische Gutfreigabe.

Fälliger Nachweis fehlt: keine endgültige Gutbuchung, keine Erhöhung der verfügbaren OP-Gutteile für Folge-OPs. Eingabe bleibt nachvollziehbar offen; Rot wird explizit In Abklärung geführt. Konkrete Abbildung wartender Prüfungen auf bisherigen Abklärungszähler vs. separate Reservierung ist eine Integrationsentscheidung O04, keine bereits eingeführte neue Zähllogik.

Erst-/Letztstück: Zusätzlich zu Regelzyklen, kein automatisches Überspringen wegen Prozentquote. Wie das tatsächliche letzte Teil bei Ausschuss, Nachfertigung, Mengenänderung, Unter-Soll-Abschluss oder nicht mehr verfügbarem Teil feststeht, muss vor Adapterfreigabe O02 erfüllen. Bis dahin keine erfolgreiche Letztstückprüfung aus einer Sollmengen-Gleichheit erfinden. Erfolgreiche Pflichtprüfungen dürfen nicht durch einen parallelen OP-Abschluss umgangen werden; vorhandener ausdrücklicher BA-/OP-Abschluss bleibt erhalten.

## 4. Ergebnis, Korrektur und Entscheidungen

Numerische Bewertung und Prozesszustand sind getrennt. Eingriffsgrenze erreicht, aber noch innerhalb Zeichnungstoleranz: Gelb; Überschreitung der Zeichnungstoleranz hat Vorrang und bleibt Rot. Der Bereich zwischen erreichtem Eingriff und Toleranzgrenze wird im technischen Entwurf als Gelb weitergeführt, damit kein unklassifizierter Bereich entsteht; genaue Randwertkonventionen/einseitige Grenzen O03. Nicht bewertbar/fehlend ist weder Grün noch akzeptiert. Ungültige Konfiguration sperrt Bewertung statt sie auf Gut zurückfallen zu lassen.

**Gelb:** Warnereignis und OP-Korrektursperre gemeinsam erfassen, Maßnahme tatsächlich ausführen, Durchführung samt Person/Zeit/Beschreibung dokumentieren. Kein automatisches Verstellen einer Maschine. Danach wird ausschließlich diese Korrektursperre aufgehoben; andere Sperren bleiben bestehen. Der Nachprüfauftrag für das nächste Teil bleibt offen und kann von Schicht-/Mitarbeiter-/Maschinenwechsel nicht gelöscht werden. AL-Kenntnisnahme ist ein eigener Zustand, **kein Gate für Wiederanlauf**.

Sichere vorgeschlagene Auflösung für das gelbe aktuelle Teil: solange erforderliche Korrektur nicht dokumentiert ist, keine endgültige Gutfreigabe; danach bei sonst erfüllten Pflichten gutfähig, historische Gelbbewertung bleibt erhalten. Das nächste Teil wird trotzdem zwingend geprüft. So entsteht kein Zirkelschluss „nächstes Teil vor Freigabe des ersten herstellen, obwohl OP gesperrt“. Genaue Teil-/Merkmalsmenge der Nachprüfung O02; mindestens auslösendes Merkmal darf nicht entfallen. Erneutes Gelb erzeugt erneut Korrektur/Nachprüfpflicht, keine Unterdrückung als bereits bekannt.

**Rot:** kein Gutteil, In Abklärung; fachliche Entscheidung des zuständigen AL/aktuell berechtigten STV. Ausschuss bucht dasselbe Teil automatisch aus Abklärung nach Ausschuss um. Nacharbeit behält Identität/Produzent und führt zu einem dokumentierten Nacharbeits-/Nachprüfzweig, nicht zu neuem Gutteil per Klick. Zulässige Sonderfreigabe bindet genau Teil, Merkmal/Abweichung, Entscheider, Zeitpunkt, Pflichtgrund, Ergebnissatz und gegebenenfalls erlaubte Bedingungen. Roten Originalwert nicht ändern oder grün einfärben; „sonderfreigegeben“ ist eigener Entscheidungsstatus. Umfang und rechtliche/vertragliche Obergrenzen O05 bleiben zu klären, ohne das beschlossene AL-/STV-Entscheidungsrecht wieder als grundsätzlich offen zu führen.

Unveränderliche Messereignisse und additive Folge-/Korrekturereignisse. Maschinen-/Werkzeugmaßnahme ist nicht dasselbe wie eine Korrektur eines falsch erfassten Messwerts. Produktionsmitarbeiter dokumentieren ihre **Fertigungsprüfung und tatsächliche Maßnahme**, erhalten aber weiterhin kein Recht, fremde **QS-Ergebnisse zu ändern oder zu kommentieren** (Q04 aus #37). Prüfer, Produzent, Korrekturausführender und Entscheider können verschieden sein und werden getrennt erfasst.

## 5. Verbindliche Rechte- und Benachrichtigungsmatrix

Alle Rechte brauchen vertrauenswürdige aktive Identität, Firmenbindung, konkrete Aufgabe/Abteilung und aktuellen Servernachweis aus dem später abgenommenen Rollenmodell; Mitarbeiterprofil oder Clientrolle allein genügen nicht.

| Aktion | Produktion MA | AL | Berechtigter STV | PL | QS | HR / technischer Admin |
|---|---|---|---|---|---|---|
| Eigene fällige Fertigungsprüfung / tatsächliche Maßnahme | B: berechtigte Arbeit, keine QS-Überschreibrechte | Eigene bestätigte Fachauthority | Gültige operative Authority | Nicht aus Cockpitlesen allein | B: direkte QS-Erfassung wie Q02 | Kein automatisches Fachrecht |
| Qualitätscockpit | Nur arbeitsnotwendige Projektion | B: alle ausdrücklich zugeordneten Abteilungen gleichberechtigt | Nur ausdrücklich beauftragter operativer Scope, keine pauschale AL-Gesamtsicht | B: sämtliche Produktionsabteilungen der eigenen Firma | B: sämtliche Produktionsabteilungen der eigenen Firma | Kein Inhaltsrecht aus Titel |
| QS-Ergebnisse lesen | B: einschließlich Prüfername, arbeitsnotwendiger Scope | Eigener Scope | Gültiger Scope | Produktionsgesamtsicht | QS-Gesamtsicht | Nur separat bestätigte Fachrolle |
| QS-Ergebnisse ändern / kommentieren | **Verboten** | Keine Macht aus Leitungstitel allein | Keine Macht aus Vertretung allein | Keine Macht aus PL-Titel allein | Nur autorisierte additive Erfassungs-/Korrekturregeln | Keine |
| Gelb: Kenntnisnahme | Kein Leitungsnachweis | B: zuständiger AL, ohne zusätzliche Freigabe | Ersatzkenntnisnahme O06 | Kein zusätzlicher Freigabeschritt | Keine QS-Freigabepflicht | Keine |
| Rot: Ausschuss / Nacharbeit / zulässige Sonderfreigabe | Nicht erlaubt | B: zuständige operative Abteilung | B: ausdrücklich gültiger passender Vertretungsumfang | Nur bei zusätzlich bestätigter zuständiger Authority | Keine Entscheidung aus QS-Lesetitel allein | Keine |
| QS-Aufgabenberechtigungen vergeben | Keine Selbstvergabe | AL QS nur QS-zugeordnete Empfänger und eigene delegierbare Rechte | Nur passende Sondergrantfreigabe nach 8A | Kein neues Recht aus Cockpitrolle | Kein allgemeines Vergaberecht | Keine automatische Fachvergabe |

Gelb: Meldung an zuständigen AL, Kenntnisnahme dokumentieren. Rot: automatische Erstbenachrichtigung ausschließlich an zuständigen AL; **kein automatischer zusätzlicher QS-/PL-/STV-Empfänger**. Berechtigter STV kann nach seiner Authority entscheiden, ohne dadurch als neuer automatischer Empfänger festgelegt zu sein. Fehlender/mehrdeutiger/abwesender AL und Eskalation O06; keine stillschweigende Sendung an irgendeinen Admin. Sonderfreigabe ist im betreffenden Auftrag für QS sichtbar mit Name, Serverzeit und Grund, kein Push. Kanal-/Wiederhol-/Eskalationskonfiguration darf Firmenisolation und gesetzte Empfängergrenzen nicht ungeprüft erweitern.

Cockpit sortiert nach **Handlungsbedarf**, Messfarbhistorie bleibt zusätzlich erhalten. Vorschlag: offene rote Entscheidung zuerst, offene gelbe Korrektur/Nachprüfung/Kenntnisnahme danach, abgeschlossene Fälle grün gruppiert. Sonderfreigabe kann abgeschlossen gruppiert sein, zeigt dennoch roten Ursprungswert und Freigabekennzeichen. Innerhalb Kategorie neueste fachliche Aktualisierung zuerst, stabile ID als Gleichstand; reine Ansicht verlängert keine Priorität oder Aufbewahrungsfrist.

## 6. Wirkung auf vorhandene Anwendung

Der bestehende Spannungs-/Mitarbeiterzähler bleibt die einzige Mengenwahrheit. Letzte erforderliche Spannung liefert weiterhin einmalig fertige OP-Gutteile; Vorgänger-/Teilmengenregeln, Ausschusscodes, Rüstung, Maschinenabschnitte, Nachfertigung und ausdrückliche Abschlüsse werden nicht neu gebaut. Neu **fachlich beschlossen** sind die zukünftigen qualitätsbedingten Gates und automatischen Umbuchungen; technische Integration braucht den gesonderten Plan und Tests. Die frühere Formulierung „keine automatische QS-Sperre/-Buchung beschlossen“ aus #37 ist insoweit überholt, nicht rückwirkend implementiert.

Keine aktuellen Benutzer-/Rollen-/Daten-/Featureflagänderungen. Sicherheitsbefunde und fehlende echte Auth-/Session-/native Abnahme aus #32/#35/#36 bleiben Produktivblocker. Mandantenfähigkeit ist neues Ziel, **kein vorhandener Schutz durch bloßes tenant_id-Feld**. Auch Dokumente, Quellenbilder, Messwerte, Benachrichtigungen, Audits, Exporte und Backups müssen getrennt werden.
