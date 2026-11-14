# UI-/UX-Spezifikation und Referenzanalyse

Zum [Fachvertrag](../requirements/digital-quality-2026-11-14.md). **Dokumentiert, nicht gerendert oder implementiert.** Beide Anhänge am 14.11.2026 visuell geprüft, die Anweisung zusätzlich vergrößert. Bilder nicht in Git übernehmen; Firmenunterlagen dürfen später nicht in neutrale Mandantenvorlagen gelangen. Quellenhashes im [Prüfprotokoll](digital-quality-validation.json).

## 1. Verbindliches Messprotokoll-Layout (B)

Quelle `74271.png`: blauer Kopf, weiße/hellblaue ERP-Flächen, lesbare Messwerttabelle, Statusicons, Toleranzbalken, Verlauf sowie dokumentierte Werkzeugkorrektur. Dieses Bild ist als **Gestaltung verbindlich**, seine Werte/Personen-/Maschinen-/Auftragsangaben sind ausschließlich Beispiele und keine produktiven Datensätze oder bestätigten Messregeln.

| Bereich | Zukünftiger Vertrag |
|---|---|
| Kopf | Firma, Artikel/Bezeichnung, BA, OP, tatsächliche Maschine/Abschnitt, Menge/Los, aktuelle Teilkennung, Status; nicht alle Merkmale unter einem irreführenden globalen Intervall zusammenfassen |
| Messmaske | Nur fällige Merkmale, mit Fälligkeitsgrund und eigener Regel. Tabelle Pos./Merkmal/Einheit/Soll-/Toleranzfeld/untere-obere Eingriffsgrenze/Messwert/Abweichung/Balken/Ergebnis |
| Maßnahme | Bei Gelb Pflicht: tatsächlich ausgeführte Maßnahme, Beschreibung, Werkzeugreferenz/Korrektur sofern betroffen, Ausführender, Durchführung-/Serverzeit. Kein automatischer Maschinenoffsetbefehl |
| Verlauf | Pro Merkmal, nur vergleichbare Einheit/Prüf-/Vorlagenversion, echte Messungen mit Teil-/Zeitbezug; ungemessene Teile nicht interpolieren oder als Null anzeigen. Nachmessungen und Korrekturereignisse getrennt erkennbar |
| Status | Erststück/Letztstück/Nachprüfung je Grund, noch fällige Merkmale, OP-Sperrgründe, Korrektur dokumentiert vs. AL-Kenntnisnahme ausstehend. „Keine Prüfung fällig“ ohne Messaufforderung |
| Aktionen | Speichern als Messnachweis ≠ endgültig Gut buchen ≠ OP schließen. Fallentscheidung nur berechtigter AL/STV; kein zusätzlicher Gelb-Freigabebutton |

Desktop zeigt Tabelle, Maßnahme und Verlauf übersichtlich nebeneinander. Tablet große Zeilen/Bedienflächen, Smartphone Merkmalskarten mit gleichem Informationsumfang/Fälligkeitsfilter; horizontale Gesamttabelle zusätzlich für berechtigte Ansicht. Vorschlag mindestens44px Touchziele, eindeutige zugängliche Labels/Tastaturfokus, Icon **und Text** statt Farbe allein. OP-/Teilwechsel warnt bei ungesicherten Eingaben; nie Ergebnis von Teil A unter Teil B anzeigen. Scanner-/Tastatureingaben öffnen Kontext, buchen nicht durch unbeabsichtigtes Wiederholsignal.

### Fachliche Korrektur der Bildbeispiele

- „10 % (alle4 Teile)“ ist widersprüchlich: jedes vierte entspricht25 %. Pro Merkmal korrekte konfigurierte Regel anzeigen, nicht diese Beschriftung kopieren.
- ØInnen ist mit80.028 bei Eingriff ±0.030 gelb dargestellt, obwohl rechnerisch noch unter der oberen Grenze. Status aus echten Regeln berechnen, nicht aus Bildfarbe übernehmen.
- Verlaufslinien/Toleranzmarken im Bild passen nicht sämtlich zur Tabelle. Beide Ansichten müssen dieselbe Regel-/Messrevision verwenden.
- Einheit mm im Tabellenkopf darf nicht pauschal auf Rauheit/Gewinde übertragen werden. Einheit je Merkmal; qualitative Prüfung ohne erfundene Zahl. Passung6H ist kein numerischer Messwert.
- Bildname/-Zeit/Artikel32505 sind Beispiele. Das frühere Originalformular52978 bleibt eigene Referenz; seine Farben/Sollmaße werden nicht mit dem Designbeispiel zu einer Universalvorlage vermischt.

## 2. Grafische Toleranzanzeige

**B:** grafischer Nullpunkt ist die **Mitte des Toleranzfelds**, nicht zwangsläufig das Nennmaß. Bei bestätigter endlicher Unter-/Obergrenze L/U ist Mittelpunkt (L+U)/2; Auslenkung wird vom Mittelpunkt nach links oder rechts bis zum Messwert gefüllt. Nennmaßabweichung kann separat numerisch stehen, klar von der Balkenauslenkung unterschieden. Bei asymmetrischer Toleranz ist eine Messung genau im Nennmaß deshalb nicht unbedingt in Balkenmitte.

Beispiel nur zur Abnahme: Soll10.000, Grenzen9.980/10.040 → Mitte10.010. Messung10.000 liegt im Balken links der Mitte, obwohl Nennmaßabweichung0.000 ist. Links/rechts gültige Eingriffsgrenze separat markieren. Sehr weit außen liegende Werte am Displayrand mit Pfeil/Originalzahl anzeigen, rechnerisch niemals auf den Rand klemmen.

Grün innerhalb der bestätigten Eingriffsgrenzen; Gelb bei erreichtem Eingriff (und nach Entwurf bis zur Toleranzgrenze); Rot außerhalb Toleranz. Genaue inklusive/exklusive Grenzregeln O03 verbindlich festlegen, nicht aus Farben schätzen. Untere/obere Eingriffsgrenze unabhängig konfigurierbar, beide innerhalb des gültigen Toleranzfeldes; inkonsistente/fehlende Grenzen als nicht bewertbar kennzeichnen.

Bei nur oberem/unterem Grenzwert (Rundlauf/Rauheit u.a.) ist ohne fachlich bestätigte zweite Feldgrenze keine eindeutige Feldmitte ableitbar. **O03:** passende fachliche Untergrenze oder ausdrücklich eigener einseitiger Anzeigemodus; keine Null oder Fantasiegrenze still ergänzen. Nichtnumerisch: lesbares i.O./n.i.O., positiver Status ohne Scheinmesswert, negatives Ergebnis **vollständig roter Balken**. Fehlend/ungültig/noch offen neutral markiert und nicht als Grün dargestellt.

## 3. Qualitätscockpit

Eigener Bereich für AL/PL/QS, nicht Ersatz des Produktionsprotokolls oder der Startseite. Filter Firma (nur autorisierte eigene), BA/OP, Abteilung, Fallstatus, Aktualität; keine freizügige globale Firmenauswahl. Reihenfolge Rot → Gelb → Grün; innerhalb neueste fachliche Änderung zuerst. Markierungen „Entscheidung erforderlich“, „Korrektur offen“, „Nachprüfung offen“, „Kenntnisnahme offen“, „erledigt“ unterscheiden. Abgeschlossene Sonderfreigabe zeigt Originalrot und Entscheidung weiterhin, auch wenn unter erledigt gruppiert.

Roter Fall öffnet Teil/Messung/Zuordnung/Autorität und zulässige Entscheidung mit Pflichtgrund. Gelber Fall zeigt Maßnahme/Nachprüfung und getrennten Kenntnisnahmebutton, **keine Freigabe**. QS-Auftragsansicht zeigt Sonderfreigabe samt Entscheidername/Zeit/Grund ohne Push. Keine versteckte Änderung der Nachrichtenempfänger durch bloße Cockpitleserechte.

## 4. Druck/PDF (B/P)

Humbel standardmäßig **DIN A4 Querformat (297×210mm)**. Andere Firmen konfigurieren Hoch-/Querformat, Typ/Layoutversion wird am Export festgehalten. Eine breitformatige Bildschirmansicht allein ist kein Drucknachweis. Vollständiges Protokoll druckt nicht bloß die aktuell gefilterte Eingabemaske, sondern erforderliche Historie/Messungen/Fälligkeiten und Entscheidungen im freigegebenen Umfang. Wiederholte Kopf-/Spaltenzeilen, Seiten x/y, Firmen-/Artikel-/BA-/OP-/Teil-/Vorlagen-/Regelversion, Prüfername, Maßnahme/Entscheidung/Zeitzone; nichts abschneiden oder allein farblich kodieren.

PDF ist versioniertes Derivat, keine neue fachliche Freigabe. Entwurf, fehlende Letztstückprüfung und Sonderfreigabe sichtbar. Personen-/Kunden-/Firmeninhalte nur bei aktueller Exportauthority, kein ungeschützter öffentlicher URLpfad oder Token-QR. Druck-/Downloadrechte, Retention/Signaturniveau bleiben teilweise offen; keine externen Konvertierungen/Offlinecaches oder neue Löschfunktion durch diesen Auftrag.

## 5. Humbel-Prüfanweisung – nur Dreherei-Referenz

Quelle `74270.jpg`, sichtbarer Titel „Arbeitsanweisung / 2.4 Prüfanweisung“, Fußzeile als `2.4_HG_WI_Inspection Instruction_V02` lesbar. Laminierung, Reflexion, Perspektive und Schriftgröße begrenzen die Auswertung; kein autoritativer Text-/Versionsnachweis und kein generischer Firmenstandard. Nicht erfasste Randnotizen auf anderen Blättern sind nicht Teil der Analyse.

| Sichtbar erkennbare Aussage | Beweisgrenze / notwendige Bestätigung |
|---|---|
| Arbeitsbereich Drehen; Erststück und letztes Teil, Spalten100/50/10 % | Gilt nur Dreherei-Referenz; Details nicht pauschal auf Fräserei/andere Firmen übertragen |
| Maße (Durchmesser/Längen/Gewinde):100 % bei <0.03,50 % bei >0.03–<0.10,10 % bei >0.10; Bemerkung „100 % Gewindeprüfung“ | Gleichheit bei0.03/0.10, Einheit und Begriff Bezugs-Toleranzbreite offen. Ausnahmen/Kundenanforderungen vor Übernahme bestätigen |
| Form-/Lagetoleranzen: ähnliche Grenzspalten, taktile/manuelle Messung,3D-Messmaschine/Messuhr | Bemerkung „Wenn Messprg. vorhanden,1–3 Stk. Danach10 % … nur Hand Messbar … gemäss Toleranzen“ erkennbar, genaue Geltung/Anzahl/Override offen |
| Oberflächenrauheit: Erststück,10 %, letztes Teil; taktile Messung/Oberflächenrauheitsmessgerät | Keine automatische Regel aus Häkchen ohne Originalfreigabe |
| Verifizierung manueller Messung durch anderen Mitarbeiter, Verweis Erststückkontrolle V05 | Umfang/Person/Freigabe und tatsächliche Folgeversion offen; kein neu erfundenes allgemeines Vieraugen-Gate |
| Allgemein: erstes Teil Auftrag und1 Teil pro Schicht nach Erstprüfanweisung; letztes Teil unter zusätzlicher Restmengenbedingung | Formulierung mit „mehr als3 Teile … nach einem Austausch“ lesbar, Auslöser/Restmengeninterpretation nicht sicher. Erst-/Letztstück als Auftragspflicht bleibt fest, Ausnahme nicht still übernehmen |
| Bei gemessenem Teil nicht OK vorherige Teile sperren und erneut messen; nach Prozessänderung/Korrektur erneut messen | Rückwirkender Umfang/letzter sicherer Gutnachweis und Abgrenzung zum ausdrücklich beschlossenen nächsten Teil offen O02/O08 |
| Höhere Kundenanforderungen im Arbeitsplan; ansonsten Maße in Toleranzmitte; Messequipment zu Schichtbeginn prüfen/Kalibrierung bestätigen; visuell vor Spannen/nach Ausspannen | Verfahren/kalibriert vs. geprüft, Vorrangmatrix und Quellenfreigabe bestätigen. Nicht allein über Foto aktivieren |

Die hier beschlossene neue **Gelbregel mit nächstem Teil** gilt verbindlich. Die Fotoanweisung kann zusätzliche Prüfungen an korrigiertem/älterem Teil nahelegen, legt deren genaue Reichweite aber nicht belastbar fest. Diese Punkte bleiben O02/O08, ohne die Pflicht zum nächsten Teil aufzuheben. Bildgestaltungsbeispiel, interne Anweisung und fachliche Zielregeln werden als getrennte Quellen versioniert.
