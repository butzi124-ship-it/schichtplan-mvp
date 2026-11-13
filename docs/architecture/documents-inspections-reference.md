# Analyse der Referenz 52978 – kein freigegebener digitaler Prüfplan

Teil des [Gesamtkonzepts](documents-inspections.md). Quelle: im Chat referenziertes Foto `73930.jpg`, visuell am 13.11.2026 vollständig und im Messbereich vergrößert betrachtet. Bild 1157×1536, SHA-256 `15c9129e4ff61784212a082660fbb0309ad46ff8b1ebb80e2778d21919dd3780`. Das Original enthält betriebliche Angaben und Mitarbeiterkennzeichen; **weder Bild noch Ausschnitt oder persönliche Kennzeichen werden ins Git-Repository kopiert**. Nachstehende formularbezogene Angaben sind Referenzbeobachtungen, kein Produktionsdatensatz oder verifizierter Messauftrag in der App.

## 1. Sicher sichtbare Struktur

| Feld | Sicher lesbare Information | Grenze |
|---|---|---|
| Titel | `52978_Prüfprotokoll`, HUMBEL / Gear Technology | Dokumenttyp, kein allgemeines Vorlagenformat |
| Artikelnr. Humbel | **52978** | Nicht mit Kundennummer verwechseln |
| Artikelnr. Stadler | **12442231** | Eigener externer Nummernraum; keine Rechte-/Primärschlüsselquelle |
| Bezeichnung | **Gehäuse komplett T&W** | Schreibweise des fotografierten Formulars |
| Zeichnung / Index | **BU_4148038**, **f** | Index f ist keine App-Dokumentversion oder Artikelrevision |
| Betriebsauftrag / Gehäusenummer | Gelbe Felder, handschriftlich als **812481** / **581** lesbar | Fotoablesung, nicht mit Original-BA/Teilbestand abgeglichen; nicht als Test-/Importdaten übernehmen |
| Messmerkmale | 16 Gehäuse, 17 Deckel, 19 Gehäuse, 20 Deckel, 22 Gehäuse, 24 Deckel | Mess-Nr. sind keine OP-Nummern |
| Erfassende Person | Zeile Mitarbeiter; Bereich Bearbeiter (Mikrometer) | Nicht automatisch identisch mit späterem Uploader, QS-Prüfer oder Freigeber |
| Zusätzlicher Hinweis | `Mit Leuchtstift Messmaschinenresultat (20%)` | 20 % ist sichtbar, aber Bezugsgröße, Stichprobenregel und Messmaschinenverfahren sind **ungeklärt** |
| Nachweis rechts | Datum / Visum / Pers.-Nr.; Datum handschriftlich als **9.10.26** lesbar, separate Zeile Vorgesetzter | Datum ist Fotoablesung, kein serverseitiger Zeitnachweis. Keine sichere digitale Identitäts-/Signaturzuordnung aus Handschrift. Vorgesetztenzeile wirkt unausgefüllt; daraus keine historische Freigabe ableiten |
| Warnhinweis | Bei Messergebnissen außerhalb des grünen Bereichs muss der Vorgesetzte die Freigabe erteilen | Sichtbare Papierregel; digital berechtigter Freigeber und Ausnahmeverfahren bleiben D04 |
| Korrekturhinweis | Bei korrigierten Angaben soll der korrekte Wert umkreist werden | Frühere/finale Messwerte nicht sicher aus Strichen ableiten; digital additive Korrektur statt Überschreiben |

Visum und Personalnummer werden nicht als Benutzerzuordnung transkribiert. Das abgelesene Datum wird nicht als verifiziertes Erfassungsdatum importiert. Die Referenz ist kein Authentifizierungsbeleg.

## 2. Sichtbare Sollmaße und auswählbare Werte

Nach Vergrößerung sind bei beiden Ø110-Spalten **−0.016 / −0.038** lesbar. Automatische Texterkennung ist bei Vorzeichen und Nummern unzuverlässig und wird nicht als Freigabequelle verwendet. Die Einheit ist im betrachteten Tabellenkopf nicht ausdrücklich ausgeschrieben: **mm ist eine plausible, vor Vorlagenfreigabe durch QS zu bestätigende Annahme**. Rechenwerte gelten nur unter dieser Annahme, sind keine eigenständig geprüften Zeichnungstoleranzen.

| Mess-Nr. / Bauteil | Sichtbarer Solltext (Dezimalpunkt wie Foto) | Rechnerisches Intervall, sofern mm / Grenzen bestätigt | Gedruckte Auswahl, absteigend |
|---|---|---|---|
| 16 / Gehäuse | Ø230 N6 (−0.022 / −0.051) | 229.949 … 229.978 | 229.980 … 229.948, Schritt 0.002; 17 Werte |
| 17 / Deckel | Ø230 N6 (−0.022 / −0.051) | 229.949 … 229.978 | Wie Mess-Nr. 16 |
| 19 / Gehäuse | Ø110 N6 (−0.016 / −0.038) | 109.962 … 109.984 | 109.986 … 109.960, Schritt 0.002; 14 Werte |
| 20 / Deckel | Ø110 N6 (−0.016 / −0.038) | 109.962 … 109.984 | Wie Mess-Nr. 19 |
| 22 / Gehäuse | Ø100 N6 (−0.016 / −0.038) | 99.962 … 99.984 | 99.986 … 99.960, Schritt 0.002; 14 Werte |
| 24 / Deckel | Ø134.5 H7 (0.04 / 0) | 134.500 … 134.540 | 134.542 … 134.497, Schritt 0.003; 16 Werte |

Die Auswahl enthält auch Werte außerhalb der rechnerischen Toleranz. Das muss erfassbar bleiben; keine automatische Rundung/Klemmung auf den nächsten erlaubten Wert. Bei freier Messung zwischen Auswahlstufen ist D02 zu entscheiden: exakter Wert plus Quellenbezug statt erzwungener Falschangabe. Auswahl ist kein kontinuierliches grünes Intervall und kein Ersatz für die vollständige numerische Bewertung.

## 3. Farben und handschriftliche Markierungen

Sichtbar sind grüne, orange/braune sowie rote bzw. rot hervorgehobene Auswahlwerte. Z.B. sind 229.980/229.948 und 134.542/134.497 rot hervorgehoben; innen liegende Werte können orange sein. Damit ist **grün ≠ vollständiger Zeichnungstoleranzbereich**. Die grüne Auswahl wirkt bei 16/17 etwa 229.970–229.958, bei 19/20 109.982–109.972, bei 22 99.978–99.970, bei 24 134.533–134.512. Diese visuellen Farbablesungen sind **nur Reviewhinweise**, keine freigegebenen numerischen Grün-/Orange-Grenzen. Licht, Druck und Auswahlstufen erlauben keine sichere allgemeine Intervall-/Randwertregel.

Markierungen stehen in mehreren Spalten, einzelne wirken mehrfach umkreist/gestrichen. Kein abschließender Istwert, keine Maschinenmessung oder erfolgreiche Freigabe wird daraus behauptet. Eine leere Unterschriftszeile oder sichtbares Datum beweist weder Ablehnung noch Freigabe.

## 4. Konsequenzen und noch benötigte QS-Bestätigung

1. Original-PDF/Zeichnung und genehmigte Formularrevision für Einheit, Vorzeichen, Messpositionen, Grenzinklusion und Farb-/Warnzonen; Farbbedeutung Orange/Rot und genaue Freigabeobergrenzen bestätigen (D02/D03).
2. Mess-Nr. zu Prüfverfahren/Messmittel, Anzahl Teile/Messpunkte und Bedeutung der 20-%-Angabe klären (D06). Messmaschinenwert und Mikrometerwert als getrennte Belege ermöglichen.
3. Zulässige Prüfer, Einreicher und Vorgesetztenfunktion, Vieraugenprinzip sowie Ausnahmefall außerhalb Grün gegenüber außerhalb Zeichnungstoleranz getrennt entscheiden (D04/D05).
4. Nachweis von ursprünglichem und korrigiertem Wert, Pflichtgrund/Neumessung und Freigabe einer korrigierten Version definieren (D07). Papier-Umkreisung nicht als digitales Löschrecht übersetzen.
5. Bezeichnung Gehäusenummer betrifft hier die Baugruppe mit Gehäuse-/Deckelmerkmalen. Eindeutigkeit pro Artikel/BA/Los und Einzelteil-/Baugruppenbeziehungen D08; keine globale Seriennummernregel aus einem Foto ableiten.

Alle diese Anforderungen sind in den [Entscheidungs- und Testplan](documents-inspections-rollout.md) übertragen. Das Foto wird nicht als allgemein vollständiger oder freigegebener Messplan ausgegeben.
