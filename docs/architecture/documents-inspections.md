# Dokumentenverwaltung und digitale Prüfprotokolle – Fach- und Technikkonzept

> **Fortschreibung 14.11.2026:** Für dynamische Fertigungsprüfungen gelten die jüngeren [verbindlichen Qualitätsregeln](../requirements/digital-quality-2026-11-14.md) und die [Änderungs-/Entscheidungsmatrix](digital-quality-acceptance.md). Insbesondere Gelb ohne Vorgesetztenfreigabe, Rotentscheidung AL/berechtigter STV, Qualitäts-Gutteilgate, Nummerierung, Cockpit und Firmenisolation ersetzen abweichende Zielannahmen des Grundentwurfs aus #37. Bestehende Implementierung und historische Nachweise werden dadurch nicht geändert. Das bisherige JSON-Prüfprotokoll bleibt historisch auf #37 bezogen; aktuelle Ergebnisse siehe [neues Prüfprotokoll](digital-quality-validation.json).

**Reiner Konzeptentwurf, Version 0.17.2, Stand 13.11.2026. Keine Implementierung oder Produktivfreigabe.** Auftrag aus dem aktuellen Chat; Übergabebezug [Issue #33](https://github.com/butzi124-ship-it/schichtplan-mvp/issues/33). Eigener Branch `docs/documents-inspections-concept`, gestapelt auf PR #36, `c44824a8386f764ae966dd62541d76b75da1861a` (0.17.1). Bestehende PRs bleiben unverändert. Nur Dokumentation und vorgeschriebene APP_VERSION/VERSION_LOG-Metadaten ändern sich.

## Lesepfad und Verbindlichkeit

| Datei | Gegenstand |
|---|---|
| [Originalformular](documents-inspections-reference.md) | Visuell belegte Referenz, Rechenableitungen und Unsicherheiten |
| [Datenmodell und Storage](documents-inspections-model.md) | Vorgeschlagene Relationen, Versionen, Verknüpfung, Transaktionen und Dateien |
| [Rechte und Freigaben](documents-inspections-permissions.md) | Beschlossene QS-Regeln, Status-/Korrekturvertrag, Sicherheitsgrenzen |
| [Entscheidungen und Abnahme](documents-inspections-rollout.md) | Offene Fachfragen, Implementierungsphasen, Test-/Recoveryplan |
| [Prüfprotokoll](documents-inspections-validation.json) | Tatsächlich ausgeführte Bestandsprüfungen, Quellenstand und Blockaden |

**B** = ausdrücklich beschlossene Regel dieses Auftrags oder fortgeltendes 8A-Konzept; **I** = im Repository beobachtet, kein Produktionsinstallationsbeweis; **P** = Vorschlag zur späteren Freigabe; **O** = offene Entscheidung. Alle neuen Tabellen-, Bucket-, Aktions- und Endpunktnamen sind **P**, keine existierenden Originalobjekte. Der Fotoeinzelfall definiert keine universelle Prüfvorlage.

## 1. Tatsächlicher Bestand und Integrationsgrenzen

Am 13.11.2026 Remote-Refs und offene PRs über GitHub neu gelesen. PR #36/#35/#34 stehen offen als Draft auf `c44824a…` / `c5384669…` / `66795417…`. `main` steht auf `e764cf2e93382e847627d8bc7c802b02ff70c375`. 20 offene Draft-PRs: #15/#16, #18–#32 sowie #34–#36. Kein neuerer Dokumenten-/Prüfprotokoll-PR in dieser vollständigen offenen Liste. Issue #33 und dessen leere Kommentarliste berücksichtigt. Keine Actions neu gestartet; keine historische CI als aktueller Abnahmelauf übernommen.

| Quelle am Basiscommit | Beobachtetes Verhalten (I) | Anschluss des Konzepts (P) |
|---|---|---|
| `js/article-master.js`, Artikelmigration, [Artikel-Sicherheit](../articles/SECURITY.md) | Artikelidentität aus Abteilung + Nummer, Revisionen, `drawing_number`, `drawing_revision`, `inspection_notes`, Referenztexte, unabhängige Artikelscopes | Dateien an genaue Artikelrevision binden. Freitextreferenzen erhalten, nicht unkontrolliert als Download-URL ausführen oder automatisch importieren |
| `js/operations.js`, [Operationsvertrag](../articles/OPERATIONS.md), Operations-SQL | Prüfanweisungen und Zeichnungs-/Arbeitsanweisungstexte im OP-Plan; eingefrorene `production_order_article_snapshots` und `production_order_operations` | Dokumentversionsmanifest beim zukünftigen BA-Snapshot zusätzlich einfrieren, bestehende Snapshots nicht verändern. Zeichnungsindex ≠ Artikelrevision ≠ Dokumentversion ≠ Vorlagenversion |
| `js/ba-entry.js`, `js/op-routing.js`, `js/setup-times.js` | BA-Anlage, OP-Maschinenabschnitte, abteilungsübergreifende Verantwortung und persönliche Rüstung vorhanden | Dokument-/Prüfkontext daran verlinken, keine erneute BA-/OP-/Maschinen- oder Rüstimplementierung |
| `js/op-quantities.js`, `js/ba-lifecycle.js`, Quantity-/Lifecycle-SQL | Bestehende Gut-/Ausschuss-/Klärmengen, Nachfertigung, Abschluss- und Verantwortungskommandos | Bestehender Code bleibt unverändert; zukünftige qualitätsbedingte Gutteilprüfung/Abklärung→Ausschussumbuchung und Gelb-OP-Sperre jetzt fachlich nach F02/F05–F07 beschlossen. Keine automatische historische BA-Wiederöffnung |
| `schichtplan_mvp.js`: `normalizeProductionChecklistTemplateFromDb`, `normalizeProductionOrderChecklistFromDb`, `renderProductionOrderChecklistSection`, `renderProductionProtocolTab`, `openProductionProtocolForOrder` | Abschluss-Checklisten mit checked/Person/Zeit und Produktionsprotokoll existieren. Sie sind kein versioniertes dimensionales Messprotokoll | Bestehende Checklisten/Produktionsereignisse unverändert lassen, später verknüpfte QS-Nachweise ergänzen. Keine Gleichsetzung eines Checkhäkchens mit QS-Freigabe |
| `production_qa_causes`, `production_station_events.qa_cause_id` und `renderProductionProtocolEntry` | Bestehender QS-Ursachenkatalog und Qualitätsereignisse werden mit Menge, Mitarbeiter und Notiz im Produktionsprotokoll angezeigt | Ursachen-/Mengen-/Ereignishistorie erhalten; neue dimensionsbezogene Prüfung separat verlinken, nicht als Ersatzursache oder zweiter Ausschusszähler implementieren |
| Suche in aktuellen JS-/SQL-Pfaden | Kein allgemeiner geschützter Dokumenten-Upload-/Vorlagen-/Messwertdienst in diesen geprüften Modulen gefunden | Neues begrenztes Modul; kein Beweis, dass im Produktivprojekt kein Storage existiert. Original-Buckets/Policies/Dateien hier **UNVERIFIED** |
| [8A](modular-permissions-8a.md), [8B1](../security/permissions-8b1.md), [8B2](../security/permissions-8b2.md) | Modulares Rollenmodell konzipiert; Identitäts-/Sitzungskandidaten privat, isoliert, deaktiviert. Kein echter freigegebener JWT-/Auth-Adapter | QS als Modulprofil mit expliziten Aktionen, keine achte globale Organisationsrangstufe und kein Alias `admin`. Keine Freischaltung durch dieses Konzept |

Der [Originalbericht #32](../qa/security-step7-2026-11-12.md) enthält den datierten Katalog vom 12.11.2026; vollständiger Restore und tatsächliche spätere Installationen bleiben offen. Neue BA-/Artikel-/Authoritytabellen aus Review-SQL sind nicht als vorhandenes Produktionsschema zu bezeichnen. Die historischen Sicherheitsberichte #25/#28–#32 bleiben gültig; ihre offenen Rechte-/Cache-/Owner-/Recoveryrisiken werden hier nicht behoben.

## 2. Fachlicher Dokumentumfang

Zentrale Dokumentübersicht mit denselben Nachweisen aus Artikel, BA, OP und optional Maschine. Eine Datei nicht für jeden Einstieg kopieren. Dokumenttypen zunächst Prüfprotokoll, Zeichnung, Rüstunterlage, Arbeitsanweisung, Messbericht; weitere Typen kontrolliert über einen versionierten Typenkatalog. Typkonfiguration darf weder beliebiges HTML/SQL noch neue Rollen-/Freigaberechte erzeugen.

Je Dokument: stabile ID, Typ, Titel, fachlich verantwortlicher Bereich, Klassifikation, Sprache, Quelle, exakte Version, Status, Verknüpfungen, Urheber/Uploader, Serverzeit und Änderungshistorie. Revisionsgrund bei Nachfolgeständen. Aktuelle freigegebene Version und historisch auftragsgültige Version unterscheidbar anzeigen. Suche nach zugelassenen Metadaten (Artikel, BA, OP, Teil-/Seriennummer, Typ, Status); Suchtreffer, Zähler und Vorschaubilder unterliegen denselben Leserechten wie Inhalte. OCR/Volltext nur spätere Option, kein externer KI-/OCR-Dienst ohne separate Freigabe.

PDF und Bilder sind vorgesehene Eingänge. Office-Dateien: kontrollierte Allowlist z.B. DOCX/XLSX nach Format-/Größen-/Inhaltsprüfung; keine automatische Makroausführung. Legacy-/makrohaltige oder verschlüsselte Dateien zunächst Quarantäne bzw. bewusste Ablehnung nach D12. Office-/PDF-Vorschau als isoliert erzeugtes Derivat; Original unverändert behalten. Weder Dateiendung noch Browser-MIME ist ein ausreichender Sicherheitsnachweis.

## 3. Flexible Prüfvorlagen

Eine Vorlage besitzt eine unveränderliche freigegebene Version. Ihre Merkmale können numerisch, Auswahlwert, Ja/Nein, Textbeobachtung oder dokumentenbasierter Nachweis sein. Je Merkmal: stabile ID, Nummer/Bezeichnung, Bauteil/Prüfposition, Verfahren, Messmittelanforderung, Einheit, Sollwert, numerische Unter-/Obergrenze, Vergleichsregel, Auflösung, Auswahlwerte, getrennte interne Warn-/Freigabezonen, Pflicht-/N/A-Regel und Stichprobenplan. Keine frei programmierbaren Skripte/Formeln; spätere Berechnungen aus geprüftem begrenztem Regelkatalog.

Sollangabe und geprüfte numerische Grenzdefinition gemeinsam versionieren. Dezimalwerte exakt speichern/rechnen, keine binäre Float-Rundung als Toleranzentscheidung. Einheit/Dezimaltrennzeichen validieren; Rohwert und normalisierte Dezimaldarstellung erhalten. Anzeige nicht vor Bewertung runden. Passungsangabe wie N6 bleibt Quelleninformation; ihre numerischen Grenzen nicht aus einem ungeprüften Normkatalog erraten. Falsche/mangelhafte Definition blockiert Vorlagenfreigabe.

Ein Prüfvorgang bindet genaue Vorlage, Artikel-/Zeichnungsstand, BA, OP sofern zutreffend, Teil-/Gehäuse-/Losidentität, Prüfschritt und optional damaligen Maschinenabschnitt. Wiederholungsmessung ersetzt keine erste Messung. Wiederkehrende Prüfungen (Erstteil, Stichprobe, Endprüfung, Nachprüfung) brauchen eigene Sequenz/Anlass. Unbekannte Daten sichtbar als unbekannt; keine erfundenen Seriennummern, Sollwerte oder Mitarbeiter.

Messmittel: ID oder begründete externe Referenz, Verfahren und dokumentierter Kalibrier-/Eignungsnachweis zum Messzeitpunkt; keine neue Werkzeugverwaltung daraus entwickeln. Zukünftige Messmittelregister-Integration separat entscheiden. Fehlender/abgelaufener Nachweis erzeugt Klärbedarf nach D06; keine unbemerkte Bestätigung als gültig.

## 4. Zwei gleichberechtigte Erfassungsvarianten – keine Auswahl beschlossen

| Kriterium | A: direkte digitale Werte | B: bestehendes Dokument mit Scan/Foto |
|---|---|---|
| Bedienung | Touch-Auswahl/Zifferneingabe, sofortige Vollständigkeits-/Wertehinweise; Handschuh-/Messplatzpilot nötig | Gewohntes Papier, anschließend Zuordnung und Qualitätskontrolle des Uploads; doppelte Arbeit möglich |
| Nachvollziehbarkeit | Actor/Serverzeit/Version je Buchung, exakte Korrekturketten | Originalseiten und Dateihash; Handschrift/Datum/Visum nicht automatisch verifizierte Identität. Uploader ist nicht zwingend Prüfer |
| Aufwand | Vorlagenpflege und Grenzvalidierung initial höher; wiederkehrende Eingabe leichter | Anfangs geringere Strukturierung, Scan-/Lesbarkeits-/Vollständigkeitsprüfung laufend nötig |
| Sicherheit | API-Manipulation, fremde Ziel-IDs, Grenz-/Vorlagenänderung verhindern | Dateischadcode, unleserliche Seiten, falsche BA-/Teilzuordnung, manipulierte Unterschrift/EXIF berücksichtigen |
| Auswertung | Strukturierte Merkmals-/Zeitreihen nach geprüfter Vergleichbarkeit | Nur Metadaten zuverlässig; OCR/Volltext ist kein freigegebener Messwert. Verifizierte Transkription separat |
| Offline/Fehler | Ohne freigegebenen Offlinevertrag nur Arbeitsspeicher; keine behauptete Speicherung bei Netzfehler | Papier kann weiter genutzt werden, digital erst nach bestätigtem Upload/Zuordnung; Originalaufbewahrung entscheiden |
| Druck/Archiv | Reproduzierbarer Bericht mit Bewertungs-/Vorlagenstand | Original + ggf. Druckderivat, Scanqualität und Seitenfolge sichtbar |

Beide können denselben Prüfcontainer und Freigabevertrag benutzen. Optionaler Hybrid ist P: Scan als Primärbeleg, bestätigte Transkription daneben mit Quellenbezug. Niemals unterschiedliche widersprüchliche Wahrheiten still zusammenführen. Datenherkunft A/B/Transkription sichtbar; Dokumentvollständigkeit allein ist kein i.O.-Messnachweis. Variante erst nach Pilot und D01 entscheiden.

## 5. Vorschlag für die spätere Bedienung

Keine UI wird jetzt gebaut. Für die Qualitätsprüfung ist das [freigegebene Messprotokoll-Layout](digital-quality-ui.md) einschließlich Humbel-A4-Querformat verbindlich; bestehende Navigation bleibt erhalten. Geplante Einstiegspunkte: eigener Bereich Dokumente und QS-Arbeitsvorrat sowie kontextuelle Register an Artikel/BA/OP. Bestehende Navigation/Barcodeeingänge nicht ersetzen oder erweitern, ohne deren geprüften Verträge zu erhalten. Scan öffnet nur einen autorisierten Kontext und führt keine Buchung/Freigabe aus.

Desktop: gefilterte Liste, Dokument-/Prüfdetails und geschützte Vorschau nebeneinander. Tablet: große Merkmalskarten, Soll/Zeichnungstoleranz/interner Freigabebereich getrennt. Smartphone: ein Merkmal je Schritt, persistenter Artikel-/BA-/Teilkontext, große Bedienflächen, Dezimaltastatur plus bedienbare Alternativen, eindeutige Labels, Fokusführung und Status zusätzlich als Text/Icon. Keine alleinige Rot-Grün-Kodierung. Handschriftvorlage nicht als winzige Fotoeingabe nachbauen.

Erfassung: exakten Kontext bestätigen → Variante/Version wählen → messen oder Original vollständig hochladen → Prüfer/Datum/Seiten prüfen → bewusst einreichen → spätere zuständige Prüfung/Freigabe. Produktionsansicht ausschließlich lesend, ohne Kommentar-/Änderungsaktionen an QS-Ergebnissen. Fehler lassen ungesendete Eingaben im aktuellen Arbeitsspeicher, kennzeichnen sie als ungesichert und behaupten keinen Commit. Logout/Rechtewechsel bereinigt sensible Daten; kein dauerhafter Offlineentwurf ohne eigene Freigabe.

PDF-/Druckausgabe: fachliche und technische Versionskennungen, Teil-/BA-/OP-Bezug, Prüfername nach B, Messwerte und Grenzdefinition, Abweichungen, Entscheidung/Korrekturverweise, Serverzeit/Zeitzone, Seiten x/y, Erzeugungsstand. Entwurf/ungeprüfter Scan deutlich kennzeichnen. PDF ist ein versioniertes Derivat; Export ist nicht automatisch Freigabe. Druck/Dateikopie können nachträglich nicht zurückgerufen werden, QR nur geschützte Referenz ohne Token/Datei-URL. Druckformat, Downloadrechte und mögliche Archivformate nach D10/D11.

## 6. Abnahmescope dieses Auftrags

Ergebnis sind zusammenhängende Fachspezifikation, Fotoanalyse, Modell, Matrix, Workflow, offene Entscheidungen und Phasen-/Testplan. Kein SQL, Storagebucket, RPC, Interface, neuer Grant, Featureflag oder Löschjob wird angelegt. Version/Log werden ausschließlich nach AGENTS.md erhöht. Tatsächliche bestehende Prüfungen und nicht ausgeführte zukünftige Abnahmen getrennt im Prüfprotokoll; keine Produktivfreigabe.
