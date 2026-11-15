# Fachentscheidungen, Integrationsphasen und Sicherheitsabnahme

> **Fortschreibung 14.11.2026:** Für dynamische Fertigungsprüfungen gelten die jüngeren [verbindlichen Qualitätsregeln](../requirements/digital-quality-2026-11-14.md) und die [Änderungs-/Entscheidungsmatrix](digital-quality-acceptance.md). Insbesondere Gelb ohne Vorgesetztenfreigabe, Rotentscheidung AL/berechtigter STV, Qualitäts-Gutteilgate, Nummerierung, Cockpit und Firmenisolation ersetzen abweichende Zielannahmen des Grundentwurfs aus #37. Bestehende Implementierung und historische Nachweise werden dadurch nicht geändert. Das bisherige JSON-Prüfprotokoll bleibt historisch auf #37 bezogen; aktuelle Ergebnisse siehe [neues Prüfprotokoll](digital-quality-validation.json).

Teil des [Konzepts](documents-inspections.md). Neue Modulabnahmen sind **geplant, nicht implementiert oder bestanden**. Die ursprünglichen #37-Dokumentations-/Bestandsprüfungen stehen im [historischen Prüfprotokoll](documents-inspections-validation.json); der neue Nachweis ist oben verlinkt. Kein Produktivzugriff oder laufender Geschäftsprozess wird verändert.

## 1. Offene Entscheidungen mit sicherer Zwischenposition

Q01–Q06 in der [Matrix](documents-inspections-permissions.md) sind beschlossen und werden nicht erneut geöffnet. Alle anderen 8A-/8B-Fach-/Securityblocker gelten fort. Zuständigkeiten nachfolgend sind vorgeschlagene Entscheidungsbeteiligte, keine neue Ernennung/Vollmacht.

| ID | Konkrete Entscheidung / benötigter Nachweis | Beteiligte / bis dahin |
|---|---|---|
| D01 | Variante A, B oder Hybrid pro Dokument-/Prüfart; Pilot, tatsächliche Papieroriginale und wer Scans hochladen darf | QS/Produktion; keine Variante verbindlich auswählen, kein produktiver Upload |
| D02 | Originalformular-/Zeichnungsstand, Einheiten, numerische Grenzinklusion, Auflösung/Rundung und zulässige Zwischenwerte | QS/Zeichnungsverantwortung; Foto nicht ungeprüft in eine aktive Vorlage umsetzen |
| D03 | Farb-/Balkensemantik nach F04 beschlossen, konkrete reale Grenzwerte/inklusive Grenzen/einseitige Felder O03 | QS; keine aktiven Grenzwerte aus Bildbeispielen übernehmen |
| D04 | **Teilweise entschieden F05/F06:** Gelb Korrektur/Dokumentation/Nachprüfung, AL nur Kenntnisnahme; Rot zuständiger AL/berechtigter STV entscheidet | Zulässigkeits-/Vieraugen-/Sonderfreigabedetails O05 bleiben offen; keine Gelb-Freigabe mehr vorsehen |
| D05 | Vorlagenautoren/-freigeber, QS-Grantissuer/Empfänger, belastbare gültige QS-Zuordnung auch bei Mehrfachabteilung; wer pflegt diese Zuordnung? | QS/Berechtigungsverantwortung; Q05 bleibt fest: AL QS nur QS-zugeordnete Personen, keine automatische Fähigkeit aus Mitgliedschaft |
| D06 | Bedeutung „Messmaschinenresultat (20 %)“, Prüftakt, Los-/Stichprobenplan, Messmittel-/Kalibriernachweis und ungültige Messmittel | QS/Messtechnik; kein ausgedachter 20-%-Algorithmus oder Messmittelregister |
| D07 | Eigene/fremde Korrektur, Transkription, Wiederholungsmessung, Bestätigung, Betroffenheit und historische Freigaben | QS; keine allgemeine Edit-/Kommentar-API, Produktionsverbot Q04 erhalten |
| D08 | **Teilweise entschieden F08/F09:** OP-lokale Nummern vs. vorhandene Identität, spätere Gravur, keine unbelegte Rückverfolgung | Namespace/Mehrfachspannungen/Legacybindungen nach O04/O07 |
| D09 | **Teilweise entschieden F10:** Cockpit PL/QS alle eigenen Produktionsabteilungen, AL alle eigenen, MA arbeitsnotwendig | Feld-/Entwurfs-/History-/Datenschutzdetails erhalten; QS-Prüfername weiterhin Q03 |
| D10 | Druck/Download/Massenexport, Dateikopie, Sign-URL-Restlaufzeit, Weitergabe/Offline, Geräte-/Scanqualität, maximale Warte-/Uploadzeiten | QS/Betreiber/Security; private Vermittlung vorgeschlagen, kein öffentlicher Bucket, kein sensibler Offlinecache |
| D11 | Aufbewahrungs-/Hold-/Löschmatrix je Dokumentklasse, Eigentum/Freigeber, betriebliche/vertragliche/rechtliche Pflichten, Archivformat, Originalpapier, RPO/RTO und Storage-/DB-Recovery | QS/Datenschutz/Betreiber; keine Frist aus Vertretungslogs übertragen, keine Löschung/Automatik |
| D12 | Zulässige Office-/Bild-/PDF-Formate, Größen/Seiten/Quoten, Makro-/verschlüsselte Dateien, Renderer/AV/OCR-Dienst und Datenstandort | Betreiber/Security/QS; Allowlist/Quarantäne, keine externen Konvertierungen oder aktive Inhalte ohne Freigabe |
| D13 | **Teilweise entschieden F02/F05–F07/F11:** Gutteilgate, Gelb-OP-Sperre, Umbuchung und Benachrichtigung fachlich beschlossen | Spannungs-/Letztstück-/historische/Folge-OP-/Versanddetails O02/O04/O08, keine Implementierung |
| D14 | Audit-/Prüferidentitäts-/Signaturniveau, dokumentierte Vertretung, historische Namen, revisionssichere externe Archivanforderung | QS/Betreiber/Datenschutz; Scanvisum ist kein digital geprüfter Signaturnachweis, DB-Append-only keine DBA-WORM-Garantie |

Zusätzlich vor Aktivierung: 8A O01/O09/O11/O12 (unabhängige Identitäten/Enrollment/administrative Befugnisse/Konflikte), O06/O08 (modulare Aktions-/Bereichsverträge) und O16 (Betrieb/Originalrestore) sowie offene 8B2-Verifier-/Owner-/native Fragen abschließen. Ernennungs-, HR-/TA- und Sonderrollenregeln aus 8A unverändert. Die neuen Q-Regeln ergänzen das Modulkonzept, verändern keine historischen Prüfberichte.

## 2. Kleine, separat testbare Phasen (P)

| Phase | Ergebnis und Abgrenzung | Voraussetzungen / Exit |
|---|---|---|
| K0 – dieser Auftrag | Fotoanalyse, Variantevergleich, Modell/Matrix/Tests, keine Implementierung | Dokumentkonsistenz/Bestand/Version geprüft, Draft-PR |
| K1 – Verträge und Testfixtures | QS bestätigt Formular-/Grenz-/Identitäts-/Rollenverträge, synthetische Artikel-/BA-/OP-/Dateifixtures; keine Aktivierung | D02–D09 soweit betroffen und Auth-Trustvertrag, positive/negative Beispiele unabhängig reviewt |
| K2 – isoliertes Dokumentfundament | Private Metadaten/Version-/Linkmodell und Quarantäne-Upload mit Scan/Prüfung/Finalize, keine Prüfentscheidung | D10–D12, sichere 8B-Identität/Sessionbrücke; native Storage/RLS/Upload/Fehler-/Recoverytests bestehen |
| K3 – Dokumentleser und Pilotablage | Versionierte Originale/Derivate, Artikel-/BA-/OP-/Maschinenlinks, datensparsame QS-/Produktionsprojektionen | K2, bestätigte Rollen-/Downloadverträge, Browsermatrix und Widerruf ohne Lecks; noch kein QS-Abschlussgate |
| K4 – Vorlagen und Erfassung | Zuerst eine bestätigte Vorlage, dann unterschiedliche Merkmaltypen/Verfahren; gewählte Variante oder ausdrücklich gleichberechtigte Pilotvarianten | D01/D02/D03/D06/D08; Grenzen/Sampling/Scanvollständigkeit nachweisen, keine automatische Ausnahmefreigabe |
| K5 – Entscheidungen und Korrekturen | Nur beschlossene normale/abweichende Freigaben, additive Korrekturhistorie und Vieraugenbedingungen | D04/D07/D14, atomare Audit-/Version-/Widerrufsraces; kein umgehbarer Altpfad |
| K6 – gezielte spätere Fachintegration | Jetzt nach F02/F05–F07/F11 beschlossene QS-Gates/-Umbuchung/-Benachrichtigungen später separat technisch in bestehende Fachkommandos integrieren | D13; vollständige positive BA-/OP-/Rüst-/Mengen-/Scannerregression. Kein neuer Zähler/Produktionsprozess als Nebeneffekt |
| K7 – Betriebsabnahme | Vollständiger isolierter Originalrestore plus Storagebytes, Last/Geräte/Druck/Recovery | D11, native/API-/Sicherheits-/Browser-/Restorebelege, gemessene RPO/RTO und Freigabeplan |
| Separat – Produktiveinsatz | Reviewtes konkretes Migrations-/Pilotpaket, Wartungs-/Stopplan | Eigene ausdrückliche Migration-/Merge-/Deployment-/Aktivierungsfreigaben; keine automatische Aktivierung aus diesem Konzept |

K1 ist der erste sicher abgrenzbare nächste Auftrag. K2 kann erst mit nachgewiesenem sicheren Auth-/Sessionpfad abgeschlossen werden; 8B2 ist keine fertige Browserautorität. Fachmodule bleiben bis gezielt abgenommenen Adaptern unverändert. Kein paralleles ungesichertes Dokumentensystem als Abkürzung.

## 3. Geplanter Sicherheits- und Regressionstestplan

**Alle folgenden N-Tests: BLOCKED_NOT_TESTED – neue Funktionen existieren in diesem Dokumentationsauftrag nicht.** Nicht als aktuelle Testergebnisse zählen. Native Prüfungen ausschließlich in ausdrücklich zugelassener isolierter PostgreSQL-/GoTrue-/PostgREST-/Storageumgebung mit künstlichen Daten. PGlite eignet sich ergänzend für SQL-Verträge, ersetzt keine JWT-/Storage-/Mehrverbindungs-/Originalrestoreabnahme.

| Gruppe | Konkrete Positiv-/Negativfälle und überprüfbares Ergebnis |
|---|---|
| N01 Identität/Session | Anon, widersprüchliche Auth-/Personbindung, inaktiv/gesperrt, abgelaufen/widerrufen, manipulierte Claims/Prüfer-ID, alter Epoch; jede neue Aktion/Download verweigert, aktive bestätigte Person erlaubt |
| N02 QS-Quersicht | QS liest BA/OP in A/B/C über zugelassene Produktion; kein Umplanen, Rüsten, Buchen, Mitarbeiter-/Medizinlesen oder versteckte SELECT*-Felder. Maschinenübergabe erhält Kontext ohne neue Steuerungsrechte |
| N03 Produktionsleser | Erlaubtes Ergebnis + Prüfername sichtbar; POST/PATCH/DELETE/Kommentare/Anhänge/Korrekturumgehung an QS-Ergebnissen denied. Entwurf/Fremdscope/Suche/History entsprechend D09 |
| N04 QS-Vergabe | AL QS nur aktive belegte QS-Zuordnung, eigene Rechte/Scope; Mitgliedschaft allein keine Rechte. Nicht-QS/Zukünftige/abgelaufene Zuordnung, Selbstfreigabe, fremder AL, unfreigegebener STV, manipulierte Mehrfachbatch denied; kein Teilerfolg |
| N05 Dokumentrechte | Liste/Details/Original/Vorschau/Derivat/Range/Export/URL/S3/Copy/Move/Realtime/GraphQL identisch geschützt; UUID-/BA-/OP-Tausch, Link an Fremdbereich und Metadatenleck über Zähler denied |
| N06 Vorlagen/Grenzen | Genau freigegebene Version; verschiedene Merkmalsarten, exakte Dezimalwerte/Einheiten/Komma, Werte exakt auf/unter/über Grenze, Null/N/A/fehlend, Auswahlzwischenwert, sich überlappende Zonen. Keine grüne Freigabe aus Zeichnungstoleranz allein |
| N07 Foto-/Papierpfad | Ganze Seiten/Reihenfolge/Lesbarkeit, falscher BA/Teil, abweichende Handzeichen, Scan-Uploader ≠ Prüfer, OCR-Zahl falsch und unbestätigt; Datei allein erzeugt kein bestanden. Original und korrigierte Transkription nachvollziehbar |
| N08 Datei-Angriffe | MIME-/Endungs-/Magic-Byte-Konflikt, aktives PDF/SVG/HTML, Office-Makro/Netzreferenz, Zip-/Bildbombe, Passwortdatei, oversized, Path Traversal, Doppelfile, EXIF-Leak; Quarantäne/Ablehnung ohne Serverausführung, öffentlicher Zugriff denied |
| N09 Transaktionen | Einreichen gegen Vorlagenrückzug, zwei Freigaben, Korrektur gegen Entscheidung, Session-/Grantwiderruf während Lockwartezeit, Versionkonflikt und Requestreplay. Echte getrennte PG-Verbindungen; Audit-/Outboxfehler rollt Fachänderung zurück |
| N10 DB/Storage-Fehler | Abbruch vor/nach Upload, Timeout, doppelte Finalisierung, Hash-/Längenwechsel, AV-/Konvertierungsausfall, DB-Commitfehler und Objectverlust: kein sichtbares ungeprüftes Asset, alte Version bleibt, keine verwaisten freigegebenen Links. Recovery ohne heimliches Überschreiben |
| N11 Unveränderlichkeit | Direkter DML-/TRUNCATE-/RPC-/Owner-/EXECUTE-/Default-/Namensspoofingtest; Original/Messung/Entscheidung/Audit nicht durch Appakteure überschreibbar, berechtigte additive Korrektur vollständig; DBA-Grenze separat |
| N12 Oberfläche/Datenschutz | Desktop/Tablet/Smartphone, Touch/Handschuhpilot, Screenreader/Fokus, Farbe plus Text, Decimalinput, Netzverlust, zwei Tabs, Logout/Loginwechsel, verspätete Antworten und deaktivierte Accounts; kein alter Inhalt/Dateicache im neuen Kontext |
| N13 Druck/Export | Seiten/Versionen/Quellen/Prüfername/Entwurfskennzeichnung, unterschiedliche Papierformate, keine abgeschnittenen Werte, keine Token-QRs, erneute Rechteprüfung vor Jobauslieferung, keine Log-/CSV-Formel-/Metadatenlecks |
| N14 Restore/Retention | DB+Bytes+ACLs+Manifest unabhängig isoliert wiederherstellen, Hash-/FK-/Korrekturketten prüfen, fehlende Datei erkannt, Widerrufe erhalten. Holds/Fristen/Originale schützen; Löschtests erst separat genehmigt, keine automatisch unsichere Policywiederherstellung |
| N15 Bestehende Fachmodule | Sämtliche bisherigen Tests/251 Legacy-Chromium-Identitäten und Visualreferenzen erhalten; Artikelrevisionen, unveränderliche BA/OP, Maschinenwechsel, eigene Rüstung/Mengen/Nachfertigung/Klärfälle/Abschlüsse, Abschlusscheckliste/Produktionsprotokoll, Werkzeuge/Journal/QR und Schicht-/Personalrechte unverändert |

Alle späteren Testprozesse erhalten feste Prozess-/Gesamttimeouts und Vorabchecks; TERM/KILL kontrolliert, keine Endlosschleife/automatische Snapshotkorrektur. Failures bleiben FAIL, fehlende Laufzeit wird BLOCKED_NOT_TESTED. Kein Skip/Filter/Toleranz-/Assertionstrick zur Grünfärbung. Ergebnisse getrennt nach unverändertem Bestand, neuem isoliertem Kandidaten, integrierter App und Originalrestore führen.

## 4. Migrations-/Freigabecheckliste für später

1. Originalschema/Storagekonfiguration/API-/Default-/Owner-/Bucket-/Funktionsabhängigkeiten sicher und vollständig nachweisen. Kritische employees-/Legacyrechte aus #32 und nicht fertige 8B-Verifier-/Sessionintegration bleiben eigenständige Sperren.
2. D01–D14 je betroffener Phase entscheiden; keine unklare Freigabe- oder Löschfunktion freischalten. Zweck-/Feld-/Bereichsprofile und System-/HR-/QS-/Produktionsgrenzen bestätigen.
3. Additives reviewtes SQL-/Storage-/Workerpaket erst später erstellen, Quellen-/Versions-/Zielguards, geschlossene Defaults, Postconditions und bekannte idempotente Wiederholung. Unbekannte Drift/Legacyzuordnung führt zum Abbruch, nicht Reparatur durch offene Grants.
4. Vollständige Original-/Storage-Restoreprobe, Rollen-/API-/Parallel-/Geräte-/Druck-/Fehlerabnahme und Last mit genauem Paket-SHA. Keine produktiven Dokumente in Git/CI/öffentlichem Testbucket.
5. Betreiber legt Backupverschlüsselung/Schlüsselverwahrung, PITR, RPO/RTO, Wartung, Schreibruhe, Verantwortliche und Stopkriterien fest. DB- und Storagekonsistenz separat prüfen; Revision-/Audit-/Produktionshistorie schützen.
6. Erst separat freigegebener kontrollierter Einsatz: geschlossene Zusatzobjekte → begrenzte Upload-/Readpfade → bestätigte Vorlagen/Grants → Pilot → spätere Fachintegration. Niemals bereits öffentliche Datei plus nachträglich hinzuzufügende Policy.
7. Bei Fehler vor Commit Rollback; nichttransaktionale Dateien in Quarantäne erhalten. Nach Commit sicherer Wartungszustand/Vorwärtskorrektur oder genehmigter Recoveryabgleich, keine Wiederaktivierung unsicherer Legacyrechte. Keine neueren Protokolle durch alten Restore überschreiben.

**Ergebnis: keine Produktivfreigabe.** Dieses Konzept kann separat fachlich reviewt werden; weder Fotoanalyse noch grüne Bestands-QA ersetzt Freigabeverträge, native Implementierungsabnahme oder sichere Betriebsnachweise.
