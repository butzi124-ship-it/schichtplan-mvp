# Qualitätsprüfung – Entscheidungsregister, Änderungen und Abnahme

Zum [Fachvertrag](../requirements/digital-quality-2026-11-14.md). **Dokumentation, keine Implementierung.** B/P/O-Notation gilt weiter. Neue Szenarien A01–A15 sind geplant und **BLOCKED_NOT_TESTED**, bis ein separat beauftragter Kandidat tatsächlich existiert; grüne Bestandsprüfungen sind keine Prüfung dieser neuen Funktionen.

## 1. Änderungsübersicht gegenüber PR #37

| Bisheriger Punkt | Jetzt verbindlich entschieden | Was weiterhin offen bleibt |
|---|---|---|
| D01 Variante A/B | Dynamische automatisierte Gutteilprüfung braucht auswertbare Einzelprüfungen | Allgemeine Dokumentenerfassung A/B/Hybrid weiter offen; Scan-/Transkriptionsworkflow |
| D02/D03 Grenzen/Layout | Blau-weißes freigegebenes Bildlayout, Balkenmitte = Toleranzfeldmitte; Grün/Gelb/Rot, untere/obere Eingriffsgrenze je Merkmal; negative qualitative Prüfung voller Rotbalken | Echte Grenzwerte, Grenzinklusion und einseitige Felder O03, Originalangaben52978 nicht verändert |
| D04 Freigaberechte | Gelb tatsächliche Korrektur/Dokumentation + nächstes Teil, OP-Sperre bis Dokumentation; AL nur Kenntnisnahme. Rot Entscheidung durch zuständigen AL oder berechtigten STV | Zulässigkeitsobergrenzen, Vieraugen-/Betroffenheits-/Nacharbeitsdetails O05, kein neuer QS-/PL-Override |
| D05 QS-Zuordnung/Vergabe | Q05 weiterhin fest: AL QS nur QS-zugeordnete Empfänger/eigene delegierbare Rechte | Vertrauenswürdige Zuordnung, Vorlagenpflege und erlaubte Regelkonfiguration O09/O10 |
| D06 Intervalle | Eigene Fälligkeit pro Merkmal,100/50/20/10 % und weitere Prozent-/Zyklusregeln, Erst-/Letzt-/Nachprüfung zusätzlich | Exakte Prozent-/Start-/Zählbasis-/Vorrangregeln O01/O02;20%-Notiz des alten Fotos bleibt eigene ungeklärte Referenz |
| D07 Korrektur | Werkzeug-/Prozesskorrektur dokumentiert, Originalmessung unveränderbar | Fremde QS-Messwertkorrektur/Transkription weiterhin gesondert; Produktion darf QS-Ergebnisse nicht kommentieren |
| D08 Teilenummern | OP-Übernahme Ja/Nein, bei Nein OP-lokal ab1, bei Ja vorhandene Nummer je Buchung/Eindeutigkeit, später gravierte Nummer; keine erfundene Rückverfolgbarkeit | Namespace, Moduswechsel, Nacharbeits-/Spannungszuordnung O04/O07 |
| D09 Sichtweite | Cockpit PL/QS alle Produktionsabteilungen, AL alle eigenen, Produktion arbeitsnotwendig; QS-Sonderfreigabe mit Entscheidername/Zeit/Grund | Vollständige Feld-/History-/Entwurfssicht, erlaubte Doppelrollen unverändert zu entscheiden |
| D10 Druck/Export | Humbel DIN A4 quer, Firmenlayout Hoch/Quer konfigurierbar | Export-/Download-/Offline-/Signatur-/Archivberechtigungen |
| D11/D12/D14 Retention/Dateien/Signatur | Unverändert geltende Sicherheitsgrenzen, keine automatische Übernahme5-Jahresfrist | Fristen/Holds/Dateiformate/Recovery/Signaturniveau weiterhin offen |
| D13 Lifecyclekopplung | Qualitätsabhängige Gutbuchung, OP-Gelbsperre, automatische Abklärung→Ausschussumbuchung und Benachrichtigung fachlich beschlossen | Exakte Spannungs-/Altpfad-/Letztstück-/historische Korrekturadapter O02/O04/O08; kein Produktionscode umgesetzt |
| Neu Mandantenfähigkeit | Firmeneinstellungen getrennt; keine Humbel-Anweisung/-Merkmale/-Intervalle für andere Firmen | Gesamtsystem-/Trust-/Storage-/Betriebsisolation O10, keine bloße Firmenpfadlösung |

Q01–Q06 sowie Organisations-/Ernennungs-/Delegations-/HR-/TA-Grenzen aus8A erhalten. Neue Produktionsprüfungen erlauben eigene Maßnahmendokumentation, nicht Manipulation fremder QS-Ergebnisse. PL-Cockpit-Gesamtsicht ist kein neuer roter Entscheidungs-/Ernennungsgrant. Die ältere Papierregel „außerhalb Grün Vorgesetztenfreigabe“ wird nicht gelöscht oder als falsch bezeichnet; für den **neuen digitalen Gelbablauf** gilt der jüngere explizite Beschluss ohne Freigabegate.

Historische Prüfprotokolle (`documents-inspections-validation.json`,8A-/8B-Nachweise) bleiben unverändert und gelten nur für ihre damaligen Quellen-/Commitstände. Neue Dokumenthashes und Ergebnisse im [aktuellen Prüfprotokoll](digital-quality-validation.json), kein Überschreiben alter Nachweise.

## 2. Verbleibende offene Punkte – keine stillen Freigaben

| ID | Konkrete Entscheidung / nächste Handlung |
|---|---|
| O01 | QS/Produktion bestätigt je Regel: Prozentalgorithmus, Prüfphase, gezählte Grundgesamtheit, freie Zyklen, Vorrang zwischen Firmen-/Abteilungs-/Artikel-/Merkmalsregeln und Verhalten bei Wechsel. Gleichheit0.03/0.10, Einheit/Toleranzbreite und Ausnahmen der Dreherei-Anweisung mit Originalquelle klären |
| O02 | Erst-/Letztstückidentifikation bei Schichtwechsel, Teilmengen/Nachfertigung/Mengenänderung/Unter-Soll-Abschluss; welches nächste Teil bei paralleler Fertigung als Nachprüfung reserviert wird, welche Merkmale nachzuprüfen sind. Verpflichtendes nächstes Teil bleibt fest; Fotozusatz korrigiertes Teil ggf. zusätzlich |
| O03 | Numerische Toleranzgrenzinklusion, Rundungs-/Messunsicherheitsregeln, obere/untere Eingriffsgrenze, einseitige Feldgrenzen. Gelb bei erreichtem und innerhalb Toleranz überschrittenem Eingriff als lückenloser Vorschlag bestätigen; reale Werte nicht aus Bild kopieren |
| O04 | Bestehender Zähler/Spannungen: Zuordnung technischer Teilevorgänge, reservierter Folge/Warteprüfung zu Altzählungen; OP-Gut erst letzte Spannung, Nacharbeit ohne Doppelzählung. Ungesicherte Legacy-/Parallelbuchungswege und Korrektur bereits gezählter Teile vollständig vertraglich schließen |
| O05 | Rot: Kriterien zulässiger Sonderfreigabe (ggf. Kunde), unabhängige Entscheidung bei Eigenbetroffenheit, Ersatzentscheid bei fehlender Authority, Nacharbeitsabschluss/-nachprüfung. AL-/berechtigter-STV-Zuständigkeit ist bereits beschlossen, nicht wieder offen |
| O06 | Meldungskanal, Zustellung/Retry/SLA, fehlender/mehrdeutiger/abwesender AL, Vertretungskenntnisnahme und gesonderte Eskalation. Rot zunächst ausschließlich AL, QS kein Sonderfreigabe-Push, keine versteckten weiteren Empfänger |
| O07 | Nummern-Namespace und Normalisierung (führende Nullen), Gravur-/Moduswechsel, Nachprüf-/Nacharbeitsbuchung, mehrere Mitarbeitende/Spannungen, Nachweis einer rückwirkenden physischen Zuordnung. Unbelegte Links bleiben ausdrücklich unbekannt |
| O08 | Umfang rückwirkender Sperren/erneuter Prüfung seit letztem sicheren Gutteil bei Rot; Auswirkung auf Folge-OP/Versand und nachträglich beanstandete bereits abgeschlossene BA. Keine automatische historische Wiederöffnung/Mengenvernichtung |
| O09 | Wer pflegt/freigibt Firmenvorlagen, Intervalle, Grenzwerte, Layouts und Nachrichteneinstellungen? Versionsfreigabe/Änderung mit offenen Teilen. Keine globalen Admin- oder Selbstvergabevollmachten erfinden; QS-Zuordnungs-/Delegationsgrundregeln erhalten |
| O10 | Betreiber/Security: belastbare Firmen-/Auth-/Rollenbindung, gemischte/Mehrfirmenbenutzer und Wechsel, vollständige Alt-/API-/Storage-/Worker-/Export-/Backupisolation, erstmalige Migration bestehender Humbel-Daten. Keine automatische Tenantzuweisung aus Profil/Hostname |
| O11 | Fortgeltende #37-Datenschutz-/Export-/Retention-/Signatur-/Dateiformatfragen, Originalrestore, Backup/PITR, messbare RPO/RTO, minimaler Owner, reale Auth-/Session-/Widerrufsabnahme und getrennte Betriebsfreigaben |

Fehlende Entscheidung blockiert den betroffenen neuen Adapter/Regeltyp, nicht die Dokumentation bereits entschiedener Ziele. Die vorhandene App bleibt in diesem Auftrag unverändert; daraus folgt keine Freigabe ihrer historischen Sicherheitsrisiken.

## 3. Test- und Abnahmekriterien für spätere Implementierung

| ID | Zwingende positive/negative Fälle | Erforderlicher Nachweis |
|---|---|---|
| A01 Layout/Druck | Verbindliche Blau-Weiß-Struktur, Merkmale/Verlauf/Maßnahme/Status, Touch/Fokus/Screenreader; Humbel A4 quer, fremde Firma Hoch/Quer; Mehrseiten ohne abgeschnittene Werte | Desktop/Tablet/Mobil, geprüfter PDF-/Druckexport, keine pauschale Snapshotaktualisierung |
| A02 Balken | Asymmetrisches Toleranzfeld: Mittelpunkt10.010 statt Nennmaß10.000; Auslenkung beide Richtungen, echte Grenzlabels, kein Abschneiden der Bewertung. Qualitativ negativ voller Rotbalken, offen neutral | Numerische Grenzfälle + UI/Export, keine Beispieldatenverwechslung |
| A03 Intervalle | N=1/2/5/10 je Merkmal,30 % nach bestätigtem Algorithmus, freie Zyklen, unterschiedliche Bereiche/Artikel; bei Teil10 Vereinigung N2/N5 ohne Duplikat; keine Fälligkeit keine Maske | Serverseitige Fälligkeitsberechnung, versioniertes Manifest, manipulierte Pflichtliste denied |
| A04 Sonderprüfungen | Erst-/Letzt-/Nachprüfung unabhängig Quote, Schicht-/Mitarbeiter-/Maschinenwechsel, Nachfertigung/Mengenänderung, Abbruch vor letztem Teil; offene Pflicht kann nicht per OP-Abschluss umgangen werden | Native Transaktionen + vorhandene Lifecycle-Positivregression |
| A05 Automatisches Gut | Fehlende/rote/ungültige Pflicht verhindert Gut, alle nötigen Werte bestanden genau einmal Gut; fehlendes Profil kein „ohne Prüfung“. Prüfpflichtig-nichtfällig keine manuelle Gutumgehung | REST/RPC-/Altpfad-Direktangriffe und UI |
| A06 Grenzbewertung | Exakt an/knapp um beide Eingriffs- und Toleranzgrenzen, Dezimalkomma/Einheiten/NaN/falscher Regelstand; Rot vor Gelb; einseitige unbekannte Grenze nicht als0 ergänzt | Exakte Dezimal-/SQL-/API-/Exportverträge |
| A07 Gelb | OP sofort zugehörig gesperrt, tatsächliche Maßnahme/Begründung/Person/Zeit nötig; Dokumentation hebt nur passende Sperre auf; nächstes Teil Pflicht. AL-Kenntnisnahme weder Ersatz der Maßnahme noch zusätzliche Freigabe | Native Paralleltests, UI, Outbox. Wiederholtes Gelb/andere offene Sperren bleiben wirksam |
| A08 Rot | In Abklärung, nur zuständiger AL/gültiger STV entscheidet; fremder AL, abgelaufener STV, bloßer PL/QS/TA denied. Sonderfreigabe zeigt unveränderten roten Wert plus Actor/Zeit/Grund | Echte Rollen/JWT/API/RLS und begrenzte Sonderfreigaberegeln |
| A09 Mengen/Produzent | Ohne Prüfpflicht manuelle3-Wege-Wahl erhalten; Abklärung→Ausschuss −1/+1 atomar; Retry/neue Request-ID/Doubleclick keine zweite Zählwirkung; Prüfer/AL wechselt Produzent nicht | Bestehender Zähler + parallele PostgreSQL-Verbindungen, injizierter Audit-/Zählfehlerrollback |
| A10 Teileidentität | Nein OP1-Teil1 und OP2-Teil1 ohne behauptete Physikbindung; Ja jeder Buchungsschritt prüft vorhandene Nummer; gleichzeitige Duplikate denied; Nachprüfung kein neues Teil; spätere Gravur belegt, unbekannte Vorgeschichte sichtbar | Unique/FK/Transaktions-/UI- und Mehrmitarbeiter-/Spannungsfälle |
| A11 Cockpit | Rot/Gelb/Grün + jeweils neueste fachliche Aktualität; Sonderfreigabe historisch rot, abgeschlossen sichtbar. AL A/B gleichberechtigt/C denied, PL/QS alle eigenen Produktionsbereiche, MA minimal | Projektionen/Suche/Zähler/Export/History prüfen, keine fremden Personen-/Gesundheitsdaten |
| A12 Nachrichten | Rot genau zuständiger AL, kein PL-/QS-/STV-Automatikempfänger; Sonderfreigabe im QS-Auftrag mit Name/Zeit/Grund ohne Push; Gelb Kenntnisnahme getrennt, Zustellfehler kein still verlorener Auftrag | Deduplizierte Outbox/Retry und atomarer Fehlerrollback, fehlender AL nach entschiedenem Verfahren |
| A13 Mandanten | Firmen A/B mit gleichen Artikel-/BA-/Teilnummern/Benutzernamen; fremde ID/URL/Pfad/Filter/FK/Realtime/Worker/Suche/Export/Backup denied; neutrale Vorlage enthält keine Humbel-Inhalte | Native RLS/API/Storage/Owner-/Default-/Definer- sowie Wiederherstellungstests, keine PGlite-Freigabe allein |
| A14 Session/Audit | Deaktivierung/Widerruf während Mess-/Entscheid-/Zähl-/Downloadrequest, zwei Tabs, Stale-/Replaykonflikt, immutable Originale, additive Korrektur; App-TA ohne fachliche Vollmacht | Echt verifizierte Auth/JWT/Session und separate Ablehnungsprotokollierung, keine ungeprüften Claims |
| A15 Bestand/Recovery | Artikel/BA/OP/Spannungen/Maschinenwechsel/Rüsten/Mengen/Klärfälle/Nachfertigung/Abschlüsse/Scanner/Planung/Dashboard unverändert außerhalb genehmigtem Qualitätsadapter; DB+Datei-Restore ohne verlorene Audits/Produzenten | Vollständige bestehende Suites/Identitäten/Assertions und genehmigter Originalrestore mit synthetischen Testdaten |

Alle A-Tests sind **BLOCKED_NOT_TESTED: Zielsystem nicht implementiert**. Tatsächliche aktuelle Prüfungen ausschließlich Syntax, QA, Bestandsintegration und Dokumentkonsistenz gemäß neuem Prüfprotokoll. Fehlende Browser-/Container-/native Laufzeit ausdrücklich getrennt, keine Tests löschen/skippen/filtern/lockern. Jeder gestartete Testprozess mit festem Timeout, kein unbegrenzter Retry/Dispatch/CI-Warteschleife.

## 4. Empfohlene kleine Entwicklungsphasen

1. **Vertragsabschluss:** O01–O05/O07/O09 für Fälligkeit, Grenzkonvention und Zählintegration anhand synthetischer Szenarien entscheiden; bestätigte Ziele nicht erneut öffnen. O10 als Querschnitt für spätere Mehrfirmenfähigkeit planen.
2. **Isoliertes Teile-/Regelmodell:** Versioniertes Firmen-/Merkmals-/Fälligkeits-/Teilemanifest mit rein lesendem Simulationsvertrag; keine neuen Produktivgrants, keine bestehenden Mengen ändern.
3. **Atomarer Zähleradapter:** Exakt eine geschützte Integration des vorhandenen Spannungszählers mit wartenden Prüfungen/Resultaten/Abklärung, Idempotenz und Produzentenbindung. Vor erster Aktivierung alle Alternativpfade abnehmen.
4. **Gelb und Nachprüfung:** Maßnahme, OP-Sperre, verpflichtendes nächstes Teil, Kenntnisnahme und AL-Meldung nativ mit Rennen testen, ohne Freigabegate.
5. **Rot und Cockpit:** Zuständige AL-/STV-Entscheidungen, Ausschussumbuchung, Nacharbeit/Sonderfreigabe, globale Firmen-QS/PL-Sicht und genau begrenzte Meldungen.
6. **UI/Export/Dokumente:** Freigegebenes Bildlayout nach Feldverträgen, Verlauf, A4-Querformat und Firmenlayouts, sichere private Dokumente/Originale nach #37; Variante B nur über geprüften Ersatznachweis.
7. **Gesamtabnahme:** Mandantenisolation des gesamten betroffenen Bestands, echte Auth/API/RLS/Storage/Parallel-/Browser-/Drucktests und isolierte Recovery; danach eigenes geprüftes Migrations-/Pilotpaket mit separaten Freigaben.

Keine Phase gilt durch diesen Dokumentations-PR als implementiert oder produktiv freigegeben. Neue QS-Gates sind spätere fachlich beschlossene Änderungen, kein Anlass zum erneuten Entwickeln von Stückzähler, BA-/OP-Phasen, Werkzeugverwaltung oder Startseite.
