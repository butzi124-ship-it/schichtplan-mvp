# Verbindliches Referenzdesign – Paket 1

Basis: v0.19.3, PR #44, `4eb014e86e3d7f9a7a1d53055cb650f15ec8bc5b`. Auftrag: freigegebenes Nutzerbild vom 15.11.2026, Fertigungsübersicht und zugehörige BA-/OP-Ansichten. Bezug: #44 und bestehende Dokumentationsarbeit #40. Kein Merge, Deployment, Produktionszugriff oder neue Referenzfreigabe.

## Umsetzung

`css/production-design.css` definiert gemeinsame Farben, Karten, Status-Badges, Tabellen, Fortschrittsanzeigen, Formularfelder und Dialogflächen. `js/production-design.js` stellt ausschließlich bereits freigegebene, geladene Daten dar. Die Zugriffsauswahl verbleibt in den bestehenden Funktionen. Die dunkelblaue Navigation erhält Vektoricons, Marke und aus bereits sichtbaren Produktionsaktionen abgeleitete Untermenüs. Der obere Suchbereich filtert ausdrücklich nur die bereits geladenen BA-/Artikelzeilen der Fertigung.

Die Fertigungsübersicht besitzt fünf aus sichtbaren Daten abgeleitete Karten, Abteilungs-/Maschinen-/Statusfilter, Auftragstabelle und auswählbare Detailvorschau. Gutmengen stammen aus der bestehenden Berechnung fertiger Gutteile. Tabellenzeilen und Vorschau öffnen den bestehenden Auftrag; keine neue Mengenbuchung. Die Maschinenübersicht bleibt direkt erreichbar. Vorhandene BA-Erstanlage, Snapshots, OP-Formulare, Zähler, Rüst-, Mengen- und Abschlussbereiche verwenden die gemeinsamen Darstellungsregeln. Produktionsflags bleiben unverändert ausgeschaltet; isolierte Tests schalten sie ausschließlich in ihren Fixtures ein.

## Bewusste Unterschiede zur Bildvorlage

- Keine Beispielzahlen, Mitarbeiteridentität, Bauteilfotos oder fiktiven OPs übernommen. Statt ungeprüfter OP-Kennzahlen zeigt die Übersicht echte aktive BA, laufende/pausierte BA, sichtbare Maschinen und BA mit Abklärmengen.
- Kein erfundener Trend, Benachrichtigungszähler, 30-Sekunden-Auto-Update oder globaler Suchdienst. Suche betrifft nur sichtbare Fertigungsdaten. Bestehende Sitzungsanzeige bleibt erhalten.
- Fertige Gutteile werden gegen das vorhandene Auftragsziel dargestellt; keine Neudefinition der Mengenlogik.
- Kein künstlicher Rüsten-/Geplant-Status für Legacy-BA; existierende Statuswerte werden unverändert beschriftet. Echte OP-Arbeitszustände bleiben in den bestehenden freigegebenen OP-Ansichten.
- Vorhandene Datenerfassungsfelder und Sicherheits-/Konflikthinweise bleiben erhalten. Auf Tablet steht die Detailkarte unter der Tabelle; auf Smartphone wird die Tabelle in einem beschrifteten horizontalen Scrollbereich bedient.
- Keine neuen Dokumenten-, Export-, Planungs- oder Abschlussaktionen allein aufgrund der Vorlage.

## Prüfung und Abnahmegrenze

Alle bestehenden Tests und Bildtoleranzen bleiben erhalten. Exakte reversible Darstellungsdeltas stellen die ursprünglichen geschützten Dateien wieder her; zusätzliche Negativtests verwerfen fremde Änderungen. Die neuen Browserfälle erzeugen Aufnahmen der laufenden Anwendung für Übersicht, BA-Details, 6M-Dialog, BA-Formular und OP-Snapshot je Gerät/Browser. Sie prüfen echte Klicks, Filter, Erhalt der Daten und den bestehenden isolierten SQL-Ablauf.

Lokal bestanden: Syntax, QA-Schutz, Reporttests, Dokumentationstests/-struktur. Lokale Browserstarts und Containervisuals sind umgebungsbedingt blockiert. Der verbindliche Nachweis des aktuellen Commits wird mit Geräten, Fehlern, CI und tatsächlichen Screenshots im Draft-PR ergänzt. Dieses Dokument behauptet keine noch nicht ausgeführte Sichtabnahme. Alte Referenzen werden auch bei Designabweichungen nicht ersetzt; neue Ist-Bilder dienen ausschließlich der gesonderten Nutzerabnahme. Weitere Module erst nach dieser Freigabe.
