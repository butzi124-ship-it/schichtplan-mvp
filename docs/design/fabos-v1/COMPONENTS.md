# Komponenten- und Interaktionsregeln

## Navigation und Shell
- Desktop ab 801 px: vertikale Navy-Navigation links, Breite laut bestehender Tokens (264 px; unter 1201 px 220 px). Sichtbarer aktiver Zustand in Primärblau, Fokus klar erkennbar. Hauptfläche hell, weiße Karten.
- Tablet/Mobil bis 800 px: bestehendes Menü-Toggle samt `aria-expanded` und Escape-Handling behalten. Navigation darf als responsives Raster/ausklappbares Menü umfließen; keine abgeschnittenen langen Wörter. **Nicht** ohne Freigabe in eine neue Routenarchitektur umbauen.
- Der „Zum Inhalt“-Skip-Link ist nur bei Fokus sichtbar und immer per Tastatur erreichbar. Menüreihenfolge entspricht vorhandenen Rechten; keine erfundenen Navigationspunkte.
- Header enthält bestehende Session-Informationen, ggf. Stillstand sowie Utility-Funktionen **im normalen Dokumentfluss**, nicht über inhaltlichen Buttons.
- Angezeigter Markenname in den SVGs ist FABOS als **Zielbild**, im aktuellen Quellcode bleibt die bestehende Brandinglogik bis zu separater Freigabe unangetastet.

## Oberflächen
- Hintergrund `#F3F7FC`, Karten weiß mit dünnem `#DCE5F1`-Rand, 12 px Radius und sehr dezentem Schatten.
- Formulare: sichtbare Labels und Validierung, mindestens 44 px große Eingabefelder/Buttons, ausreichend Abstand und echte Fehlermeldungen.
- Primäraktionen: blau mit weißem Text, Destruktives Rot, Warn-/Ausschussaktionen signalstark; vorhandene Zählerfarben und fachliche Semantik nicht ändern.
- Status nie nur in Farbe ausdrücken: ergänzende Beschriftung/Icon und valide Live-Regionen. Tabellen per eigenem horizontalen Scrollbereich, nie Dokumentüberlauf.
- Leere Bereiche und fehlende Daten als ehrliche Leerzustände ausgeben; keine erfundenen Termine, Fortschritte, Berechtigungen, Mitarbeiter oder Maschinenwerte.

## Produktionsübersicht (`js/production-design.js`)
- Oben Seitentitel, danach KPI-Karten mit **nur aus vorhandenen Daten berechneten** Werten.
- Filter Abteilung, Maschine, Status, Suche; darunter echte BA-/OP-Tabelle und Details. Maschinenkarten als responsives Kartenraster.
- Desktop: Tabelle und Detail möglichst nebeneinander; Tablet/Mobil: untereinander. Filter und Knöpfe dürfen sauber umbrechen; echte Befehle und bestehende Rollen unverändert.
- Nicht ungeprüft „automatische Fertigungsplanung“, „KI-Kalkulation“ oder sonstige Zukunftsmodule anzeigen.

## Fertigungszähler
- Aktive Maschine, BA-Tabs, Spannungs-/Stationstabs, Mitarbeiter-Kontext und Mengenzähler in klarer Reihenfolge.
- Gutteile (+/−) Blau; Ausschuss Rot; In Abklärung Orange, entsprechend der **vorhandenen** App-Semantik.
- Mindest-Touchfläche 44×44 px; Abschluss-Schaltfläche mindestens 52 px hoch.
- Änderungen am Zähler führen nicht zu unbeabsichtigten Scrollsprüngen; kein horizontaler Overflow. Mengen- und Berechtigungslogik bleibt unverändert.
- **6M-Dialog:** Overlay mit klarer Bezeichnung, Ursachenwahl, Bemerkung, Abbrechen/Speichern. Bei engem Bildschirm vollständig scrollbar und Tasten jederzeit erreichbar.
- **Pflichtcheckliste:** Status-/Pflichtlogik, Zugriffsrechte und Sperrung bleiben fachlich unverändert; auch bei vollständiger Menge kein unzulässiger Abschluss.

## Werkzeugverwaltung
- Vorhandene Unterbereiche z. B. Werkzeugliste, Bestand, Entnahme/Einlagerung, Nachbestellen, Journal, QR-Etiketten weiterverwenden, **soweit in der geladenen Rolle vorhanden**.
- Filter und Aktionen können umbrechen. Desktop: tabellarisch, Mobil: kontrollierter Tabellen-Scroll beziehungsweise geeignete Kartenansicht, aber ohne Daten-/Funktionsverlust.
- Im aktuellen Designpaket **keine** CostOS- oder KI-Werkzeugauswahl implementieren.

## Dashboard und Login
- Login: gut lesbare weiße Formular-Karte, nicht vorhandene Auth-Flows erfinden; Test-Login nur in ausdrücklich isolierter Testumgebung.
- Dashboard: vorhandene Schnellzugriffe aus `js/dashboard-home.js` und legitime Leerzustände; keine künstlich erzeugten Kennzahlen. Auf schmalen Bildschirmen ein- oder zweispaltige Karten.

## Accessibility und Qualität
- Kontrast: für normalen Text mindestens WCAG-AA-Kontrastziel 4,5:1 prüfen; große Schrift und UI-Grenzen entsprechend.
- Fokus sichtbar: 3-px-Outline gemäß bestehender Regel. Tastaturbedienung und Screenreader-Beschriftungen intakt.
- Keine `display:none`-Umgehungen für fehlgeschlagene fachliche Tests; keine pauschalen `overflow:hidden`-Regeln, die wichtige Inhalte abschneiden.
- Touch-Skalierung darf einen Code128 nicht unlesbar machen: Generator und unabhängigen Decoder unverändert prüfen.
