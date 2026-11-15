# Design Teil 1 – Korrekturen v0.19.3

Basis: PR #43, Commit 67c3fc33a860a9534ed88ef3218ae4307e2e8bf8. Korrekturbranch `fix/design-part1-display`, Ziel `design/shared-foundation-part1`. Keine Änderung an main, Geschäftslogik, Datenbank, Rechten oder Produktionsflags. Kein Deployment, keine Umbenennung.

## Fünf Korrekturen

1. Ausschuss-/Klärfallzähler verwenden ein gemeinsames, nach verfügbarer Kartenbreite umbrechendes Raster. Mindestens 44-px-Schaltflächen; Touchansicht behält ihre 60-px-Flächen. Schmale Tablet-Spannungskarten zeigen die beiden Zähler untereinander, statt Buttons durch Nachbarkarten zu verdecken.
2. Vorhandene Hilfe- und Versionsbedienelemente werden mit ihren unveränderten Handlern in eine gemeinsame Fußleiste im normalen Dokumentfluss übernommen. Keine fixierten Überlagerungen über Überschriften, Zählern oder Formularen. Eine vollständig verborgene Fußleiste bleibt ausgeblendet.
3. Die Höhe des ausschließlich synthetischen Vorschauhinweises wird per ResizeObserver gemessen. Die Desktop-/Tablet-Seitennavigation beginnt darunter; Umbrüche des Hinweises werden berücksichtigt. Kein Zugriff auf Daten oder Autorität.
4. Der zugängliche Menüname Sicherheitsunterweisungen bleibt unverändert. Eine semantische Umbruchstelle nach Sicherheits verhindert mitten im Wort abgeschnittene Navigation. Kein unkontrolliertes overflow-wrap:anywhere in Menüs.
5. Die Fußleiste liegt auch hinter Werkzeugtabellen im normalen Fluss; Tabellenaktionen haben keine schwebenden Hilfs-/Versionsüberlagerungen mehr.

## Bei Auswertung der inzwischen abgeschlossenen Basis-CI festgestellte Oberflächenregressionen

PR #43 hatte 760 bestandene und 53 fehlgeschlagene Browserfälle, nicht nur rote Visuals. Hinzugefügtes role=status am Sitzungsfeld machte bestehende Artikelstatus-Selektoren mehrdeutig; das Feld verwendet jetzt aria-live ohne konkurrierende Statusrolle. Aufgabe/Mitarbeiter/Frist erhalten eindeutige zugängliche Namen. Nach asynchronem Neuzeichnen wird der zuletzt fokussierte Navigationsbutton wieder fokussiert, solange kein anderes Feld den Fokus übernommen hat. Bestehende Browsertests bleiben unverändert; diese Fehler werden im Produkt behoben, nicht durch Lockerung von Selektoren verdeckt.

## Prüfungen und Freigabegrenze

- npm ci, Syntaxprüfung, test:ci (2), test:qa (30), test:documentation (5), docs:check und qa:check lokal bestanden.
- Ursprüngliche 251 Legacy-Identitäten, Business-/Renderer-Hashes und alle 33 Referenzen erhalten. Neue exakte reversible Präsentationsdelta-Prüfung; bestehende Schutzwerte/Assertions unverändert.
- test:agents lokal vor Seitenaufruf blockiert: gepinnte Browser fehlen. test:e2e wird vollständig aufgerufen; kein Browsernachweis aus fehlgeschlagenen Starts. Lokaler Visual-Container fehlt, test:visual meldet BLOCKED. Keine vorgetäuschte Containerkennung.
- Neue additive Browserfälle prüfen alle fünf Korrekturen in fünf bestehenden Browser-/Geräteprojekten: Kartengrenzen, elementFromPoint, echte Klicks auf Ausschuss/Abklärung mit anschließendem Abbruch ohne Buchung, Hilfe-/Versionsdialog, Hinweisabstand, Menütext und Tabellenaktionen. Screenshots und Traces werden in CI bereitgestellt.
- Separater zusätzlicher Design-Display-Workflow führt neue Bedienungsfälle und ursprüngliche Designfälle aus. Er ersetzt keine vollständige Funktionspipeline.
- Die vollständige Pipeline führt weiterhin die ursprünglichen 33 Bildvergleiche mit unveränderten Referenzen und Toleranzen aus. Designabweichungen gegenüber dem alten Layout werden weiter als fehlgeschlagen gemeldet.
- Ein zusätzlicher Sammler kopiert ausschließlich die tatsächlich erzeugten Ist-Bilder in `artifacts/visual/reference-proposals/` samt SHA-/Versionsmanifest und ausdrücklichem Freigabehinweis. Keine Datei unter `tests/visual-snapshots` wird ersetzt. Fehlende Vorschläge schlagen fehl, statt Vollständigkeit vorzutäuschen.

Erst nach Sichtprüfung der neuen CI-Bilder und tatsächlicher Bedienungsnachweise kann Abnahmebereitschaft beurteilt werden. Ein grüner Aufnahmelauf ist keine Freigabe. Aktuelle Ergebnisse und konkrete verbleibende Befunde werden im Draft-PR dokumentiert; dieses Dokument behauptet vor Ausführung keine grüne neue CI.

## Korrektur nach erstem CI-Nachweis

Erster Korrekturcommit 52629125: 44/50 fokussierte Browserfälle bestanden. Alle fünf geometrischen Trefferflächen-/Dialogprüfungen erreichten die abschließende Schreibprüfung; deren synthetischer Auftrag war noch nicht initialisiert und löste die bestehende reguläre Checklistenanlage aus. Die neue Fixture wird wie die unveränderte Visual-Fixture vorinitialisiert; die Assertion auf keine Buchung bleibt unverändert. Zusätzlich zeigte der unveränderte mobile Logouttest eine CSS-Spezifitätskollision: der mobile Gridstil überlagerte die hidden-Klasse. #tabs.hidden hat nun Vorrang. Version 0.19.2 enthält diese Korrekturen. Die drei ursprünglichen Fokusfehler und fünf Beschriftungsfehler waren im ersten fokussierten Lauf bereits bestanden.

Beide strikten visuellen Durchläufe werden auch bei Referenzabweichungen ausgeführt (always für die Wiederholung); kein Fehler wird dadurch als bestanden gewertet. Die neuen CI-Ergebnisse sind separat am aktuellen Head zu prüfen.

## Weiterer vorhandener schwebender Bereich

Die zusätzliche Mitarbeiter-Sichtprüfung zeigte den bestehenden Stillstand-Timer über Schichtaktionen. Er wird mit seinem unveränderten stopDowntimeTimer-Handler in den App-Kopf im normalen Fluss versetzt. Eine zusätzliche Prüfung in allen fünf Projekten sichert Position, Trefferfläche und frei erreichbare Abwesenheitsaktionen. Der vorherige Head ca24e7be bestand bereits alle 50 fokussierten Fälle; der endgültige Head muss die nun 55 Fälle erneut bestehen.
