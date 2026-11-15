# Gemeinsame Designgrundlage – Arbeitspaket 1

Version 0.19.0. Grundlage: PR #42, unveränderter Originalcommit dc6a286add800d782da2bdaf65dd5dd0c19b9bb0. Dieser Entwicklungsbranch zielt auf `feature/test-preview-documentation-issue40`, nicht auf main. Verwandte Vorarbeiten: PR #27 und Issue #40.

## Verbindliche Gestaltung und Umfang

Dunkelblaue Navigation, helle Karten, dezente blaue Akzente und Systemschrift bleiben die Grundlage. Die Navigation gilt jetzt in jedem vorhandenen Modul, statt nur im Dashboard. Anmeldeseite, Startübersicht, Meine Schichten und To-Do/Meine Aufgaben verwenden gemeinsame Abstände, Typografie und Bedienelemente. Der sichtbare Menüname To-Do bleibt für bestehende Abläufe erhalten. Der Titel im App-Kopf ist Humbel-App; keine Umbenennung auf FABOS.

Desktop verwendet eine feste Seitennavigation; bis 800 px wird dieselbe zulässige Navigation als zweispaltiges Menü im Dokumentfluss dargestellt. Das Menü ist zunächst geöffnet; ein beschrifteter Menübutton kann es ein-/ausblenden. Escape im mobilen Menü schließt es und setzt den Fokus zurück. Desktop-Navigation, aktive Seite, Fokuswiederherstellung nach erneutem Rendern, Sprunglink und beschriftete Tastatur-Scrollbereiche sind gemeinsame Präsentationsfunktionen. Es gibt keine neue Berechtigungsentscheidung, Geschäftsaktion, Speicherung oder Netzwerkverbindung im Shell-Modul.

## Wiederverwendbare Komponenten

| Datei / Klasse | Verwendung |
|---|---|
| css/app-shell.css | Designvariablen für Farben, Abstände, Radien, Schatten, Seitenleistenbreite; globaler App-Rahmen und Navigation |
| js/app-shell.js | Synchronisiert ausschließlich vorhandene DOM-Elemente, aktiven Menüeintrag, Menübutton und Tastaturfokus |
| h-card, h-page, h-page-header | Helle Inhaltskarte, Seitenraster, Überschriftenbereich |
| h-page-title, h-page-description | Gemeinsame Überschrift und erläuternder Text |
| h-field-panel, h-field-label | Formularbereich und sichtbare Feldbeschriftung |
| h-table-scroll | Begrenzte, beschriftete, per Tastatur erreichbare Tabelle |
| h-login, h-app-header | Anmeldung und App-Kopf |
| css/home-dashboard.css | Nur Dashboard-Komponenten; keine modulabhängigen Navigationsregeln mehr |

Die alten konkurrierenden Navigationsregeln wurden aus dem injizierten App-CSS und der Dashboard-CSS entfernt. Fachliche Warn-/Erfolgsfarben bleiben bestehen. Andere Module erhalten zunächst den gemeinsamen Rahmen; ihre fachlichen Detailmasken sind nicht in diesem Paket vollständig umgestaltet.

## Funktionserhalt und Schutz

Anmelde-/Abmeldehandler, Rollenfilter, Sitzungsverwaltung, Datenbankzugriffe, Geschäftsberechnungen, Task-Aktionen, Schichtfilter und Formular-IDs bleiben erhalten. Die Änderungen an den vorhandenen Renderern betreffen ausschließlich HTML-Klassen, Überschriften/Beschreibungen und zugängliche Beschriftungen. Platzhalter und deaktivierte Funktionen bleiben ausdrücklich solche.

`design-foundation-patches.json` beschreibt die exakten reversiblen Präsentationsänderungen. Die neue Policy normalisiert ausschließlich diese eindeutigen Änderungen vor den bestehenden Schutzprüfungen. Ursprüngliche Geschäfts-/Renderer-Hashes, vorhandene Testassertionen, alle 251 Legacy-Identitäten und 33 Bildreferenzen bleiben unverändert. Zusätzliche Negativtests verwerfen fehlende, doppelte oder fremde Änderungen. Der statische Paketierer kopiert jetzt auch CSS; er führt kein Deployment aus.

## Tatsächliche Validierung

- npm ci --ignore-scripts: bestanden.
- Pflichtsyntax inklusive neuer Shell-/Testdateien: bestanden.
- test:ci: 2/2 bestanden; test:qa: 27/27 bestanden.
- qa:check: bestanden, 251 Legacy-Identitäten und 33 Originalreferenzen erhalten.
- test:integration: 5/5 bestanden (isolierte Integrationsverträge; keine native Gesamt-DB-Abnahme).
- test:documentation: 5/5 bestanden; docs:check: bestanden. Hilfequellindex nach Prüfung der geänderten Darstellungsquellen aktualisiert.
- qa:agent-server: bestanden; bestätigt Werkzeugverfügbarkeit, keine Browserbedienung.
- test:agents und vollständiges test:e2e (813 registrierte Fälle, 0 bestanden, 0 übersprungen; alle Browserstarts fehlgeschlagen) wurden aufgerufen: lokal BLOCKED_NOT_TESTED wegen fehlender gepinnter Browser. System-Chromium scheitert vor Seitenaufruf mit socket() Operation not permitted. Keine bestandenen UI-Assertions daraus ableiten. Vier zusätzliche fehlende Trace-ZIP-Dateien entstanden bei gleichzeitigem Schreiben der beiden Browserreporter in dasselbe Artefaktverzeichnis; sie sind Ausführungs-/Artefaktfehler und kein nachgewiesener Appfehler. Weitere Browserläufe müssen getrennte Ausgabeverzeichnisse verwenden.
- test:visual: BLOCKED_NOT_TESTED, vorgesehener Playwright-1.62.1-noble-Container lokal nicht verfügbar. Keine Umgebungskennung vorgetäuscht und keine Referenzen automatisch aktualisiert.
- Neue zusätzliche Browserfälle in tests/design-foundation.spec.js für alle fünf Projekte: Anmeldung, modulübergreifende Navigation, Keyboard-Fokus, mobiler Menübutton/Escape, Logout, eigene Schichten/Aufgaben und beschriftete Admin-Taskfelder. Lokal nicht fachlich ausgeführt.
- Ergänzende Cloud-Browserprüfung konnte die lokale Loopback-Vorschau nicht erreichen; kein Sichtprüfungsnachweis.

Die alten Bildreferenzen enthalten die bisherige Gestaltung. Nach dieser ausdrücklich beauftragten Designänderung ist mit Abweichungen in CI zu rechnen; sie dürfen nicht als Fachfehler oder bestandene neue Designabnahme umgedeutet werden. Neue Referenzbilder erst im gepinnten Container erzeugen, Bilder prüfen und anschließend explizit versionieren. Fehlertoleranzen und bestehende Assertions nicht lockern. Tastatur-, Mobil-, Browser-/Konsolen- und tatsächliche Sichtabnahme bleiben bis zur Browserausführung offen. Aktuelle CI-Ergebnisse sind im Draft-PR zu verfolgen; dieser Bericht behauptet keine grüne neue GitHub-Gesamt-CI.

Ältere Office-Dateien unter docs/generated/0.18.1 sind historische Entwicklungsunterlagen und keine Screenshots des neuen Designs. Keine neuen Office-Unterlagen oder Veröffentlichung in diesem Paket.

## Nächstes Arbeitspaket

Personal-/Schichtmodellmasken, Abteilungs-/Maschinenverwaltung, Artikel-/BA-/OP-Detailseiten und Produktionszähler schrittweise auf die gemeinsamen Komponenten umstellen; anschließend Werkzeug-/Lager-/Scanner- und Auswertungsdetails. Bestehende fachliche Sicherheits- und Datenbankfreigabesperren bleiben unabhängig vom Design bestehen. Kein Merge, keine Datenbankänderung und kein Deployment.
