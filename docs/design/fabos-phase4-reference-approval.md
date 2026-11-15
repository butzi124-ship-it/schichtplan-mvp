# FABOS Phase 4 – ausdrücklich freigegebene Screenshot-Referenzen

## Umfang der Freigabe

Der Auftraggeber hat nach Prüfung der Desktop-, Tablet- und Smartphone-Ansichten ausschließlich die 33 stabilen Aufnahmen aus GitHub Actions `43744801271`, Artefakt `14174051024`, freigegeben. Diese Anweisung ist die Freigabegrundlage für die Übernahme. Keine weitere Designänderung, keine zusätzliche Szene, kein Merge und kein Deployment sind freigegeben. PR #46 und sein bekannter Whitespace-Fehler bleiben unberührt.

- Ausgangsversion: v0.20.10; Basis PR #49.
- Quellcommit: `fee93f207e81e8e15dec7cb248387e0bccfc3a15`.
- Tatsächlicher Quell-CI-Checkout: `a2cbcf127209f6b155967d75463673d0e2d180bf`.
- Beide Commits haben denselben Git-Baum `ed44bbf2336ae1220a43a2fd0b7367efa5dae3df`. Der CI-Checkout ist ein synthetischer PR-Merge, kein ausgeführter Merge.
- Quelle: https://github.com/butzi124-ship-it/schichtplan-mvp/actions/runs/43744801271/artifacts/14174051024
- ZIP-SHA256: `cc9e8f33fe111b269e7a3c4f7139601ccb97707bc9df1d6014cecb61eb56fee6`.
- Aufnahmeumgebung: `mcr.microsoft.com/playwright:v1.62.1-noble`.

## Prüfung vor und nach dem Kopieren

Das Artefakt wurde neu heruntergeladen und sein vollständiger ZIP-Hash geprüft. Beide Manifeste weisen jeweils genau 33 eindeutige Aufnahmen ohne fehlende Dateien aus. Ihre Pfade entsprechen exakt den bestehenden elf Szenen auf `visual-desktop`, `visual-tablet` und `visual-mobile`.

Für alle 66 Quelldateien wurden PNG-Abmessungen, SHA256 der Dateibytes und der dekodierten RGBA-Pixel mit den Manifesten abgeglichen. Alle 33 Paare sind sowohl byte- als auch pixelidentisch. Das Stabilitätsmanifest meldet `STABLE_33_OF_33` für den oben genannten Checkout. Erst nach vollständiger Prüfung wurden die 33 Dateien aus `first/` in die bisherigen Pfade unter `tests/visual-snapshots/` kopiert. Jede übernommene Datei wurde anschließend nochmals bytegenau mit der freigegebenen Quelle verglichen. Es wurden keine Bilder neu erzeugt oder als Ersatz verwendet.

`tests/visual-snapshots/provenance.json` dokumentiert die neue freigegebene Herkunft, Version, CI-/Artefaktkennungen, Manifest-Hashes, alle 33 Datei-/Pixelhashes und Bildmaße. Die Kennung und der Hash der vorherigen Provenienz bleiben dort als historische Verknüpfung erhalten; der vorherige Stand liegt unverändert in der Git-Historie. Die sonstigen geschützten Geschäfts-/Testquellhashes bleiben erhalten.

## Erfolgreiche Vergleiche ebenfalls dokumentieren

Playwright 1.62.1 schreibt `*-actual.png` und Differenzbilder normalerweise nur bei fehlgeschlagenen Screenshotvergleichen. Der bisherige Collector aus Phase 3 würde deshalb vollständig bestandene Vergleiche als fehlende Aufnahmen melden.

Die bestehenden Screenshot-Assertions werden unverändert ausgeführt. Erst nach einer bestandenen Assertion wird zusätzlich eine echte aktuelle Aufnahme desselben Ziels mit deaktivierten Animationen, verborgenem Caret und CSS-Pixelskalierung im Test-Ausgabeverzeichnis gespeichert. Bei Fehlern bleiben die tatsächlichen Fehlerbilder von Playwright erhalten. Es werden niemals Referenzdateien als angeblich neue Aufnahmen kopiert.

Ein fehlendes Differenzbild ist nur dann zulässig, wenn der Collector tatsächliche Aufnahme und Referenz dekodiert hat und identische Abmessungen sowie vollständig identische Pixel nachweist. Anderenfalls bleibt das fehlende Differenzbild ein Fehler. Beide aktuellen Bildsätze werden weiterhin separat auf Vollständigkeit, unveränderte Referenzen und Pixelidentität geprüft. Die Galerie kennzeichnet pixelidentische Vergleiche ohne ein künstliches Differenzbild zu erzeugen.

Zwei zusätzliche Collector-Tests prüfen erfolgreiche echte Aufnahmen ohne Differenzbilder sowie die weiterhin abgewiesene fehlende Differenzevidenz bei realer Pixelabweichung. Im bisherigen Negativtest wird dafür die synthetische Aufnahme tatsächlich abweichend gemacht; alle ursprünglichen Assertions bleiben unverändert. Keine Tests werden übersprungen, entfernt oder abgeschwächt.

## Versionierung und Qualitätsgrenzen

APP_VERSION und VERSION_LOG werden auf v0.20.11 mit tatsächlicher Europe/Berlin-Zeit aktualisiert; der Dokumentationsindex wird nachvollziehbar erneuert. Der zusätzliche exakte Aufnahmeblock wird im bestehenden reversiblen Schutzmanifest erfasst. Geschäftslogik, Berechtigungen, UI/CSS, elf Szenennamen, Gerätezuordnungen, Testfilter, Vergleiche und Bildtoleranzen bleiben unverändert.

Die vorhandene Versionsmetadaten-Normalisierung prüft zuerst die tatsächliche APP_VERSION und setzt ausschließlich den sichtbaren Versionslog-Text für den Bildvergleich auf den bereits verwendeten konstanten Wert. Diese Normalisierung wird nicht verändert. Falls trotzdem sichtbare Versionsabweichungen auftreten, sind sie als Fehler zu berichten; diese Freigabe erlaubt keine weiteren Referenzaufnahmen.

Erforderlich bleiben 853 bestandene Funktionstests, zwei bestandene 33er-Bildvergleiche sowie QA Integrity, Dokumentations-, Design-Display-, BA/OP- und Sicherheitsworkflows. Ein grüner Lauf ist keine Merge- oder Deployment-Freigabe. Produktive Datenbanktests bleiben gemäß bestehender Dokumentation gesondert als nicht ausgeführt auszuweisen; isolierte SQL-Tests sind keine Produktionsdatenbankprüfung.
