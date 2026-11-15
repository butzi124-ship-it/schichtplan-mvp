# Visuelle Abnahme – Nachweise und Stop-Kriterien

## Geprüfter Ausgangspunkt
- Ausgangscommit `d7c0c22c6a9ac8563d4ecd48f0b2759d42e524e3` (v0.20.6) in Draft-PR #45.
- 853/853 funktionale Browserfälle bestanden; Dokumentation, Design-Display, BA/OP und Sicherheitsprüfungen grün.
- Strikter Bildvergleich noch rot; **33/33 neue review-only Vorschläge**, keine Baseline freigegeben.
- Referenz-Testlauf: https://github.com/butzi124-ship-it/schichtplan-mvp/actions/runs/43726264797

## Vor jeder Referenzfreigabe
1. Jede der 33 Live-Aufnahmen (11 Szenen × 3 Geräte) visuell auf offensichtliche Fehler prüfen: Überlauf, abgeschnittenen Text, modale Überdeckung, leer laufende Bereiche, bedienbare Mindestgrößen und tatsächliche Daten.
2. Zwei unmittelbar wiederholte Aufnahmen müssen stabil sein. Unterschiede als Rendering- oder Fixture-Problem untersuchen statt unbesehen Baselines auszutauschen.
3. Zeige aktuelle Ansicht, alte Referenz und Diff pro Szene. **Nur** nachvollziehbare, gewünschte Designänderungen als neue Referenzen vorschlagen.
4. Neue Baselines erst nach **gesonderter ausdrücklicher Nutzerfreigabe** committen. Der vorliegende Designpaket-Upload ist **keine** solche Freigabe.
5. Die drei `.svg`-Entwürfe dürfen nie in `tests/**/__screenshots__`, Snapshot-Pfade oder den CI-Baseline-Baum kopiert werden.
6. Sicherheit/Geschäftslogik und sämtliche 251 geschützten historischen Testidentitäten unverändert. Keine Test-Assertions abschwächen oder Skip-/Focus-Regeln ändern, nur um grün zu werden.
7. `AGENTS.md`/`docs/DEVELOPMENT.md` strikt befolgen: Änderungen separater Branch + Draft-PR, Version und Änderungslog bei App-Änderung, `git diff --check`, QA, Syntax, Browser, dokumentierte Workflows. Merge/Deployment nur nach zusätzlicher Nutzerfreigabe.

## Qualitätsmerkmale nach Viewport
- **Desktop:** dunkelblaue linke Navigation, sichtbarer ausgewählter Menüpunkt, weißes Kartenraster, Tabelle plus Detail, keine absoluten Overlays über Dateneingaben.
- **Tablet:** Navigation und Filter umfließen; 6M-Dialog erreichbar/scrollbar; keine abgeschnittenen Tabellenaktionen; 44-px-Touchziele.
- **Smartphone:** kein Dokument-Horizontalüberlauf; Zähler und Aktionsflächen nicht beschnitten; klare Menütexte; Code128 inklusive Ruhezone dekodierbar; 6M und Checkliste vollständig erreichbar.
- **Alle:** Kontraste und Status semantisch, Fokus tastaturbedienbar, kein versteckter Skip-Link auf Screenshots ohne Fokus; Bedienfläche „Auftrag abschließen“ mindestens 52 px.

## Abnahme-Etiketten
- `DESIGN-GUIDE`: diese Dokumente und SVGs, keine funktionale Umsetzung.
- `IMPLEMENTIERT`: Codeänderung vorhanden, noch kein Beweis korrekter Darstellung.
- `GETESTET`: gültige automatische Tests auf exakt passendem Commit.
- `REFERENZFREIGABE`: erst nach Bildvergleich und expliziter Freigabe durch Auftraggeber.

Kriterium für „fertig“ ist nicht allein ein grüner CI-Lauf. Keine Veröffentlichung ohne getrennte Merge-/Deployment-Freigabe.
