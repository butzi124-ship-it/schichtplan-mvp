# Entwicklung und Freigabe

## Aufgaben ohne Übertragen zwischen Chats

GitHub-Issues sind das Backlog. Ein Auftrag an Codex/Work nennt Repository und Issue-Nummer; der Agent liest Issue, Kommentare, AGENTS.md und aktuellen Code. Pro Aufgabe: Issue → Arbeitsbranch → Draft-PR → CI → Review/Freigabe → Merge → Pages. Umsetzung und überprüfbare Ergebnisse bleiben im PR. Es besteht keine Abhängigkeit von einem offenen Desktop-Chat. Dieses Repository startet selbst keinen KI-Agenten und erteilt keine automatische Merge-Freigabe.

## Lokale Prüfung

Node 22 wie in CI verwenden:

```sh
npm ci
npx playwright install --with-deps chromium
node --check schichtplan_mvp.js
node --check tests/smoke.spec.js
npm run check:syntax
npm run test:ci
npm run test:e2e
git diff --check
```

Alle vorhandenen Tests und Gerätefilter bleiben erhalten: Desktop Chromium führt die gesamte Suite aus, Phone und Tablet die bisherigen Produktionstests. Mindestens 251 Browser-Testfälle sind der Ausgangsstand. CI verbietet test.only. Screenshots und Traces bleiben bei Fehlern erhalten; HTML, JSON und Zusammenfassung werden 14 Tage als Artefakte gespeichert. Der Actions-Lauf zeigt im Summary Ergebnisse je Gerät, auch instabile und übersprungene Fälle. Bei frühem Abbruch steht ausdrücklich, dass Browser-Ergebnisse fehlen. Der Statuscheck bleibt `e2e`; fehlgeschlagene Schritte werden nicht ignoriert.

Die Browser-Suite arbeitet mit synthetischen Daten und Supabase-Mocks. Sie beweist keine echten SQL-, RLS- oder Concurrency-Eigenschaften. Ein separates Supabase-Testprojekt ist optional und derzeit nicht eingerichtet. Niemals Produktionszugänge oder -daten verwenden. Firefox/WebKit oder visuelle Vergleiche nur bei entsprechendem Änderungsrisiko ergänzen.

## Verbindliche Einrichtung durch Repository-Administrator

Historischer Stand 05.11.2026: Rulesets-Liste leer, Branch-Schutz noch nicht vorhanden. Am 06.11.2026 meldet der Branch-Endpunkt `protected: true`; #6 dokumentiert den aktivierten erforderlichen e2e-Check. Die nachfolgenden Schritte sind daher eine Prüfanleitung, keine Aufforderung zum doppelten Anlegen von Regeln. Aufgabe #7 verwendet die QA-Pipeline aus PR #2 und ergänzt einen e2e-Kompatibilitätscheck, der ausschließlich nach erfolgreichem Quality Gate grün wird. Vor Integration den aktuellen Regelstand und die Pages-Quelle erneut prüfen. Der Connector bietet keine Schreibfunktion für diese Administration. Vor künftigen Releases den bestehenden Schutz anhand folgender Kriterien prüfen; keine doppelten Regelwerke anlegen:

1. Repository → Settings → Rules → Rulesets → New ruleset → New branch ruleset. Name `main-reviewed-e2e`, Enforcement status **Active**.
2. Target branches → Include default branch (`main`). Bypass list leer lassen, auch keine Administrator-Ausnahme hinzufügen.
3. **Restrict deletions**, **Block force pushes** und **Require a pull request before merging** aktivieren. Mindestens eine unabhängige Approval, stale approvals verwerfen, Gespräche auflösen lassen. Für ein Ein-Personen-Repository muss ein weiterer berechtigter Reviewer eingerichtet werden; eine Selbstfreigabe ist kein unabhängiges Review.
4. **Require status checks to pass** aktivieren, `e2e` hinzufügen, als Quelle GitHub Actions wählen. **Require branches to be up to date before merging** aktivieren. Keine Pfadfilter oder erlaubten Überspringungen hinzufügen. Den Check nach dem ersten Lauf dieses PR auswählen.
5. Speichern. Unter Rulesets prüfen, dass die Regel aktiv auf main wirkt. Alternativ klassische Branch Protection für exakt main mit denselben Anforderungen und **Do not allow bypassing the above settings** einrichten. Nicht zwei widersprüchliche Regelwerke anlegen.
6. Settings → Pages prüfen: bestehende Quelle und Pfad zunächst dokumentieren/sichern. Wenn Pages weiter aus main veröffentlicht, verhindert ausschließlich der verifizierte Branch-Schutz das Veröffentlichen ungeprüfter Änderungen. Ein separater Actions-Deployment-Workflow ist in dieser Änderung bewusst nicht eingerichtet. Die unveränderte Quelle allein garantiert keinen Schutz.

Direktlinks:
- https://github.com/butzi124-ship-it/schichtplan-mvp/settings/rules
- https://github.com/butzi124-ship-it/schichtplan-mvp/settings/pages

## Kontrollierter Abnahmetest nach Aktivierung

1. Einen separaten Probe-Branch von main erstellen und einen Draft-PR nur mit einer neuen Datei `tests/ci-gate-probe.spec.js` öffnen: `const { test, expect } = require('@playwright/test'); test('gate probe', () => expect(true).toBe(false));`. Keine vorhandenen Tests ändern.
2. Roten `e2e`-Lauf abwarten, PR bereit zur Prüfung setzen und im Merge-Bereich nachweisen, dass **e2e required / failed** das Mergen verhindert. Screenshot/CI-URL und SHA dokumentieren. Niemals auf Merge oder Bypass klicken. Scheitert der Blockierungsnachweis, Schutz korrigieren; nicht veröffentlichen.
3. Probe-Datei wieder entfernen, eine harmlose Dokumentationsverbesserung auf demselben Branch behalten und alle Prüfungen erneut ausführen. Aktuellen grünen SHA, Geräteergebnisse und Review festhalten.
4. Pages-/Deployment-Historie prüfen: roter Probe-SHA darf nicht ausgerollt worden sein. main-SHA unverändert prüfen. Probe-PR anschließend schließen oder die grüne Dokumentationsänderung nach unabhängiger Freigabe kontrolliert mergen.

Der rote Probelauf ist vor Einrichtung des Schutzes kein Beweis und wird in diesem Zustand nicht als bestanden behauptet. Issue #1 bleibt offen, bis Administration, roter Blockierungsnachweis und kontrollierte Freigabe dokumentiert sind.

## PR-Bericht und Veröffentlichung

Die PR-Vorlage beschreibt Änderungen und Risiken; GitHub Actions ergänzt je Lauf automatisch den technischen Bericht im Actions Summary und als `test-results/ci-summary.md`. Der Checkout-SHA ist bei PR-Läufen der getestete Merge-Commit; PR-Head wird separat genannt. Beide nicht verwechseln. Kein Bot-Kommentar und keine Schreibberechtigung für CI erforderlich. PR-Beschreibung bei neuem Stand aktualisieren. Actions-Links über den Checks-Bereich öffnen.

Vor Merge: vollständigen Diff prüfen, grünen aktuellen e2e, Geräteabdeckung, VERSION_LOG und unabhängiges Review prüfen. Kein Auto-Merge. Bei rotem main nach einem Merge keine weiteren Releases; bekannte funktionierende Version nur über geprüften Revert-PR wiederherstellen. Deploy-Änderungen benötigen vorher Vergleich und Sicherung der aktuellen Pages-Konfiguration.
