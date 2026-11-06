# Entwicklung, Prüfung und ausdrückliche Freigabe

Aufträge kommen aus GitHub-Issues. Ablauf: Issue → separater Branch → Draft-PR → vollständige CI → Review und ausdrückliche Nutzerfreigabe → kontrollierter Merge. Ein Merge und ein Deployment sind getrennte Entscheidungen. Kein automatischer Merge, kein direkter Push nach main, keine produktive DB-/RLS-/Auth-Änderung. Regeln aus AGENTS.md und die zusätzlichen QA-Sicherheitsregeln gelten gemeinsam.

## Vereinte Pflichtprüfungen

```sh
npm ci --ignore-scripts
node --check schichtplan_mvp.js
node --check tests/smoke.spec.js
npm run check:syntax
npm run test:ci
npm run test:qa
npm run qa:check
npm run test:agents
npm run qa:agent-server
npm run test:e2e
git diff --check
```

Visuals laufen ausschließlich im gepinnten Playwright-1.62.1-noble-Container: zweimal 33 strikte Vergleiche ohne Referenzänderung. test:e2e enthält die vollständige Chromium-Suite und die bestehenden kritischen Firefox-/WebKit-Fälle. 251 ursprüngliche Chromium-Identitäten und ihre Geschäfts-/Sicherheitsasserts bleiben geschützt. Fokus, Skip, erwartete Fehler, akzeptierte Flakiness, erzwungene Klicks oder gelockerte Bildtoleranzen dürfen keinen grünen Lauf erzeugen.

Der Report aus #1 bleibt erhalten: JSON-Reporter, APP_VERSION, VERSION_LOG/Berlin-Zeit, tatsächlicher Checkout-SHA, PR-Head, CI-Link und Geräteergebnisse. Der zusätzliche QA-Reporter liefert Fehlerdiagnosen und Artefakte. Beide verwenden denselben Testlauf; es gibt keine konkurrierenden Browserpipelines. Reports, Screenshots und Traces bleiben 14 Tage verfügbar. Ein Report ist keine Freigabe.

Für den Artikel-PR kommen alle Artikel-Browserfälle und `npm run test:article-db` hinzu. Letzteres führt ausschließlich die vorgeschlagene SQL-Datei gegen synthetische Tabellen/Auth-Shims in einer wegwerfbaren PGlite-Instanz aus. Keine Produktionsverbindung. Das ist kein vollständiges Supabase-Auth-/REST-/Realtime-/Mehrverbindungs-Ergebnis. Die allgemeine `test:db:preflight` muss ohne diese Voraussetzungen weiterhin BLOCKED_NOT_TESTED melden, nicht grün vortäuschen.

## Bestehender Schutz und tatsächliche Gate-Kette

Am 06.11.2026 gelesen: main `1ac495261bb38912507ef1cdb1205aa0eb07c66b`, aktives Ruleset `main-reviewed-e2e` (26664844), PR-Zwang, strikte Aktualität, GitHub-Actions-Check **e2e**, keine Bypass-Akteure, keine Force-Pushes oder Löschungen. Das vorhandene Ruleset wird nicht umgestellt. Die importierbare Vorlage unter docs/qa/main-ruleset.json erhält denselben Pflichtcheck.

Functional Browsers + QA Integrity + Visual Regression → Quality Gate → **e2e**. Beide Aggregatoren laufen mit `always()` und akzeptieren nur `success` ihrer vollständigen Abhängigkeiten. Die Unit-Tests führen die tatsächlichen Shell-Gates mit grünen, roten, abgebrochenen, ausgelassenen und leeren Resultaten aus. Ein roter/skipped/cancelled Schritt kann nicht kosmetisch als grün durchgereicht werden.

## Deployment ist separat gesperrt

Der Workflow deployt weder bei push noch bei pull_request – auch nicht nach main. Der Deployment-Job benötigt erfolgreiche e2e-Prüfung, main als Ref, `workflow_dispatch` und das ausdrücklich gesetzte Boolean-Input `deploy_approved=true` (Standard false). Dieses Input darf erst nach ausdrücklicher Nutzerfreigabe gesetzt/ausgelöst werden. Dieser Integrationsauftrag enthält keine solche Freigabe.

Wichtig: Eine eventuell noch aktive alte Pages-Quelle „Deploy from a branch / main“ würde den neuen Actions-Deployment-Job umgehen. Ihr aktueller Administrationszustand ist über den Connector nicht nachgewiesen. **Vor jedem tatsächlichen Merge mit möglicher Veröffentlichung** muss der Administrator unter Settings → Pages die Quelle prüfen und gesichert auf GitHub Actions umstellen oder Legacy-Deployment anderweitig stoppen. Keine Behauptung eines vollständig abgesicherten produktiven Deployments ohne diesen Nachweis. Diese PRs ändern keine Repository-Settings und führen kein Deployment aus.

## Integrationsreihenfolge

PR #2 nimmt main in seine Branch-Historie auf und erhält dessen Regeln, Vorlagen, Berichtstests und Reload-Korrektur. PR #11 übernimmt anschließend diesen geprüften QA-Stand. Beide werden gegen aktuelles main geprüft; PR #11 enthält bis zur Integration von #2 dessen QA-Abhängigkeit und darf deshalb nicht vor #2 integriert werden.

Nach einer später ausdrücklich genehmigten Integration von #2 ist #11 gegen das dann tatsächlich neue main erneut auf Aktualität/sauberen Diff zu prüfen; bei Squash/Rebase oder weiteren Änderungen sind Synchronisation und CI erneut erforderlich. Aktuelle Freigabebereitschaft bezieht sich ausschließlich auf die dokumentierten Heads und main-Basis. Keine Vorabfreigabe für künftige SHAs. Aufgabe #8 beginnt erst nach ausdrücklicher Freigabe dieser Integrationsschritte.

Produktive Artikel-Migration und Aktivierung bleiben gesondert freizugeben. Bis dahin bleibt der Artikel-Schalter aus. Alle bisherigen BA-, Mengen-, Rollen- und Datenbankinformationen bleiben erhalten.
