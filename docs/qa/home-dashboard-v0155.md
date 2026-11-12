# PR #27: zwei Browserfehler der Startseite / 0.15.5

Ausgangsbasis: PR #27, `feat/home-dashboard-v0153`, `401fbd5a6aac5c067b21f05821dc6ac9fb376040` (0.15.4). Bestehenden Draft direkt weiterführen; keine neue PR-Kette. APP_VERSION **0.15.5**, VERSION_LOG **2026-11-12 06:54 Europe/Berlin**, tatsächlicher Änderungszeitpunkt. Refs #17 und #27.

## Bestätigte Ursachen und eng begrenzte Korrektur

[Actions 43161217288](https://github.com/butzi124-ship-it/schichtplan-mvp/actions/runs/43161217288), Functional Browsers Job `130992595110`: **752 bestanden, sechs fehlgeschlagen**. Aus den tatsächlichen Logs:

- `tests/smoke.spec.js:159`, nur Desktop-Chromium: `getByRole('button', { name: 'Personalverwaltung' })` findet die Hauptnavigation und die neue Mitarbeiterkachel, deren Hinweis ebenfalls „Personalverwaltung“ lautet. Die Kachel heißt sichtbar weiterhin „Mitarbeiter“, ihr Hinweis lautet jetzt **„Stammdaten · Team“**. Damit ist ihr zugänglicher Name von der Hauptnavigation unterscheidbar; Ziel und Rollenprüfung bleiben gleich. Kein aria-hidden/aria-label-Trick zum Verbergen vorhandener Funktionen.
- `tests/smoke.spec.js:1416`, alle fünf Browserprojekte: Der bestehende `clickMain`-Vertrag prüft nach der Rückkehr auf die Startseite den Modulnamen im Inhaltsbereich. Die sichtbare Hero-Überschrift lautet jetzt wieder **„Start / Dashboard“**. Struktur, CSS und Untertitel bleiben erhalten; kein unsichtbarer Testanker.

Der gesamte Diff in `js/dashboard-home.js` besteht aus genau diesen zwei UI-Textänderungen. Keine Ereignisbehandlung, Datenfilterung, Navigation, Rollenauswahl oder Geschäftslogik geändert. Der Haupt-App-Code erhält ausschließlich Version und Log.

Die vorhandenen Browserfälle, Selektoren, Gerätefilter, Assertions und 33 Referenzbilder bleiben unverändert. Die geschützten Geschäftslogik-, HTML-, CSS- und Renderer-Hashes werden **nicht geändert**. Die zwei Textkorrekturen sind stattdessen als exakte Before/After-Paare in `assetPatches` dokumentiert: vor dem bisherigen Renderer-Hashvergleich werden ausschließlich diese zwei eindeutig vorkommenden Zeichenfolgen rückgeführt. Fehlende, doppelte oder abweichende Texte scheitern. Die bestehende Hash-Assertion bleibt unverändert aktiv.

Zwei additive Node-Regressionen führen den tatsächlichen Renderer aus und prüfen den sichtbaren Titel, den einzigen Personal-Schnellzugriff mit unverändertem Ziel sowie die Ablehnung fehlender, doppelter oder fremder Textänderungen. Die bisherige Schutzregression prüft weiterhin zusätzliche Datei-/Hook-Manipulationen. Keine bestehende Erwartung ersetzt oder gelockert.

## Lokale Prüfungen des neuen Dateistands

| Prüfung | Ergebnis |
|---|---|
| `npm ci --ignore-scripts` | PASS |
| Beide Pflicht-`node --check`, `npm run check:syntax`, zusätzliche Checks für Renderer, Policy und neue Regression | PASS |
| `npm run qa:check` | PASS: ursprüngliche Hashes, 251 Legacy-Identitäten und 33 Bildreferenzen erhalten |
| `npm run test:qa` | PASS: **18/18**, keine Skips (16 bestehende plus zwei neue Prüfungen) |
| `npm run test:integration` | PASS: **5/5**, keine Skips; isolierte PGlite-/Manifestverträge |
| `npm run test:ci` | PASS: **2/2** |
| `npm run qa:agent-server` | PASS: offizieller Playwright-MCP-Handshake |
| Gezielte sechs bestehende Browserfälle | **BLOCKED_NOT_TESTED**: zunächst exakt diese sechs Fälle aufgerufen; alle sechs scheitern vor Seitenaufruf wegen fehlender Playwright-Binaries |
| Anschließend `npm run test:e2e`, unveränderte vollständige Matrix | **BLOCKED_NOT_TESTED**: 758 Fälle in Desktop-Chromium, Mobile-Chromium, Tablet-Chromium, Firefox und WebKit; Browserstarts wegen fehlender gepinnter Binaries blockiert. Kein bestandener Browsernachweis für den neuen Stand |
| `npm run test:agents` | **BLOCKED_NOT_TESTED**: Seed aufgerufen; gleicher fehlender Browser |
| `npm run test:visual` | **BLOCKED_NOT_TESTED**, Exit 2: vorgeschriebener Playwright-1.62.1-noble-Container fehlt; Referenzen unverändert |
| `npm run test:db:preflight` / erneute native Abnahme | **BLOCKED_NOT_TESTED**, Exit 2: keine Docker-/Podman-Laufzeit, kein geprüfter isolierter Loopback-/Auth-/RLS-Teststack verfügbar |
| `git diff --check` | PASS |

Gezielter Browseraufruf vor der vollständigen Matrix:

```sh
npm run test:e2e -- tests/smoke.spec.js --grep 'loads, logs in locally, and navigates modules and subtabs|responsive defaults and explicit preference survive navigation and reload'
npm run test:e2e
```

Der Filter betrifft nur den ausdrücklich verlangten ersten Diagnoseaufruf. Weder Projektkonfiguration noch vollständiger Regressionstest wurden eingeschränkt. Keine Wiederholungsschleife und kein Ausweichen auf die Live-App.

## Historische Evidenz, Grenzen und Abschluss

Im angegebenen Ausgangslauf bestanden QA Integrity und zweimal Visual Regression; Functional Browsers war mit den oben genannten sechs Fehlern rot. Die bereits erfolgreiche native Datenbankabnahme ist vom Auftraggeber für den Ausgangsstand bestätigt; sie wird nicht als erneut ausgeführter Test des neuen Commits ausgegeben. Hier keine Neuentwicklung der QA-, Visual- oder DB-Abnahme.

Der neue Commit korrigiert beide nachgewiesenen UI-Ursachen. Eine Behauptung „alle sechs Browserfälle jetzt grün“ wäre ohne deren Browserausführung falsch. Die aktuelle GitHub-CI wird nach dem Push separat kontrolliert; keine Actions-Abfrage oder Warte-/Dispatch-Schleife nach Sicherung.

Die offenen Sicherheits- und Produktivfreigabeblocker aus #25 bleiben bestehen. Keine Änderungen an Supabase, RLS, Auth, Berechtigungen, BA/OP, Schichtplanlogik, Werkzeugverwaltung, Feature-Flags, Workflows oder Pages. Kein Merge, Deployment, produktiver Datenbankzugriff oder Migration.
