# Humbel App: verbindlicher Entwicklungsablauf

- Auftrag und Abnahmekriterien aus dem GitHub-Issue lesen, einschließlich Kommentaren. Aktuellen Repository-Stand als Grundlage verwenden.
- Pro Aufgabe separater Branch und Draft-PR mit Issue-Link. Nie direkt nach main pushen. Keine automatischen Merges oder Deployments.
- Änderungen eng auf den Auftrag begrenzen. Keine Produktionsdaten, Datenbankschemas, RLS oder Auth verändern. Keine Produktionsdaten in localStorage spiegeln; nur vorgesehene UI-Präferenzen.
- Rechte admin, department_admin und employee/tool_scanner erhalten. Keine temporären Debug-Overlays hinterlassen.
- Bei jeder Änderung APP_VERSION erhöhen und VERSION_LOG mit tatsächlichem Zeitpunkt Europe/Berlin aktualisieren. Keine alten Zeitstempel übernehmen.
- Pflicht: npm ci, node --check schichtplan_mvp.js, node --check tests/smoke.spec.js, npm run check:syntax, npm run test:ci, npm run test:e2e, git diff --check.
- Keine bestehenden Tests entfernen, überspringen oder Filter einschränken, um grün zu werden. Bei Fehlern Ursache klären und Regression prüfen.
- PR enthält Änderungen, Version/Zeit, Tests je Gerät, CI-Link, Commit und Risiken. Fehlende Prüfungen ausdrücklich ausweisen. Grün bedeutet nicht automatisch freigegeben.
- Erst nach nachgewiesenem Branch-Schutz, grünem e2e für den aktuellen PR-Stand und Review/Freigabe darf kontrolliert gemergt werden.
- Details und noch manuelle Administrationsschritte: docs/DEVELOPMENT.md.

# Humbel QA Safety

- App business logic must not change during QA work. App version/log changes are allowed.
- Use local file pages with tests/helpers/supabase-mock.js or mock-page.js. Never browse the live app from automated tests.
- No production Supabase write, schema, RLS or auth-user changes. Only explicit, verified isolated test databases may ever receive DB write tests.
- Preserve all 251 legacy Chromium tests and their business/security assertions.
- A correct failing assertion exposes a defect: report it and keep it failing. Never skip, fixme, delete, weaken, mark expected-failure, narrow coverage, force/synthesize clicks, or relax visual tolerances to get green.
- The healer may only repair equivalent locators/synchronization after proving the expected behavior. No automatic application changes, merges or deployments.
- Never update snapshots automatically. Linux CI generates initial references once; review the images and version them before comparisons can pass. Later updates require an explicitly authorized design change and image review.
- After any fix run qa:check, test:qa, test:agents, test:e2e and the CI-container test:visual. A blocked DB preflight is NOT a passed database test.
- Use development branches and PRs. Never bypass required checks or push release changes directly to main.
