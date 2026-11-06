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
