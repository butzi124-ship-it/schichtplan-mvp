# Automated QA and Release Safeguards (0.5.17)

## Coverage and architecture

- All 251 legacy Chromium identities are versioned in tests/qa/legacy-tests.json: desktop 91, phone 89, tablet 71. No business assertion was removed or weakened.
- Firefox and WebKit each run the 71 critical order-navigation, sequential-counting and touch-view scenarios on desktop. These are desktop browser engines, NOT physical iOS/Android certification.
- Supabase mock data/client moved unchanged from smoke.spec.js to tests/helpers/supabase-mock.js so the official agent seed and tests share one client. Test-only selection logic now derives default view from device settings instead of the project name.
- Tests deny external network by default; local files and fulfilled CDN fixtures are the only allowed resources. No live Supabase SDK is loaded.
- 33 screenshots cover 11 scenes on desktop, phone and tablet. Fixed time, local Tailwind, locale/timezone and reduced motion stabilize rendering. Only the exact Linux Playwright 1.62.1 noble container is accepted for comparisons; Windows screenshots are not substitutes.
- HTML report, named failures, diagnostic hypotheses, actual/expected/diff images, traces, SHA and per-project counts are uploaded even on failure. A hypothesis is not a verified root cause. The healer must inspect evidence, propose a fix and re-run all regressions.
- Skipped, focused, expected-failure and flaky tests cannot silently produce green. Normal visual checks use updateSnapshots=none. Hashes protect app business code, test source and approved images against quiet weakening; intentional future changes require explicit review and a corresponding manifest update.

### Reviewed presentation fix

The first Linux image run detected tool-filter controls overflowing their grid tracks (1090px at 1024px). Six existing filter IDs now have min-width:0 and width/max-width:100%. Mobile still overflowed (399px at 390px): tools/storage change the outer wrapper to max-w-none, so the existing mobile header-wrap selector no longer matched. Its two selectors now accept both existing wrapper classes. No filtering, booking, production or role logic changed. The exact CSS-only insertion and replacements are recorded separately in protected-files.json; reversing only those reviewed edits reproduces the original 0.5.16 business-code hash. No assertion was removed or relaxed, and the defective images were not approved.

Initial generation validates that only the expected 33 missing-reference soft failures occurred, then performs a second strict comparison. Any additional failure aborts generation. Normal CI never runs this initializer.

The first multi-browser Linux run exposed a container startup issue, not an app assertion failure: GitHub mounts /github/home as pwuser while this job runs as root, so Firefox rejected HOME ownership before opening any page. The functional container now explicitly uses HOME=/root, following the browser launch diagnostic. Firefox remains enabled with all 71 assertions/scenarios unchanged.

## Commands

```sh
npm ci --ignore-scripts
npx playwright install chromium firefox webkit
npm run test:e2e
npm run test:chromium
npm run test:cross-browser
npm run test:agents
npm run qa:agent-server
npm run test:qa
npm run qa:check
npm run test:db:preflight
```

`test:visual` is executed by CI in `mcr.microsoft.com/playwright:v1.62.1-noble` with `HUMBEL_VISUAL_ENV=playwright-1.62.1-noble`. It deliberately exits 2 outside that environment. No normal command updates references.

Initial references are generated once by Initial Visual Baselines on the development branch, uploaded as an artifact, inspected and committed with provenance. The job refuses existing references. It never commits or approves its own output. Later visual changes require an explicit design-change request, image review and a dedicated reviewed update; the healer may not rebaseline.

Initial reference review: run 41977286948 at f026821cfec911955e7786fbed68ff6c2964fa78 produced 33 images and then passed 33 strict comparisons (11 per viewport, zero failed/skipped/flaky). All three scene overview sheets and the mobile tool/checklist originals were inspected. Existing internally scrollable tabs/tables and vertically scrollable dialogs remain intentional; no global horizontal overflow was accepted. Earlier failing images were discarded. The initializer is now manual-dispatch-only and refuses to overwrite the committed images; provenance records the generating commit, run and every PNG hash.

A later independent CI comparison exposed a fixture scroll race in the 6M scene: the dialog itself matched, but automatic button scrolling left the dimmed background at different offsets (desktop trace scrollTop 253 instead of the reference's bottom-aligned background). That scene now waits for the existing pending render-scroll restoration and verifies an explicit starting camera before the genuine click/tap. Desktop/phone references were bottom-aligned; the tablet reference trace recorded scrollY=308, so its camera remains at that fixed offset. The archived tablet trace and approved PNG have identical PNG hashes to the committed reference. No PNG, screenshot tolerance, target or functional assertion was changed. The existing functional scroll-preservation scenarios remain unchanged.

## Official Playwright agents

Generated using `npx playwright init-agents --loop=codex` with the installed 1.62.1 CLI (no version upgrade needed). Playwright is now exactly pinned. Planner, generator and healer live under .codex/agents. The parent .codex/config.toml configures the MCP server because current Codex agents inherit parent MCP settings. No model override or API key was added.

Use the named agents with specs/humbel-qa.md and tests/seed.spec.ts. The seed is an executable local mock login, not an empty placeholder. Official MCP initialization and planner/generator/healer tool discovery are tested by qa:agent-server. Agent model-driven planning/healing is on demand, not an autonomous CI service. Official generated instructions allowing assertion changes/fixme were replaced with mandatory failure reporting; do not regenerate without reapplying/reviewing these restrictions. Codex may need a new session to discover newly added project agents.

## Real Supabase integration: BLOCKED / NOT TESTED

Only a catalog SELECT was performed on the existing project. schema-inventory.json contains 31 tables, 308 columns, 127 policies, constraints and triggers; NO application rows or auth users were exported. This inventory is not a restorable SQL dump and does not include all grants, sequences, types or function bodies.

Local Docker, Supabase CLI and WSL are not installed. No separate active test project was identified or authorized for writes. The inactive second project is not treated as an approved test database. No real INSERT/UPDATE, RLS/auth/session, concurrency or Realtime integration tests were run. Do not call mocks real database tests.

Required before enabling real DB tests:

1. Install a supported local Docker runtime and official Supabase CLI.
2. Obtain a reviewed schema-only dump including grants, types, sequences, function bodies, triggers, extensions and RLS. Do not include production rows or credentials. Restore only to a disposable local stack.
3. Review migrations against the catalog inventory and apply them locally. No production RLS changes are authorized by this QA task.
4. Create separate local admin/department-admin/employee/scanner auth users and deterministic fixtures. Use server-controlled role claims and real JWT sessions; service-role writes do not prove RLS enforcement.
5. Test SELECT/INSERT/UPDATE and denials for every role, expired/logout sessions and Realtime events; execute parallel increments, duplicate clicks and journal/count consistency checks. The current client-side read/modify/write path is not proof of atomic multi-device counting.
6. Add the real DB suite as a required release job only after it exists and passes. Until then the Quality Gate explicitly covers mock/browser/visual integrity, not DB integration.

`test:db:preflight` fails closed with exit 2 and a missing-prerequisites report; it accepts only loopback URLs and never makes a network request or performs writes. Hosted databases are refused, including production. Several catalog tables have RLS disabled (planner_* and tool_materials); exposure depends on grants and other controls and has not been audited here. No policy was changed.

## Release gate and separate deployment approval

Current main protection was read on 06.11.2026: active `main-reviewed-e2e` ruleset 26664844, strict up-to-date PRs and required GitHub Actions check **e2e**, no bypass. Do not replace the required context with Quality Gate. Functional/Integrity/Visual must all succeed; Quality Gate then feeds the fail-closed required e2e job. Unit tests execute the actual gate scripts against failed, cancelled, skipped and missing dependencies.

All Issue #1 development/safety rules, templates, JSON/commit/version report and report tests are preserved alongside this QA suite. The merged main reload precondition and the separately asserted, fixed visual build metadata are exact reversible fixture patches: all legacy business assertions, identities and 33 PNG references remain unchanged.

Deployment no longer follows a main push automatically. It requires e2e success, a main workflow_dispatch, and deploy_approved=true (default false), only after explicit owner authorization. No workflow is dispatched and no deployment is performed by this integration task.

Before an actual merge, an administrator must verify the Pages source. A legacy branch-based Pages build can bypass this workflow; it must be disabled or switched to GitHub Actions before claiming controlled publication. The connector does not prove current Pages settings. No settings were mutated here. See docs/DEVELOPMENT.md for the exact integration sequence and limits.

## Limits

This substantially reduces routine regression work; it cannot guarantee zero manual validation. Live DB/RLS, real multi-device concurrency, external CDN failures, physical camera/scanner hardware, print devices and actual mobile Safari remain outside demonstrated coverage. Existing historical rows are not modified. App business logic is unchanged; version/log and the reviewed filter-width/header-selector CSS are the only app edits.

Sources: https://playwright.dev/docs/test-agents ; https://playwright.dev/docs/test-snapshots ; https://playwright.dev/docs/docker ; https://learn.chatgpt.com/docs/agent-configuration/subagents ; https://docs.github.com/en/pages/getting-started-with-github-pages/using-custom-workflows-with-github-pages
