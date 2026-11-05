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
