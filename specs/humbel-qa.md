# Humbel QA Plan

Seed: tests/seed.spec.ts via playwright.agents.config.js. The seed uses the same mock data/client as the 251 legacy tests, with external network denied.

1. Planner: inventory login, modules, BA/article navigation, station selection and all production constraints. Save additive scenarios under specs/.
2. Generator: add focused tests using the mock fixture, not new implementations of app calculations. Use explicit expected quantities and genuine pointer/touch actions.
3. Healer: reproduce failure; inspect first failing assertion, console, screenshot, trace and hit testing; fix only equivalent selectors/synchronization. If the application is wrong, stop with diagnosis and a proposed app fix. Do not edit business assertions or reference images.

## Required Invariants

- One station retains existing behavior. Multiple steps use last-step good parts plus current scrap_total for remaining quantity; lifetime and unresolved clarification never clear multi-step remaining quantity.
- A step consumes at most predecessor good parts, across all employees. Earlier-step corrections cannot invalidate downstream counts.
- No negative quantities, double bookings or booking without a valid 6M cause. Recheck capacity when a cause dialog saves.
- Completion requires permission, running/paused status, stations, consistent quantities, complete checklist, no remaining quantity and explicit confirmation.
- Employee/scanner cannot gain production access; department admins remain limited to own departments.
- No writes during navigation, view switching or blocked completion. Never write production_counts in production-order tests.
- Keep real Chromium touch, scroll and animation checks; Firefox/WebKit run the critical production suites.

## Visual References

Login, production overview, touch view, article tabs, station tabs, good counters, scrap/clarification, 6M cause dialog, checklist, completion-ready dialog and tools: desktop 1280x720, phone 390x844 and tablet 1024x768. Generate and compare only in the pinned Linux CI image. No automatic rebaselining.

## Real DB Testing: Blocked

Catalog inventory exists, but is not a restorable migration. Docker/local Supabase, full schema/grants/functions/triggers/RLS, reviewed fixtures and isolated auth users are missing. Do not substitute mocks and call them real DB tests. See docs/qa/README.md.
