const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { localTestUrl, prerequisites } = require("../../scripts/db-preflight");
const { sourceViolations, appBusiness } = require("../../scripts/qa-policy");
const Reporter = require("../../scripts/qa-reporter");

test("DB preflight rejects production, hosted, credentialed and malformed targets", () => {
  for (const url of ["https://itntnpzqzyjfdnjdshvm.supabase.co", "https://other.supabase.co", "http://localhost.evil.test", "file:///tmp/db", "http://user:pass@localhost", "", "invalid"]) assert.equal(localTestUrl(url), false);
  for (const url of ["http://127.0.0.1:54321", "http://localhost:54321", "http://[::1]:54321"]) assert.equal(localTestUrl(url), true);
  assert(prerequisites({}).length > 0);
});
test("test policy rejects disabled, focused, expected-failure and forced actions", () => {
  for (const source of ["test.skip('x')", "test.fixme('x')", "test.only('x')", "test.fail(true)", "button.tap({ force: true })"]) assert(sourceViolations(source).length > 0);
  assert.deepEqual(sourceViolations("await button.tap(); await expect(button).toBeDisabled();"), []);
});
test("app normalization permits only version/log differences, not business edits", () => {
  const a = 'const APP_VERSION = "0.5.16";\nconst VERSION_LOG = [\n{version: "a"}\n];\nfunction count(){return 5;}';
  const b = a.replace("0.5.16", "0.5.17").replace('version: "a"', 'version: "b"');
  assert.equal(appBusiness(a), appBusiness(b));
  assert.notEqual(appBusiness(a), appBusiness(b.replace("return 5", "return 6")));
});
for (const outcome of ["skipped", "unexpected", "flaky"]) {
  test(`reporter fails closed on ${outcome} and records the named failing test`, () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "humbel-qa-"));
    const previous = process.env.QA_ARTIFACT_DIR;
    process.env.QA_ARTIFACT_DIR = dir;
    try {
      const reporter = new Reporter();
      reporter.onBegin({}, { allTests: () => [{ parent: { project: () => ({ name: "fixture" }) },
        outcome: () => outcome, expectedStatus: "passed", titlePath: () => ["file", "blocked completion"], location: { file: "fixture.js", line: 1 },
        results: [{ status: "failed", errors: [{ message: "expected disabled" }], attachments: [] }] }] });
      assert.deepEqual(reporter.onEnd({ status: "passed" }), { status: "failed" });
      const report = JSON.parse(fs.readFileSync(path.join(dir, "summary.json"), "utf8"));
      assert.equal(report.failures[0].test, "blocked completion");
      assert.equal(report.status, "failed");
    } finally {
      if (previous === undefined) delete process.env.QA_ARTIFACT_DIR; else process.env.QA_ARTIFACT_DIR = previous;
      assert(path.resolve(dir).startsWith(path.resolve(os.tmpdir()) + path.sep));
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
}
