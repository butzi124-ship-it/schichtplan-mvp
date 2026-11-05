const fs = require("fs");
const path = require("path");

class QaReporter {
  onBegin(config, suite) { this.suite = suite; }
  onEnd(result) {
    const counts = {};
    const failures = [];
    let unsafe = false;
    for (const test of this.suite.allTests()) {
      const project = test.parent.project()?.name || "unknown";
      const outcome = test.outcome();
      counts[project] ||= { passed: 0, failed: 0, skipped: 0, flaky: 0 };
      const bad = outcome !== "expected" || test.expectedStatus !== "passed";
      const key = outcome === "expected" ? bad ? "failed" : "passed" : outcome === "unexpected" ? "failed" : outcome;
      counts[project][key]++;
      if (bad) unsafe = true;
      if (bad) {
        failures.push({ project, test: test.titlePath().slice(1).join(" > "), location: test.location,
          outcome, diagnosis: diagnose(test.results.flatMap(r => r.errors.map(e => e.message || String(e))).join("\n")), attempts: test.results.map(r => ({ status: r.status,
            errors: r.errors.map(e => e.message || String(e)),
            artifacts: r.attachments.filter(a => a.path).map(a => ({ name: a.name, path: a.path })) })) });
      }
    }
    const dir = process.env.QA_ARTIFACT_DIR || "artifacts/functional";
    fs.mkdirSync(dir, { recursive: true });
    const report = { sha: process.env.GITHUB_SHA || "local", status: unsafe ? "failed" : result.status, runnerStatus: result.status, counts, failures };
    fs.writeFileSync(path.join(dir, "summary.json"), JSON.stringify(report, null, 2));
    const lines = ["# QA Report", `Commit: ${report.sha}`, `Runner result: ${result.status}`, "",
      ...Object.entries(counts).map(([p, c]) => `- ${p}: ${c.passed} passed, ${c.failed} failed, ${c.skipped} skipped, ${c.flaky} flaky`),
      "", ...failures.flatMap(f => [`## ${f.project}: ${f.test}`, `Source: ${f.location.file}:${f.location.line}`, `Diagnostic hypothesis: ${f.diagnosis}`,
        ...f.attempts.flatMap(a => ["```text", ...a.errors.map(e => e.slice(0, 12000).replace(/```/g, "'''")), "```",
          ...a.artifacts.map(x => `- ${x.name}: ${x.path}`)]),
        "Diagnosis required: inspect trace, actual/expected/diff images and the first failing assertion. Do not weaken assertions or refresh references. Re-run the full suite after a fix."] )];
    fs.writeFileSync(path.join(dir, "summary.md"), lines.join("\n"));
    if (unsafe) return { status: "failed" };
  }
}
module.exports = QaReporter;

function diagnose(error) {
  if (/intercepts pointer|pointer events/i.test(error)) return "Hit target obstructed: compare elementFromPoint, bounds, overflow and layout/visual viewport before changing selectors. Fix layout, never force the click.";
  if (/toHaveScreenshot|snapshot|pixel/i.test(error)) return "Visual mismatch or missing reference: inspect expected/actual/diff in the pinned environment. Do not refresh references to hide a regression.";
  if (/timeout/i.test(error)) return "Locator or precondition timed out: inspect the earliest trace step and console/network evidence. Preserve the assertion and expected behavior.";
  return "Assertion/runtime failure: compare observed versus specified behavior using the attached trace. A business defect must remain failing until an authorized fix is verified.";
}
