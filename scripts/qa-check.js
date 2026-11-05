const fs = require("fs");
const path = require("path");
const assert = require("node:assert/strict");
const YAML = require("yaml");
const { listTests } = require("./list-tests");
const { hash, appBusiness, sourceViolations } = require("./qa-policy");
const protectedFiles = require("../tests/qa/protected-files.json");
let app = fs.readFileSync("schichtplan_mvp.js", "utf8").replace(/\r\n/g, "\n");
if (protectedFiles.reviewedPresentationPatch) {
  assert(app.includes(protectedFiles.reviewedPresentationPatch), "Reviewed CSS patch must remain exact");
  app = app.replace(protectedFiles.reviewedPresentationPatch, "");
}
for (const replacement of protectedFiles.reviewedPresentationReplacements || []) {
  assert.equal(app.split(replacement.after).length - 1, 1, "Reviewed CSS replacement must remain exact and unique");
  app = app.replace(replacement.after, replacement.before);
}
assert.equal(hash(appBusiness(app)), protectedFiles.appBusinessSha256, "App business logic changed in a QA-only release");
for (const [file, expected] of Object.entries(protectedFiles.tests)) {
  assert.equal(hash(fs.readFileSync(file, "utf8").replace(/\r\n/g, "\n")), expected, `${file}: protected regression source changed; explicit human review required`);
  assert.deepEqual(sourceViolations(fs.readFileSync(file, "utf8")), [], `${file}: forbidden skip/force/failure annotation`);
}
const actual = new Set(listTests().map(t => `${t.project}:${t.title}`));
const legacy = require("../tests/qa/legacy-tests.json").tests;
assert.equal(legacy.length, 251);
for (const test of legacy) assert(actual.has(`${test.project}:${test.title}`), `Missing legacy test: ${test.project} ${test.title}`);
const workflow = YAML.parse(fs.readFileSync(".github/workflows/e2e.yml", "utf8"));
assert.deepEqual(workflow.jobs["quality-gate"].needs, ["functional", "integrity", "visual"]);
assert.equal(workflow.jobs["quality-gate"].if, "always()");
assert.equal(workflow.jobs.deploy.needs, "quality-gate");
assert.equal(workflow.jobs.deploy.if, "github.event_name == 'push' && github.ref == 'refs/heads/main'");
assert.equal(workflow.jobs.visual.container, "mcr.microsoft.com/playwright:v1.62.1-noble");
assert(!fs.readFileSync(".github/workflows/e2e.yml", "utf8").includes("--update-snapshots"), "Normal CI must never update snapshots");
const root = "tests/visual-snapshots";
assert(fs.existsSync(path.join(root, "provenance.json")), "BLOCKED: initial Linux references must be generated, reviewed and committed");
const provenance = JSON.parse(fs.readFileSync(path.join(root, "provenance.json"), "utf8"));
assert.equal(provenance.files.length, 33);
assert.equal(provenance.image, "mcr.microsoft.com/playwright:v1.62.1-noble");
for (const ref of provenance.files) assert.equal(hash(fs.readFileSync(path.join(root, ref.file))), ref.sha256, `Unreviewed reference change: ${ref.file}`);
console.log(`QA integrity passed: ${legacy.length} legacy identities retained, business logic unchanged, strict release chain and 33 versioned references verified.`);
