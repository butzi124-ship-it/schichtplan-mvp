const fs = require("fs");
const path = require("path");
const assert = require("node:assert/strict");
const YAML = require("yaml");
const { listTests } = require("./list-tests");
const { hash, appBusiness, sourceViolations } = require("./qa-policy");
const protectedFiles = require("../tests/qa/protected-files.json");
let app = fs.readFileSync("schichtplan_mvp.js", "utf8").replace(/\r\n/g, "\n");
for (const {before, after} of [...require("../tests/qa/quantity-integration-patches.json").patches].reverse()) {
  assert.equal(app.split(after).length - 1, 1, "Quantity hook must remain exact and unique");
  app = app.replace(after, before);
}
for (const {before, after} of [...require("../tests/qa/setup-integration-patches.json").patches].reverse()) {
  assert.equal(app.split(after).length - 1, 1, "Setup hook must remain exact and unique");
  app = app.replace(after, before);
}
for (const {before, after} of [...require("../tests/qa/ba-entry-integration-patches.json").patches].reverse()) {
  assert.equal(app.split(after).length - 1, 1, "BA entry hook must remain exact and unique");
  app = app.replace(after, before);
}
for (const {before, after} of [...require("../tests/qa/forecast-integration-patches.json").patches].reverse().concat([...require("../tests/qa/readiness-integration-patches.json").patches].reverse()).concat([...require("../tests/qa/operations-integration-patches.json").patches].reverse()).concat(require("../tests/qa/article-integration-patches.json").patches)) {
  assert.equal(app.split(after).length - 1, 1, "Article integration hook must remain exact and unique");
  app = app.replace(after, before);
}
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
  let checkedSource = fs.readFileSync(file, "utf8").replace(/\r\n/g, "\n");
  for (const navigation of [require("../tests/qa/legacy-navigation-patch.json"), require("../tests/qa/visual-metadata-patch.json")]) {
    if (file !== navigation.file) continue;
    assert.equal(checkedSource.split(navigation.after).length - 1, 1, "Explicit test fixture patch must remain exact");
    checkedSource = checkedSource.replace(navigation.after, navigation.before);
  }
  for (const change of [...require("../tests/qa/quantity-integration-patches.json").filePatches].reverse()) {
    if (change.file !== file) continue;
    assert.equal(checkedSource.split(change.after).length - 1, 1, "Explicit quantity file adapter must remain exact and unique");
    checkedSource = checkedSource.replace(change.after, change.before);
  }
  assert.equal(hash(checkedSource), expected, `${file}: protected regression source changed; explicit human review required`);
  assert.deepEqual(sourceViolations(fs.readFileSync(file, "utf8")), [], `${file}: forbidden skip/force/failure annotation`);
}
const actual = new Set(listTests().map(t => `${t.project}:${t.title}`));
const legacy = require("../tests/qa/legacy-tests.json").tests;
assert.equal(legacy.length, 251);
for (const test of legacy) assert(actual.has(`${test.project}:${test.title}`), `Missing legacy test: ${test.project} ${test.title}`);
const workflow = YAML.parse(fs.readFileSync(".github/workflows/e2e.yml", "utf8"));
assert.deepEqual(workflow.jobs["quality-gate"].needs, ["functional", "integrity", "visual"]);
assert.equal(workflow.jobs["quality-gate"].if, "always()");
assert.equal(workflow.jobs.deploy.needs, "e2e");
assert.equal(workflow.jobs.e2e.needs, "quality-gate");
assert.equal(workflow.jobs.e2e.if, "always()");
assert.equal(workflow.jobs.deploy.if, "github.event_name == 'workflow_dispatch' && github.ref == 'refs/heads/main' && inputs.deploy_approved == true");
assert.equal(workflow.jobs.visual.container, "mcr.microsoft.com/playwright:v1.62.1-noble");
assert(!fs.readFileSync(".github/workflows/e2e.yml", "utf8").includes("--update-snapshots"), "Normal CI must never update snapshots");
const initializer = YAML.parse(fs.readFileSync(".github/workflows/visual-baselines.yml", "utf8"));
assert.deepEqual(Object.keys(initializer.on), ["workflow_dispatch"], "Reference initialization must require explicit dispatch");
const root = "tests/visual-snapshots";
assert(fs.existsSync(path.join(root, "provenance.json")), "BLOCKED: initial Linux references must be generated, reviewed and committed");
const provenance = JSON.parse(fs.readFileSync(path.join(root, "provenance.json"), "utf8"));
assert.equal(provenance.files.length, 33);
assert.equal(provenance.image, "mcr.microsoft.com/playwright:v1.62.1-noble");
for (const ref of provenance.files) assert.equal(hash(fs.readFileSync(path.join(root, ref.file))), ref.sha256, `Unreviewed reference change: ${ref.file}`);
console.log(`QA integrity passed: ${legacy.length} legacy identities retained, legacy business logic preserved outside explicit article hooks, strict release chain and 33 versioned references verified.`);
