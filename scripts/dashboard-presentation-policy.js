const assert = require("node:assert/strict");
const { hash } = require("./qa-policy");
const reviewed = require("../tests/qa/dashboard-presentation-patches.json");

function reverseDashboardPresentation(source) {
  for (const { before, after } of [...reviewed.patches].reverse()) {
    assert.equal(source.split(after).length - 1, 1,
      "Reviewed dashboard presentation hook must remain exact and unique");
    source = source.replace(after, before);
  }
  return source;
}

function verifyDashboardFiles(readFile) {
  for (const [file, expected] of Object.entries(reviewed.files)) {
    assert.equal(hash(readFile(file)), expected,
      `${file}: reviewed dashboard presentation source changed; explicit review required`);
  }
}

function reverseDashboardFile(file, source) {
  for (const change of [...reviewed.filePatches].reverse()) {
    if (file !== change.file) continue;
    assert.equal(source.split(change.after).length - 1, 1,
      "Reviewed dashboard asset hook must remain exact and unique");
    source = source.replace(change.after, change.before);
  }
  return source;
}

module.exports = { reverseDashboardPresentation, verifyDashboardFiles, reverseDashboardFile };
