const assert = require('node:assert/strict');
const reviewed = require('../tests/qa/reference-production-patches.json');
function reverseReferenceFile(file, source) {
  for (const patch of [...reviewed.patches].reverse()) {
    if (patch.file !== file) continue;
    assert.equal(source.split(patch.after).length - 1, 1, `${file}: reference presentation delta must remain exact and unique`);
    source = source.replace(patch.after, patch.before);
  }
  return source;
}
module.exports = { reverseReferenceFile };
