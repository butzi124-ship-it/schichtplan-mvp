const assert = require('node:assert/strict');
const reviewed = require('../tests/qa/design-display-patches.json');
function reverseDisplayFile(file, source) {
  source = require('./reference-production-policy').reverseReferenceFile(file, source);
  for (const patch of [...reviewed.patches].reverse()) {
    if (patch.file !== file) continue;
    assert.equal(source.split(patch.after).length - 1, 1,
      `${file}: display correction must be exact and unique`);
    source = source.replace(patch.after, patch.before);
  }
  return source;
}
module.exports = { reverseDisplayFile };
