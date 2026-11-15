const assert = require('node:assert/strict');
const { hash } = require('./qa-policy');
const reviewed = require('../tests/qa/design-foundation-patches.json');

function reverseDesignFile(file, source) {
  for (const patch of [...reviewed.patches].reverse()) {
    if (patch.file !== file) continue;
    assert.equal(source.split(patch.after).length - 1, 1,
      `${file}: design presentation change must be exact and unique`);
    source = source.replace(patch.after, patch.before);
  }
  return source;
}

function verifyDesignFiles(readFile) {
  for (const [file, expected] of Object.entries(reviewed.files)) {
    assert.equal(hash(readFile(file)), expected, `${file}: unreviewed design foundation change`);
  }
}
module.exports = { reverseDesignFile, verifyDesignFiles };
