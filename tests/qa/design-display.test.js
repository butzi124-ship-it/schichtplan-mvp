const fs = require('node:fs'), cp = require('node:child_process');
const { test } = require('node:test'), assert = require('node:assert/strict');
const { reverseDisplayFile } = require('../../scripts/design-display-policy');
const { hash, appBusiness } = require('../../scripts/qa-policy');
const manifest = require('./design-display-patches.json');
test('display fixes restore the exact parent files apart from version metadata', () => {
  for (const file of new Set(manifest.patches.map(p => p.file))) {
    const restored = reverseDisplayFile(file, fs.readFileSync(file, 'utf8'));
    assert.equal(hash(file.endsWith('.js') && file === 'schichtplan_mvp.js' ? appBusiness(restored) : restored), manifest.baseFilesSha256[file]);
  }
});
test('display guards reject a missing or duplicated fix and keep unrelated changes visible', () => {
  for (const patch of manifest.patches) {
    const source = fs.readFileSync(patch.file, 'utf8');
    assert.throws(() => reverseDisplayFile(patch.file, source.replace(patch.after, '')));
    assert.throws(() => reverseDisplayFile(patch.file, source.replace(patch.after, patch.after + patch.after)));
    assert.notEqual(reverseDisplayFile(patch.file, source + '\nUNREVIEWED'), reverseDisplayFile(patch.file, source));
  }
});
test('proposal collector writes only review copies, not accepted reference images', () => {
  const vm = require('node:vm');
  const writes = [];
  vm.runInNewContext(fs.readFileSync('scripts/collect-design-proposals.js', 'utf8'), {
    require: name => name === 'node:fs' ? {
      readFileSync: () => 'const APP_VERSION = "0.19.1";', mkdirSync() {}, existsSync: () => false,
      writeFileSync: (p, data) => writes.push([p, data]),
    } : require(name), process: { env: {} }, console: { log() {} },
  });
  assert(writes.length >= 2);
  assert(writes.every(([p]) => p.startsWith('artifacts/visual/reference-proposals/')));
  const m = JSON.parse(writes.find(([p]) => p.endsWith('manifest.json'))[1]);
  assert.equal(m.approvalRequired, true);
  assert.equal(m.referencesChanged, false);
  assert.equal(m.missing.length, 33);
});
