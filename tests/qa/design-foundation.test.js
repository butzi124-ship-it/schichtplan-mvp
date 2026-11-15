const fs = require('node:fs');
const assert = require('node:assert/strict');
const { test } = require('node:test');
const { reverseDesignFile, verifyDesignFiles } = require('../../scripts/design-foundation-policy');
const { appBusiness, hash } = require('../../scripts/qa-policy');
const manifest = require('./design-foundation-patches.json');

test('design changes exactly restore the original app business logic, markup and dashboard CSS', () => {
  for (const file of new Set(manifest.patches.map(p => p.file))) {
    const restored = reverseDesignFile(file, fs.readFileSync(file, 'utf8'));
    assert.equal(hash(file.endsWith('.js') ? appBusiness(restored) : restored), manifest.baseFilesSha256[file]);
  }
});

test('design normalization rejects missing, repeated or changed deltas and keeps unrelated edits visible', () => {
  for (const file of new Set(manifest.patches.map(p => p.file))) {
    const source = fs.readFileSync(file, 'utf8');
    for (const patch of manifest.patches.filter(p => p.file === file)) {
      for (const value of ['', patch.after + patch.after, patch.before]) {
        assert.throws(() => reverseDesignFile(file, source.replace(patch.after, value)));
      }
    }
    assert.notEqual(reverseDesignFile(file, source + '\nUNREVIEWED'), reverseDesignFile(file, source));
  }
});

test('new shell assets are fingerprinted and cannot change silently', () => {
  verifyDesignFiles(file => fs.readFileSync(file));
  assert.deepEqual(Object.keys(manifest.files).sort(), ['css/app-shell.css', 'js/app-shell.js']);
  for (const changed of Object.keys(manifest.files)) {
    assert.throws(() => verifyDesignFiles(file => file === changed ? Buffer.from('changed') : fs.readFileSync(file)));
  }
});

test('shared shell has no backend, authority, persistence or credential dependency', () => {
  const source = fs.readFileSync('js/app-shell.js', 'utf8');
  assert(!/\b(?:fetch|XMLHttpRequest|supabaseClient|localStorage|sessionStorage|currentUser|currentEmployeeRecord|state)\b/.test(source));
  const css = fs.readFileSync('css/home-dashboard.css', 'utf8');
  assert(!css.includes('humbel-home-mode'));
  assert(!css.includes('#tabs'));
});

test('static packaging includes shared CSS as well as JS without any deployment', () => {
  const vm = require('node:vm');
  const copies = [];
  const fileSystem = { mkdirSync() {}, copyFileSync() {}, writeFileSync() {}, cpSync(from, to) { copies.push([from, to]); } };
  vm.runInNewContext(fs.readFileSync('scripts/stage-pages.js', 'utf8'), {
    require: name => name === 'fs' ? fileSystem : require(name), process: { env: {} },
  });
  assert(copies.some(([from, to]) => from === 'css' && to === '_site/css'));
});
