const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), os = require('node:os'), path = require('node:path'), cp = require('node:child_process');
const { PNG } = require('pngjs');
const collector = path.resolve('scripts/collect-design-proposals.js');
const scenes = ['login', 'production-overview', 'touch-production', 'article-tabs', 'station-tabs', 'good-counter', 'scrap-clarification', '6m-dialog', 'mandatory-checklist', 'completion-ready', 'tool-management'];
const projects = ['visual-desktop', 'visual-tablet', 'visual-mobile'];
const output = 'artifacts/visual/reference-proposals';
function fixture() {
  const cwd = fs.mkdtempSync(path.join(os.tmpdir(), 'fabos-capture-'));
  const png = new PNG({ width: 2, height: 2 }); png.data.fill(255);
  const bytes = PNG.sync.write(png);
  fs.writeFileSync(path.join(cwd, 'schichtplan_mvp.js'), 'const APP_VERSION = "0.20.10";');
  for (const project of projects) for (const scene of scenes) {
    const dir = path.join(cwd, 'test-results/visual', scene + '-' + project);
    fs.mkdirSync(dir, { recursive: true });
    for (const kind of ['actual', 'diff']) fs.writeFileSync(path.join(dir, scene + '-' + kind + '.png'), bytes);
    const ref = path.join(cwd, 'tests/visual-snapshots', project, scene + '.png');
    fs.mkdirSync(path.dirname(ref), { recursive: true }); fs.writeFileSync(ref, bytes);
  }
  const baseline = () => projects.flatMap(p => scenes.map(s => fs.readFileSync(path.join(cwd, 'tests/visual-snapshots', p, s + '.png')).toString('hex')));
  const original = baseline();
  return { cwd, bytes, baseline, original, cleanup: () => fs.rmSync(cwd, { recursive: true, force: true }) };
}
function collect(f, run) {
  return cp.spawnSync(process.execPath, [collector], { cwd: f.cwd, encoding: 'utf8', env: { ...process.env, DESIGN_CAPTURE_RUN: run, GITHUB_SHA: 'same-reviewed-checkout' } });
}
function json(f, file) { return JSON.parse(fs.readFileSync(path.join(f.cwd, output, file), 'utf8')); }
test('both 33-image runs, references and diffs survive output replacement and compare identical pixels', () => {
  const f = fixture();
  try {
    assert.equal(collect(f, 'first').status, 0);
    // Playwright replaces this directory on its next run; preserved copies must survive.
    const first = fs.readFileSync(path.join(f.cwd, output, 'first/visual-desktop/login.png'));
    fs.rmSync(path.join(f.cwd, 'test-results/visual'), { recursive: true });
    for (const project of projects) for (const scene of scenes) {
      const dir = path.join(f.cwd, 'test-results/visual', scene + '-' + project);fs.mkdirSync(dir, { recursive: true });
      for (const kind of ['actual', 'diff']) fs.writeFileSync(path.join(dir, scene + '-' + kind + '.png'), f.bytes);
    }
    assert.equal(collect(f, 'second').status, 0);
    const result = json(f, 'stability.json');
    assert.equal(result.status, 'STABLE_33_OF_33'); assert.equal(result.comparisons.length, 33);
    assert(result.comparisons.every(x => x.identicalPixels && x.referenceUnchanged));
    assert.deepEqual(fs.readFileSync(path.join(f.cwd, output, 'first/visual-desktop/login.png')), first);
    for (const run of ['first', 'second']) {
      const m = json(f, run + '/manifest.json'); assert.equal(m.proposals.length, 33); assert.equal(m.missing.length, 0);
      assert(m.approvalRequired); assert.equal(m.referencesChanged, false);
      for (const p of m.proposals) for (const file of [p.file, p.reference, p.diff]) assert(fs.existsSync(path.join(f.cwd, output, run, file)));
    }
    const html = fs.readFileSync(path.join(f.cwd, output, 'index.html'), 'utf8');
    assert.equal((html.match(/<section>/g) || []).length, 33);
    assert(html.includes('Bisherige Referenz'));assert(html.includes('Differenz zur Referenz'));
    assert.deepEqual(f.baseline(), f.original);
  } finally { f.cleanup(); }
});
test('one changed capture fails stability instead of accepting a new reference', () => {
  const f = fixture();
  try {
    assert.equal(collect(f, 'first').status, 0);
    const png = PNG.sync.read(f.bytes);png.data[0] = 0;
    fs.writeFileSync(path.join(f.cwd, 'test-results/visual/login-visual-desktop/login-actual.png'), PNG.sync.write(png));
    assert.equal(collect(f, 'second').status, 1);
    const result = json(f, 'stability.json');assert.equal(result.status, 'UNSTABLE_OR_INCOMPLETE');
    assert.equal(result.comparisons.filter(x => !x.identicalPixels).length, 1);
    assert.deepEqual(f.baseline(), f.original);
  } finally { f.cleanup(); }
});
test('missing or ambiguous screenshots and missing comparison evidence fail closed', () => {
  const f = fixture();
  try {
    const duplicate = path.join(f.cwd, 'test-results/visual/duplicate-visual-desktop');fs.mkdirSync(duplicate);
    fs.writeFileSync(path.join(duplicate, 'login-actual.png'), f.bytes);
    fs.rmSync(path.join(f.cwd, 'test-results/visual/good-counter-visual-tablet/good-counter-actual.png'));
    const different = PNG.sync.read(f.bytes);different.data[0] = 0;
    fs.writeFileSync(path.join(f.cwd, 'test-results/visual/station-tabs-visual-mobile/station-tabs-actual.png'), PNG.sync.write(different));
    fs.rmSync(path.join(f.cwd, 'test-results/visual/station-tabs-visual-mobile/station-tabs-diff.png'));
    fs.rmSync(path.join(f.cwd, 'tests/visual-snapshots/visual-mobile/article-tabs.png'));
    assert.equal(collect(f, 'first').status, 1);
    const m = json(f, 'first/manifest.json');assert.equal(m.proposals.length, 31);assert.equal(m.missing.length, 4);
    assert(m.missing.includes('visual-desktop/login.png'));
    assert(m.missing.includes('visual-mobile/article-tabs-reference.png'));
    assert.equal(collect(f, 'second').status, 1);assert.equal(json(f, 'stability.json').status, 'UNSTABLE_OR_INCOMPLETE');
  } finally { f.cleanup(); }
});
test('successful real captures need no diff only when their decoded pixels equal the reference', () => {
  const f = fixture();
  try {
    for (const project of projects) for (const scene of scenes) fs.rmSync(path.join(f.cwd, 'test-results/visual', scene + '-' + project, scene + '-diff.png'));
    assert.equal(collect(f, 'first').status, 0);assert.equal(collect(f, 'second').status, 0);
    const m = json(f, 'second/manifest.json');assert.equal(m.proposals.length, 33);assert.deepEqual(m.missing, []);
    assert(m.proposals.every(p => p.identicalToReference === true && p.diff === null));
    assert.equal(json(f, 'stability.json').status, 'STABLE_33_OF_33');
    const html = fs.readFileSync(path.join(f.cwd, output, 'index.html'), 'utf8');
    assert(!html.includes('second/null'));assert(!html.includes('second/undefined'));
    assert.deepEqual(f.baseline(), f.original);
  } finally { f.cleanup(); }
});
test('missing diff for a genuine pixel mismatch remains an error even with a current capture', () => {
  const f = fixture();
  try {
    const png = PNG.sync.read(f.bytes);png.data[0] = 0;
    fs.writeFileSync(path.join(f.cwd, 'test-results/visual/login-visual-desktop/login-actual.png'), PNG.sync.write(png));
    fs.rmSync(path.join(f.cwd, 'test-results/visual/login-visual-desktop/login-diff.png'));
    assert.equal(collect(f, 'first').status, 1);
    assert(json(f, 'first/manifest.json').missing.includes('visual-desktop/login-diff.png'));
    assert.deepEqual(f.baseline(), f.original);
  } finally { f.cleanup(); }
});
