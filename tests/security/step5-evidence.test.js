const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const crypto = require('node:crypto');
const read = p => JSON.parse(fs.readFileSync(p, 'utf8'));
const hash = p => crypto.createHash('sha256').update(fs.readFileSync(p)).digest('hex');
const root = 'docs/qa/';

test('step5 evidence distinguishes endpoint success from blocked original-schema verification', () => {
  const e = read(root + 'security-step5-original-evidence.json');
  assert.equal(e.authoritativeRestore, false);
  assert.equal(e.productionDataRead, false);
  assert.equal(e.productionChanges, false);
  for (const endpoint of ['tables', 'extensions', 'migrations']) {
    const o = e.observations.find(x => x.endpoint === endpoint);
    assert.equal(o.status, 'BLOCKED_NOT_TESTED');
    assert.equal(o.schemaStatus, 'UNVERIFIED');
    assert.match(o.errorClass, /28P01/);
    assert(Number.isFinite(Date.parse(o.startedAt)));
    assert(Date.parse(o.endedAt) >= Date.parse(o.startedAt));
  }
  assert(!Object.hasOwn(e, 'tables'), 'Authentication failure must never become an empty verified catalog');
});

test('step5 enriches every original access without losing its identity or claiming complete integration', () => {
  const p = read(root + 'security-step5-access-plan.json');
  const original = read(p.sourceInventory);
  assert.equal(hash(p.sourceInventory), p.sourceInventorySha256);
  assert.equal(p.accesses.length, original.accesses.length);
  assert.equal(new Set(p.accesses.map(x => x.id)).size, original.accesses.length);
  original.accesses.forEach((old, i) => {
    for (const key of Object.keys(old)) assert.deepEqual(p.accesses[i][key], old[key]);
    assert.equal(p.accesses[i].serverEnforcementCurrent, 'UNVERIFIED');
    assert.equal(p.accesses[i].fullyIntegratedApplication, false);
    assert(p.accesses[i].requiredContract && p.accesses[i].requiredChange);
    assert.equal(typeof p.accesses[i].operationArgumentAtBaseline, 'string');
    assert(p.accesses[i].operationArgumentAtBaseline.length > 0);
  });
});

test('step5 deviation matrix covers historical tables and preserves exact SQL input provenance', () => {
  const m = read(root + 'security-step5-deviations.json');
  const oldPath = root + 'production-catalog-2026-11-08.json';
  assert.equal(hash(oldPath), m.historicalCatalogSha256);
  assert.equal(m.currentOriginalSchemaVerified, false);
  assert.equal(m.authoritativeRestore, false);
  const historical = m.objects.filter(x => x.historicalEvidence);
  assert.deepEqual(historical.map(x => x.object), read(oldPath).tables.map(x => x.name));
  for (const o of m.objects) assert.match(o.actualVerifiedDefinition, /UNVERIFIED/);
  for (const input of m.repositorySqlInputs) assert.equal(hash(input.path), input.sha256, input.path);
});

test('step5 prepares evidence without activating the isolated client or changing candidate SQL', () => {
  const html = fs.readFileSync('schichtplan_mvp.html', 'utf8');
  assert(!html.includes('js/secure-access.js'));
  const m = read(root + 'security-step5-deviations.json');
  for (const input of m.repositorySqlInputs.filter(x => /security-step[34]/.test(x.path))) {
    assert.match(fs.readFileSync(input.path, 'utf8'), /ISOLATED/);
    assert.equal(hash(input.path), input.sha256);
  }
});
