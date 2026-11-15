const {test, mock} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {PGlite} = require('@electric-sql/pglite');
const {clockSQL, transformSQL, install, prepare, patches} = require('../db-integration/forecast-clock');
const clock = require('../fixtures/forecast-clock.json');

// All SQL except clock reads must survive the adapter byte-for-byte, including
// publication/expiry checks and strings that happen to contain clock keywords.
test('native clock adapter only substitutes clock reads, including stored SQL bodies', () => {
  const source = "select now(), pg_catalog.clock_timestamp(), CURRENT_DATE, current_timestamp, transaction_timestamp(), statement_timestamp(); -- now() current_date\nselect 'now() current_date', \"current_date\", my_schema.now(); /* clock_timestamp() */";
  const changed = transformSQL(source);
  assert.equal(changed, "select humbel_test_clock.now(), humbel_test_clock.clock_timestamp(), humbel_test_clock.today(), humbel_test_clock.now(), humbel_test_clock.now(), humbel_test_clock.statement_timestamp(); -- now() current_date\nselect 'now() current_date', \"current_date\", my_schema.now(); /* clock_timestamp() */");
  const sql = fs.readFileSync('supabase/review/issue-10-production-forecast.sql', 'utf8');
  assert.equal(transformSQL(sql).replaceAll('humbel_test_clock.now()', 'now()').replaceAll('humbel_test_clock.clock_timestamp()', 'clock_timestamp()').replaceAll('humbel_test_clock.statement_timestamp()', 'statement_timestamp()'), sql);
  for (const file of ['supabase/migrations/20261106023103_article_master_phase1.sql','supabase/migrations/20261106053328_operations_phase2_review.sql',...['issue-9-machine-readiness.sql','ba-entry-phase1.sql','op-routing-phase2.sql','setup-phase3.sql','quantity-phase4.sql'].map(f=>'supabase/review/'+f)]) {
    const source = fs.readFileSync(file, 'utf8');
    assert.equal(transformSQL(source).replaceAll("humbel_test_clock.today()", 'current_date').replaceAll('humbel_test_clock.now()', 'now()').replaceAll('humbel_test_clock.clock_timestamp()', 'clock_timestamp()').replaceAll('humbel_test_clock.statement_timestamp()', 'statement_timestamp()'), source);
  }
});

test('native SQL clock is shared across connections, progresses and keeps UTC/Berlin days and transaction time coherent', async () => {
  mock.timers.enable({apis:['Date'], now:Date.parse('2031-01-01T23:45:00Z')});
  const db = new PGlite();
  // Thin pg.Client-shaped adapters exercise the actual prototype clock hook.
  // These are PGlite contracts, not a claim of native JWT/PostgREST coverage.
  const config = {isolated:true,host:'127.0.0.1',port:55432,database:'humbel_isolated_test',sentinel:'a'.repeat(64)};
  class Connection {
    async query(text, values) {
      if (text.includes('current_database() name')) return {rows:[{name:config.database,sentinel:config.sentinel}]};
      if (text === clockSQL) return db.exec(text);
      return db.query(text, values);
    }
  }
  try {
    await db.exec('create role anon; create role authenticated; create role service_role;');
    const first = new Connection(), second = new Connection();
    const meta = await install(first, config);
    assert.equal(meta.base, clock.instant);
    const read = async c => (await c.query("select now() t,clock_timestamp() wall,current_date::text as day")).rows[0];
    for (const zone of ['UTC', 'America/Los_Angeles', 'Pacific/Kiritimati']) {
      await db.query("select set_config('TimeZone',$1,false)", [zone]);
      const row = await read(first);
      assert.equal(row.t.toISOString(), new Date(clock.instant).toISOString());
      assert.equal(row.wall.getTime(), row.t.getTime());
      assert.equal(row.day, '2026-11-07');
    }
    mock.timers.tick(1000);
    assert.equal((await read(second)).t.getTime(), Date.parse(clock.instant) + 1000);
    await db.exec('begin');
    const start = await read(first);
    mock.timers.tick(3000);
    const later = await read(second);
    assert.equal(later.t.getTime(), start.t.getTime());
    assert.equal(later.wall.getTime() - start.wall.getTime(), 3000);
    await db.exec('commit');
    assert.equal((await read(second)).t.getTime(), Date.parse(clock.instant) + 4000);
    await db.exec('set role authenticated');
    assert.equal((await read(first)).day, '2026-11-07');
    await assert.rejects(db.query('select * from humbel_test_clock.anchor'), e => e.code === '42501');
    await db.exec('reset role');
  } finally {
    try { await db.close(); } finally { mock.timers.reset(); }
  }
});

test('native clock refuses unverified database targets before installing anything', async () => {
  let writes = 0;
  const good = {isolated:true,host:'127.0.0.1',port:55432,database:'humbel_isolated_test',sentinel:'a'.repeat(64)};
  for (const change of [{isolated:false},{host:'example.supabase.co'},{port:5432},{database:'postgres'},{sentinel:''},{}]) {
    const db = {query:async () => {writes++;return {rows:[{name:'postgres',sentinel:'wrong'}]};}};
    await assert.rejects(install(db, {...good, ...change}));
  }
  assert.equal(writes, 1); // only the final mismatched actual-target read is allowed
});

test('native source preparation retains every assertion and rejects unknown pinned source', () => {
  // The actual pinned files are checked separately in a disposable checkout.
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'humbel-clock-guard-'));
  try {
    for (const patch of patches) {
      const file = path.join(root, patch.file);
      fs.mkdirSync(path.dirname(file), {recursive:true});fs.writeFileSync(file, patch.before);
    }
    assert.throws(() => prepare(root), /Pinned native source changed/);
    for (const patch of patches) assert.equal(fs.readFileSync(path.join(root, patch.file), 'utf8'), patch.before);
    const workflow = fs.readFileSync('.github/workflows/ba-entry-isolated.yml', 'utf8');
    assert.ok(workflow.includes('ref: 25e9a6914e38fc791d9ecad13d92a45e30d19bb3'));
    assert.ok(workflow.includes('node tests/db-integration/forecast-clock.js prepare'));
    assert.ok(workflow.includes('docker exec humbel-gateway node scripts/db-integration/run.js'));
  } finally {fs.rmSync(root, {recursive:true,force:true});}
});
