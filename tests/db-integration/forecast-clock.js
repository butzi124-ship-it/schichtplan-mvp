// Test harness only. Never imported by application SQL, HTML or Edge entrypoints.
// The pinned native suite retains its assertions; only its clock sources are adapted.
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const clock = require('../fixtures/forecast-clock.json');
const schema = 'humbel_test_clock';

// Unlike a process-local Date mock, stored SQL functions also govern PostgREST and
// the as_of/observed_at inputs consumed by the real Edge runtime. JWT clocks stay real.
// Elapsed real time preserves ordering, transaction time and expiry semantics.
const clockSQL = `
 create schema ${schema};
 revoke all on schema ${schema} from public;
 grant usage on schema ${schema} to anon, authenticated, service_role;
 create table ${schema}.anchor(real_at timestamptz not null, test_at timestamptz not null);
 insert into ${schema}.anchor values(pg_catalog.clock_timestamp(),'${clock.instant}');
 revoke all on ${schema}.anchor from public, anon, authenticated, service_role;
 ${['now','clock_timestamp','statement_timestamp'].map(name => `
 create function ${schema}.${name}() returns timestamptz language sql ${name === 'clock_timestamp' ? 'volatile' : 'stable'} security definer set search_path='' as $$
 select test_at + (pg_catalog.${name}() - real_at) from ${schema}.anchor
 $$;
 revoke all on function ${schema}.${name}() from public;
 grant execute on function ${schema}.${name}() to anon, authenticated, service_role;
 `).join('')}
 create function ${schema}.today() returns date language sql stable security definer set search_path='' as $$
 select (${schema}.now() at time zone '${clock.zone}')::date
 $$;
 revoke all on function ${schema}.today() from public;
 grant execute on function ${schema}.today() to anon, authenticated, service_role;
`;

function transformSQL(text) {
  // Process dollar-quoted function bodies, but retain quoted literals/identifiers
  // and comments byte-for-byte. No expressions other than clock reads change.
  return text.replace(/'(?:''|[^'])*'|"(?:""|[^"])*"|--[^\n]*|\/\*[\s\S]*?\*\/|(?<![\w.])(?:pg_catalog\.)?(now|clock_timestamp|transaction_timestamp|statement_timestamp)\s*\(\s*\)|\b(current_date|current_timestamp)\b/gi,
    (token, fn, keyword) => fn ? `${schema}.${fn.toLowerCase() === 'transaction_timestamp' ? 'now' : fn.toLowerCase()}()`
      : keyword ? keyword.toLowerCase() === 'current_date' ? `${schema}.today()` : `${schema}.now()`
      : token);
}

async function install(db, config) {
  assert.equal(config.isolated, true);
  assert.equal(config.host, '127.0.0.1');
  assert.equal(config.port, 55432);
  assert.equal(config.database, 'humbel_isolated_test');
  assert.match(config.sentinel, /^[a-f0-9]{64}$/);
  const target = (await db.query("select current_database() name,current_setting('humbel.isolated_sentinel',true) sentinel")).rows[0];
  assert.equal(target.name, config.database);
  assert.equal(target.sentinel, config.sentinel);
  await db.query(clockSQL);
  // Every independent pg.Client in this test process shares this prototype. RPC
  // connections execute the clock references already installed in stored SQL.
  const proto = Object.getPrototypeOf(db), original = proto.query;
  proto.query = function (query, ...args) {
    if (typeof query === 'string') query = transformSQL(query);
    else if (query?.text) query = {...query, text: transformSQL(query.text)};
    return original.call(this, query, ...args);
  };
  console.log(`Isolated application SQL clock: ${clock.instant}, ${clock.zone}; real JWT clocks unchanged.`);
  return {base: clock.instant, zone: clock.zone, mode: 'anchored-elapsed', sqlClockReferencesOnly: true, jwtClock: 'real'};
}

// Fail closed if the reviewed infrastructure changes. Do not vendor or replace its
// tests. Apply two exact clock-only patches to its disposable checkout before run.
const patches = [
  {file: 'scripts/db-integration/run.js', sha256: '0015fd8cedad4437ff614aa3ec2abba861f6eed6549bbaa512875535dd65d052',
   before: ' let expected=31;',
   after: " results.testClock=await require('../../tests/db-integration/forecast-clock').install(db,config());\n let expected=31;"},
  {file: 'scripts/db-integration/seed.js', sha256: '4a13640baa1d0a10599d0ce338a74324fe78367cc00c995cf8e863f4050cb56f',
   before: 'new Date(Date.now()-10000).toISOString()',
   after: "new Date((await db.query('select clock_timestamp() as observed')).rows[0].observed.getTime()-10000).toISOString()"},
];
function prepare(root) {
  const prepared = patches.map(patch => {
    const file = path.join(root, patch.file), source = fs.readFileSync(file, 'utf8');
    assert.equal(crypto.createHash('sha256').update(source).digest('hex'), patch.sha256, 'Pinned native source changed: ' + patch.file);
    assert.equal(source.split(patch.before).length - 1, 1, 'Native clock patch must be unique');
    const changed = source.replace(patch.before, patch.after);
    assert.equal(changed.replace(patch.after, patch.before), source);
    return {file, changed};
  });
  for (const {file, changed} of prepared) fs.writeFileSync(file, changed);
}
module.exports = {clockSQL, transformSQL, install, prepare, patches};
if (require.main === module) {
  assert.equal(process.argv[2], 'prepare');
  prepare(process.cwd());
}
