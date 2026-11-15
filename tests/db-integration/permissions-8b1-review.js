// Review/test orchestrator only. No network, connection strings or application imports.
const fs = require('node:fs');
const crypto = require('node:crypto');
const assert = require('node:assert/strict');
const baseline = require('./security-step7-review');
const prior = require('./security-step6-review');
const file = 'supabase/review/permissions-8b1-isolated.sql';
const hash = value => crypto.createHash('sha256').update(value).digest('hex');
const manifest = () => JSON.parse(fs.readFileSync('tests/db-integration/permissions-8b1-package.json', 'utf8'));
function source() {
 const s = fs.readFileSync(file, 'utf8');
 assert.equal(hash(s), manifest().inputSha256, 'Unreviewed 8B1 source');
 assert.equal((s.match(/^begin;$/gm)||[]).length, 1);
 assert.equal((s.match(/^commit;$/gm)||[]).length, 1);
 return s.replace(/^begin;\n/m, '').replace(/^commit;\s*$/m, '');
}
async function snapshot(db) {
 const queries = [
  "select n.nspowner,n.nspacl::text from pg_namespace n where n.nspname='humbel_permissions'",
  "select c.relname,c.relkind,c.relowner,c.relacl::text,c.relrowsecurity,c.relforcerowsecurity from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname='humbel_permissions' order by c.relname",
  "select c.relname,a.attname,a.atttypid::regtype::text,a.attnotnull,a.attacl::text,pg_get_expr(d.adbin,d.adrelid) def from pg_attribute a join pg_class c on c.oid=a.attrelid join pg_namespace n on n.oid=c.relnamespace left join pg_attrdef d on d.adrelid=a.attrelid and d.adnum=a.attnum where n.nspname='humbel_permissions' and a.attnum>0 and not a.attisdropped order by c.relname,a.attnum",
  "select c.relname,k.conname,pg_get_constraintdef(k.oid) def,k.convalidated from pg_constraint k join pg_class c on c.oid=k.conrelid join pg_namespace n on n.oid=c.relnamespace where n.nspname='humbel_permissions' order by 1,2",
  "select p.proname,p.proowner,p.proacl::text,p.prosecdef,p.proconfig::text,pg_get_functiondef(p.oid) def from pg_proc p join pg_namespace n on n.oid=p.pronamespace where n.nspname='humbel_permissions' order by 1,p.oid",
  "select * from pg_policies where schemaname='humbel_permissions' order by tablename,policyname",
  "select c.relname,t.tgname,pg_get_triggerdef(t.oid) def from pg_trigger t join pg_class c on c.oid=t.tgrelid join pg_namespace n on n.oid=c.relnamespace where n.nspname='humbel_permissions' and not t.tgisinternal order by 1,2",
  "select code from humbel_permissions.roles order by code"
 ];
 const rows=[]; for (const q of queries) rows.push((await db.query(q)).rows);
 return hash(JSON.stringify(rows));
}
async function postconditions(db) {
 assert.deepEqual((await db.query('select * from humbel_permissions.runtime')).rows, [{singleton:true,enabled:false}]);
 assert.deepEqual((await db.query('select code from humbel_permissions.roles order by code')).rows.map(x=>x.code), manifest().roles.slice().sort());
 const tables=(await db.query("select c.relname,c.relrowsecurity from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname='humbel_permissions' and c.relkind='r' order by 1")).rows;
 assert.equal(tables.length,10);
 for (const t of tables) {
  assert.equal(t.relrowsecurity,true,t.relname);
  for (const role of ['anon','authenticated','service_role']) for (const p of ['SELECT','INSERT','UPDATE','DELETE','TRUNCATE','REFERENCES','TRIGGER','MAINTAIN']) {
   assert.equal((await db.query('select has_table_privilege($1,$2,$3) v',[role,'humbel_permissions.'+t.relname,p])).rows[0].v,false,t.relname+'/'+role+'/'+p);
  }
 }
 for (const role of ['anon','authenticated','service_role']) {
  assert.equal((await db.query("select has_schema_privilege($1,'humbel_permissions','USAGE') v",[role])).rows[0].v,false);
  for (const fn of ['actor_context()','can_perform(text,text,uuid)']) assert.equal((await db.query('select has_function_privilege($1,$2,\'EXECUTE\') v',[role,'humbel_permissions.'+fn])).rows[0].v,false);
 }
 assert.equal((await db.query("select count(*)::int n from pg_policies where schemaname='humbel_permissions'")).rows[0].n,0);
}
async function apply(db) {
 const body=source(); // Verify artifact before any DB change.
 await db.exec('begin');
 try {
  // Identical target boundary to the existing isolated harness, never hosted URLs.
  await db.exec(prior.preflight);
  await db.exec('select pg_advisory_xact_lock(160081)');
  const old=(await db.query('select version,input_hash,catalog_hash from humbel_security.review_step7')).rows;
  assert.equal(old.length,1);assert.equal(old[0].version,'step7-v1');
  assert.equal(old[0].input_hash,hash([...prior.sources(),baseline.source()].join('\n')),'Step7 input provenance');
  await baseline.postconditions(db);
  const oldHash=await prior.snapshot(db);assert.equal(oldHash,old[0].catalog_hash,'Step7 baseline drift');
  if ((await db.query("select to_regnamespace('humbel_permissions') present")).rows[0].present) {
   const marker=(await db.query('select * from humbel_permissions.installation')).rows;
   assert.equal(marker.length,1);assert.equal(marker[0].version,'8b1-v1');
   assert.equal(marker[0].input_hash,hash(body));assert.equal(marker[0].baseline_hash,oldHash);
   await postconditions(db);assert.equal(await snapshot(db),marker[0].schema_hash,'8B1 drift; no automatic repair');
   await db.exec('commit');return {status:'ISOLATED_8B1_VERIFIED_NOOP',productionReady:false};
  }
  await db.exec(body);
  for (const table of ['identities','sessions','memberships','organization_roles','leaderships','capabilities','module_grants']) assert.equal((await db.query('select count(*)::int n from humbel_permissions.'+table)).rows[0].n,0);
  assert.equal(await prior.snapshot(db),oldHash,'Existing security/business schema changed');
  await db.exec('create table humbel_permissions.installation(version text primary key,input_hash text not null,baseline_hash text not null,schema_hash text not null);alter table humbel_permissions.installation enable row level security;revoke all on humbel_permissions.installation from public,anon,authenticated,service_role');
  await postconditions(db);
  await db.query('insert into humbel_permissions.installation values($1,$2,$3,$4)',['8b1-v1',hash(body),oldHash,await snapshot(db)]);
  await db.exec('commit');return {status:'ISOLATED_8B1_INSTALLED',productionReady:false};
 } catch(e) {await db.exec('rollback');throw e;}
}
module.exports={apply,source,snapshot,postconditions,manifest};
