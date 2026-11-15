// Network-free isolated review installer, no application import/connection setup.
const fs=require('node:fs'),crypto=require('node:crypto'),assert=require('node:assert/strict');
const base=require('./permissions-8b1-review');
const root=require('./security-step6-review');
const hash=s=>crypto.createHash('sha256').update(s).digest('hex');
const manifest=()=>JSON.parse(fs.readFileSync('tests/db-integration/permissions-8b2-package.json','utf8'));
function source(){const s=fs.readFileSync('supabase/review/permissions-8b2-isolated.sql','utf8');assert.equal(hash(s),manifest().inputSha256,'Unreviewed8B2 source');assert.equal((s.match(/^begin;$/gm)||[]).length,1);assert.equal((s.match(/^commit;$/gm)||[]).length,1);return s.replace(/^begin;\n/m,'').replace(/^commit;\s*$/m,'');}
async function snapshot(db){
 const queries=[
  "select nspowner,nspacl::text from pg_namespace where nspname='humbel_identity_admin'",
  "select c.relname,c.relkind,c.relowner,c.relacl::text,c.relrowsecurity from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname='humbel_identity_admin' order by 1",
  "select c.relname,a.attname,a.atttypid::regtype::text,a.attnotnull,a.attacl::text,pg_get_expr(d.adbin,d.adrelid) def from pg_attribute a join pg_class c on c.oid=a.attrelid join pg_namespace n on n.oid=c.relnamespace left join pg_attrdef d on d.adrelid=a.attrelid and d.adnum=a.attnum where n.nspname='humbel_identity_admin' and a.attnum>0 and not a.attisdropped order by 1,a.attnum",
  "select c.relname,k.conname,pg_get_constraintdef(k.oid) def,k.convalidated from pg_constraint k join pg_class c on c.oid=k.conrelid join pg_namespace n on n.oid=c.relnamespace where n.nspname='humbel_identity_admin' order by 1,2",
  "select p.proname,pg_get_function_identity_arguments(p.oid) args,p.proowner,p.proacl::text,p.prosecdef,p.proconfig::text,pg_get_functiondef(p.oid) def from pg_proc p join pg_namespace n on n.oid=p.pronamespace where n.nspname='humbel_identity_admin' order by 1,2",
  "select c.relname,t.tgname,pg_get_triggerdef(t.oid) def from pg_trigger t join pg_class c on c.oid=t.tgrelid join pg_namespace n on n.oid=c.relnamespace where n.nspname='humbel_identity_admin' and not t.tgisinternal order by 1,2",
  "select * from pg_policies where schemaname='humbel_identity_admin' order by tablename,policyname"
 ];const rows=[];for(const q of queries)rows.push((await db.query(q)).rows);return hash(JSON.stringify(rows));
}
async function postconditions(db){
 assert.deepEqual((await db.query('select * from humbel_identity_admin.runtime')).rows,[{singleton:true,enabled:false}]);
 assert.deepEqual((await db.query('select * from humbel_permissions.runtime')).rows,[{singleton:true,enabled:false}]);
 assert.equal((await db.query("select count(*)::int n from pg_tables where schemaname='humbel_identity_admin' and rowsecurity")).rows[0].n,9);
 assert.equal((await db.query("select count(*)::int n from pg_policies where schemaname='humbel_identity_admin'")).rows[0].n,0);
 // All effective schema/table/column/function privileges, including PUBLIC inheritance.
 const bad=await db.query(`select count(*)::int n from pg_class c join pg_namespace n on n.oid=c.relnamespace
 cross join (values('anon'),('authenticated'),('service_role')) r(role)
 cross join (values('SELECT'),('INSERT'),('UPDATE'),('DELETE'),('TRUNCATE'),('REFERENCES'),('TRIGGER'),('MAINTAIN')) p(priv)
 where n.nspname='humbel_identity_admin' and c.relkind='r' and has_table_privilege(r.role,c.oid,p.priv)`);assert.equal(bad.rows[0].n,0);
 assert.equal((await db.query(`select count(*)::int n from pg_attribute a join pg_class c on c.oid=a.attrelid join pg_namespace n on n.oid=c.relnamespace
 cross join (values('anon'),('authenticated'),('service_role')) r(role) cross join(values('SELECT'),('INSERT'),('UPDATE'),('REFERENCES')) p(priv)
 where n.nspname='humbel_identity_admin' and c.relkind='r' and a.attnum>0 and not a.attisdropped and has_column_privilege(r.role,c.oid,a.attnum,p.priv)`)).rows[0].n,0);
 for(const role of ['anon','authenticated','service_role']){
  assert.equal((await db.query("select has_schema_privilege($1,'humbel_identity_admin','USAGE') v",[role])).rows[0].v,false);
  assert.equal((await db.query("select count(*)::int n from pg_proc p join pg_namespace n on n.oid=p.pronamespace where n.nspname='humbel_identity_admin' and has_function_privilege($1,p.oid,'EXECUTE')",[role])).rows[0].n,0);
 }
}
async function apply(db){
 const body=source();
 // Verify an ALREADY installed foundation; never bootstrap Phase1 implicitly.
 assert.notEqual((await db.query("select to_regclass('humbel_permissions.installation') marker")).rows[0].marker,null,'8B1 prerequisite not installed; no implicit bootstrap');
 await base.apply(db);
 const legacy=await root.snapshot(db),foundation=await base.snapshot(db);
 await db.exec('begin');try{
  await db.exec(root.preflight);await db.exec('select pg_advisory_xact_lock(160082)');
  if((await db.query("select to_regnamespace('humbel_identity_admin') v")).rows[0].v){
   const r=(await db.query('select * from humbel_identity_admin.installation')).rows;assert.equal(r.length,1);assert.equal(r[0].version,'8b2-v1');assert.equal(r[0].input_hash,hash(body));assert.equal(r[0].legacy_hash,legacy);assert.equal(r[0].foundation_hash,foundation);await postconditions(db);assert.equal(await snapshot(db),r[0].schema_hash,'8B2 drift');await db.exec('commit');return{status:'ISOLATED_8B2_VERIFIED_NOOP'};
  }
  await db.exec(body);
  assert.equal(await root.snapshot(db),legacy,'Legacy schema changed');assert.equal(await base.snapshot(db),foundation,'Phase1 schema changed');
  for(const t of ['operators','scopes','session_attestations','identity_versions','requests','session_tombstones','audit'])assert.equal((await db.query('select count(*)::int n from humbel_identity_admin.'+t)).rows[0].n,0);
  await db.exec('create table humbel_identity_admin.installation(version text primary key,input_hash text not null,legacy_hash text not null,foundation_hash text not null,schema_hash text not null);alter table humbel_identity_admin.installation enable row level security;revoke all on humbel_identity_admin.installation from public,anon,authenticated,service_role');
  await postconditions(db);await db.query('insert into humbel_identity_admin.installation values($1,$2,$3,$4,$5)',['8b2-v1',hash(body),legacy,foundation,await snapshot(db)]);await db.exec('commit');return{status:'ISOLATED_8B2_INSTALLED'};
 }catch(e){await db.exec('rollback');throw e;}
}
module.exports={apply,source,snapshot,postconditions,manifest};
