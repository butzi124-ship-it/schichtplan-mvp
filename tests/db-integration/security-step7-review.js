// Isolated only, caller supplies disposable local DB; no connection or URL lookup.
const fs=require('node:fs'),assert=require('node:assert/strict'),crypto=require('node:crypto');
const prior=require('./security-step6-review');
const manifest=()=>JSON.parse(fs.readFileSync('tests/db-integration/security-step7-package.json','utf8'));
const hash=s=>crypto.createHash('sha256').update(s).digest('hex');
function source(){const m=manifest(),s=fs.readFileSync('supabase/review/security-step7-isolated.sql','utf8');assert.equal(hash(s),m.inputSha256,'Unreviewed Step7 input');assert.equal((s.match(/^begin;$/gm)||[]).length,1);assert.equal((s.match(/^commit;$/gm)||[]).length,1);return s.replace(/^begin;\n/m,'').replace(/^commit;\s*$/m,'');}
async function postconditions(db){
 await prior.postconditions(db);
 for(const t of manifest().tables){
  for(const role of ['anon','authenticated'])for(const p of ['TRUNCATE','REFERENCES','TRIGGER','MAINTAIN'])assert.equal((await db.query('select has_table_privilege($1,$2,$3) v',[role,'public.'+t,p])).rows[0].v,false,t+'/'+role+'/'+p);
  const columns=(await db.query('select attname from pg_attribute where attrelid=$1::regclass and attnum>0 and not attisdropped',['public.'+t])).rows;
  for(const {attname} of columns)for(const role of ['anon','authenticated'])assert.equal((await db.query('select has_column_privilege($1,$2,$3,\'REFERENCES\') v',[role,'public.'+t,attname])).rows[0].v,false,t+'/'+role+'/REFERENCES column');
 }
 for(const p of ['SELECT','INSERT','UPDATE','DELETE'])assert.equal((await db.query("select has_table_privilege('anon','public.employees',$1) v",[p])).rows[0].v,false,'employees anon '+p);
 assert.equal((await db.query("select relrowsecurity from pg_class where oid='public.employees'::regclass")).rows[0].relrowsecurity,true);
 assert.equal((await db.query("select count(*)::int n from pg_policies where schemaname='public' and tablename='employees' and (qual='true' or with_check='true')")).rows[0].n,0);
 assert.equal((await db.query("select count(*)::int n from pg_default_acl d cross join lateral aclexplode(d.defaclacl) a where d.defaclobjtype='r' and d.defaclnamespace='public'::regnamespace and a.grantee in(0,'anon'::regrole,'authenticated'::regrole) and a.privilege_type in('TRUNCATE','REFERENCES','TRIGGER','MAINTAIN')")).rows[0].n,0,'Future API maintenance defaults');
 for(const fn of ['is_admin()','current_employee_id()','current_app_employee_id()','current_app_employee_role()','can_access_production_department(uuid)','set_updated_at()'])assert.equal((await db.query("select has_function_privilege('anon',$1,'EXECUTE') v",['public.'+fn])).rows[0].v,false,fn);
}
async function apply(db){
 const sql=[...prior.sources(),source()],checksum=hash(sql.join('\n'));
 await db.exec('begin');
 try{
  await db.exec(prior.preflight);
  await db.exec('lock table '+manifest().tables.map(t=>'public.'+t).join(',')+' in access exclusive mode');
  if((await db.query("select to_regclass('humbel_security.review_step7') marker")).rows[0].marker){
   const rows=(await db.query('select version,input_hash,catalog_hash from humbel_security.review_step7')).rows;
   assert.equal(rows.length,1);assert.equal(rows[0].version,'step7-v1');assert.equal(rows[0].input_hash,checksum,'Different Step7 package');await postconditions(db);assert.equal(await prior.snapshot(db),rows[0].catalog_hash,'Step7 catalog drift');
   await db.exec('commit');return{status:'ISOLATED_STEP7_VERIFIED_NOOP',authoritativeRestore:false};
  }
  assert.equal((await db.query("select to_regnamespace('humbel_security') present")).rows[0].present,null,'Fresh isolated baseline required; no implicit upgrade/downgrade');
  for(const body of sql)await db.exec(body);
  await postconditions(db);
  await db.exec('create table humbel_security.review_step7(version text primary key,input_hash text not null,catalog_hash text not null);alter table humbel_security.review_step7 enable row level security;revoke all on humbel_security.review_step7 from public,anon,authenticated');
  await db.query('insert into humbel_security.review_step7 values($1,$2,$3)',['step7-v1',checksum,await prior.snapshot(db)]);
  await db.exec('commit');return{status:'ISOLATED_STEP7_INSTALLED',authoritativeRestore:false};
 }catch(e){await db.exec('rollback');throw e;}
}
module.exports={apply,source,postconditions,manifest};
