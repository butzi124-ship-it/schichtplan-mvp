// Network-free orchestrator. Caller supplies ONLY an approved disposable local DB.
// This is review/test infrastructure, never imported by application or deployment.
const fs=require('node:fs'),crypto=require('node:crypto'),assert=require('node:assert/strict');
const files=['supabase/review/security-step3-isolated.sql','supabase/review/security-step4-integration-isolated.sql','supabase/review/security-step6-isolated.sql'];
const tables=['employees','departments','planner_absence_replacements','planner_absences','planner_assignments','planner_availability','planner_saturday_requests','planner_shift_cancellations','planner_sick_leaves','planner_swaps','planner_vacations','planner_special_days','planner_special_day_assignments','planner_tasks','tool_materials'];
const ten=tables.filter(t=>!['employees','departments','planner_special_days','planner_special_day_assignments','planner_tasks'].includes(t));
const hash=s=>crypto.createHash('sha256').update(s).digest('hex');
function sources(){
 const manifest=JSON.parse(fs.readFileSync('tests/db-integration/security-step6-package.json','utf8'));
 return files.map(path=>{const raw=fs.readFileSync(path,'utf8');assert.equal(hash(raw),manifest.inputs[path],'Unreviewed input: '+path);
  assert.equal((raw.match(/^begin;$/gm)||[]).length,1);assert.equal((raw.match(/^commit;$/gm)||[]).length,1);
  return raw.replace(/^begin;\r?\n/m,'').replace(/^commit;\s*$/m,'');});
}
const preflight=`
 set local lock_timeout='5s';set local statement_timeout='30s';
 do $$begin
 if not (pg_catalog.current_database()='humbel_isolated_test' or (pg_catalog.current_database()='postgres' and pg_catalog.version() like '%(PGlite %'))
 or pg_catalog.length(coalesce(pg_catalog.current_setting('humbel.isolated_sentinel',true),''))<32
 then raise exception 'ISOLATED STEP6 ONLY' using errcode='42501';end if;end$$;
 select pg_catalog.pg_advisory_xact_lock(160006);
 lock table ${tables.map(t=>'public.'+t).join(',')} in access exclusive mode;
`;
// Definition fingerprint, not only a marker. Captures policies, all table/column ACLs,
// columns/defaults/constraints, roles/memberships, schema/default ACLs, triggers and
// relevant definer bodies/owners/EXECUTE. Historical catalog != original restore.
async function snapshot(db){
 const statements=[
  ["select n.nspname,c.relname,c.relkind,c.relowner,c.relacl::text,c.relrowsecurity,c.relforcerowsecurity from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname in('public','humbel_security') order by 1,2"],
  ["select n.nspname,c.relname,a.attname,a.atttypid::regtype::text,a.attnotnull,a.attidentity,a.attgenerated,a.attacl::text,pg_get_expr(d.adbin,d.adrelid) as default_expr from pg_attribute a join pg_class c on c.oid=a.attrelid join pg_namespace n on n.oid=c.relnamespace left join pg_attrdef d on d.adrelid=a.attrelid and d.adnum=a.attnum where n.nspname in('public','humbel_security') and a.attnum>0 and not a.attisdropped order by 1,2,a.attnum"],
  ["select schemaname,tablename,policyname,permissive,roles::text,cmd,qual,with_check from pg_policies where schemaname in('public','humbel_security') order by 1,2,3"],
  ["select n.nspname,p.proname,pg_get_function_identity_arguments(p.oid) args,p.proowner,p.proacl::text,p.prosecdef,p.proconfig::text,pg_get_functiondef(p.oid) definition from pg_proc p join pg_namespace n on n.oid=p.pronamespace where n.nspname in('public','humbel_security') and p.prokind in('f','p') order by 1,2,3"],
  ["select n.nspname,c.relname,t.tgname,t.tgenabled,pg_get_triggerdef(t.oid) definition from pg_trigger t join pg_class c on c.oid=t.tgrelid join pg_namespace n on n.oid=c.relnamespace where n.nspname in('public','humbel_security') and not t.tgisinternal order by 1,2,3"],
  ["select n.nspname,c.relname,k.conname,pg_get_constraintdef(k.oid) definition from pg_constraint k join pg_class c on c.oid=k.conrelid join pg_namespace n on n.oid=c.relnamespace where n.nspname in('public','humbel_security') order by 1,2,3"],
  ["select rolname,rolsuper,rolinherit,rolcreaterole,rolcreatedb,rolcanlogin,rolreplication,rolbypassrls from pg_roles order by rolname"],
  ["select roleid,member,grantor,admin_option,inherit_option,set_option from pg_auth_members order by roleid,member,grantor"],
  ["select nspname,nspowner,nspacl::text from pg_namespace where nspname in('public','humbel_security') order by 1"],
  ["select defaclrole,defaclnamespace,defaclobjtype,defaclacl::text from pg_default_acl order by 1,2,3"],
  ["select n.nspname,c.relname,pg_get_viewdef(c.oid,false) definition,c.reloptions::text from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname in('public','humbel_security') and c.relkind in('v','m') order by 1,2"],
  ["select n.nspname,c.relname,i.indexrelid::regclass::text,pg_get_indexdef(i.indexrelid) definition,i.indisvalid,i.indisready from pg_index i join pg_class c on c.oid=i.indrelid join pg_namespace n on n.oid=c.relnamespace where n.nspname in('public','humbel_security') order by 1,2,3"],
  ["select n.nspname,c.relname,s.seqtypid,s.seqstart::text,s.seqincrement::text,s.seqmax::text,s.seqmin::text,s.seqcache::text,s.seqcycle from pg_sequence s join pg_class c on c.oid=s.seqrelid join pg_namespace n on n.oid=c.relnamespace where n.nspname in('public','humbel_security') order by 1,2"]
 ];
 const result=[];for(const [sql] of statements)result.push((await db.query(sql)).rows);
 return hash(JSON.stringify(result));
}
async function postconditions(db){
 for(const t of ten){
  assert.equal((await db.query('select relrowsecurity from pg_class where oid=$1::regclass',['public.'+t])).rows[0].relrowsecurity,true,t+' RLS');
  for(const role of ['anon','authenticated'])for(const privilege of ['SELECT','INSERT','UPDATE','DELETE','TRUNCATE','REFERENCES','TRIGGER','MAINTAIN']){
   const expected=role==='authenticated'&&privilege==='SELECT';
   assert.equal((await db.query('select has_table_privilege($1,$2,$3) v',[role,'public.'+t,privilege])).rows[0].v,expected,`${t}/${role}/${privilege}`);
   assert.equal((await db.query('select has_table_privilege($1,$2,$3) v',[role,'public.'+t,privilege+' WITH GRANT OPTION'])).rows[0].v,false,`${t}/${role}/${privilege} grant option`);
  }
  const cols=(await db.query('select attname from pg_attribute where attrelid=$1::regclass and attnum>0 and not attisdropped',['public.'+t])).rows;
  for(const {attname} of cols)for(const privilege of ['SELECT','INSERT','UPDATE','REFERENCES']){
   assert.equal((await db.query("select has_column_privilege('anon',$1,$2,$3) v",['public.'+t,attname,privilege])).rows[0].v,false,t+' anon column');
   const expected=privilege==='SELECT'||(t==='tool_materials'&&((privilege==='INSERT'&&['name','sort_order','is_active'].includes(attname))||(privilege==='UPDATE'&&['name','is_active','updated_at'].includes(attname))));
   assert.equal((await db.query("select has_column_privilege('authenticated',$1,$2,$3) v",['public.'+t,attname,privilege])).rows[0].v,expected,t+' authenticated column '+attname+'/'+privilege);
   for(const role of ['anon','authenticated'])assert.equal((await db.query('select has_column_privilege($1,$2,$3,$4) v',[role,'public.'+t,attname,privilege+' WITH GRANT OPTION'])).rows[0].v,false,t+' column grant option');
  }
 }
}
async function apply(db){
 const sql=sources(),inputHash=hash(sql.join('\n'));
 await db.exec('begin');
 try{
  await db.exec(preflight);
  const installed=(await db.query("select to_regclass('humbel_security.review_installation') marker")).rows[0].marker;
  if(installed){
   const rows=(await db.query('select version,input_hash,catalog_hash from humbel_security.review_installation')).rows;
   assert.equal(rows.length,1,'Unknown review history');assert.equal(rows[0].version,'step6-v1');assert.equal(rows[0].input_hash,inputHash,'Different installed package');
   await postconditions(db);assert.equal(await snapshot(db),rows[0].catalog_hash,'Catalog drift: abort, never repair blindly');
   await db.exec('commit');return{status:'VERIFIED_ISOLATED_NOOP',authoritativeRestore:false};
  }
  assert.equal((await db.query("select to_regnamespace('humbel_security') present")).rows[0].present,null,'Unknown partial/previous candidate: use a fresh isolated baseline');
  for(const body of sql)await db.exec(body);
  await postconditions(db);
  await db.exec('create table humbel_security.review_installation(version text primary key,input_hash text not null,catalog_hash text not null);alter table humbel_security.review_installation enable row level security;revoke all on humbel_security.review_installation from public,anon,authenticated;');
  await db.query('insert into humbel_security.review_installation values($1,$2,$3)',['step6-v1',inputHash,await snapshot(db)]);
  await db.exec('commit');return{status:'ISOLATED_REVIEW_INSTALLED',authoritativeRestore:false};
 }catch(error){await db.exec('rollback');throw error;}
}
module.exports={apply,postconditions,snapshot,sources,preflight,tables,ten};
