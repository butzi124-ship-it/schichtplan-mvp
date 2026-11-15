const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),crypto=require('node:crypto');
const {fixture,ids}=require('./fixture');
const review=require('../db-integration/security-step6-review');
const deny=p=>assert.rejects(p,e=>e.code==='42501');
async function ready(f){await f.db.query("select set_config('humbel.isolated_sentinel','synthetic-step6-isolated-sentinel-at-least-32',false)");}
async function enroll(f){
 for(const label of ['admin','leader','worker','other','deputy','scanner']){
  const role=label==='admin'?'admin':label==='leader'?'department_admin':label==='scanner'?'tool_scanner':'employee';
  await f.db.query("insert into humbel_security.principals values($1,$1,$2,$3,true,now()+interval '1 day','Independent synthetic Step6 approval')",[ids[label],role,label==='other'?ids.d2:ids.d1]);
 }
}
async function withFixture(fn){const f=await fixture();try{await fn(f);}finally{await f.db.close();}}
const legacy={
 planner_assignments:"(shift_id,shift_date,assigned_user) values('legacy','2026-11-12','worker')",
 planner_absences:"(absence_key,absence_date,user_name) values('legacy','2026-11-12','worker')",
 planner_availability:"(availability_key,shift_id,shift_date,user_name,status) values('legacy','legacy','2026-11-12','worker','yes')",
 planner_saturday_requests:"(request_key,shift_id,shift_date,user_name) values('legacy','legacy','2026-11-14','worker')",
 planner_shift_cancellations:"(shift_id,shift_date,reason) values('legacy','2026-11-12','Synthetic confidential reason')",
 planner_sick_leaves:"(user_name,from_date,to_date) values('worker','2026-11-12','2026-11-13')",
 planner_vacations:"(user_name,from_date,to_date) values('worker','2026-11-12','2026-11-13')",
 planner_swaps:"(user_a,user_b,start_date) values('worker','other','2026-11-12')",
 planner_absence_replacements:"(shift_id,shift_date,source_type,source_id,absent_user,from_date,to_date,mode,replacement_user) values('legacy','2026-11-12','sick','synthetic','worker','2026-11-12','2026-11-13','replace','deputy')",
 tool_materials:"(name) values('Synthetic legacy material')"
};
async function rows(f){const out={};for(const t of review.ten)out[t]=(await f.db.query('select * from public.'+t+' order by id')).rows;return out;}

test('STEP6 BASELINE ONLY: all 20 supplied grant combinations are reproduced, without claiming original restore',async()=>withFixture(async f=>{
 for(const t of review.ten){
  assert.equal((await f.db.query('select relrowsecurity from pg_class where oid=$1::regclass',['public.'+t])).rows[0].relrowsecurity,false);
  for(const role of ['anon','authenticated'])for(const p of ['SELECT','INSERT','UPDATE','DELETE'])assert.equal((await f.db.query('select has_table_privilege($1,$2,$3) v',[role,'public.'+t,p])).rows[0].v,true);
 }
}));

test('STEP6: missing sentinel aborts before mutation; unknown prior candidate is not overwritten',async()=>withFixture(async f=>{
 const before=await review.snapshot(f.db);await deny(review.apply(f.db));assert.equal(await review.snapshot(f.db),before);
 await ready(f);await f.db.exec('create schema humbel_security');
 await assert.rejects(review.apply(f.db),/Unknown partial/);assert.equal((await f.db.query("select to_regclass('humbel_security.principals') x")).rows[0].x,null);
}));

test('STEP6 CORRECTED: all ten RLS/ACL/column postconditions; old permissive policies and column grants removed',async()=>withFixture(async f=>{
 await ready(f);await f.db.exec('grant update(user_name) on public.planner_sick_leaves to anon,authenticated;create policy synthetic_true on public.planner_assignments for all to public using(true) with check(true)');
 assert.equal((await review.apply(f.db)).status,'ISOLATED_REVIEW_INSTALLED');await review.postconditions(f.db);
 assert.equal((await f.db.query("select count(*)::int n from pg_policies where policyname='synthetic_true'")).rows[0].n,0);
 for(const t of review.ten)await deny(f.asUser(null,d=>d.query('select * from public.'+t),{anonymous:true}));
 await deny(f.asUser(ids.worker,d=>d.query('select * from humbel_security.review_installation')));
}));

test('STEP6 CORRECTED: legacy data survives install and verified repeated installation byte-equivalently',async()=>withFixture(async f=>{
 for(const [t,sql] of Object.entries(legacy))await f.db.exec('insert into public.'+t+sql);
 const before=await rows(f);await ready(f);await review.apply(f.db);assert.deepEqual(await rows(f),before);
 const first=await review.snapshot(f.db);assert.equal((await review.apply(f.db)).status,'VERIFIED_ISOLATED_NOOP');assert.equal(await review.snapshot(f.db),first);assert.deepEqual(await rows(f),before);
 await enroll(f);
 for(const t of review.ten.filter(t=>t!=='tool_materials')){
  for(const who of [ids.worker,ids.leader,ids.admin])assert.deepEqual((await f.asUser(who,d=>d.query('select * from public.'+t))).rows,[]);
  await deny(f.asUser(ids.admin,d=>d.query('delete from public.'+t)));
 }
 assert.deepEqual(await rows(f),before);
}));

test('STEP6 CORRECTED: real material CRUD payloads remain allowed only to admin; identity/timestamps/delete are bounded',async()=>withFixture(async f=>{
 await ready(f);await review.apply(f.db);await enroll(f);
 const id=(await f.asUser(ids.admin,d=>d.query("insert into public.tool_materials(name,sort_order,is_active) values('Synthetic reviewed material',10,true) returning *"))).rows[0].id;
 for(const who of [ids.worker,ids.leader,ids.scanner]){
  assert.equal((await f.asUser(who,d=>d.query('select id,name,sort_order,is_active from public.tool_materials where id=$1',[id]))).rows.length,1);
  await deny(f.asUser(who,d=>d.query("insert into public.tool_materials(name) values('Denied')")));
  assert.deepEqual((await f.asUser(who,d=>d.query("update public.tool_materials set name='Denied' where id=$1 returning id",[id]))).rows,[]);
 }
 await f.asUser(ids.admin,d=>d.query("update public.tool_materials set name='Synthetic renamed',updated_at=now() where id=$1 returning *",[id]));
 await f.asUser(ids.admin,d=>d.query('update public.tool_materials set is_active=false,updated_at=now() where id=$1',[id]));
 for(const field of ['id=gen_random_uuid()','created_at=now()','sort_order=100'])await deny(f.asUser(ids.admin,d=>d.query('update public.tool_materials set '+field+' where id=$1',[id])));
 await deny(f.asUser(ids.admin,d=>d.query('delete from public.tool_materials where id=$1',[id])));
 assert.equal((await f.db.query('select is_active from public.tool_materials where id=$1',[id])).rows[0].is_active,false);
}));

test('STEP6 CORRECTED: own AL plan, explicit deputy expiry/revocation and distinct publish/read survive combined installation',async()=>withFixture(async f=>{
 await ready(f);await review.apply(f.db);await enroll(f);assert.equal((await f.assign(ids.leader)).rows[0].result.version,1);
 for(const who of [ids.worker,ids.scanner,ids.orphan])await deny(f.assign(who,crypto.randomUUID()));
 await deny(f.assign(ids.leader,crypto.randomUUID(),0,ids.d2,ids.other));
 await f.db.query("insert into humbel_security.delegations values($1,$2,$3,now()-interval '1 day',now()+interval '1 day',false,'Synthetic explicit delegation')",[ids.deputy,ids.d1,ids.leader]);
 assert.equal((await f.assign(ids.deputy,crypto.randomUUID())).rows[0].result.version,1);
 for(const change of ["revoked=true","revoked=false,valid_until=now()-interval '1 second'"]){await f.db.query('update humbel_security.delegations set '+change+' where user_id=$1',[ids.deputy]);await deny(f.assign(ids.deputy,crypto.randomUUID()));}
 await deny(f.asUser(ids.leader,d=>d.query('select public.security_publish_assignment($1,1)',[ids.assignment])));
 for(const [who,read,publish] of [[ids.worker,true,false],[ids.admin,false,true]])await f.db.query("insert into humbel_security.schedule_grants values($1,$2,$3,$4,now()-interval '1 day',now()+interval '1 day',false,'Synthetic test scope')",[who,ids.d1,read,publish]);
 await f.asUser(ids.admin,d=>d.query('select public.security_publish_assignment($1,1)',[ids.assignment]));
 const result=(await f.asUser(ids.worker,d=>d.query('select * from public.security_published_schedule()'))).rows;
 assert.equal(result.length,1);assert.deepEqual(Object.keys(result[0]).sort(),['assignment_id','employee_id','shift_date','shift_id']);
 await assert.rejects(f.assign(ids.leader,ids.assignment,0),e=>e.code==='40001');
}));

test('STEP6: inherited privileges fail postconditions and roll back all three stages, policies and grants',async()=>withFixture(async f=>{
 await ready(f);await f.db.exec('create role synthetic_extra;grant update on public.planner_assignments to synthetic_extra;grant synthetic_extra to authenticated');
 const before=await review.snapshot(f.db);await assert.rejects(review.apply(f.db),/planner_assignments\/authenticated\/UPDATE/);
 assert.equal(await review.snapshot(f.db),before);assert.equal((await f.db.query("select to_regnamespace('humbel_security') x")).rows[0].x,null);
}));

test('STEP6: DDL failure after Step3 rolls back everything without data loss',async()=>withFixture(async f=>{
 await ready(f);await f.db.exec("insert into public.tool_materials(name) values('Preserved');create function public.security_actor_context() returns text language sql as $$select 'collision'::text$$");
 const before=await review.snapshot(f.db),data=await rows(f);await assert.rejects(review.apply(f.db),e=>e.code==='42723');
 assert.equal(await review.snapshot(f.db),before);assert.deepEqual(await rows(f),data);
}));

test('STEP6: repeated package rejects policy/column/role drift and changed history without opening rights',async()=>withFixture(async f=>{
 await ready(f);await review.apply(f.db);
 await f.db.exec('grant update(id) on public.tool_materials to authenticated');await assert.rejects(review.apply(f.db),/authenticated column id\/UPDATE/);
 await f.db.exec('revoke update(id) on public.tool_materials from authenticated');
 await f.db.exec("alter policy scoped_read on public.planner_assignments using(true)");await assert.rejects(review.apply(f.db),/Catalog drift/);
 await f.db.exec("update humbel_security.review_installation set input_hash='unreviewed'");await assert.rejects(review.apply(f.db),/Different installed package/);
}));

test('STEP6: manual original findings have explicit provenance; readonly metadata template executes only in synthetic DB',async()=>withFixture(async f=>{
 const evidence=JSON.parse(fs.readFileSync('docs/qa/security-step6-original-evidence.json','utf8'));
 assert.equal(evidence.originalSchemaComplete,false);assert.equal(evidence.effectivePrivileges.length,20);assert.match(evidence.source,/not independently reexecuted/);
 assert.equal(evidence.externalDataApiReachability,'UNVERIFIED');
 const query=fs.readFileSync('supabase/review/security-step6-metadata-readonly.sql','utf8');
 const before=await review.snapshot(f.db);const results=await f.db.exec(query),data=results.find(r=>r.rows?.[0]?.metadata_only)?.rows[0].metadata_only;
 assert(data);assert.equal(data.relations.filter(r=>r.schema==='public'&&r.kind==='r').length,31);
 assert.equal(data.effective_table_rights.length,62);assert.equal(data.policies.length,127);assert.equal(data.columns.filter(c=>c.schema==='public').length,308);
 assert.equal(data.effective_column_rights.length,616);assert(data.functions.length>=6);assert(data.indexes.length>0);assert(data.dependencies.length>0);
 assert.equal(await review.snapshot(f.db),before);
}));

test('STEP6: unknown native runner fails before changes; new independent workflow preserves Step3/4 acceptance chain',()=>{
 const os=require('node:os'),path=require('node:path'),YAML=require('yaml'),{prepare}=require('../db-integration/security-step6-runner');
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'step6-native-'));
 try{fs.mkdirSync(path.join(dir,'scripts/db-integration'),{recursive:true});const file=path.join(dir,'scripts/db-integration/run.js');fs.writeFileSync(file,'unknown');assert.throws(()=>prepare(dir),/Exact Step4/);assert.equal(fs.readFileSync(file,'utf8'),'unknown');}finally{fs.rmSync(dir,{recursive:true,force:true});}
 const w=YAML.parse(fs.readFileSync('.github/workflows/security-step6-isolated.yml','utf8')),job=w.jobs['security-step6-native'],steps=job.steps.map(s=>s.run||'');
 assert.equal(job.needs,'security-contract');assert(steps.findIndex(s=>s.includes('security-step4-runner.js'))<steps.findIndex(s=>s.includes('security-step6-runner.js')));
 assert(steps.some(s=>s==='docker exec humbel-gateway node scripts/db-integration/run.js ba-op'));
 assert(job.steps.some(s=>s.if==='always()'&&s.run?.includes('runtime.js stop')));
 const native=fs.readFileSync('tests/db-integration/security-step6-extension.js','utf8');assert(native.indexOf("security-step4-extension")<native.indexOf('sources()[2]'));
 assert(!fs.readFileSync('schichtplan_mvp.html','utf8').includes('secure-access.js'));
});

test('STEP6: same-signature view body drift cannot be hidden behind a valid installation marker',async()=>withFixture(async f=>{
 await ready(f);await f.db.exec('create view public.synthetic_review_view as select 1 as value');await review.apply(f.db);
 await f.db.exec('create or replace view public.synthetic_review_view as select 2 as value');
 await assert.rejects(review.apply(f.db),/Catalog drift/);
}));

test('STEP6: inherited TRUNCATE cannot bypass row policies or survive migration postconditions',async()=>withFixture(async f=>{
 await ready(f);await f.db.exec('create role synthetic_truncate;grant truncate on public.planner_assignments to synthetic_truncate;grant synthetic_truncate to authenticated');
 const before=await review.snapshot(f.db);await assert.rejects(review.apply(f.db),/planner_assignments\/authenticated\/TRUNCATE/);
 assert.equal(await review.snapshot(f.db),before);
}));
