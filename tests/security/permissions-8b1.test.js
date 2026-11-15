const {test}=require('node:test'),assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto');
const {fixture,ids}=require('./permissions-8b1-fixture');
const review=require('../db-integration/permissions-8b1-review');
const prior=require('../db-integration/security-step6-review');
const deny=p=>assert.rejects(p,e=>e.code==='42501');
async function use(fn){const f=await fixture();try{await fn(f);}finally{await f.db.close();}}
async function ready(f){await f.identity('worker');const r=await f.role('worker');await f.membership('worker',ids.d1);const g=await f.grant('worker',r);return {r,g};}

test('8B1 DEFAULT: private additive install preserves existing schema and employees; empty assignments, seven names, replay verified',async()=>use(async f=>{
 await review.postconditions(f.db);const before=await prior.snapshot(f.db);
 assert.equal((await f.db.query('select count(*)::int n from humbel_permissions.identities')).rows[0].n,0);
 for(const t of ['memberships','organization_roles','capabilities','module_grants'])assert.equal((await f.db.query('select count(*)::int n from humbel_permissions.'+t)).rows[0].n,0);
 assert.equal((await review.apply(f.db)).status,'ISOLATED_8B1_VERIFIED_NOOP');assert.equal(await prior.snapshot(f.db),before);
 const rootHash=(await f.db.query('select input_hash from humbel_security.review_step7')).rows[0].input_hash;
 await f.db.exec("update humbel_security.review_step7 set input_hash='unrecognized-package'");await assert.rejects(review.apply(f.db),/Step7 input provenance/);
 await f.db.query('update humbel_security.review_step7 set input_hash=$1',[rootHash]);
 for(const who of [ids.worker,ids.admin])for(const table of ['identities','organization_roles','memberships','module_grants'])await deny(f.asUser(who,d=>d.query('select * from humbel_permissions.'+table)));
 await deny(f.asUser(ids.admin,d=>d.query('select * from humbel_permissions.actor_context()')));
 await deny(f.asUser(null,d=>d.query('select * from humbel_permissions.actor_context()'),{anonymous:true}));
}));

test('8B1 GATE: verified legacy admin, seven roles and membership do not grant implicit actions or appoint anyone',async()=>use(async f=>{
 await ready(f);await f.identity('admin');for(const c of review.manifest().roles)await f.role('admin',c,['department_head','deputy'].includes(c)?ids.d1:null);
 await f.expose();assert.deepEqual(await f.context('worker'),[]);assert.equal(await f.call('worker'),false);
 await f.enabled();assert.equal(await f.call('worker'),true);assert.equal(await f.call('admin'),false);
 assert.equal(await f.call('worker',null),false);assert.equal(await f.call('worker',ids.d2),false);
 await deny(f.asUser(ids.worker,d=>d.exec('update humbel_permissions.runtime set enabled=true')));
}));

test('8B1 IDENTITY: forged roles/IDs, orphan, missing or wrong sessions and ambiguous employee binding fail closed',async()=>use(async f=>{
 await ready(f);await f.expose();await f.enabled();
 assert.equal(await f.call('worker',ids.d1,{role:'admin',employee_id:ids.admin,user_metadata:{role:'executive'},app_metadata:{role:'executive'}}),true);
 assert.equal(await f.call('worker',ids.d2,{role:'admin',employee_id:ids.admin}),false);
 for(const claims of [{session_id:randomUUID()},{session_id:'not-uuid'},{session_id:null},{is_anonymous:true}])assert.deepEqual(await f.context('worker',claims),[]);
 assert.deepEqual(await f.context('orphan'),[]);assert.deepEqual(await f.context('missing'),[]);
 const otherSession=randomUUID();await f.identity('other');assert.deepEqual(await f.context('worker',{session_id:f.sessions.other}),[]);
 await f.db.query('update humbel_permissions.identities set employee_id=$1 where user_id=$2',[ids.orphan,ids.worker]);assert.deepEqual(await f.context('worker'),[]);
 await f.db.query('update humbel_permissions.identities set employee_id=$1 where user_id=$1',[ids.worker]);
 await f.db.exec('alter table public.employees drop constraint employees_auth_user_id_key');
 await f.db.query('update public.employees set auth_user_id=$1 where id=$2',[ids.worker,ids.orphan]);assert.deepEqual(await f.context('worker'),[]);
}));

test('8B1 REVOCATION: active/blocked/root mismatch, session expiry and epoch changes invalidate old contexts without local-role fallback',async()=>use(async f=>{
 await ready(f);await f.expose();await f.enabled();assert.equal(await f.call('worker'),true);
 for(const change of ['enabled=false','blocked=true',"valid_until=now()-interval '1 hour'",'binding_version=2','authorization_epoch=2',`primary_department_id='${ids.d2}'`]){
  await f.db.exec('begin');try{await f.db.query('update humbel_permissions.identities set '+change+' where user_id=$1',[ids.worker]);await f.db.query("select set_config('request.jwt.claims',$1,true)",[JSON.stringify({sub:ids.worker,session_id:f.sessions.worker})]);assert.equal((await f.db.query("select humbel_permissions.can_perform('synthetic','read',$1) v",[ids.d1])).rows[0].v,false);}finally{await f.db.exec('rollback');}
 }
 for(const change of ['revoked=true',"valid_until=now()-interval '1 hour'", "valid_from=now()+interval '1 hour'"]){await f.db.exec('begin');try{await f.db.exec('update humbel_permissions.sessions set '+change);await f.db.query("select set_config('request.jwt.claims',$1,true)",[JSON.stringify({sub:ids.worker,session_id:f.sessions.worker})]);assert.equal((await f.db.query('select count(*)::int n from humbel_permissions.actor_context()')).rows[0].n,0);}finally{await f.db.exec('rollback');}}
 await f.db.query('update public.employees set is_active=false where id=$1',[ids.worker]);assert.equal(await f.call('worker'),false);
 await f.db.query("update public.employees set is_active=true,role='admin' where id=$1",[ids.worker]);assert.equal(await f.call('worker'),false);
}));

test('8B1 MULTI: membership alone grants nothing, explicit A/B heads are equal; unknown C, duplicate and missing head links denied',async()=>use(async f=>{
 await f.identity('leader');await f.membership('leader',ids.d1);await f.membership('leader',ids.d2);
 const a=await f.role('leader','department_head',ids.d1),b=await f.role('leader','department_head',ids.d2);
 await f.expose();await f.enabled();assert.equal(await f.call('leader'),false);
 await f.grant('leader',a);await f.grant('leader',b,ids.d2);
 assert.equal(await f.call('leader'),false);
 for(const [dept,id]of [[ids.d1,a],[ids.d2,b]])await f.db.query('insert into humbel_permissions.leaderships values($1,$2,$3,\'department_head\')',[dept,id,ids.leader]);
 assert.equal(await f.call('leader'),true);assert.equal(await f.call('leader',ids.d2),true);assert.equal(await f.call('leader',randomUUID()),false);
 await assert.rejects(f.membership('leader',ids.d1),e=>e.code==='23505');
 await f.role('leader','department_head',ids.d1);assert.equal(await f.call('leader'),false);assert.equal(await f.call('leader',ids.d2),true);
}));

test('8B1 GRANTS: finite/unbounded, scope, revocation, primary epoch and dependent role remain separate; deputies unresolved',async()=>use(async f=>{
 const {r,g}=await ready(f);await f.expose();await f.enabled();assert.equal(await f.call('worker'),true);
 for(const change of ['enabled=false','revoked=true','primary_epoch=2',"valid_until=now()-interval '1 hour'", "valid_from=now()+interval '1 hour'"]){await f.db.exec('begin');try{await f.db.query('update humbel_permissions.module_grants set '+change+' where id=$1',[g]);await f.db.query("select set_config('request.jwt.claims',$1,true)",[JSON.stringify({sub:ids.worker,session_id:f.sessions.worker})]);assert.equal((await f.db.query("select humbel_permissions.can_perform('synthetic','read',$1) v",[ids.d1])).rows[0].v,false);}finally{await f.db.exec('rollback');}}
 for(const [table,change] of [
  ['memberships','enabled=false'],['memberships','revoked=true'],['memberships',"valid_until=now()-interval '1 hour'"],
  ['organization_roles','enabled=false'],['organization_roles',"valid_until=now()-interval '1 hour'"],
  ['organization_roles',"valid_from=now()+interval '1 hour'"],['capabilities','enabled=false']]){
  await f.db.exec('begin');try{await f.db.exec('update humbel_permissions.'+table+' set '+change);
   await f.db.query("select set_config('request.jwt.claims',$1,true)",[JSON.stringify({sub:ids.worker,session_id:f.sessions.worker})]);
   assert.equal((await f.db.query("select humbel_permissions.can_perform('synthetic','read',$1) v",[ids.d1])).rows[0].v,false);
  }finally{await f.db.exec('rollback');}
 }
 await f.db.query('update public.departments set active=false where id=$1',[ids.d1]);assert.equal(await f.call('worker'),false);
 await f.db.query('update public.departments set active=true where id=$1',[ids.d1]);
 await f.db.query('update humbel_permissions.organization_roles set revoked=true where id=$1',[r]);assert.equal(await f.call('worker'),false);
 await f.identity('deputy');const d=await f.role('deputy','deputy',ids.d1);await f.membership('deputy',ids.d1);await f.grant('deputy',d);assert.equal(await f.call('deputy'),false);
}));

test('8B1 NO ESCALATION: direct REST-equivalent SQL/column writes, private function overload and public/temp spoofing cannot create authority',async()=>use(async f=>{
 await ready(f);await f.expose();await f.enabled();
 for(const q of ["update public.employees set role='admin' where id='"+ids.worker+"'", "update public.employees set auth_user_id='"+ids.orphan+"' where id='"+ids.worker+"'"]){const rows=await f.asUser(ids.worker,d=>d.query(q+' returning id'));assert.deepEqual(rows.rows,[]);}
 for(const q of ["insert into humbel_permissions.organization_roles(id,user_id,role_code,valid_from,evidence) values(gen_random_uuid(),'"+ids.worker+"','executive',now(),'forged')",'delete from humbel_permissions.identities',"create function humbel_permissions.fake() returns bool language sql as 'select true'"] )await deny(f.asUser(ids.worker,d=>d.exec(q)));
 await f.asUser(ids.worker,async d=>{await d.exec("create temporary table identities(enabled bool);insert into identities values(true)");assert.equal((await d.query("select humbel_permissions.can_perform('synthetic','read',$1) v",[ids.d2])).rows[0].v,false);});
 await assert.rejects(f.asUser(ids.worker,d=>d.query('select humbel_permissions.actor_context($1)',[ids.admin])),e=>e.code==='42883');
}));

test('8B1 DATA: constraints and failed batches preserve existing/no-login staff, primary department and grants; install drift rolls back',async()=>use(async f=>{
 await f.db.query("insert into public.employees(display_name,role,auth_user_id) values('No login synthetic','employee',null)");
 const before=(await f.db.query('select * from public.employees order by id')).rows;await ready(f);
 await f.db.exec('begin');await assert.rejects(async()=>{await f.db.query("insert into humbel_permissions.memberships values($1,$2,true,false,now(),null,'test')",[ids.worker,ids.d2]);await f.db.query("insert into humbel_permissions.organization_roles(id,user_id,role_code,valid_from,evidence) values($1,$2,'unrecognized',now(),'bad')",[randomUUID(),ids.worker]);},e=>e.code==='23503');await f.db.exec('rollback');
 assert.equal((await f.db.query('select count(*)::int n from humbel_permissions.memberships where department_id=$1',[ids.d2])).rows[0].n,0);
 assert.deepEqual((await f.db.query('select * from public.employees order by id')).rows,before);
 await f.db.exec('grant select on humbel_permissions.identities to authenticated');await assert.rejects(review.apply(f.db),/identities\/authenticated\/SELECT/);
 assert.deepEqual((await f.db.query('select * from public.employees order by id')).rows,before);
}));

test('8B1 ATOMIC INSTALL: wrong target and injected install failure leave no new schema, unchanged legacy definitions/data',async()=>{
 const {originalFixture}=require('./step7-fixture');const baseline=require('../db-integration/security-step7-review');const f=await originalFixture();
 try{
  await baseline.apply(f.db);const oldHash=await prior.snapshot(f.db),oldRows=(await f.db.query('select * from public.employees order by id')).rows;
  await f.db.exec("select set_config('humbel.isolated_sentinel','',false)");await deny(review.apply(f.db));
  assert.equal((await f.db.query("select to_regnamespace('humbel_permissions') v")).rows[0].v,null);
  await f.db.exec("select set_config('humbel.isolated_sentinel','synthetic-only-8b1-failure-sentinel-over-32',false)");
  const faulty={query:(...a)=>f.db.query(...a),exec:s=>f.db.exec(s.includes('create schema humbel_permissions;')?s+'\nselect 1/0;':s)};
  await assert.rejects(review.apply(faulty),e=>e.code==='22012');
  assert.equal((await f.db.query("select to_regnamespace('humbel_permissions') v")).rows[0].v,null);
  assert.equal(await prior.snapshot(f.db),oldHash);assert.deepEqual((await f.db.query('select * from public.employees order by id')).rows,oldRows);
  await review.apply(f.db);await f.db.exec('alter table humbel_permissions.identities add column unexpected text');
  await assert.rejects(review.apply(f.db),/8B1 drift/);
 }finally{await f.db.close();}
});
