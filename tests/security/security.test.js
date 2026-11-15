const {test,after}=require('node:test');
const assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto');
const {fixture,ids}=require('./fixture');
let f;
after(async()=>{if(f)await f.db.close();});
const denied=promise=>assert.rejects(promise,e=>e.code==='42501');
const readSchedule=who=>f.asUser(who,d=>d.query('select * from public.security_published_schedule()'));

test('BASELINE ONLY: captured permissive policies reproduce foreign role escalation and anon DML',async()=>{
 f=await fixture();
 await f.asUser(ids.worker,d=>d.query("update public.employees set role='admin' where id=$1",[ids.other]));
 assert.equal((await f.db.query('select role from public.employees where id=$1',[ids.other])).rows[0].role,'admin');
 await f.db.query("update public.employees set role='employee' where id=$1",[ids.other]);
 assert.equal((await f.asUser(null,d=>d.query("select has_table_privilege('anon','public.planner_assignments','UPDATE') allowed"),{anonymous:true})).rows[0].allowed,true);
});

test('CORRECTED: all ten RLS flags active; anon table/column rights and private authority access denied',async()=>{
 await f.db.exec('grant update(role) on public.employees to anon,authenticated');
 await f.install();
 assert.equal((await f.db.query("select has_column_privilege('anon','public.employees','role','UPDATE') allowed")).rows[0].allowed,false);
 const tables=['planner_absence_replacements','planner_absences','planner_assignments','planner_availability','planner_saturday_requests',
 'planner_shift_cancellations','planner_sick_leaves','planner_swaps','planner_vacations','tool_materials'];
 for(const table of tables){
  assert.equal((await f.db.query('select relrowsecurity from pg_class where oid=$1::regclass',['public.'+table])).rows[0].relrowsecurity,true);
  assert.equal((await f.db.query("select has_table_privilege('anon',$1,'SELECT,INSERT,UPDATE,DELETE') allowed",['public.'+table])).rows[0].allowed,false);
  await denied(f.asUser(null,d=>d.query('select * from public.'+table),{anonymous:true}));
 }
 await denied(f.asUser(ids.worker,d=>d.query('select * from humbel_security.principals')));
 await denied(f.asUser(ids.admin,d=>d.query("update humbel_security.principals set role='admin' where user_id=$1",[ids.worker])));
});

test('CORRECTED: unauthenticated, spoofed, unenrolled, inactive and mismatched actors fail closed',async()=>{
 await denied(f.asUser(null,d=>d.query('select public.security_own_contacts()'),{anonymous:true}));
 await denied(f.asUser(null,d=>d.query('select public.security_own_contacts()')));
 assert.equal((await f.asUser(ids.orphan,d=>d.query('select public.is_admin() allowed'),{claims:{user_metadata:{role:'admin'}}})).rows[0].allowed,false);
 await denied(f.assign(ids.orphan,randomUUID()));
 await f.db.query('update public.employees set is_active=false where id=$1',[ids.worker]);
 try{await denied(f.asUser(ids.worker,d=>d.query('select public.security_own_contacts()')));}finally{await f.db.query('update public.employees set is_active=true where id=$1',[ids.worker]);}
 await f.db.query("update public.employees set role='admin' where id=$1",[ids.worker]);
 try{assert.equal((await f.asUser(ids.worker,d=>d.query('select public.is_admin() allowed'))).rows[0].allowed,false);}finally{await f.db.query("update public.employees set role='employee' where id=$1",[ids.worker]);}
});

test('CORRECTED: ordinary staff cannot change own/foreign protected fields, create admins, or change leaders',async()=>{
 for(const target of [ids.worker,ids.other]){
  for(const field of ["role='admin'",`auth_user_id='${ids.orphan}'`,`department_id='${ids.d2}'`,"is_active=false"]){
   const r=await f.asUser(ids.worker,d=>d.query('update public.employees set '+field+' where id=$1 returning id',[target]));
   assert.deepEqual(r.rows,[]);
  }
 }
 await denied(f.asUser(ids.worker,d=>d.query("insert into public.employees(display_name,role) values('unauthorized','admin')")));
 assert.deepEqual((await f.asUser(ids.worker,d=>d.query('delete from public.employees where id=$1 returning id',[ids.other]))).rows,[]);
 assert.deepEqual((await f.asUser(ids.leader,d=>d.query('update public.departments set leader_employee_id=$1 where id=$2 returning id',[ids.worker,ids.d1]))).rows,[]);
 assert.deepEqual((await f.asUser(ids.worker,d=>d.query('select id from public.employees'))).rows.map(r=>r.id),[ids.worker]);
});

test('CORRECTED: own-department planning works; employee, scanner, foreign-department and foreign-subject writes denied',async()=>{
 const r=await f.assign(ids.leader);assert.equal(r.rows[0].result.version,1);
 await denied(f.assign(ids.worker,randomUUID()));
 await denied(f.assign(ids.scanner,randomUUID()));
 await denied(f.assign(ids.leader,randomUUID(),0,ids.d2,ids.other));
 await denied(f.assign(ids.leader,randomUUID(),0,ids.d1,ids.other));
 await denied(f.asUser(ids.leader,d=>d.query("update public.planner_assignments set assigned_user='other' where id=$1",[ids.assignment])));
 await denied(f.asUser(ids.leader,d=>d.query("update public.planner_shift_cancellations set reason='unauthorized'")));
 const admin=await f.assign(ids.admin,randomUUID(),0,ids.d2,ids.other);assert.equal(admin.rows[0].result.version,1);
});

test('CORRECTED: deputy needs explicit, current same-department delegation to an active verified leader',async()=>{
 await denied(f.assign(ids.deputy,randomUUID()));
 await f.db.query("insert into humbel_security.delegations values($1,$2,$3,now()-interval '1 day',now()+interval '1 day',false,'Synthetic approved planning delegation')",[ids.deputy,ids.d1,ids.leader]);
 assert.equal((await f.assign(ids.deputy,randomUUID())).rows[0].result.version,1);
 await denied(f.assign(ids.deputy,randomUUID(),0,ids.d2,ids.other));
 for(const mutation of ["revoked=true","valid_until=now()-interval '1 second'","valid_from=now()+interval '1 hour'"]){
  await f.db.query('update humbel_security.delegations set '+mutation+' where user_id=$1',[ids.deputy]);
  await denied(f.assign(ids.deputy,randomUUID()));
  await f.db.query("update humbel_security.delegations set revoked=false,valid_from=now()-interval '1 day',valid_until=now()+interval '1 day' where user_id=$1",[ids.deputy]);
 }
 await f.db.query('update humbel_security.principals set enabled=false where user_id=$1',[ids.leader]);
 try{await denied(f.assign(ids.deputy,randomUUID()));}finally{await f.db.query('update humbel_security.principals set enabled=true where user_id=$1',[ids.leader]);}
});

test('CORRECTED: publication, planning and scoped schedule reading are separate; no health/profile fields leak',async()=>{
 assert.deepEqual((await readSchedule(ids.worker)).rows,[]);
 await denied(f.asUser(ids.leader,d=>d.query('select public.security_publish_assignment($1,1)',[ids.assignment])));
 for(const [who,read,publish] of [[ids.worker,true,false],[ids.admin,false,true]])await f.db.query("insert into humbel_security.schedule_grants values($1,$2,$3,$4,now()-interval '1 day',now()+interval '1 day',false,'Synthetic field/scope approval')",[who,ids.d1,read,publish]);
 assert.deepEqual((await readSchedule(ids.worker)).rows,[]);
 await f.asUser(ids.admin,d=>d.query('select public.security_publish_assignment($1,1)',[ids.assignment]));
 const rows=(await readSchedule(ids.worker)).rows;assert.equal(rows.length,1);
 assert.deepEqual(Object.keys(rows[0]).sort(),['assignment_id','employee_id','shift_date','shift_id']);
 assert.equal(rows[0].employee_id,ids.worker);assert.deepEqual((await readSchedule(ids.other)).rows,[]);
 await denied(f.asUser(ids.worker,d=>d.query('select public.security_publish_assignment($1,1)',[ids.assignment])));
 await f.assign(ids.leader,ids.assignment,1);assert.deepEqual((await readSchedule(ids.worker)).rows,[]);
});

test('CORRECTED: approved own contact selfservice works without mutable security fields or foreign target',async()=>{
 await denied(f.asUser(ids.worker,d=>d.query('select public.security_own_contacts($1)',[{synthetic_contact:'value'}])));
 await f.db.query("insert into humbel_security.contact_fields values('synthetic_contact','Test-only field approval; no production field decision')");
 const r=await f.asUser(ids.worker,d=>d.query('select public.security_own_contacts($1) data',[{synthetic_contact:'value'}]));assert.deepEqual(r.rows[0].data,{synthetic_contact:'value'});
 for(const key of ['role','auth_user_id','department_id','is_active','employee_id','permissions'])await denied(f.asUser(ids.worker,d=>d.query('select public.security_own_contacts($1)',[{[key]:'admin'}])));
 assert.deepEqual((await f.asUser(ids.other,d=>d.query('select public.security_own_contacts() data'))).rows[0].data,{});
 await denied(f.asUser(ids.scanner,d=>d.query('select public.security_own_contacts()')));
 assert.equal((await f.db.query('select role from public.employees where id=$1',[ids.worker])).rows[0].role,'employee');
});

test('CORRECTED: materials remain readable for staff/scanners; only verified admin edits reference data',async()=>{
 const material=randomUUID();await f.db.query("insert into public.tool_materials(id,name) values($1,'Synthetic material')",[material]);
 for(const user of [ids.worker,ids.leader,ids.scanner])assert.equal((await f.asUser(user,d=>d.query('select id from public.tool_materials where id=$1',[material]))).rows.length,1);
 await denied(f.asUser(ids.worker,d=>d.query("insert into public.tool_materials(name) values('unauthorized')")));
 await f.asUser(ids.admin,d=>d.query("update public.tool_materials set name='Reviewed synthetic material' where id=$1",[material]));
 assert.equal((await f.db.query('select name from public.tool_materials where id=$1',[material])).rows[0].name,'Reviewed synthetic material');
});

test('CORRECTED: unknown legacy bindings and unresolved private workflows stay quarantined',async()=>{
 const id=randomUUID();await f.db.query("insert into public.planner_sick_leaves(id,user_name,from_date,to_date) values($1,'unresolved','2026-11-12','2026-11-13')",[id]);
 for(const user of [ids.worker,ids.leader,ids.admin])assert.deepEqual((await f.asUser(user,d=>d.query('select * from public.planner_sick_leaves where id=$1',[id]))).rows,[]);
 await denied(f.asUser(ids.worker,d=>d.query("insert into public.planner_availability(shift_id,shift_date,user_name,status) values('s','2026-11-12','other','yes')")));
 await assert.rejects(f.assign(ids.leader,ids.assignment,0),e=>e.code==='40001');
 await denied(f.assign(ids.leader,randomUUID(),null));
});

test('CORRECTED: audit failure rolls back the plan and version; repeated stale command cannot overwrite',async()=>{
 const id=randomUUID();
 await f.db.exec("create function humbel_security.synthetic_abort() returns trigger language plpgsql as $$begin raise exception 'Synthetic audit failure';end$$;create trigger synthetic_abort before insert on humbel_security.audit for each row execute function humbel_security.synthetic_abort()");
 try{await assert.rejects(f.assign(ids.leader,id));assert.equal((await f.db.query('select id from public.planner_assignments where id=$1',[id])).rows.length,0);assert.equal((await f.db.query('select row_id from humbel_security.record_bindings where row_id=$1',[id])).rows.length,0);}
 finally{await f.db.exec('drop trigger synthetic_abort on humbel_security.audit;drop function humbel_security.synthetic_abort()');}
});

test('CORRECTED: verified admin retains staff administration without automatically enrolling a new privilege holder',async()=>{
 const updated=await f.asUser(ids.admin,d=>d.query("update public.employees set first_name='Synthetic revised name' where id=$1 returning id",[ids.worker]));
 assert.equal(updated.rows[0].id,ids.worker);
 const r=await f.asUser(ids.leader,d=>d.query("update public.employees set role='admin' where id=$1 returning id",[ids.leader]));assert.deepEqual(r.rows,[]);
 await f.asUser(ids.admin,d=>d.query("update public.employees set role='admin' where id=$1",[ids.worker]));
 try{assert.equal((await f.asUser(ids.worker,d=>d.query('select public.is_admin() allowed'))).rows[0].allowed,false);}
 finally{await f.db.query("update public.employees set role='employee' where id=$1",[ids.worker]);}
});

test('CORRECTED: explicit private read grants require mapped scope and expire; they confer no raw writing',async()=>{
 const id=randomUUID();
 await f.db.query("insert into public.planner_sick_leaves(id,user_name,from_date,to_date) values($1,'Synthetic privacy fixture','2026-11-12','2026-11-13')",[id]);
 await f.db.query("insert into humbel_security.record_bindings values('planner_sick_leaves',$1,$2,$3,1,'Synthetic explicit mapping')",[id,ids.d1,ids.worker]);
 await f.db.query("insert into humbel_security.record_grants values($1,$2,'planner_sick_leaves',true,false,now()-interval '1 day',now()+interval '1 day',false,'Synthetic privacy approval only')",[ids.worker,ids.d1]);
 assert.equal((await f.asUser(ids.worker,d=>d.query('select id from public.planner_sick_leaves where id=$1',[id]))).rows.length,1);
 assert.deepEqual((await f.asUser(ids.leader,d=>d.query('select id from public.planner_sick_leaves where id=$1',[id]))).rows,[]);
 await denied(f.asUser(ids.worker,d=>d.query("update public.planner_sick_leaves set user_name='other' where id=$1",[id])));
 await f.db.query("update humbel_security.record_grants set valid_until=now()-interval '1 second' where user_id=$1",[ids.worker]);
 assert.deepEqual((await f.asUser(ids.worker,d=>d.query('select id from public.planner_sick_leaves where id=$1',[id]))).rows,[]);
});
