// Native-only candidate verification. Original diagnosis and BA/OP cases run first.
const fs=require('node:fs'),assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto');
const {rpc,http,connect,sqlAs,signup}=require('../../scripts/db-integration/client');
async function run(context){
 await require('./integration-dispatch').run(context,'ba-op');
 const {db,s,check,results}=context,p=s.people;
 assert.equal(results.tests.filter(t=>t.status==='failed').length,0,'Baseline/BA-OP failures block candidate installation');
 const baseline=results.securityStatus,start=results.tests.length;
 await db.query(fs.readFileSync('supabase/review/security-step3-isolated.sql','utf8'));
 // Explicit known synthetic fixture enrollment, NEVER a bootstrap algorithm.
 const expected={manager:'admin',reviewer:'department_admin',worker:'employee',outsider:'employee',scanner:'tool_scanner'};
 for(const [name,role] of Object.entries(expected)){
  const row=(await db.query('select role,department_id,auth_user_id from public.employees where id=$1',[p[name].employee_id])).rows[0];
  assert.equal(row.role,role);assert.equal(row.auth_user_id,p[name].id);
  await db.query("insert into humbel_security.principals values($1,$2,$3,$4,true,now()+interval '1 day','Independent known synthetic fixture enrollment')",[p[name].id,p[name].employee_id,role,row.department_id]);
 }
 const deputy=await signup('synthetic-step3-deputy');deputy.employee_id=randomUUID();
 await db.query("insert into public.employees(id,auth_user_id,display_name,role,department_id) values($1,$2,'Synthetic step3 deputy','employee',$3)",[deputy.employee_id,deputy.id,s.dept]);
 await db.query("insert into humbel_security.principals values($1,$2,'employee',$3,true,now()+interval '1 day','Independent synthetic deputy enrollment')",[deputy.id,deputy.employee_id,s.dept]);
 await db.query("notify pgrst,'reload schema'");await new Promise(r=>setTimeout(r,1000));
 const assign=(user,id,version=0,dept=s.dept,subject=p.worker.employee_id)=>rpc(user,'security_assignment_command',{
  p_id:id,p_department:dept,p_subject:subject,p_shift_id:'step3-'+id,p_date:'2026-11-12',p_expected_version:version});
 const deny=promise=>assert.rejects(promise,e=>e.code==='42501'||[401,403].includes(e.status));
 const assignment=randomUUID();
 await check('Step3 CORRECTED: ten protected tables deny real anonymous JWT REST/SQL access and all raw private authority access',async()=>{
  for(const table of ['planner_absence_replacements','planner_absences','planner_assignments','planner_availability','planner_saturday_requests','planner_shift_cancellations','planner_sick_leaves','planner_swaps','planner_vacations','tool_materials']){
   assert.equal((await db.query('select relrowsecurity from pg_class where oid=$1::regclass',['public.'+table])).rows[0].relrowsecurity,true);
   assert.equal((await db.query("select has_table_privilege('anon',$1,'SELECT,INSERT,UPDATE,DELETE') allowed",['public.'+table])).rows[0].allowed,false);
   assert.ok([401,403].includes((await http(null,'/'+table+'?limit=1')).status));
  }
  await assert.rejects(sqlAs(p.worker,c=>c.query('select * from humbel_security.principals')),e=>e.code==='42501');
  await deny(rpc(null,'security_own_contacts',{}));
 });
 await check('Step3 CORRECTED: user_metadata spoof cannot change own/foreign role, department, identity or leader',async()=>{
  for(const target of [p.worker,p.outsider]){
   const before=(await db.query('select role,department_id,auth_user_id,is_active from public.employees where id=$1',[target.employee_id])).rows[0];
   const response=await http(p.worker,'/employees?id=eq.'+target.employee_id,{method:'PATCH',body:{role:'admin',department_id:s.foreign,auth_user_id:p.scanner.id,is_active:false},headers:{prefer:'return=representation'}});
   assert.ok([200,204,401,403].includes(response.status));
   assert.deepEqual((await db.query('select role,department_id,auth_user_id,is_active from public.employees where id=$1',[target.employee_id])).rows[0],before);
  }
  const r=await http(p.reviewer,'/departments?id=eq.'+s.dept,{method:'PATCH',body:{leader_employee_id:p.worker.employee_id},headers:{prefer:'return=representation'}});
  assert.ok([200,204,401,403].includes(r.status));
  assert.equal((await db.query('select leader_employee_id from public.departments where id=$1',[s.dept])).rows[0].leader_employee_id,p.reviewer.employee_id);
 });
 await check('Step3 CORRECTED: own-department AL writes audited plan; employee/scanner and foreign-subject writes denied',async()=>{
  assert.equal((await assign(p.reviewer,assignment)).version,1);
  for(const user of [p.worker,p.scanner])await deny(assign(user,randomUUID()));
  await deny(assign(p.reviewer,randomUUID(),0,s.foreign,p.outsider.employee_id));
  await deny(assign(p.reviewer,randomUUID(),0,s.dept,p.outsider.employee_id));
  assert.ok([401,403].includes((await http(p.reviewer,'/planner_assignments?id=eq.'+assignment,{method:'PATCH',body:{assigned_user:'Synthetic outsider'}})).status));
 });
 await check('Step3 CORRECTED: explicit deputy is bounded by department, time, principal and revocation',async()=>{
  await deny(assign(deputy,randomUUID()));
  await db.query("insert into humbel_security.delegations values($1,$2,$3,now()-interval '1 day',now()+interval '1 day',false,'Synthetic approved delegation')",[deputy.id,s.dept,p.reviewer.id]);
  assert.equal((await assign(deputy,randomUUID())).version,1);
  await deny(assign(deputy,randomUUID(),0,s.foreign,p.outsider.employee_id));
  await db.query('update humbel_security.delegations set revoked=true where user_id=$1',[deputy.id]);await deny(assign(deputy,randomUUID()));
  await db.query("update humbel_security.delegations set revoked=false,valid_until=now()-interval '1 second' where user_id=$1",[deputy.id]);await deny(assign(deputy,randomUUID()));
 });
 await check('Step3 CORRECTED: publish permission is distinct; staff sees only explicitly released minimal schedule',async()=>{
  assert.deepEqual(await rpc(p.worker,'security_published_schedule',{}),[]);
  await deny(rpc(p.reviewer,'security_publish_assignment',{p_id:assignment,p_expected_version:1}));
  await db.query("insert into humbel_security.schedule_grants values($1,$3,true,false,now()-interval '1 day',now()+interval '1 day',false,'Synthetic read approval'),($2,$3,false,true,now()-interval '1 day',now()+interval '1 day',false,'Synthetic publish approval')",[p.worker.id,p.manager.id,s.dept]);
  await rpc(p.manager,'security_publish_assignment',{p_id:assignment,p_expected_version:1});
  const rows=await rpc(p.worker,'security_published_schedule',{});assert.equal(rows.length,1);assert.equal(rows[0].employee_id,p.worker.employee_id);
  assert.deepEqual(Object.keys(rows[0]).sort(),['assignment_id','employee_id','shift_date','shift_id']);
  assert.deepEqual(await rpc(p.outsider,'security_published_schedule',{}),[]);
  await deny(rpc(p.worker,'security_publish_assignment',{p_id:assignment,p_expected_version:1}));
 });
 await check('Step3 CORRECTED: own approved contact field works, protected/foreign fields denied, tool reference access retained',async()=>{
  await deny(rpc(p.worker,'security_own_contacts',{p_values:{synthetic_contact:'value'}}));
  await db.query("insert into humbel_security.contact_fields values('synthetic_contact','Test-only approval; production fields undecided')");
  assert.deepEqual(await rpc(p.worker,'security_own_contacts',{p_values:{synthetic_contact:'value'}}),{synthetic_contact:'value'});
  for(const field of ['role','auth_user_id','department_id','employee_id','permissions'])await deny(rpc(p.worker,'security_own_contacts',{p_values:{[field]:'admin'}}));
  assert.deepEqual(await rpc(p.outsider,'security_own_contacts',{}),{});
  for(const user of [p.worker,p.scanner])assert.equal((await http(user,'/tool_materials?limit=1')).status,200);
  assert.ok([401,403].includes((await http(p.worker,'/tool_materials',{method:'POST',body:{name:'Unauthorized synthetic material'}})).status));
 });
 await check('Step3 CORRECTED: independent PostgreSQL connections serialize stale plan writes without lost updates',async()=>{
  const connections=await Promise.all([connect(),connect()]);
  try{
   const outcomes=await Promise.allSettled(connections.map(connection=>sqlAs(p.reviewer,c=>c.query('select public.security_assignment_command($1,$2,$3,$4,$5,1) result',[assignment,s.dept,p.worker.employee_id,'step3-'+assignment,'2026-11-12']),{connection})));
   assert.equal(outcomes.filter(x=>x.status==='fulfilled').length,1);
   assert.equal(outcomes.find(x=>x.status==='rejected').reason.code,'40001');
   assert.equal((await db.query("select row_version::int v from humbel_security.record_bindings where table_name='planner_assignments' and row_id=$1",[assignment])).rows[0].v,2);
   assert.deepEqual(await rpc(p.worker,'security_published_schedule',{}),[]);
  }finally{await Promise.all(connections.map(c=>c.end()));}
 });
 await check('Step3 CORRECTED: existing BA entry, cross-department OP machine binding and legacy counters survive hardening',async()=>{
  const made=await rpc(p.worker,'ba_entry_command',{p_action:'create',p_ba_number:'819900',p_article_number:'SYNTH-ACCEPT',p_quantity:2});
  assert.equal(made.op_rows.length,3);
  const target=(await db.query("select id from public.production_machines where machine_code='SYN-FOREIGN'")).rows[0].id;
  const route=await rpc(p.reviewer,'op_routing_command',{p_operation_id:made.op_rows[0].id,p_action:'fix',p_request_id:randomUUID(),p_expected_version:0,p_machine_id:target,p_reason:'Synthetic verified cross-department route'});
  assert.equal(route.executing_department_id,s.foreign);assert.equal(route.binding,true);
  await deny(rpc(p.worker,'op_routing_command',{p_operation_id:made.op_rows[0].id,p_action:'select',p_request_id:randomUUID(),p_expected_version:route.row_version,p_machine_id:s.machines[0],p_reason:'Synthetic unauthorized rebind'}));
  assert.equal((await db.query('select good_qty from public.production_station_counts where order_id=$1',[s.legacy])).rows[0].good_qty,17);
  assert.equal((await db.query('select scrap_total from public.production_order_stations where id=$1',[s.station])).rows[0].scrap_total,3);
 });
 results.step3CandidateStatus=results.tests.slice(start).every(t=>t.status==='passed')?'NATIVE_STRUCTURAL_MIRROR_CANDIDATE_VERIFIED__NO_PRODUCTION_CLEARANCE':'CANDIDATE_FAILED';
 results.securityStatus=baseline;
 results.limitations.push('Step3 is an isolated candidate, not an authoritative restore. Contact field choices, legacy bindings, UI adapters, private workflows, original ACL/owner/default-grant and full session/Realtime/tool regression remain release blockers.');
}
module.exports={run};
