// Native Phase 2 extension composed after the unchanged Phase 1 cases in the pinned isolated stack.
const fs=require('node:fs'),assert=require('node:assert/strict');const {randomUUID}=require('node:crypto');
const {rpc,signup,http,connect,sqlAs}=require('../../scripts/db-integration/client');
async function run(context){
 await require('./ba-entry-extension').run(context);
 const {db,s,check,results}=context,p=s.people;
 await db.query(fs.readFileSync('supabase/review/op-routing-phase2.sql','utf8'));await db.query("notify pgrst,'reload schema'");await new Promise(r=>setTimeout(r,1200));
 const leader=await signup('synthetic-routing-leader'),deputy=await signup('synthetic-routing-deputy');leader.employee_id=randomUUID();deputy.employee_id=randomUUID();
 await db.query("insert into public.employees(id,auth_user_id,display_name,role,department_id) values($1,$2,'Synthetic receiving leader','department_admin',$3),($4,$5,'Synthetic receiving deputy','employee',$3)",[leader.employee_id,leader.id,s.foreign,deputy.employee_id,deputy.id]);
 await db.query("update public.employees set role='department_admin' where id=$1",[p.reviewer.employee_id]);
 await db.query('update public.departments set leader_employee_id=$1 where id=$2',[p.reviewer.employee_id,s.dept]);await db.query('update public.departments set leader_employee_id=$1 where id=$2',[leader.employee_id,s.foreign]);
 const target=randomUUID();await db.query("insert into public.production_machines(id,name,machine_code,department_id) values($1,'Synthetic foreign suitable machine','SYN-FOREIGN',$2)",[target,s.foreign]);
 await db.query("insert into public.production_routing_capabilities values($1,'SYNTHETIC'),($2,'SYNTHETIC')",[s.machines[1],target]);
 await db.query("insert into public.production_routing_abort_reasons values('synthetic-change','Synthetic configured reason',false,true)");
 await db.query("insert into public.production_routing_authorities values($1,$2,'leader',null,now()-interval '1 day',now()+interval '1 day',false,$3),($4,$5,'leader',null,now()-interval '1 day',now()+interval '1 day',false,$3),($6,$5,'deputy',$4,now()-interval '1 day',now()+interval '1 day',false,$3),($3,$2,'admin',null,now()-interval '1 day',now()+interval '1 day',false,$3),($3,$5,'admin',null,now()-interval '1 day',now()+interval '1 day',false,$3)",[p.reviewer.id,s.dept,p.manager.id,leader.id,s.foreign,deputy.id]);
 await db.query("insert into public.production_readiness_grants values($1,$2,true,true,now()+interval '1 day',false,$3),($4,$5,true,true,now()+interval '1 day',false,$3)",[p.worker.id,s.machines[1],p.manager.id,p.outsider.id,target]);
 const create=async ba=>rpc(p.worker,'ba_entry_command',{p_action:'create',p_ba_number:ba,p_article_number:'SYNTH-ACCEPT',p_quantity:7});
 const command=(u,op,action,machine,version=0,extra={})=>rpc(u,'op_routing_command',{p_operation_id:op,p_action:action,p_request_id:randomUUID(),p_expected_version:version,p_machine_id:machine,p_reason:'Synthetic routing integration evidence',...extra});
 let snap,op,route;
 await check('Phase 2 defaults off and raw routing metadata / anon access is denied',async()=>{
  assert.deepEqual((await rpc(p.worker,'op_routing_view')).rows,[]);
  for(const table of ['production_operation_routing','production_operation_routing_events','production_operation_activity_segments','production_routing_authorities'])assert.ok([401,403].includes((await http(p.worker,'/'+table+'?limit=1')).status));
  assert.ok([401,403].includes((await http(null,'/rpc/op_routing_view',{method:'POST',body:{}})).status));
  await db.query('update public.production_routing_settings set enabled=true');snap=await create('812240');op=snap.op_rows[0].id;
 });
 await check('Real employee JWT selects own suitable machine; original OP/BA/counter data remains unchanged',async()=>{
  const row=(await rpc(p.worker,'op_routing_view')).rows.find(x=>x.operation_id===op);assert.equal(row.machine_id,s.machines[0]);
  route=await command(p.worker,op,'select',s.machines[1]);assert.equal(route.executing_department_id,s.dept);
  assert.equal((await db.query('select good_qty from public.production_station_counts where order_id=$1',[s.legacy])).rows[0].good_qty,17);
  assert.equal((await db.query('select scrap_total from public.production_order_stations where id=$1',[s.station])).rows[0].scrap_total,3);
  assert.deepEqual((await db.query('select operations from public.production_order_article_snapshots where order_id=$1',[snap.order_id])).rows[0].operations,snap.operations);
 });
 await check('AL cross-department binding transfers operative authority and notifications; old assign RPC cannot bypass binding',async()=>{
  route=await command(p.reviewer,op,'fix',target,route.row_version);
  assert.equal(route.binding,true);assert.equal(route.executing_department_id,s.foreign);
  assert((await rpc(leader,'op_routing_view')).notifications.some(n=>n.operation_id===op));
  assert.equal((await rpc(p.reviewer,'op_routing_view')).rows.find(r=>r.operation_id===op).can_manage,false);
  await assert.rejects(command(p.worker,op,'select',s.machines[1],route.row_version));await assert.rejects(command(p.outsider,op,'select',target,route.row_version));
  const e=(await db.query('select * from public.production_operation_execution where operation_id=$1',[op])).rows[0];
  await assert.rejects(rpc(p.manager,'machine_readiness_command',{p_operation_id:op,p_action:'assign',p_request_id:randomUUID(),p_expected_version:e.row_version,p_machine_id:s.machines[0],p_reason:'Synthetic bypass must fail'}),e=>e.code==='42501');
 });
 await check('Receiving employee uses existing physical handover/start API; return interrupts all active setup segments at one instant',async()=>{
  let e=(await db.query('select * from public.production_operation_execution where operation_id=$1',[op])).rows[0];
  e=await rpc(p.outsider,'machine_readiness_command',{p_operation_id:op,p_action:'handover',p_request_id:randomUUID(),p_expected_version:e.row_version,p_quantity:7,p_reason:'Synthetic physical foreign-machine receipt'});
  e=await rpc(p.outsider,'machine_readiness_command',{p_operation_id:op,p_action:'start',p_request_id:randomUUID(),p_expected_version:e.row_version,p_reason:'Synthetic explicit start'});assert.equal(e.work_status,'running');
  const section=(await db.query('select id from public.production_operation_machine_sections where operation_id=$1 and closed_at is null',[op])).rows[0].id;
  await db.query("insert into public.production_operation_activity_segments(section_id,user_id,activity,started_at) values($1,$2,'setup_internal',now()-interval '2 minutes'),($1,$3,'setup_external',now()-interval '1 minute')",[section,p.outsider.id,leader.id]);
  await assert.rejects(command(leader,op,'return',s.machines[0],route.row_version),e=>e.code==='22023');
  await db.query("update public.production_routing_authorities set valid_until=now()-interval '1 second' where user_id=$1",[deputy.id]);await assert.rejects(command(deputy,op,'return',s.machines[0],route.row_version,{p_abort_code:'synthetic-change'}));
  await db.query("update public.production_routing_authorities set valid_until=now()+interval '1 day' where user_id=$1",[deputy.id]);
  route=await command(deputy,op,'return',s.machines[0],route.row_version,{p_abort_code:'synthetic-change'});
  const rows=(await db.query('select ended_at from public.production_operation_activity_segments where section_id=$1',[section])).rows;assert(rows.every(r=>r.ended_at));assert.equal(rows[0].ended_at.getTime(),rows[1].ended_at.getTime());
  assert.equal((await db.query('select work_status from public.production_operation_execution where operation_id=$1',[op])).rows[0].work_status,'paused');
  assert((await rpc(p.reviewer,'op_routing_view')).notifications.some(n=>n.operation_id===op&&n.kind==='returned'));
  await assert.rejects(db.query("insert into public.production_operation_activity_segments(section_id,user_id,activity,started_at) values($1,$2,'production',now())",[section,p.worker.id]),e=>e.code==='PT409');
 });
 await check('Two independent PostgreSQL routing requests: duplicate idempotency and stale-version exclusion',async()=>{
  const fresh=await create('812241'),oid=fresh.op_rows[0].id,request=randomUUID();
  const c1=await connect(),c2=await connect();
  try{
   const params=[oid,'select',request,0,s.machines[1],'Synthetic concurrent machine choice',null];
   const replies=await Promise.all([c1,c2].map(c=>sqlAs(p.worker,c=>c.query('select public.op_routing_command($1,$2,$3,$4,$5,$6,$7) result',params),{connection:c})));
   assert.deepEqual(replies[0].rows[0].result,replies[1].rows[0].result);
   assert.equal(Number((await db.query('select count(*) n from public.production_operation_routing_events where operation_id=$1',[oid])).rows[0].n),1);
   await assert.rejects(command(p.worker,oid,'select',s.machines[0],0),e=>e.code==='PT409'&&e.status===409);
  }finally{await c1.end();await c2.end();}
 });
 await check('Notification/audit failures roll back machine sections and active time interruption atomically',async()=>{
  const fresh=await create('812242'),oid=fresh.op_rows[0].id;
  await db.query("create function public.synthetic_routing_notify_abort() returns trigger language plpgsql as $$begin raise exception 'Synthetic notify failure';end$$;create trigger synthetic_routing_notify_abort before insert on public.production_operation_routing_notifications for each row execute function public.synthetic_routing_notify_abort()");
  try{await assert.rejects(command(p.reviewer,oid,'fix',target));assert.equal(Number((await db.query('select count(*) n from public.production_operation_routing where operation_id=$1',[oid])).rows[0].n),0);}
  finally{await db.query('drop trigger synthetic_routing_notify_abort on public.production_operation_routing_notifications;drop function public.synthetic_routing_notify_abort()');}
  const r=await command(p.worker,oid,'select',s.machines[0]);const section=(await db.query('select id from public.production_operation_machine_sections where operation_id=$1 and closed_at is null',[oid])).rows[0].id;
  await db.query("insert into public.production_operation_activity_segments(section_id,user_id,activity,started_at) values($1,$2,'setup_internal',now()-interval '1 minute')",[section,p.worker.id]);
  await db.query("create function public.synthetic_routing_audit_abort() returns trigger language plpgsql as $$begin raise exception 'Synthetic audit failure';end$$;create trigger synthetic_routing_audit_abort before insert on public.production_operation_routing_events for each row execute function public.synthetic_routing_audit_abort()");
  try{await assert.rejects(command(p.worker,oid,'select',s.machines[1],r.row_version,{p_abort_code:'synthetic-change'}));assert.equal((await db.query('select machine_id from public.production_operation_execution where operation_id=$1',[oid])).rows[0].machine_id,s.machines[0]);assert.equal((await db.query('select ended_at from public.production_operation_activity_segments where section_id=$1',[section])).rows[0].ended_at,null);}
  finally{await db.query('drop trigger synthetic_routing_audit_abort on public.production_operation_routing_events;drop function public.synthetic_routing_audit_abort()');}
 });
 await check('Mandatory machine and mutable-role escalation cannot authorize a routing change',async()=>{
  const fresh=await create('812243'),oid=fresh.op_rows[0].id;await db.query("insert into public.production_operation_machine_constraints values($1,$2,'Synthetic confirmed mandatory constraint')",[oid,s.machines[0]]);
  await assert.rejects(command(p.reviewer,oid,'fix',target),e=>e.code==='42501');
  await db.query("update public.employees set role='admin' where id=$1",[p.worker.employee_id]);try{await assert.rejects(command(p.worker,oid,'fix',s.machines[0]),e=>e.code==='42501');}finally{await db.query("update public.employees set role='employee' where id=$1",[p.worker.employee_id]);}
 });
 results.limitations.push('Phase 2: trusted AL/deputy and machine-suitability provisioning needs separate approval; Phase 3 timer UI and Phase 4 counter-event attribution are not implemented. No production/authoritative restore claim.');
}
module.exports={run};
