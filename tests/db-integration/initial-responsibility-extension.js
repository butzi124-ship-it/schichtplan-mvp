// Native-only continuation: all 52 existing Phase 1–5 cases remain unchanged.
const fs=require('node:fs'),assert=require('node:assert/strict'),{randomUUID}=require('node:crypto');
const {rpc,http,connect,sqlAs,signup}=require('../../scripts/db-integration/client');
async function run(context){
 await require('./lifecycle-extension').run(context);
 const {db,s,check,results}=context,p=s.people;let n=818000;
 await db.query(fs.readFileSync('supabase/review/ba-initial-responsibility.sql','utf8'));await db.query("notify pgrst,'reload schema'");await new Promise(r=>setTimeout(r,1200));
 const create=(ba=String(n++),user=p.worker)=>rpc(user,'ba_entry_command',{p_action:'create',p_ba_number:ba,p_article_number:'SYNTH-ACCEPT',p_quantity:2});
 const view=async(snap,user=p.manager)=>(await rpc(user,'ba_lifecycle_view')).rows.find(r=>r.order_id===snap.order_id);
 const events=async snap=>(await db.query("select * from public.production_lifecycle_events where order_id=$1 and action='initial_responsibility'",[snap.order_id])).rows;
 const command=async(snap,user,action,data)=>rpc(user,'ba_lifecycle_command',{p_order_id:snap.order_id,p_action:action,p_request_id:randomUUID(),p_expected_version:(await view(snap)).row_version,p_data:data});
 const restore=()=>db.query("update public.production_routing_authorities set revoked=false,authority='leader',valid_from=now()-interval '1 day',valid_until=now()+interval '1 day' where user_id=$1 and department_id=$2",[p.reviewer.id,s.dept]);
 await check('Initial responsibility defaults off; existing BA not backfilled; raw and anonymous access denied',async()=>{
  const snap=await create();assert.equal(snap.responsibility_status,'pending_rule');assert.equal((await events(snap)).length,0);
  assert.ok([401,403].includes((await http(p.worker,'/production_initial_responsibility_settings?limit=1')).status));
  assert.ok([401,403].includes((await http(null,'/rpc/ba_entry_command',{method:'POST',body:{p_action:'create',p_ba_number:String(n++),p_article_number:'SYNTH-ACCEPT',p_quantity:2}})).status));
  await db.query('update public.production_initial_responsibility_settings set enabled=true');
  const opened=await rpc(p.worker,'ba_entry_command',{p_action:'open',p_ba_number:snap.ba_number});assert.equal(opened.responsibility_status,'open');assert.equal((await events(snap)).length,0);
 });
 await check('Real first-OP worker JWT supplies three fields; verified first-OP AL is saved and audited without manufacturing mutations',async()=>{
  const snap=await create();assert.equal(snap.principal_user_id,p.reviewer.id);assert.equal(snap.principal_department_id,s.dept);assert.equal(snap.responsibility_status,'assigned');assert.equal(snap.selected_op,10);assert.equal(snap.due_date,null);
  const ev=await events(snap);assert.equal(ev.length,1);assert.equal(ev[0].actor_id,p.worker.id);assert.equal(ev[0].payload.rule,'first_op_department_v1');assert.equal(ev[0].payload.authority.user_id,p.reviewer.id);assert.equal(ev[0].payload.outcome,'assigned');assert.equal(ev[0].before_value.principal_user_id,null);assert.equal(ev[0].after_value.principal_user_id,p.reviewer.id);
  for(const table of ['production_operation_routing','production_operation_machine_sections','production_order_stations','production_station_events']){const column=table.startsWith('production_operation_')?'operation_id':'order_id';const filter=column==='operation_id'?'operation_id in(select id from public.production_order_operations where order_id=$1)':'order_id=$1';assert.equal(Number((await db.query('select count(*) n from public.'+table+' where '+filter,[snap.order_id])).rows[0].n),0);}
  assert.ok((await rpc(p.reviewer,'ba_lifecycle_view')).notifications.some(x=>x.order_id===snap.order_id&&x.kind==='responsibility_assigned'));
 });
 await check('Expired, revoked, future and absent trusted leader grants leave BA open despite department_admin profile',async()=>{
  for(const patch of ["valid_until=now()-interval '1 second'","revoked=true","valid_from=now()+interval '1 hour',valid_until=now()+interval '2 hours'","authority='admin'"]){await db.query('update public.production_routing_authorities set '+patch+' where user_id=$1 and department_id=$2',[p.reviewer.id,s.dept]);try{const snap=await create();assert.equal(snap.responsibility_status,'open');assert.equal(snap.principal_user_id,null);assert.equal((await events(snap))[0].payload.outcome,'missing_valid_leader');}finally{await restore();}}
 });
 await check('Inactive or scanner leader and missing trusted department pointer never assigned',async()=>{
  for(const patch of ["is_active=false","role='tool_scanner'"]){await db.query('update public.employees set '+patch+' where id=$1',[p.reviewer.employee_id]);try{assert.equal((await create()).responsibility_status,'open');}finally{await db.query("update public.employees set is_active=true,role='department_admin' where id=$1",[p.reviewer.employee_id]);}}
  await db.query('update public.departments set leader_employee_id=null where id=$1',[s.dept]);try{const snap=await create();assert.equal(snap.responsibility_status,'open');assert.equal((await events(snap))[0].payload.outcome,'leader_mapping_mismatch');}finally{await db.query('update public.departments set leader_employee_id=$1 where id=$2',[p.reviewer.employee_id,s.dept]);}
 });
 await check('Contradictory independent leader authorities are ambiguous even with a preferred department pointer',async()=>{
  const other=await signup('synthetic-initial-ambiguous');await db.query("insert into public.employees(id,auth_user_id,display_name,role,department_id) values($1,$2,'Synthetic competing leader','department_admin',$3)",[randomUUID(),other.id,s.dept]);await db.query("insert into public.production_routing_authorities values($1,$2,'leader',null,now()-interval '1 day',now()+interval '1 day',false,$3)",[other.id,s.dept,p.manager.id]);
  try{const snap=await create();assert.equal(snap.responsibility_status,'open');assert.equal((await events(snap))[0].payload.outcome,'ambiguous_leader');assert.equal((await events(snap))[0].payload.candidate_count,2);}finally{await db.query('delete from public.production_routing_authorities where user_id=$1',[other.id]);}
 });
 await check('Missing first responsibility is recoverable only through existing audited admin assignment with catalog reason',async()=>{
  await db.query('update public.production_routing_authorities set revoked=true where user_id=$1 and department_id=$2',[p.reviewer.id,s.dept]);let snap;try{snap=await create();}finally{await restore();}
  await assert.rejects(command(snap,p.worker,'responsibility',{user_id:p.reviewer.id,reason_code:'synthetic-transfer'}),e=>e.code==='42501');await assert.rejects(command(snap,p.reviewer,'responsibility',{user_id:p.reviewer.id,reason_code:'synthetic-transfer'}),e=>e.code==='42501');await assert.rejects(command(snap,p.manager,'responsibility',{user_id:p.reviewer.id}),e=>e.code==='22023');await command(snap,p.manager,'responsibility',{user_id:p.reviewer.id,reason_code:'synthetic-transfer'});assert.equal((await view(snap)).principal_user_id,p.reviewer.id);assert.equal((await events(snap)).length,1);
 });
 await check('Two native connections create one BA, one initial assignment and one immutable event',async()=>{
  const ba=String(n++),c1=await connect(),c2=await connect();try{const both=await Promise.all([c1,c2].map(connection=>sqlAs(p.worker,c=>c.query("select public.ba_entry_command('create',$1,'SYNTH-ACCEPT',2,null) result",[ba]),{connection})));const one=both[0].rows[0].result,two=both[1].rows[0].result;assert.equal(one.order_id,two.order_id);assert.equal([one,two].filter(x=>!x.existing).length,1);assert.equal(one.principal_user_id,p.reviewer.id);assert.equal(two.principal_user_id,p.reviewer.id);assert.equal((await events(one)).length,1);}finally{await c1.end();await c2.end();}
 });
 await check('Locked authority revocation commits before assignment; stale valid grant cannot authorize selection',async()=>{
  const c1=await connect(),c2=await connect();try{await c1.query('begin');await c1.query('update public.production_routing_authorities set revoked=true where user_id=$1 and department_id=$2',[p.reviewer.id,s.dept]);const pid=(await c2.query('select pg_backend_pid() pid')).rows[0].pid;const pending=sqlAs(p.worker,c=>c.query("select public.ba_entry_command('create',$1,'SYNTH-ACCEPT',2,null) result",[String(n++)]),{connection:c2});pending.catch(()=>{});let blocked=false;for(let i=0;i<200&&!blocked;i++){blocked=(await db.query("select wait_event_type='Lock' blocked from pg_stat_activity where pid=$1",[pid])).rows[0]?.blocked===true;if(!blocked)await new Promise(r=>setTimeout(r,10));}assert.ok(blocked,'Creation must wait for the independently provisioned authority lock');await c1.query('commit');const snap=(await pending).rows[0].result;assert.equal(snap.responsibility_status,'open');assert.equal((await events(snap))[0].payload.outcome,'missing_valid_leader');}finally{await c1.query('rollback');await restore();await c1.end();await c2.end();}
 });
 await check('Audit failure rolls back BA, immutable snapshot and initial assignment atomically',async()=>{
  const ba=String(n++);await db.query("create function public.synthetic_initial_abort() returns trigger language plpgsql as $$begin if new.action='initial_responsibility' then raise exception 'Synthetic initial audit failure';end if;return new;end$$;create trigger synthetic_initial_abort before insert on public.production_lifecycle_events for each row execute function public.synthetic_initial_abort()");try{await assert.rejects(create(ba));assert.equal(Number((await db.query('select count(*) n from public.production_orders where ba_number=$1',[ba])).rows[0].n),0);assert.equal(Number((await db.query('select count(*) n from public.production_order_article_snapshots where ba_number=$1',[ba])).rows[0].n),0);}finally{await db.query('drop trigger synthetic_initial_abort on public.production_lifecycle_events;drop function public.synthetic_initial_abort()');}
 });
 await check('Existing open/create retries never reassign or bypass required later responsibility transfer',async()=>{
  const snap=await create();await db.query('update public.departments set leader_employee_id=null where id=$1',[s.dept]);try{const repeat=await create(snap.ba_number);assert.equal(repeat.existing,true);assert.equal(repeat.principal_user_id,p.reviewer.id);assert.equal((await events(snap)).length,1);assert.equal((await view(snap)).principal_user_id,p.reviewer.id);}finally{await db.query('update public.departments set leader_employee_id=$1 where id=$2',[p.reviewer.employee_id,s.dept]);}
  await assert.rejects(rpc(p.worker,'ba_entry_command',{p_action:'create',p_ba_number:String(n++),p_article_number:'SYNTH-ACCEPT',p_quantity:2,p_principal_user_id:p.worker.id}),e=>e.status>=400);await assert.rejects(create(String(n++),p.scanner),e=>['42501','22023'].includes(e.code));
 });
 results.phase5.initialResponsibility='first_op_department_v1';
 results.initialResponsibility={implemented:true,rule:'first_op_department_v1',syntheticActivationOnly:true,previousCasesPreserved:52,additionalNativeCases:10};
 results.limitations.push('Automatic first-OP AL rule supersedes earlier pending-rule limitation only for new Phase-1 entry BAs while its isolated-only gate is enabled; no existing BA backfill or production activation. PR15/16 ownership and independent provisioning review remain open.');
}
module.exports={run};
