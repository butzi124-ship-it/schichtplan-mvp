const {test,before,after}=require('node:test'),assert=require('node:assert/strict'),{randomUUID}=require('node:crypto');
const {setupDatabase,worker,reviewer,manager,outsider,machine,machineA2,foreignMachine,leaderB,deputy}=require('../helpers/setup-db');
let api,n=813000;before(async()=>api=await setupDatabase());after(async()=>api?.db.close());
const cmd=(u,action,op=null,run=null,extra={})=>api.rpc(u,'setup_time_command',{p_action:action,p_request_id:randomUUID(),p_operation_id:op,p_run_id:run,...extra});
const view=u=>api.rpc(u,'setup_time_view');
async function fresh(){const s=await api.seed(String(n++)),op=s.op_rows[0].id;await api.rpc(worker,'op_routing_command',{p_operation_id:op,p_action:'select',p_request_id:randomUUID(),p_expected_version:0,p_machine_id:machine,p_reason:'Synthetic explicit assignment'});return op;}
const start=(u,op,run=null,mode='external')=>cmd(u,'start',op,run,{p_mode:mode});
const rows=()=>api.db.query('select * from public.production_operation_activity_segments order by started_at');
test('explicit internal/external switches split personal intervals; target and quantities unchanged',async()=>{
 const op=await fresh(),before=(await api.db.query('select * from public.production_station_counts')).rows;
 const r=await start(worker,op,null,'internal');await start(worker,op,r.run_id);await cmd(worker,'end_own',op,r.run_id);
 const a=(await rows()).rows.filter(x=>x.setup_run_id===r.run_id);assert.equal(a.length,2);assert.equal(a[0].ended_at.getTime(),a[1].started_at.getTime());assert.equal(a[0].activity,'setup_internal');assert.equal(a[1].activity,'setup_external');
 assert.deepEqual((await api.db.query('select * from public.production_station_counts')).rows,before);
 await cmd(worker,'finish',op,r.run_id);
});
test('two employees share run; personal finish retains colleague; participant shared finish ends all at same instant',async()=>{
 const op=await fresh(),r=await start(worker,op);assert.equal((await start(reviewer,op)).run_id,r.run_id);
 await cmd(worker,'end_own',op,r.run_id);assert((await rows()).rows.some(a=>a.user_id===reviewer&&a.setup_run_id===r.run_id&&!a.ended_at));
 await start(worker,op,r.run_id,'internal');await cmd(worker,'finish',op,r.run_id);
 const active=(await rows()).rows.filter(a=>a.setup_run_id===r.run_id);assert(active.every(a=>a.ended_at));assert.equal(active.at(-1).ended_at.getTime(),active.at(-2).ended_at.getTime());
 await assert.rejects(start(worker,op,r.run_id),e=>e.code==='PT409');
});
test('general pause interrupts all own production/setup only, return offers open run and never restarts',async()=>{
 const op=await fresh(),r=await start(worker,op);await start(reviewer,op);await cmd(worker,'pause');assert.equal((await view(worker)).paused,true);
 assert((await rows()).rows.some(a=>a.user_id===reviewer&&a.setup_run_id===r.run_id&&!a.ended_at));
 await assert.rejects(start(worker,op,r.run_id));await cmd(worker,'return');assert.equal((await view(worker)).resume_run,r.run_id);
 assert(!(await rows()).rows.some(a=>a.user_id===worker&&!a.ended_at));await start(worker,op,r.run_id);await cmd(reviewer,'finish',op,r.run_id);assert.equal((await view(worker)).resume_run,null);
 const section=(await api.db.query('select id from public.production_operation_machine_sections where operation_id=$1 and closed_at is null',[op])).rows[0].id;
 await api.db.query("insert into public.production_operation_activity_segments(section_id,user_id,activity,started_at) values($1,$2,'production',clock_timestamp())",[section,worker]);await cmd(worker,'pause');assert(!(await rows()).rows.some(a=>a.user_id===worker&&!a.ended_at));await cmd(worker,'return');
});
test('logout and explicit workstation takeover affect predecessor alone and do not start successor',async()=>{
 const op=await fresh(),r=await start(worker,op);await start(reviewer,op);await cmd(manager,'takeover',op,null,{p_previous_user:worker});
 assert(!(await rows()).rows.some(a=>a.user_id===worker&&!a.ended_at));assert((await rows()).rows.some(a=>a.user_id===reviewer&&!a.ended_at));
 await cmd(reviewer,'logout');assert.equal((await view(worker)).rows.flatMap(x=>x.runs).find(x=>x.id===r.run_id).ended_at,null);await cmd(worker,'finish',op,r.run_id);
});
test('machine production remains running during external setup; internal mode blocks running machine',async()=>{
 const op=await fresh();await api.db.query("update public.production_operation_execution set work_status='running' where operation_id=$1",[op]);const r=await start(worker,op);
 await assert.rejects(start(worker,op,r.run_id,'internal'),e=>e.code==='PT409');assert.equal((await api.db.query('select work_status from public.production_operation_execution where operation_id=$1',[op])).rows[0].work_status,'running');await cmd(worker,'finish',op,r.run_id);
 await api.db.query("update public.production_operation_execution set work_status='paused' where operation_id=$1",[op]);
});
test('paused shared setup still requires configured machine abort; history retained and explicit new run required',async()=>{
 const op=await fresh(),r=await start(worker,op);await cmd(worker,'pause_own',op,r.run_id);
 const route=(await api.rpc(worker,'op_routing_view')).rows.find(x=>x.operation_id===op);assert.equal(route.active_setup,true);
 const p={p_operation_id:op,p_action:'select',p_request_id:randomUUID(),p_expected_version:route.row_version,p_machine_id:machineA2,p_reason:'Synthetic move'};
 await assert.rejects(api.rpc(worker,'op_routing_command',p),e=>e.code==='22023');await api.rpc(worker,'op_routing_command',{...p,p_abort_code:'synthetic-machine-change'});
 assert.equal((await view(worker)).rows.flatMap(x=>x.runs).find(x=>x.id===r.run_id).end_reason,'machine_change');await assert.rejects(start(worker,op,r.run_id));
 const next=await start(worker,op);assert.notEqual(next.run_id,r.run_id);await cmd(worker,'finish',op,next.run_id);
});
test('correction needs reason, original remains immutable, stale audit and overlaps rejected',async()=>{
 const op=await fresh(),r=await start(worker,op);await cmd(worker,'end_own',op,r.run_id);await cmd(worker,'finish',op,r.run_id);
 const a=(await view(worker)).rows.flatMap(x=>x.runs).find(x=>x.id===r.run_id).segments[0];const {can_correct,correction_count,...effective}=a;
 const original=(await rows()).rows.find(x=>x.id===a.id);const change={started_at:a.started_at,ended_at:a.ended_at,activity:'setup_internal'};
 await assert.rejects(cmd(worker,'correct',op,r.run_id,{p_segment_id:a.id,p_expected:effective,p_change:change}));
 await cmd(worker,'correct',op,r.run_id,{p_segment_id:a.id,p_expected:effective,p_change:change,p_reason:'Synthetic mode correction'});
 assert.deepEqual((await rows()).rows.find(x=>x.id===a.id),original);
 await assert.rejects(cmd(worker,'correct',op,r.run_id,{p_segment_id:a.id,p_expected:effective,p_change:change,p_reason:'Stale'}),e=>e.code==='PT409');
 await assert.rejects(api.asUser(worker,()=>api.db.query('delete from public.production_setup_corrections')),e=>e.code==='42501');
});
test('foreign employee and mutable admin role cannot finish another run',async()=>{
 const op=await fresh(),r=await start(worker,op);await assert.rejects(start(outsider,op),e=>e.code==='42501');await assert.rejects(cmd(outsider,'finish',op,r.run_id),e=>e.code==='42501');
 await api.db.query("update public.employees set role='admin' where auth_user_id=$1",[outsider]);try{await assert.rejects(cmd(outsider,'finish',op,r.run_id),e=>e.code==='42501');}finally{await api.db.query("update public.employees set role='employee' where auth_user_id=$1",[outsider]);}
 await cmd(manager,'finish',op,r.run_id);
});
test('idempotent duplicate close and pause events; reused request payload rejected',async()=>{
 const op=await fresh(),r=await start(worker,op);const p={p_action:'finish',p_request_id:randomUUID(),p_operation_id:op,p_run_id:r.run_id};
 const a=await api.rpc(worker,'setup_time_command',p),b=await api.rpc(worker,'setup_time_command',p);assert.deepEqual(a,b);
 await assert.rejects(api.rpc(worker,'setup_time_command',{...p,p_action:'pause'}),e=>e.code==='22023');await cmd(worker,'pause');await cmd(worker,'return');
 assert.equal((await view(worker)).resume_run,null);
});
test('time metrics distinguish sum, wall duration and interval union without double machine time',async()=>{
 const op=await fresh(),r=await start(worker,op,null,'internal');await start(reviewer,op,r.run_id,'internal');await cmd(reviewer,'finish',op,r.run_id);
 const run=(await view(worker)).rows.flatMap(x=>x.runs).find(x=>x.id===r.run_id);
 assert(Number(run.person_minutes)>=Number(run.active_calendar_minutes));assert.equal(Number(run.machine_stop_minutes),Number(run.active_calendar_minutes));assert(Number(run.calendar_minutes)>=Number(run.active_calendar_minutes));
});
test('correction rejects overlapping personal intervals and audit failure rolls back a start',async()=>{
 const op=await fresh(),r=await start(worker,op);await start(worker,op,r.run_id,'internal');await cmd(worker,'finish',op,r.run_id);
 const run=(await view(worker)).rows.flatMap(x=>x.runs).find(x=>x.id===r.run_id),a=run.segments[0],b=run.segments[1];const {can_correct,correction_count,...expected}=a;
 await assert.rejects(cmd(worker,'correct',op,r.run_id,{p_segment_id:a.id,p_expected:expected,p_change:{started_at:a.started_at,ended_at:b.ended_at,activity:a.activity},p_reason:'Synthetic overlapping correction'}),e=>e.code==='22023');
 const next=await fresh();await api.db.exec("create function public.synthetic_setup_failure() returns trigger language plpgsql as $$begin raise exception 'Synthetic audit failure';end$$;create trigger synthetic_setup_failure before insert on public.production_setup_events for each row execute function public.synthetic_setup_failure()");
 try{await assert.rejects(start(worker,next));assert.equal((await api.db.query('select count(*)::int n from public.production_setup_runs r join public.production_operation_machine_sections x on x.id=r.section_id where x.operation_id=$1',[next])).rows[0].n,0);}finally{await api.db.exec('drop trigger synthetic_setup_failure on public.production_setup_events;drop function public.synthetic_setup_failure()');}
});
test('executing deputy and admin correction authority follows transferred department; expired delegation denied',async()=>{
 const op=await fresh(),route=(await api.rpc(worker,'op_routing_view')).rows.find(x=>x.operation_id===op);
 await api.rpc(reviewer,'op_routing_command',{p_operation_id:op,p_action:'fix',p_request_id:randomUUID(),p_expected_version:route.row_version,p_machine_id:foreignMachine,p_reason:'Synthetic transfer'});
 const r=await start(outsider,op);await api.db.query("update public.production_routing_authorities set valid_until=now()-interval '1 second' where user_id=$1",[deputy]);
 await assert.rejects(cmd(deputy,'finish',op,r.run_id),e=>e.code==='42501');await assert.rejects(cmd(reviewer,'finish',op,r.run_id),e=>e.code==='42501');
 await api.db.query("update public.production_routing_authorities set valid_until=now()+interval '1 day' where user_id=$1",[deputy]);await cmd(deputy,'finish',op,r.run_id);
});
test('server gate, scanner and raw history access fail closed',async()=>{
 const op=await fresh();await api.db.query("update public.employees set role='tool_scanner' where auth_user_id=$1",[worker]);try{assert.deepEqual((await view(worker)).rows,[]);await assert.rejects(start(worker,op),e=>e.code==='42501');}finally{await api.db.query("update public.employees set role='employee' where auth_user_id=$1",[worker]);}
 await api.db.exec('update public.production_setup_settings set enabled=false');try{assert.deepEqual((await view(manager)).rows,[]);await assert.rejects(start(manager,op),e=>e.code==='42501');}finally{await api.db.exec('update public.production_setup_settings set enabled=true');}
 for(const table of ['production_setup_events','production_setup_runs','production_setup_corrections'])await assert.rejects(api.asUser(worker,()=>api.db.query('select * from public.'+table)),e=>e.code==='42501');
});
test('stale takeover cannot change another employees pause state; terminal BA still allows safe personal finish',async()=>{
 const op=await fresh(),r=await start(worker,op);await cmd(worker,'pause');await assert.rejects(cmd(reviewer,'takeover',op,null,{p_previous_user:worker}),e=>e.code==='PT409');assert.equal((await view(worker)).paused,true);await cmd(worker,'return');await start(worker,op,r.run_id);
 await api.db.query("update public.production_operation_execution set work_status='cancelled' where operation_id=$1",[op]);await cmd(worker,'end_own',op,r.run_id);await cmd(worker,'finish',op,r.run_id);assert((await rows()).rows.filter(a=>a.setup_run_id===r.run_id).every(a=>a.ended_at));
});
test('existing production start and resume reject internal machine occupation; external is independent',async()=>{
 const op=await fresh(),r=await start(worker,op,null,'internal');const execution=(await api.db.query('select row_version from public.production_operation_execution where operation_id=$1',[op])).rows[0];
 for(const action of ['start','resume'])await assert.rejects(api.rpc(worker,'machine_readiness_command',{p_operation_id:op,p_action:action,p_request_id:randomUUID(),p_expected_version:execution.row_version,p_reason:'Synthetic blocked production'}),e=>e.code==='PT409');
 await cmd(worker,'finish',op,r.run_id);
});
test('own historical corrections remain visible after department transfer; management corrections follow current responsibility',async()=>{
 const op=await fresh(),r=await start(worker,op);await cmd(worker,'finish',op,r.run_id);
 const route=(await api.rpc(worker,'op_routing_view')).rows.find(x=>x.operation_id===op);
 await api.rpc(reviewer,'op_routing_command',{p_operation_id:op,p_action:'fix',p_request_id:randomUUID(),p_expected_version:route.row_version,p_machine_id:foreignMachine,p_reason:'Synthetic transfer after setup'});
 const own=(await view(worker)).rows.flatMap(x=>x.runs).find(x=>x.id===r.run_id).segments[0];assert.equal(own.can_correct,true);
 assert.equal((await view(reviewer)).rows.flatMap(x=>x.runs).find(x=>x.id===r.run_id).segments[0].can_correct,false);
 assert.equal((await view(leaderB)).rows.flatMap(x=>x.runs).find(x=>x.id===r.run_id).segments[0].can_correct,true);
 const {can_correct,correction_count,...expected}=own,extra={p_segment_id:own.id,p_expected:expected,p_change:{started_at:own.started_at,ended_at:own.ended_at,activity:'setup_internal'},p_reason:'Synthetic own historical correction'};
 await assert.rejects(cmd(reviewer,'correct',op,r.run_id,extra),e=>e.code==='42501');await cmd(worker,'correct',op,r.run_id,extra);
});
