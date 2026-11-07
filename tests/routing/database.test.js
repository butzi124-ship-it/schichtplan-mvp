const {test,before,after}=require('node:test');const assert=require('node:assert/strict');const {randomUUID}=require('node:crypto');
const {routingDatabase,manager,worker,reviewer,outsider,department,foreignDepartment,machine,foreignMachine,leaderB,deputy,machineA2,order}=require('../helpers/routing-db');
let api,n=812200;
before(async()=>api=await routingDatabase());after(async()=>api?.db.close());
const view=uid=>api.rpc(uid,'op_routing_view');
const command=(uid,op,action,target,version=0,extra={})=>api.rpc(uid,'op_routing_command',{p_operation_id:op,p_action:action,p_request_id:randomUUID(),p_expected_version:version,p_machine_id:target,p_reason:'Synthetic explicit reason',...extra});
const fresh=async()=>{const s=await api.seed(String(n++));return{snapshot:s,op:s.op_rows[0].id};};
test('article default is preselected; employee selects suitable own-department machine without mutating article, BA or counters',async()=>{
 const {snapshot,op}=await fresh();const before=(await api.db.query('select * from public.production_station_counts')).rows;
 const row=(await view(worker)).rows.find(r=>r.operation_id===op);assert.equal(row.machine_id,machine);assert.equal(row.default_machine_id,machine);
 const r=await command(worker,op,'select',machineA2);assert.equal(r.machine_id,machineA2);assert.equal(r.executing_department_id,department);
 assert.deepEqual((await api.db.query('select * from public.production_station_counts')).rows,before);
 assert.equal((await api.db.query('select actual_time_minutes from public.production_order_stations where order_id=$1',[order])).rows[0].actual_time_minutes,42);
 const frozen=(await api.db.query('select operations from public.production_order_article_snapshots where order_id=$1',[snapshot.order_id])).rows[0].operations;
 assert.equal(frozen[0].machine_id,machine);assert.equal((await api.db.query('select count(*)::int n from public.production_operation_machine_sections where operation_id=$1',[op])).rows[0].n,2);
});
test('AL binds across departments; recipient is notified, operative authority transfers, original AL keeps read-only overview; return notifies origin',async()=>{
 const {op,snapshot}=await fresh();const r=await command(reviewer,op,'fix',foreignMachine);
 assert.equal(r.binding,true);assert.equal(r.executing_department_id,foreignDepartment);
 const target=(await view(leaderB));assert(target.notifications.some(n=>n.kind==='transferred'&&n.operation_id===op));
 assert.equal((await view(reviewer)).rows.find(x=>x.operation_id===op).can_manage,false);
 await assert.rejects(command(worker,op,'select',machineA2,r.row_version),e=>e.code==='42501');
 await assert.rejects(command(outsider,op,'select',foreignMachine,r.row_version),e=>e.code==='42501');
 const returned=await command(leaderB,op,'return',machine,r.row_version);assert.equal(returned.original_department_id,department);assert.equal(returned.executing_department_id,department);
 assert((await view(reviewer)).notifications.some(n=>n.kind==='returned'&&n.operation_id===op));
 assert.equal((await api.db.query('select department_id from public.production_orders where id=$1',[snapshot.order_id])).rows[0].department_id,department);
});
test('active deputy may return; expired/revoked delegation or mismatched principal cannot authorize',async()=>{
 const {op}=await fresh();const r=await command(reviewer,op,'fix',foreignMachine);
 await api.db.query("update public.production_routing_authorities set valid_until=now()-interval '1 second' where user_id=$1",[deputy]);
 await assert.rejects(command(deputy,op,'return',machine,r.row_version),e=>e.code==='42501');
 await api.db.query("update public.production_routing_authorities set valid_until=now()+interval '1 day',revoked=true where user_id=$1",[deputy]);
 await assert.rejects(command(deputy,op,'return',machine,r.row_version),e=>e.code==='42501');
 await api.db.query('update public.production_routing_authorities set revoked=false,principal_user_id=$1 where user_id=$2',[reviewer,deputy]);
 await assert.rejects(command(deputy,op,'return',machine,r.row_version),e=>e.code==='42501');
 await api.db.query('update public.production_routing_authorities set principal_user_id=$1 where user_id=$2',[leaderB,deputy]);
 assert.equal((await command(deputy,op,'return',machine,r.row_version)).executing_department_id,department);
});
test('technical mandatory binding, suitability, inactive target and foreign normal-employee changes fail closed',async()=>{
 const {op}=await fresh();await api.db.query("insert into public.production_operation_machine_constraints values($1,$2,'Synthetic confirmed mandatory machine')",[op,machine]);
 await assert.rejects(command(reviewer,op,'fix',foreignMachine),e=>e.code==='42501');
 await assert.rejects(command(worker,op,'select',foreignMachine),e=>e.code==='42501');
 await api.db.query('update public.production_machines set active=false where id=$1',[machineA2]);
 try{await assert.rejects(command(worker,op,'select',machineA2),e=>e.code==='42501');}finally{await api.db.query('update public.production_machines set active=true where id=$1',[machineA2]);}
 const other=await fresh();await api.db.query('delete from public.production_routing_capabilities where machine_id=$1',[machineA2]);
 try{await assert.rejects(command(reviewer,other.op,'fix',machineA2),e=>e.code==='42501');}finally{await api.db.query("insert into public.production_routing_capabilities values($1,'TEST')",[machineA2]);}
});
test('setup switch requires configured abort reason, closes all personal segments at one instant, retains cumulative execution and never starts new timers',async()=>{
 const {op}=await fresh();const r=await command(worker,op,'select',machine);
 const section=(await api.db.query('select id from public.production_operation_machine_sections where operation_id=$1 and closed_at is null',[op])).rows[0].id;
 await api.db.query("insert into public.production_operation_activity_segments(section_id,user_id,activity,started_at) values($1,$2,'setup_internal',now()-interval '5 min'),($1,$3,'setup_external',now()-interval '3 min')",[section,worker,manager]);
 await api.db.query("update public.production_operation_execution set good_quantity=17,work_status='running',started_at=now()-interval '1 hour',started_here=true where operation_id=$1",[op]);
 await assert.rejects(command(worker,op,'select',machineA2,r.row_version),e=>e.code==='22023');
 const next=await command(worker,op,'select',machineA2,r.row_version,{p_abort_code:'synthetic-machine-change'});
 const segments=(await api.db.query('select * from public.production_operation_activity_segments where section_id=$1',[section])).rows;assert.equal(segments.length,2);assert(segments.every(s=>s.ended_at));assert.equal(segments[0].ended_at.getTime(),segments[1].ended_at.getTime());
 const ex=(await api.db.query('select * from public.production_operation_execution where operation_id=$1',[op])).rows[0];assert.equal(ex.good_quantity,17);assert.equal(ex.work_status,'paused');assert.equal(ex.started_here,false);
 assert.equal((await api.db.query('select count(*)::int n from public.production_operation_activity_segments where ended_at is null')).rows[0].n,0);assert.equal(next.machine_id,machineA2);
 await assert.rejects(api.db.query("update public.production_operation_activity_segments set ended_at=now() where section_id=$1",[section]),e=>e.code==='42501');
});
test('return interrupts active production without marking OP complete; legacy assign API cannot bypass current binding',async()=>{
 const {op}=await fresh();const r=await command(reviewer,op,'fix',foreignMachine);
 const section=(await api.db.query('select id from public.production_operation_machine_sections where operation_id=$1 and closed_at is null',[op])).rows[0].id;
 await api.db.query("insert into public.production_operation_activity_segments(section_id,user_id,activity,started_at) values($1,$2,'production',now()-interval '2 min')",[section,outsider]);
 await api.db.query("update public.production_operation_execution set work_status='running',started_at=now()-interval '2 min',good_quantity=11 where operation_id=$1",[op]);
 await command(leaderB,op,'return',machine,r.row_version);
 const e=(await api.db.query('select * from public.production_operation_execution where operation_id=$1',[op])).rows[0];assert.equal(e.work_status,'paused');assert.equal(e.good_quantity,11);
 await assert.rejects(api.rpc(manager,'machine_readiness_command',{p_operation_id:op,p_action:'assign',p_request_id:randomUUID(),p_expected_version:e.row_version,p_machine_id:machineA2,p_reason:'Bypass attempt'}),e=>e.code==='42501');
});
test('duplicate request is idempotent, stale update rejected, and notification failure rolls back assignment, sections and audit',async()=>{
 const {op}=await fresh();const id=randomUUID(),params={p_request_id:id};const r=await command(reviewer,op,'fix',foreignMachine,0,params);assert.deepEqual(await command(reviewer,op,'fix',foreignMachine,0,params),r);
 await assert.rejects(command(reviewer,op,'fix',machine,0),e=>['42501','40001'].includes(e.code));
 const other=await fresh();await api.db.exec("create function public.synthetic_notify_abort() returns trigger language plpgsql as $$begin raise exception 'Synthetic notification failure';end$$;create trigger synthetic_notify_abort before insert on public.production_operation_routing_notifications for each row execute function public.synthetic_notify_abort()");
 try{await assert.rejects(command(reviewer,other.op,'fix',foreignMachine));assert.equal((await api.db.query('select count(*)::int n from public.production_operation_routing where operation_id=$1',[other.op])).rows[0].n,0);}
 finally{await api.db.exec('drop trigger synthetic_notify_abort on public.production_operation_routing_notifications;drop function public.synthetic_notify_abort()');}
});
test('old untracked readiness workflow still works; cross-department tracked work uses receiving department authority',async()=>{
 const {op}=await fresh();const r=await command(reviewer,op,'fix',foreignMachine);
 let e=(await api.db.query('select * from public.production_operation_execution where operation_id=$1',[op])).rows[0];
 e=await api.rpc(outsider,'machine_readiness_command',{p_operation_id:op,p_action:'handover',p_request_id:randomUUID(),p_expected_version:e.row_version,p_quantity:48,p_reason:'Synthetic physical receipt'});
 e=await api.rpc(leaderB,'machine_readiness_command',{p_operation_id:op,p_action:'assign',p_request_id:randomUUID(),p_expected_version:e.row_version,p_machine_id:foreignMachine,p_reason:'Receiving AL retains existing plan API'});
 e=await api.rpc(outsider,'machine_readiness_command',{p_operation_id:op,p_action:'start',p_request_id:randomUUID(),p_expected_version:e.row_version,p_reason:'Synthetic explicit work start'});assert.equal(e.work_status,'running');
 const old=await fresh();e=(await api.db.query('select * from public.production_operation_execution where operation_id=$1',[old.op])).rows[0];
 e=await api.rpc(manager,'machine_readiness_command',{p_operation_id:old.op,p_action:'handover',p_request_id:randomUUID(),p_expected_version:e.row_version,p_quantity:48,p_reason:'Untracked original path'});assert.equal(e.work_status,'ready');
});
test('server flag off, scanner, mutable role escalation and raw table writes cannot open routing authority',async()=>{
 const {op}=await fresh();await api.db.query("update public.employees set role='admin' where auth_user_id=$1",[worker]);
 try{await assert.rejects(command(worker,op,'fix',foreignMachine),e=>e.code==='42501');}finally{await api.db.query("update public.employees set role='employee' where auth_user_id=$1",[worker]);}
 await api.db.query("update public.employees set role='tool_scanner' where auth_user_id=$1",[worker]);
 try{assert.deepEqual((await view(worker)).rows,[]);await assert.rejects(command(worker,op,'fix',foreignMachine),e=>e.code==='42501');}finally{await api.db.query("update public.employees set role='employee' where auth_user_id=$1",[worker]);}
 await api.db.exec('update public.production_routing_settings set enabled=false');try{assert.deepEqual((await view(manager)).rows,[]);await assert.rejects(command(manager,op,'fix',foreignMachine));}finally{await api.db.exec('update public.production_routing_settings set enabled=true');}
 await assert.rejects(api.asUser(worker,()=>api.db.query('select * from public.production_operation_routing_events')),e=>e.code==='42501');
});

test('switch closes only this OP; an unrelated personal segment remains active and closed sections reject late time starts',async()=>{
 const one=await fresh(),two=await fresh();
 const a=await command(worker,one.op,'select',machine);await command(reviewer,two.op,'fix',machineA2);
 const sections=(await api.db.query('select id,operation_id from public.production_operation_machine_sections where operation_id=any($1::uuid[]) and closed_at is null',[[one.op,two.op]])).rows;
 const sa=sections.find(s=>s.operation_id===one.op).id,sb=sections.find(s=>s.operation_id===two.op).id;
 await api.db.query("insert into public.production_operation_activity_segments(section_id,user_id,activity,started_at) values($1,$2,'setup_internal',now()),($3,$4,'production',now())",[sa,worker,sb,reviewer]);
 await command(worker,one.op,'select',machineA2,a.row_version,{p_abort_code:'synthetic-machine-change'});
 assert.equal((await api.db.query('select ended_at from public.production_operation_activity_segments where section_id=$1',[sb])).rows[0].ended_at,null);
 await assert.rejects(api.db.query("insert into public.production_operation_activity_segments(section_id,user_id,activity,started_at) values($1,$2,'production',now())",[sa,manager]),e=>e.code==='40001');
});
