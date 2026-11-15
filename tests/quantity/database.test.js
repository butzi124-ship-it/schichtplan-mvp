const {test,before,after}=require('node:test'),assert=require('node:assert/strict'),{randomUUID}=require('node:crypto');
const {quantityDatabase,worker,reviewer,manager,outsider,machine,machineA2,foreignMachine}=require('../helpers/quantity-db');
const {blank}=require('../../js/operations');
let api,n=814000;before(async()=>api=await quantityDatabase());after(async()=>api?.db.close());
const view=()=>api.rpc(worker,'op_quantity_view');
const cmd=(u,op,action,extra={})=>api.rpc(u,'op_quantity_command',{p_operation_id:op,p_action:action,p_request_id:randomUUID(),...extra});
async function fresh(count=3){
 const ba=String(n++);let r=await api.rpc(manager,'operations_article_command',{p_action:'create',p_department_id:require('../helpers/quantity-db').department,p_article_number:ba,p_data:{designation:'Synthetic phase 4'}});
 const plan=await api.rpc(manager,'operation_plan',{p_revision_id:r.id,p_operations:[{...blank(10),machine_id:machine,machine_group:'TEST',activity:'A'},{...blank(20,10),machine_id:machine,machine_group:'TEST',activity:'B'}],p_expected_version:r.row_version});
 await api.rpc(reviewer,'operations_article_command',{p_action:'release',p_revision_id:r.id,p_expected_version:plan.row_version});
 await api.db.query("insert into public.article_master_assignments(user_id,department_id,article_number,first_op_number,first_op_evidence,expires_at,issued_by) values($1,$2,$3,10,'Synthetic phase 4 assignment',now()+interval '1 day','00000000-0000-0000-0000-000000000001')",[worker,require('../helpers/quantity-db').department,ba]);
 const s=await api.rpc(worker,'ba_entry_command',{p_action:'create',p_ba_number:ba,p_article_number:ba,p_quantity:100});
 for(const op of s.op_rows){await api.rpc(worker,'op_routing_command',{p_operation_id:op.id,p_action:'select',p_request_id:randomUUID(),p_expected_version:0,p_machine_id:machine,p_reason:'Synthetic selection'});await cmd(worker,op.id,'configure',{p_new_count:op.position===1?count:1,p_note:'Synthetic complete required plan'});}
 await cmd(reviewer,s.op_rows[0].id,'join');
 const people=[];for(const u of [worker,reviewer]){const row=(await api.db.query('select pe.id from public.production_order_employees pe join public.employees e on e.id=pe.employee_id where pe.order_id=$1 and e.auth_user_id=$2',[s.order_id,u])).rows[0];people.push(row.id);}
 return{...s,people,rows:(await view()).rows.filter(r=>r.order_id===s.order_id)};
}
async function book(s,index,station=0,u=worker,extra={}){const r=(await view()).rows.find(r=>r.operation_id===s.op_rows[index].id);return cmd(u,r.operation_id,'count',{p_station_id:r.stations[station].id,p_order_employee_id:s.people[u===worker?0:1],p_kind:'good',p_delta:1,p_section_id:r.section_id,...extra});}
test('new OP zero, 3 tensions 100/100/98 contribute 98 only, same OP persists',async()=>{
 const s=await fresh();assert.equal(s.rows[0].finished_good,0);
 // Bulk fixture quantities model accumulated legacy button presses, not a second store.
 for(const [i,qty] of [100,100,98].entries())await api.db.query('insert into public.production_station_counts(order_id,station_id,order_employee_id,good_qty) values($1,$2,$3,$4)',[s.order_id,s.rows[0].stations[i].id,s.people[i===1?1:0],qty]);
 let r=(await view()).rows.find(r=>r.operation_id===s.op_rows[0].id);assert.equal(r.finished_good,98);await book(s,0,2);r=(await view()).rows.find(r=>r.operation_id===s.op_rows[0].id);assert.equal(r.finished_good,99);assert.equal((await view()).rows.find(r=>r.operation_id===s.op_rows[1].id).finished_good,0);
});
test('parallel next OP keeps excess booking, one conflict and AL notification; correction remains possible',async()=>{
 const s=await fresh(1);await book(s,0);await book(s,1);await book(s,1);
 let r=(await view()).rows.find(r=>r.operation_id===s.op_rows[1].id);assert.equal(r.finished_good,2);assert.equal(r.available,1);assert(r.stations[0].conflict_id);
 await book(s,1);assert.equal((await api.db.query('select count(*)::int n from public.production_quantity_conflicts where operation_id=$1',[r.operation_id])).rows[0].n,1);
 assert((await api.rpc(reviewer,'op_quantity_view')).notifications.some(n=>n.operation_id===r.operation_id));
 await book(s,0,0,worker,{p_delta:-1});r=(await view()).rows.find(r=>r.operation_id===s.op_rows[1].id);assert.equal(r.available,0);assert.equal(r.finished_good,3);
});
test('idempotent retry increments once; changed request, foreign person, foreign BA and scanner denied',async()=>{
 const s=await fresh(1),r=s.rows[0],params={p_operation_id:r.operation_id,p_action:'count',p_request_id:randomUUID(),p_station_id:r.stations[0].id,p_order_employee_id:s.people[0],p_kind:'good',p_delta:1,p_section_id:r.section_id};
 const a=await api.rpc(worker,'op_quantity_command',params);assert.deepEqual(await api.rpc(worker,'op_quantity_command',params),a);
 await assert.rejects(api.rpc(worker,'op_quantity_command',{...params,p_delta:-1}),e=>e.code==='22023');await assert.rejects(api.rpc(outsider,'op_quantity_command',{...params,p_request_id:randomUUID()}),e=>e.code==='42501');
 await assert.rejects(book(s,0,0,worker,{p_order_employee_id:s.people[1]}),e=>e.code==='42501');const other=await fresh(1);await assert.rejects(book(s,0,0,worker,{p_station_id:other.rows[0].stations[0].id}),e=>e.code==='22023');
 await api.db.query("update public.employees set role='tool_scanner' where auth_user_id=$1",[worker]);try{await assert.rejects(api.rpc(worker,'op_quantity_command',{...params,p_request_id:randomUUID()}),e=>e.code==='42501');}finally{await api.db.query("update public.employees set role='employee' where auth_user_id=$1",[worker]);}
});
test('machine move retains OP quantity, section totals split and stale machine booking rejected',async()=>{
 const s=await fresh(1),r=s.rows[0];await book(s,0);let route=(await api.rpc(worker,'op_routing_view')).rows.find(v=>v.operation_id===r.operation_id);
 await api.rpc(worker,'op_routing_command',{p_operation_id:r.operation_id,p_action:'select',p_request_id:randomUUID(),p_expected_version:route.row_version,p_machine_id:machineA2,p_reason:'Synthetic counter machine move'});
 await assert.rejects(book(s,0,0,worker,{p_section_id:r.section_id}),e=>e.code==='PT409');await book(s,0);const next=(await view()).rows.find(v=>v.operation_id===r.operation_id);assert.equal(next.finished_good,2);assert.deepEqual(next.sections.map(x=>x.finished_good).sort(),[1,1]);
});
test('6M scrap / in-clarification preserve event meanings and lifetime, never increase finished good',async()=>{
 const s=await fresh(1),cause=(await api.db.query("insert into public.production_qa_causes(group_id,group_label,reason_id,reason_label) values('M','Mensch','AA','1AA') returning id")).rows[0].id;
 await assert.rejects(book(s,0,0,worker,{p_kind:'scrap'}),e=>e.code==='22023');for(const kind of ['scrap','clarify']){await book(s,0,0,worker,{p_kind:kind,p_qa_cause_id:cause});await book(s,0,0,worker,{p_kind:kind,p_delta:-1});}
 const r=(await view()).rows.find(v=>v.operation_id===s.rows[0].operation_id);assert.equal(r.finished_good,0);assert.equal(r.stations[0].scrap_total,0);assert.equal(r.stations[0].scrap_lifetime,1);assert.equal(r.stations[0].clarify_lifetime,1);
 const events=(await api.db.query('select event_type,qa_cause_id from public.production_station_events where order_id=$1 order by created_at',[s.order_id])).rows;assert.deepEqual(events.map(e=>e.event_type),['scrap','scrap_correction','clarify','clarify_correction']);assert.equal(events[0].qa_cause_id,cause);
});
test('audit failure rolls back count, event, attribution and conflict',async()=>{
 const s=await fresh(1);await api.db.exec("create function public.synthetic_quantity_abort() returns trigger language plpgsql as $$begin raise exception 'Synthetic audit failure';end$$;create trigger synthetic_quantity_abort before insert on public.production_quantity_events for each row execute function public.synthetic_quantity_abort()");
 try{await assert.rejects(book(s,1));assert.equal((await view()).rows.find(v=>v.operation_id===s.rows[1].operation_id).finished_good,0);assert.equal((await api.db.query('select count(*)::int n from public.production_station_events where order_id=$1',[s.order_id])).rows[0].n,0);}finally{await api.db.exec('drop trigger synthetic_quantity_abort on public.production_quantity_events;drop function public.synthetic_quantity_abort()');}
});
test('raw legacy grants cannot change tracked counts; immutable plan/history and disabled gate fail closed',async()=>{
 const s=await fresh(1);await book(s,0);await api.db.exec('grant select,update on public.production_station_counts to authenticated');await assert.rejects(api.asUser(worker,()=>api.db.query('update public.production_station_counts set good_qty=999 where station_id=$1',[s.rows[0].stations[0].id])),e=>e.code==='42501');
 await assert.rejects(cmd(worker,s.rows[0].operation_id,'configure',{p_new_count:1,p_note:'Replace'}),e=>e.code==='PT409');await api.db.exec('update public.production_quantity_settings set enabled=false');try{await assert.rejects(cmd(worker,s.rows[0].operation_id,'count',{p_station_id:s.rows[0].stations[0].id,p_order_employee_id:s.people[0],p_kind:'good',p_delta:1,p_section_id:s.rows[0].section_id}),e=>e.code==='42501');assert.deepEqual((await view()).rows,[]);}finally{await api.db.exec('update public.production_quantity_settings set enabled=true');}
});
test('partial predecessor goods permit explicit partial physical handover and start, without completing predecessor',async()=>{
 const s=await fresh(1);await book(s,0);const op=s.op_rows[1].id;
 let r=(await api.rpc(worker,'machine_readiness_view')).rows.find(r=>r.operation_id===op);assert.equal(r.quantity_linked,true);assert.equal(r.predecessor_released,true);assert.equal(r.received_quantity,0);
 const ready=await api.rpc(worker,'machine_readiness_command',{p_operation_id:op,p_action:'handover',p_request_id:randomUUID(),p_expected_version:r.row_version,p_quantity:1,p_reason:'Synthetic physical partial handover'});assert.equal(ready.received_quantity,1);
 const running=await api.rpc(worker,'machine_readiness_command',{p_operation_id:op,p_action:'start',p_request_id:randomUUID(),p_expected_version:ready.row_version,p_reason:'Synthetic partial work start'});assert.equal(running.work_status,'running');
 await assert.rejects(api.rpc(worker,'machine_readiness_command',{p_operation_id:op,p_action:'complete',p_request_id:randomUUID(),p_expected_version:running.row_version,p_quantity:100,p_reason:'Cannot bypass Phase 5'}),e=>e.code==='42501');
 assert.notEqual((await api.db.query('select work_status from public.production_operation_execution where operation_id=$1',[s.op_rows[0].id])).rows[0].work_status,'completed');
 await assert.rejects(api.db.query("update public.production_orders set target_quantity=1 where id=$1",[s.order_id]),e=>e.code==='42501');
});
test('retroactive upstream correction opens downstream conflict, plan labels retained and adoption remains explicit',async()=>{
 const s=await fresh(1);await book(s,0);await book(s,1);let row=(await view()).rows.find(r=>r.operation_id===s.op_rows[1].id);assert.equal(row.stations[0].conflict_id,null);
 await book(s,0,0,worker,{p_delta:-1});row=(await view()).rows.find(r=>r.operation_id===s.op_rows[1].id);assert(row.stations[0].conflict_id);assert.equal(row.finished_good,1);
 await book(s,0);assert((await view()).rows.find(r=>r.operation_id===s.op_rows[1].id).stations[0].conflict_id); // No automatic Phase 5 resolution.
 await assert.rejects(api.asUser(worker,()=>api.db.query('delete from public.production_quantity_events')),e=>e.code==='42501');
});
test('forged untracked order_id cannot smuggle raw counts into a mapped station',async()=>{
 const s=await fresh(1);await api.db.exec('grant insert on public.production_station_counts to authenticated');
 await assert.rejects(api.asUser(worker,()=>api.db.query('insert into public.production_station_counts(order_id,station_id,order_employee_id,good_qty) values($1,$2,$3,99)',[require('../helpers/quantity-db').order,s.rows[0].stations[0].id,s.people[0]])),e=>e.code==='42501');assert.equal((await view()).rows.find(r=>r.operation_id===s.rows[0].operation_id).finished_good,0);
});
test('AL adoption preserves 1AA, historical employee counts and unknown machine origin without inventing distribution',async()=>{
 const snap=await api.seed(String(n++)),op=snap.op_rows[0].id;await api.rpc(worker,'op_routing_command',{p_operation_id:op,p_action:'select',p_request_id:randomUUID(),p_expected_version:0,p_machine_id:machine,p_reason:'Synthetic legacy adoption'});
 const st=(await api.db.query("insert into public.production_order_stations(order_id,station_no,op_number,name,scrap_total) values($1,1,'1AA','Unchanged tension',2) returning id",[snap.order_id])).rows[0].id;
 const person=(await api.db.query("insert into public.production_order_employees(order_id,employee_id,employee_name) select $1,id,'Legacy Worker' from public.employees where auth_user_id=$2 returning id",[snap.order_id,worker])).rows[0].id;
 await api.db.query('insert into public.production_station_counts(order_id,station_id,order_employee_id,good_qty) values($1,$2,$3,17)',[snap.order_id,st,person]);
 const args={p_station_ids:[st],p_new_count:0,p_note:'Synthetic verified historical mapping'};await assert.rejects(cmd(worker,op,'configure',args),e=>e.code==='42501');await cmd(reviewer,op,'configure',args);
 const row=(await view()).rows.find(r=>r.operation_id===op);assert.equal(row.finished_good,17);assert.equal(row.stations[0].op_number,'1AA');assert.equal(row.stations[0].scrap_total,2);assert.equal(row.sections.reduce((n,s)=>n+s.finished_good,0),0);assert.equal(row.stations[0].counts[0].good_qty,17);
});
