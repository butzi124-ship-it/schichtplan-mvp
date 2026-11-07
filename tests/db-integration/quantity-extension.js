// Native isolated PostgreSQL/JWT/PostgREST only, chained with every prior phase.
const fs=require('node:fs'),assert=require('node:assert/strict'),{randomUUID}=require('node:crypto');
const {rpc,http,connect,sqlAs}=require('../../scripts/db-integration/client');
async function run(context){
 await require('./setup-extension').run(context);
 const {db,s,check,results}=context,p=s.people;let n=814900;
 await db.query(fs.readFileSync('supabase/review/quantity-phase4.sql','utf8'));await db.query("notify pgrst,'reload schema'");await new Promise(r=>setTimeout(r,1200));
 const command=(u,op,action,extra={})=>rpc(u,'op_quantity_command',{p_operation_id:op,p_action:action,p_request_id:randomUUID(),...extra});
 const fresh=async()=>{const snap=await rpc(p.worker,'ba_entry_command',{p_action:'create',p_ba_number:String(n++),p_article_number:'SYNTH-ACCEPT',p_quantity:100});const op=snap.op_rows[0].id;await rpc(p.worker,'op_routing_command',{p_operation_id:op,p_action:'select',p_request_id:randomUUID(),p_expected_version:0,p_machine_id:s.machines[0],p_reason:'Synthetic phase 4 assignment'});await command(p.worker,op,'configure',{p_new_count:3,p_note:'Synthetic complete plan'});return(await rpc(p.worker,'op_quantity_view')).rows.find(r=>r.operation_id===op);};
 const book=(r,extra={})=>command(p.worker,r.operation_id,'count',{p_station_id:r.stations[2].id,p_order_employee_id:r.own_person_id,p_kind:'good',p_delta:1,p_section_id:r.section_id,...extra});
 await check('Phase 4 defaults off, real anonymous/raw denials, no production activation',async()=>{
  assert.deepEqual((await rpc(p.worker,'op_quantity_view')).rows,[]);assert.ok([401,403].includes((await http(null,'/rpc/op_quantity_view',{method:'POST',body:{}})).status));
  for(const table of ['production_quantity_plans','production_quantity_attributions','production_quantity_events','production_quantity_conflicts'])assert.ok([401,403].includes((await http(p.worker,'/'+table+'?limit=1')).status));
  await db.query('update public.production_quantity_settings set enabled=true');
 });
 await check('Native existing counter, exact final tension, retained legacy labels and 6M history',async()=>{
  const r=await fresh();assert.equal(r.finished_good,0);await book(r,{p_station_id:r.stations[0].id});await book(r,{p_station_id:r.stations[1].id});await book(r);
  const next=(await rpc(p.worker,'op_quantity_view')).rows.find(v=>v.operation_id===r.operation_id);assert.equal(next.finished_good,1);assert.equal(next.stations.reduce((n,s)=>n+s.good,0),3);
  assert.equal((await db.query('select good_qty from public.production_station_counts where order_id=$1',[s.legacy])).rows[0].good_qty,17);
 });
 await check('Native partial predecessor goods enable explicit partial handover/start before predecessor completion',async()=>{
  const r=await fresh();await book(r);const op=(await db.query('select id from public.production_order_operations where order_id=$1 and predecessor=$2 order by position limit 1',[r.order_id,r.op_number])).rows[0].id;
  await rpc(p.worker,'op_routing_command',{p_operation_id:op,p_action:'select',p_request_id:randomUUID(),p_expected_version:0,p_machine_id:s.machines[1],p_reason:'Synthetic follow-up selection'});
  await command(p.worker,op,'configure',{p_new_count:1,p_note:'Synthetic follow-up tension'});
  let v=(await rpc(p.worker,'machine_readiness_view')).rows.find(x=>x.operation_id===op);assert.equal(v.received_quantity,0);assert.equal(v.quantity_available,1);
  v=await rpc(p.worker,'machine_readiness_command',{p_operation_id:op,p_action:'handover',p_request_id:randomUUID(),p_expected_version:v.row_version,p_quantity:1,p_reason:'Synthetic physical partial handover'});
  v=await rpc(p.worker,'machine_readiness_command',{p_operation_id:op,p_action:'start',p_request_id:randomUUID(),p_expected_version:v.row_version,p_reason:'Synthetic partial production'});assert.equal(v.work_status,'running');
  await assert.rejects(rpc(p.worker,'machine_readiness_command',{p_operation_id:op,p_action:'complete',p_request_id:randomUUID(),p_expected_version:v.row_version,p_quantity:100,p_reason:'Synthetic prohibited Phase 5 bypass'}),e=>e.code==='42501');
 });
 await check('Two native connections: same request once; distinct simultaneous presses both retained',async()=>{
  const r=await fresh(),c1=await connect(),c2=await connect();
  try{
   const args=[r.operation_id,'count',randomUUID(),r.stations[2].id,r.own_person_id,'good',1,r.section_id,null,null,null,null];
   const sql='select public.op_quantity_command($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12) result';
   const replies=await Promise.all([c1,c2].map(c=>sqlAs(p.worker,c=>c.query(sql,args),{connection:c})));assert.deepEqual(replies[0].rows[0].result,replies[1].rows[0].result);
   await Promise.all([c1,c2].map(c=>sqlAs(p.worker,c=>c.query(sql,args.map((v,i)=>i===2?randomUUID():v)),{connection:c})));
   const next=(await rpc(p.worker,'op_quantity_view')).rows.find(v=>v.operation_id===r.operation_id);assert.equal(next.finished_good,3);assert(next.stations[2].conflict_id);
   assert.equal(Number((await db.query('select count(*) n from public.production_quantity_conflicts where operation_id=$1',[r.operation_id])).rows[0].n),1);
  }finally{await c1.end();await c2.end();}
 });
 await check('Native stale machine section returns HTTP409; current section keeps OP total',async()=>{
  const r=await fresh();await book(r);const route=(await rpc(p.worker,'op_routing_view')).rows.find(v=>v.operation_id===r.operation_id);
  await rpc(p.worker,'op_routing_command',{p_operation_id:r.operation_id,p_action:'select',p_request_id:randomUUID(),p_expected_version:route.row_version,p_machine_id:s.machines[1],p_reason:'Synthetic phase 4 move'});
  await assert.rejects(book(r),e=>e.code==='PT409'&&e.status===409);const next=(await rpc(p.worker,'op_quantity_view')).rows.find(v=>v.operation_id===r.operation_id);await book(next);assert.equal((await rpc(p.worker,'op_quantity_view')).rows.find(v=>v.operation_id===r.operation_id).finished_good,2);
 });
 await check('Native count audit failure rolls back original counter and event; outsider cannot count',async()=>{
  const r=await fresh();await assert.rejects(command(p.outsider,r.operation_id,'count',{p_station_id:r.stations[2].id,p_order_employee_id:r.own_person_id,p_kind:'good',p_delta:1,p_section_id:r.section_id}),e=>e.code==='42501');
  await db.query("create function public.synthetic_quantity_abort() returns trigger language plpgsql as $$begin raise exception 'Synthetic audit failure';end$$;create trigger synthetic_quantity_abort before insert on public.production_quantity_events for each row execute function public.synthetic_quantity_abort()");
  try{await assert.rejects(book(r));assert.equal((await rpc(p.worker,'op_quantity_view')).rows.find(v=>v.operation_id===r.operation_id).finished_good,0);}finally{await db.query('drop trigger synthetic_quantity_abort on public.production_quantity_events;drop function public.synthetic_quantity_abort()');}
 });
 results.limitations.push('Phase 4 is detection only, no Phase 5 resolution/completion/quantity-change API. Historical unassigned quantities are never fabricated into machine sections. Original ownership/ACL restore and activation remain separately blocked.');
}
module.exports={run};
