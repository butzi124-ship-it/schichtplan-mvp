// Real JWT/PostgREST and independent connections in the pinned ephemeral stack only.
const fs=require('node:fs'),assert=require('node:assert/strict'),{randomUUID}=require('node:crypto');
const {rpc,http,connect,sqlAs}=require('../../scripts/db-integration/client');
async function run(context){
 await require('./routing-extension').run(context);
 const {db,s,check,results}=context,p=s.people;let n=813200;
 await db.query(fs.readFileSync('supabase/review/setup-phase3.sql','utf8'));await db.query("notify pgrst,'reload schema'");await new Promise(r=>setTimeout(r,1200));
 const cmd=(u,action,op=null,run=null,extra={})=>rpc(u,'setup_time_command',{p_action:action,p_request_id:randomUUID(),p_operation_id:op,p_run_id:run,...extra});
 const fresh=async()=>{const snap=await rpc(p.worker,'ba_entry_command',{p_action:'create',p_ba_number:String(n++),p_article_number:'SYNTH-ACCEPT',p_quantity:7});const op=snap.op_rows[0].id;await rpc(p.worker,'op_routing_command',{p_operation_id:op,p_action:'select',p_request_id:randomUUID(),p_expected_version:0,p_machine_id:s.machines[0],p_reason:'Synthetic phase 3 explicit adoption'});return op;};
 const start=(u,op,run=null,mode='external')=>cmd(u,'start',op,run,{p_mode:mode});
 const segments=async rid=>(await db.query('select * from public.production_operation_activity_segments where setup_run_id=$1 order by started_at',[rid])).rows;
 await check('Phase 3 defaults off; anonymous/scanner/raw-table access denied with real JWT',async()=>{
  assert.deepEqual((await rpc(p.worker,'setup_time_view')).rows,[]);await assert.rejects(cmd(p.worker,'pause'));
  assert.ok([401,403].includes((await http(null,'/rpc/setup_time_view',{method:'POST',body:{}})).status));
  for(const table of ['production_setup_runs','production_setup_events','production_setup_corrections','production_setup_interruptions']) assert.ok([401,403].includes((await http(p.worker,'/'+table+'?limit=1')).status));
  await db.query('update public.production_setup_settings set enabled=true');await cmd(p.worker,'logout');await cmd(p.reviewer,'logout');
 });
 await check('Native mode switches, personal stop and shared finish preserve originals and quantities',async()=>{
  const op=await fresh(),r=await start(p.worker,op,null,'internal');await start(p.worker,op,r.run_id);assert.equal((await start(p.reviewer,op)).run_id,r.run_id);
  await cmd(p.worker,'end_own',op,r.run_id);assert((await segments(r.run_id)).some(a=>a.user_id===p.reviewer.id&&!a.ended_at));
  await start(p.worker,op,r.run_id,'internal');await cmd(p.worker,'finish',op,r.run_id);const a=await segments(r.run_id);
  assert(a.every(a=>a.ended_at));assert.equal(a.at(-1).ended_at.getTime(),a.at(-2).ended_at.getTime());await assert.rejects(start(p.worker,op,r.run_id));
  assert.equal((await db.query('select good_qty from public.production_station_counts where order_id=$1',[s.legacy])).rows[0].good_qty,17);assert.equal((await db.query('select scrap_total from public.production_order_stations where id=$1',[s.station])).rows[0].scrap_total,3);
 });
 await check('Native pause/logout/takeover are personal; return never automatically starts a timer',async()=>{
  const op=await fresh(),r=await start(p.worker,op);await start(p.reviewer,op);await cmd(p.worker,'pause');await cmd(p.worker,'return');
  assert.equal((await rpc(p.worker,'setup_time_view')).resume_run,r.run_id);assert(!(await segments(r.run_id)).some(a=>a.user_id===p.worker.id&&!a.ended_at));
  await start(p.worker,op,r.run_id);await cmd(p.reviewer,'takeover',op,null,{p_previous_user:p.worker.id});assert((await segments(r.run_id)).some(a=>a.user_id===p.reviewer.id&&!a.ended_at));
  await cmd(p.reviewer,'logout');await cmd(p.worker,'finish',op,r.run_id);assert.equal((await rpc(p.worker,'setup_time_view')).resume_run,null);
 });
 await check('Two connections: duplicate start and concurrent pause/shared close cannot double-close or lose colleague',async()=>{
  const op=await fresh(),c1=await connect(),c2=await connect(),request=randomUUID();
  try{
   const args=['start',request,op,null,'external',null,null,null,null,null];
   const replies=await Promise.all([c1,c2].map(c=>sqlAs(p.worker,c=>c.query('select public.setup_time_command($1,$2,$3,$4,$5,$6,$7,$8,$9,$10) result',args),{connection:c})));
   assert.deepEqual(replies[0].rows[0].result,replies[1].rows[0].result);const rid=replies[0].rows[0].result.run_id;assert.equal((await segments(rid)).length,1);await start(p.reviewer,op,rid);
   await Promise.all([sqlAs(p.worker,c=>c.query("select public.setup_time_command('pause',$1)",[randomUUID()]),{connection:c1}),sqlAs(p.reviewer,c=>c.query("select public.setup_time_command('finish',$1,$2,$3)",[randomUUID(),op,rid]),{connection:c2})]);
   assert((await segments(rid)).every(a=>a.ended_at));assert.equal((await rpc(p.worker,'setup_time_view')).resume_run,null);await cmd(p.worker,'return');
  }finally{await c1.end();await c2.end();}
 });
 await check('Two connections: routing against personal logout closes old run once and retains machine sections',async()=>{
  const op=await fresh(),r=await start(p.worker,op),route=(await rpc(p.worker,'op_routing_view')).rows.find(x=>x.operation_id===op),c1=await connect(),c2=await connect();
  try{
   await Promise.all([sqlAs(p.worker,c=>c.query("select public.setup_time_command('logout',$1)",[randomUUID()]),{connection:c1}),sqlAs(p.worker,c=>c.query("select public.op_routing_command($1,'select',$2,$3,$4,$5,$6)",[op,randomUUID(),route.row_version,s.machines[1],'Synthetic concurrent machine move','synthetic-change']),{connection:c2})]);
   assert((await segments(r.run_id)).every(a=>a.ended_at));assert.equal((await db.query('select end_reason from public.production_setup_runs where id=$1',[r.run_id])).rows[0].end_reason,'machine_change');
   const next=await start(p.worker,op);assert.notEqual(next.run_id,r.run_id);await cmd(p.worker,'finish',op,next.run_id);
  }finally{await c1.end();await c2.end();}
 });
 await check('Native correction keeps original immutable; requires reason, denies foreign edits and stale requests',async()=>{
  const op=await fresh(),r=await start(p.worker,op);await cmd(p.worker,'finish',op,r.run_id);const original=(await segments(r.run_id))[0];const v=(await rpc(p.worker,'setup_time_view')).rows.flatMap(x=>x.runs).find(x=>x.id===r.run_id).segments[0];const {can_correct,correction_count,...expected}=v;
  const extra={p_segment_id:v.id,p_expected:expected,p_change:{started_at:v.started_at,ended_at:v.ended_at,activity:'setup_internal'}};
  await assert.rejects(cmd(p.worker,'correct',op,r.run_id,extra));await assert.rejects(cmd(p.outsider,'correct',op,r.run_id,{...extra,p_reason:'Synthetic foreign attempt'}));
  await cmd(p.worker,'correct',op,r.run_id,{...extra,p_reason:'Synthetic audited correction'});assert.deepEqual((await segments(r.run_id))[0],original);
  await assert.rejects(cmd(p.worker,'correct',op,r.run_id,{...extra,p_reason:'Synthetic stale correction'}),e=>e.code==='PT409'&&e.status===409);
 });
 await check('Native audit failure rolls back timer start and finish atomically',async()=>{
  const op=await fresh();await db.query("create function public.synthetic_setup_abort() returns trigger language plpgsql as $$begin raise exception 'Synthetic setup audit failure';end$$;create trigger synthetic_setup_abort before insert on public.production_setup_events for each row execute function public.synthetic_setup_abort()");
  try{await assert.rejects(start(p.worker,op));assert.equal(Number((await db.query('select count(*) n from public.production_setup_runs r join public.production_operation_machine_sections x on x.id=r.section_id where x.operation_id=$1',[op])).rows[0].n),0);}finally{await db.query('drop trigger synthetic_setup_abort on public.production_setup_events;drop function public.synthetic_setup_abort()');}
 });
 await check('Native external setup allows running production while internal setup rejects machine occupation',async()=>{
  const op=await fresh();await db.query("update public.production_operation_execution set work_status='running' where operation_id=$1",[op]);const r=await start(p.worker,op);await assert.rejects(start(p.worker,op,r.run_id,'internal'),e=>e.code==='PT409');await cmd(p.worker,'finish',op,r.run_id);await db.query("update public.production_operation_execution set work_status='paused' where operation_id=$1",[op]);
 });
 results.limitations.push('Phase 3 requires separately approved trusted authority provisioning, original schema/ACL restore and review SQL activation. Embedded tests do not replace native concurrency evidence. No MOBATIME integration or Phase 4 counter attribution.');
}
module.exports={run};
