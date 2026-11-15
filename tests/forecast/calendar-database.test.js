const {test,mock}=require('node:test');const a=require('node:assert/strict');const {randomUUID}=require('node:crypto');const {forecastDatabase,seedForecast,manager,department}=require('../helpers/forecast-db');const {calculate}=require('../../js/production-forecast');

for(const [stamp,day,sameDay] of [
 ['2026-11-07T22:30:00Z','2026-11-07',false], // 23:30 Berlin: 53 min OP crosses midnight
 ['2026-11-07T23:30:00Z','2026-11-08',true],  // UTC/Berlin dates differ; Berlin is 00:30
 ['2026-03-29T00:30:00Z','2026-03-29',true],  // spring DST transition
 ['2026-10-25T01:30:00Z','2026-10-25',true],  // repeated autumn hour
]) test('publication respects Berlin calendar at '+stamp,async()=>{
 let isolated;
 try{
  mock.timers.enable({apis:['Date'],now:Date.parse(stamp)});isolated=await forecastDatabase();await seedForecast(isolated);
  const clock=(await isolated.db.query("select now() as transaction_time,clock_timestamp() as wall_time,(now() at time zone 'Europe/Berlin')::date::text as day")).rows[0];
  a.equal(clock.transaction_time.toISOString(),stamp.replace('Z','.000Z'));a.equal(clock.wall_time.toISOString(),stamp.replace('Z','.000Z'));a.equal(clock.day,day);
  const src=await isolated.rpc(manager,'production_forecast_input',{p_department_id:department,p_request_id:randomUUID()});
  a.equal(new Date(src.as_of).getTime(),Date.now());a.equal(src.config.start_date,day);
  const result=calculate(src),next=result.results.find(r=>r.operation_id===src.jobs.find(j=>j.op_number===20).id);
  a.equal(next.color,'green');a.equal(next.arrival_today,sameDay);a.ok(next.arrival_at);
  const arrivalDay=new Intl.DateTimeFormat('en-CA',{timeZone:'Europe/Berlin',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(next.arrival_at));
  a.equal(arrivalDay,sameDay?day:'2026-11-08');
  await isolated.rpc(null,'production_forecast_publish',{p_department_id:department,p_actor_id:manager,p_run_id:randomUUID(),p_result:result},'service_role');
  const readiness=await isolated.rpc(manager,'machine_readiness_view');
  a.equal(readiness.rows.find(r=>r.op_number===20).material_status,sameDay?'expected':'planned');
  a.ok(readiness.rows.every(r=>r.received_quantity===0&&r.work_status==='planned'));
  a.equal((await isolated.db.query('select count(*)::int n from public.production_operation_arrivals')).rows[0].n,sameDay?2:0);
 }finally{try{await isolated?.db.close();}finally{mock.timers.reset();}}
});
