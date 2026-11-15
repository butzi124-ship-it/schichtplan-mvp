const {test, mock}=require('node:test');
const assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto');
const {forecastDatabase,seedForecast,manager,department}=require('../helpers/forecast-db');
const {calculate}=require('../../js/production-forecast');
const clock=require('../fixtures/forecast-clock.json');

for(const instant of [clock.instant,'2026-11-07T23:30:00Z']) test('shared browser fixture has Berlin dates, same-day expected arrival and green/orange/red at '+instant, async()=>{
  mock.timers.enable({apis:['Date'],now:Date.parse(instant)});
  let api;
  try{
    api=await forecastDatabase();await seedForecast(api);
    assert.equal((await api.db.query('show timezone')).rows[0].TimeZone,clock.zone);
    const recalc=async()=>{
      mock.timers.tick(1000);
      const src=await api.rpc(manager,'production_forecast_input',{p_department_id:department,p_request_id:randomUUID()});
      const result=calculate(src);
      await api.rpc(null,'production_forecast_publish',{p_department_id:department,p_actor_id:manager,p_run_id:randomUUID(),p_result:result},'service_role');
      return result.results.find(r=>r.operation_id===src.jobs.find(j=>j.op_number===10).id);
    };
    assert.equal((await recalc()).color,'green');
    assert.equal((await api.rpc(manager,'machine_readiness_view')).rows.find(r=>r.op_number===20).material_status,'expected');
    const times=(await api.db.query("select to_char(now() at time zone 'Europe/Berlin','HH24:MI') s,to_char((now() at time zone 'Europe/Berlin')+interval '75 minutes','HH24:MI') e")).rows[0];
    await api.db.query('update public.shift_definitions set start_time=$1,end_time=$2',[times.s,times.e]);
    assert.equal((await recalc()).color,'orange');
    await api.db.exec("insert into public.planner_sick_leaves(user_name,from_date,to_date) values('Synthetic Operator',current_date,current_date+1)");
    assert.equal((await recalc()).color,'red');
    assert.ok((await api.rpc(manager,'machine_readiness_view')).rows.every(r=>r.received_quantity===0));
  }finally{try{await api?.db.close();}finally{mock.timers.reset();}}
});

test('explicit browser-fixture time advance expires SQL cache and derived arrival together',async()=>{
  mock.timers.enable({apis:['Date'],now:Date.parse(clock.instant)});
  let api;
  try{
    api=await forecastDatabase();await seedForecast(api);
    await api.db.exec('update public.production_forecast_settings set max_age_seconds=2');
    const source=await api.rpc(manager,'production_forecast_input',{p_department_id:department,p_request_id:randomUUID()});
    await api.rpc(null,'production_forecast_publish',{p_department_id:department,p_actor_id:manager,p_run_id:randomUUID(),p_result:calculate(source)},'service_role');
    assert.equal((await api.rpc(manager,'production_forecast_view')).rows[0].color,'green');
    assert.equal((await api.rpc(manager,'machine_readiness_view')).rows.find(r=>r.op_number===20).material_status,'expected');
    mock.timers.tick(3000);
    assert.equal((await api.rpc(manager,'production_forecast_view')).rows[0].color,'unknown');
    assert.equal((await api.rpc(manager,'machine_readiness_view')).rows.find(r=>r.op_number===20).material_status,'planned');
    assert.ok((await api.rpc(manager,'machine_readiness_view')).rows.every(r=>r.received_quantity===0));
  }finally{try{await api?.db.close();}finally{mock.timers.reset();}}
});
