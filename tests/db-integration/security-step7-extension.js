// Additive native proof only; complete atomic install is a distinct embedded test.
const assert=require('node:assert/strict'),{http}=require('../../scripts/db-integration/client');
const {source,postconditions,manifest}=require('./security-step7-review');
async function run(context){
 await require('./security-step6-extension').run(context);
 const {db,s,check,results}=context,p=s.people;assert.equal(results.tests.filter(t=>t.status==='failed').length,0);const start=results.tests.length;
 await check('Step7 native observed maintenance privilege fixture removed on all31 tables without employee data loss',async()=>{
  // Confirmed original threat replicated only in disposable native test DB.
  const before=(await db.query('select * from public.employees order by id')).rows;
  for(const t of manifest().tables)await db.query('grant truncate,references,trigger,maintain on public.'+t+' to anon,authenticated');
  await db.query('begin');try{await db.query('lock table '+manifest().tables.map(t=>'public.'+t).join(',')+' in access exclusive mode');await db.query(source());await postconditions(db);await db.query('commit');}catch(e){await db.query('rollback');throw e;}
  assert.deepEqual((await db.query('select * from public.employees order by id')).rows,before);
 });
 await check('Step7 real REST: worker cannot modify any protected own/foreign identity field or create/delete staff',async()=>{
  const before=(await db.query('select * from public.employees order by id')).rows;
  for(const target of [p.worker.employee_id,p.outsider.employee_id])for(const payload of [{role:'admin'},{auth_user_id:p.manager.id},{department_id:s.foreign},{is_active:false},{id:require('node:crypto').randomUUID()}]){
   const r=await http(p.worker,'/employees?id=eq.'+target,{method:'PATCH',body:payload,headers:{prefer:'return=representation'}});assert(!r.ok||Array.isArray(r.data)&&r.data.length===0);
  }
  assert.equal((await http(p.worker,'/employees',{method:'POST',body:{display_name:'Denied synthetic Step7',role:'admin'}})).ok,false);
  const r=await http(p.worker,'/employees?id=eq.'+p.outsider.employee_id,{method:'DELETE',headers:{prefer:'return=representation'}});assert(!r.ok||Array.isArray(r.data)&&r.data.length===0);
  assert.deepEqual((await db.query('select * from public.employees order by id')).rows,before);
 });
 await check('Step7 real REST: independently verified admin retains ordinary staff maintenance',async()=>{
  const r=await http(p.manager,'/employees?id=eq.'+p.worker.employee_id,{method:'PATCH',body:{first_name:'Synthetic Step7 admin-maintained'},headers:{prefer:'return=representation'}});assert.equal(r.ok,true);assert.equal(r.data[0].first_name,'Synthetic Step7 admin-maintained');
 });
 results.step7Status=results.tests.slice(start).every(t=>t.status==='passed')?'NATIVE_ADDITIVE_STEP7_VERIFIED__NO_PRODUCTION_CLEARANCE':'FAILED';
 results.limitations.push('Step7 original catalog captured read-only; native stack is still structural mirror, not full original restore. Existing CRUD policies outside protected15 remain open; atomic fresh-install/NOOP proof is embedded, not native.');
}
module.exports={run};
