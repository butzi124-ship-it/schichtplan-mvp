// Native, isolated-only test extension; runs AFTER all baseline vulnerability proofs.
// Never connects to or changes the hosted Supabase project.
const fs=require('node:fs');
const assert=require('node:assert/strict');
const {http}=require('../../scripts/db-integration/client');
const oldTables=['planner_absence_replacements','planner_absences','planner_assignments',
 'planner_availability','planner_saturday_requests','planner_shift_cancellations',
 'planner_sick_leaves','planner_swaps','planner_vacations','tool_materials'];
async function run({db,s,check,results}){
 const p=s.people;
 await check('Review-only legacy RLS candidate: real SQL prevents anon access, employee escalation and non-admin private-data access',async()=>{
  const candidate=fs.readFileSync('supabase/review/legacy-security-hardening-isolated.sql','utf8');
  await db.query(candidate);
  await db.query("notify pgrst,'reload schema'");
  await new Promise(resolve=>setTimeout(resolve,900));
  for(const table of oldTables){
   const flags=(await db.query('select relrowsecurity from pg_catalog.pg_class where oid=$1::regclass',['public.'+table])).rows[0];
   assert.equal(flags.relrowsecurity,true,'RLS must be active for '+table);
   const grant=(await db.query("select has_table_privilege('anon',$1,'SELECT,INSERT,UPDATE,DELETE') access",['public.'+table])).rows[0];
   assert.equal(grant.access,false,'Anonymous role must have no raw DML for '+table);
   const r=await http(null,'/'+table+'?limit=1');
   assert.ok([401,403].includes(r.status),'Anonymous REST access must be denied '+table+' '+r.status);
  }
  const restricted=await http(p.worker,'/planner_sick_leaves?limit=1');
  assert.equal(restricted.status,200);
  assert.deepEqual(restricted.data,[],'Ordinary employee must not see private sick-leave rows');
  const escalated=await http(p.worker,'/employees?id=eq.'+p.outsider.employee_id,
    {method:'PATCH',body:{role:'admin'},headers:{prefer:'return=representation'}});
  assert.ok([200,204,401,403].includes(escalated.status));
  assert.equal((await db.query('select role from public.employees where id=$1',[p.outsider.employee_id])).rows[0].role,'employee',
     'No foreign employee role escalation despite JWT user_metadata spoof');
  const payload={user_name:'Synthetic review probe',from_date:'2026-11-06',to_date:'2026-11-07'};
  const denied=await http(p.worker,'/planner_vacations',{method:'POST',body:payload});
  assert.ok([401,403].includes(denied.status),'Employee may not create arbitrary vacation records');
  const permitted=await http(p.manager,'/planner_vacations',{method:'POST',body:payload,headers:{prefer:'return=representation'}});
  assert.equal(permitted.status,201,'Authorized admin must retain planner write path: '+JSON.stringify(permitted.data));
  const reference=await http(p.worker,'/tool_materials?limit=1');
  assert.equal(reference.status,200,'Active employee may still read non-sensitive tool material catalog');
  assert.equal((await db.query('select good_qty from public.production_station_counts where order_id=$1',[s.legacy])).rows[0].good_qty,17);
  results.securityStatus='ISOLATED_HARDENING_CANDIDATE_VERIFIED__PRODUCTION_NOT_TESTED';
 });
}
module.exports={run};
