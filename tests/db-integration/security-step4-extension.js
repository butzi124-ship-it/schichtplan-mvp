// All prior baseline/BA-OP/security proofs remain mandatory before this integration test.
const fs=require('node:fs'),assert=require('node:assert/strict'),{randomUUID}=require('node:crypto');
const {rpc,http,config}=require('../../scripts/db-integration/client');
const {create}=require('../../js/secure-access');
async function run(context){
 await require('./security-step3-extension').run(context);
 const {db,s,check,results}=context,p=s.people;assert.equal(results.tests.filter(t=>t.status==='failed').length,0);
 await db.query(fs.readFileSync('supabase/review/security-step4-integration-isolated.sql','utf8'));
 await db.query("notify pgrst,'reload schema'");await new Promise(resolve=>setTimeout(resolve,900));
 let user=p.worker;const listeners=new Set();
 const client={auth:{getUser:async()=>{const response=await fetch(config().auth+'/user',{headers:{authorization:'Bearer '+user.token}});return response.ok?{data:{user:await response.json()},error:null}:{data:{user:null},error:{message:'Rejected'}};},onAuthStateChange:fn=>{listeners.add(fn);return{data:{subscription:{unsubscribe:()=>listeners.delete(fn)}}};}},
 rpc:async(name,args)=>{try{return{data:await rpc(user,name,args),error:null};}catch(error){return{data:null,error:{code:error.code}};}},
 from:table=>{assert.equal(table,'tool_materials');const b={select:columns=>{assert.equal(columns,'id,name,sort_order,is_active');return b;},eq:()=>b,order:()=>b,then:(resolve,reject)=>http(user,'/tool_materials?select=id,name,sort_order,is_active&is_active=eq.true&order=sort_order,name').then(r=>resolve(r.ok?{data:r.data,error:null}:{data:null,error:{code:r.status}}),reject)};return b;}
 };
 const a=create({client,enabled:()=>true,online:()=>true,events:null});
 const switchUser=next=>{user=next;for(const listener of listeners)listener('SIGNED_IN');};
 const start=results.tests.length;
 try{
  await check('Step4 real adapter + GoTrue JWT: server identity and safe own contacts, no protected profile fields',async()=>{
   const identity=await a.refreshIdentity();assert.equal(identity.user_id,p.worker.id);assert.equal(identity.employee_id,p.worker.employee_id);assert.equal(identity.can_admin,false);
   assert.deepEqual(await a.ownContacts({synthetic_contact:'Synthetic adapter contact'}),{synthetic_contact:'Synthetic adapter contact'});
   await assert.rejects(a.ownContacts({role:'admin'}));
   assert.equal((await db.query('select role from public.employees where id=$1',[p.worker.employee_id])).rows[0].role,'employee');
  });
  await check('Step4 real adapter: own AL plan, distinct publication and minimal employee read; foreign department denied',async()=>{
   const id=randomUUID();switchUser(p.reviewer);
   assert.equal((await a.assign({id,department:s.dept,subject:p.worker.employee_id,shiftId:'step4-'+id,date:'2026-11-12',expectedVersion:0})).version,1);
   await assert.rejects(a.assign({id:randomUUID(),department:s.foreign,subject:p.outsider.employee_id,shiftId:'foreign',date:'2026-11-12',expectedVersion:0}));
   await assert.rejects(a.publish({id,expectedVersion:1}));switchUser(p.manager);await a.publish({id,expectedVersion:1});
   switchUser(p.worker);const rows=await a.readSchedule();assert(rows.some(r=>r.assignment_id===id));
   assert(rows.every(r=>Object.keys(r).every(k=>['assignment_id','employee_id','shift_id','shift_date'].includes(k))));
   switchUser(p.outsider);assert.deepEqual(await a.readSchedule(),[]);
  });
  await check('Step4 real adapter: account change clears old contact cache; revoked identity and scanner plan denied',async()=>{
   switchUser(p.worker);await a.ownContacts();assert(a.getState().contacts);
   switchUser(p.outsider);assert.equal(a.getState().contacts,null);assert.deepEqual(await a.ownContacts(),{});
   switchUser(p.worker);await db.query('update humbel_security.principals set enabled=false where user_id=$1',[p.worker.id]);
   try{await assert.rejects(a.refreshIdentity());assert.equal(a.getState().identity,null);}finally{await db.query('update humbel_security.principals set enabled=true where user_id=$1',[p.worker.id]);}
   switchUser(p.scanner);await assert.rejects(a.readSchedule());assert(Array.isArray(await a.readMaterials()));
  });
 }finally{a.dispose();}
 results.step4IntegrationStatus=results.tests.slice(start).every(t=>t.status==='passed')?'ISOLATED_NATIVE_ADAPTER_VERIFIED__NO_PRODUCTION_CLEARANCE':'FAILED';
 results.limitations.push('Step4 adapter is isolated and disabled by default; not loaded by production HTML. Legacy broad SELECT*, planner writes, private workflows, frontend migration and authoritative restore remain open.');
}
module.exports={run};
