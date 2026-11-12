// Bridge the actual adapter to the actual SQL candidate; auth functions are synthetic.
const fs=require('node:fs');
const {fixture,ids}=require('./fixture');
const names={security_actor_context:[],security_published_schedule:[],security_own_contacts:['p_values'],
 security_assignment_command:['p_id','p_department','p_subject','p_shift_id','p_date','p_expected_version'],security_publish_assignment:['p_id','p_expected_version']};
async function accessFixture(){
 const f=await fixture();await f.install();await f.db.exec(fs.readFileSync('supabase/review/security-step4-integration-isolated.sql','utf8'));
 await f.db.query("insert into humbel_security.contact_fields values('synthetic_contact','Synthetic test-only contact approval')");
 await f.db.query("insert into humbel_security.schedule_grants values($1,$3,true,false,now()-interval '1 day',now()+interval '1 day',false,'Synthetic read scope'),($2,$3,false,true,now()-interval '1 day',now()+interval '1 day',false,'Synthetic publish scope')",[ids.worker,ids.admin,ids.d1]);
 const listeners=new Set(),calls=[];let user=ids.worker,pending=Promise.resolve();
 const enqueue=fn=>{const result=pending.then(fn);pending=result.catch(()=>{});return result;};
 const client={auth:{
  getUser:async()=>{calls.push('auth.getUser');return{data:{user:user?{id:user}:null},error:null};},
  onAuthStateChange:callback=>{listeners.add(callback);return{data:{subscription:{unsubscribe:()=>listeners.delete(callback)}}};}
 },rpc:async(name,args={})=>{
  calls.push(name);const actor=user;
  try{if(!Object.hasOwn(names,name))throw Error('Unexpected RPC');
   const keys=names[name],tableResult=name==='security_published_schedule';const result=await enqueue(()=>f.asUser(actor,d=>d.query(tableResult?'select * from public.security_published_schedule()':`select public.${name}(${keys.map((k,i)=>'$'+(i+1)).join(',')}) result`,keys.map(k=>args[k]??null))));return{data:tableResult?result.rows:result.rows[0].result,error:null};
  }catch(error){return{data:null,error:{code:error.code||'TEST_ERROR'}};}
 },from:table=>{
  calls.push('table:'+table);if(table!=='tool_materials')throw Error('Raw table fallback forbidden');let columns;
  const builder={select(value){columns=value;if(value!=='id,name,sort_order,is_active')throw Error('Explicit safe columns required');return builder;},eq(key,value){if(key!=='is_active'||value!==true)throw Error('Unexpected filter');return builder;},order(){return builder;},then(resolve,reject){
   const actor=user;enqueue(()=>f.asUser(actor,d=>d.query('select '+columns+' from public.tool_materials where is_active order by sort_order,name'))).then(r=>resolve({data:r.rows,error:null}),e=>resolve({data:null,error:{code:e.code}})).catch(reject);
  }};return builder;
 }};
 function switchUser(id,event='SIGNED_IN'){user=id;for(const listener of listeners)listener(event,{user:id?{id}:null});}
 return{...f,client,switchUser,calls,listeners};
}
module.exports={accessFixture,ids};
