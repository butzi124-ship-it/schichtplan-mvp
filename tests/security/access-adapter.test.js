const {test,before,after}=require('node:test');
const assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto');
const {accessFixture,ids}=require('./access-fixture');
const {create}=require('../../js/secure-access');
let f;before(async()=>{f=await accessFixture();});after(async()=>f?.db.close());
const options=()=>({client:f.client,enabled:()=>true,events:null});
const values=(id=randomUUID(),version=0,dept=ids.d1,subject=ids.worker)=>({id,department:dept,subject,shiftId:'secure-'+id,date:'2026-11-12',expectedVersion:version});

test('adapter defaults off, makes no auth/data requests, and never falls back to raw tables',async()=>{
 const start=f.calls.length,a=create({client:f.client,events:null});
 try{await assert.rejects(a.readSchedule());assert.equal(f.calls.length,start);assert.deepEqual(a.getState(),{identity:null,schedule:[],contacts:null,materials:[]});}finally{a.dispose();}
});

test('server identity ignores UI roles and rejects unenrolled/mismatched identities',async()=>{
 const a=create(options());
 try{
  f.switchUser(ids.worker);assert.equal((await a.refreshIdentity()).can_admin,false);
  f.switchUser(ids.orphan);await assert.rejects(a.refreshIdentity());assert.equal(a.getState().identity,null);
  await f.db.query("update public.employees set department_id=$1 where id=$2",[ids.d2,ids.worker]);
  f.switchUser(ids.worker);await assert.rejects(a.refreshIdentity());
 }finally{await f.db.query('update public.employees set department_id=$1 where id=$2',[ids.d1,ids.worker]);a.dispose();}
});

test('real SQL adapter plans own department and separates publishing and staff schedule reading',async()=>{
 const a=create(options()),v=values();
 try{
  f.switchUser(ids.leader);assert.equal((await a.assign(v)).version,1);await assert.rejects(a.publish({id:v.id,expectedVersion:1}));
  await assert.rejects(a.assign(values(randomUUID(),0,ids.d2,ids.other)));
  f.switchUser(ids.admin);await a.publish({id:v.id,expectedVersion:1});
  f.switchUser(ids.worker);const rows=await a.readSchedule();assert.equal(rows.length,1);assert.equal(rows[0].employee_id,ids.worker);await assert.rejects(a.assign(values()));
  f.switchUser(ids.other);assert.deepEqual(await a.readSchedule(),[]);
 }finally{a.dispose();}
});

test('deputy scope is refetched and expired/revoked delegation cannot reuse cached capability',async()=>{
 const a=create(options());
 try{
  await f.db.query("insert into humbel_security.delegations values($1,$2,$3,now()-interval '1 day',now()+interval '1 day',false,'Synthetic delegation')",[ids.deputy,ids.d1,ids.leader]);
  f.switchUser(ids.deputy);assert.equal((await a.assign(values())).version,1);
  await f.db.query('update humbel_security.delegations set revoked=true where user_id=$1',[ids.deputy]);await assert.rejects(a.assign(values()));assert.equal(a.getState().identity,null);
  await f.db.query("update humbel_security.delegations set revoked=false,valid_until=now()-interval '1 second' where user_id=$1",[ids.deputy]);await assert.rejects(a.assign(values()));
 }finally{a.dispose();}
});

test('own configured contacts use only the safe RPC; protected/foreign fields and scanner access denied',async()=>{
 const a=create(options());
 try{
  f.switchUser(ids.worker);assert.deepEqual(await a.ownContacts({synthetic_contact:'own'}),{synthetic_contact:'own'});
  for(const key of ['role','auth_user_id','department_id','employee_id'])await assert.rejects(a.ownContacts({[key]:'admin'}));
  f.switchUser(ids.other);assert.deepEqual(await a.ownContacts(),{});
  f.switchUser(ids.scanner);await assert.rejects(a.ownContacts());await assert.rejects(a.readSchedule());
  assert(!f.calls.some(name=>['table:employees','table:planner_assignments'].includes(name)));
 }finally{a.dispose();}
});

test('material reference reads preserve staff/scanner access with explicit columns',async()=>{
 const a=create(options());
 try{
  await f.db.query("insert into public.tool_materials(name) values('Synthetic active material')");
  for(const user of [ids.worker,ids.leader,ids.scanner]){f.switchUser(user);const rows=await a.readMaterials();assert.equal(rows.length,1);assert.deepEqual(Object.keys(rows[0]).sort(),['id','is_active','name','sort_order']);}
 }finally{a.dispose();}
});

test('session switch and same-user token change discard pending responses and all scoped cache',async()=>{
 const original=f.client.rpc;
 try{
  for(const [next,event] of [[ids.other,'SIGNED_IN'],[ids.worker,'TOKEN_REFRESHED'],[null,'SIGNED_OUT']]){
   f.switchUser(ids.worker);let release,started;
   const waiting=new Promise(resolve=>started=resolve);
   f.client.rpc=async(name,args)=>{const result=await original(name,args);if(name==='security_own_contacts'){started();await new Promise(resolve=>release=resolve);}return result;};
   const a=create(options());const pending=a.ownContacts();await waiting;f.switchUser(next,event);release();await assert.rejects(pending);
   assert.deepEqual(a.getState(),{identity:null,schedule:[],contacts:null,materials:[]});a.dispose();f.client.rpc=original;
  }
 }finally{f.client.rpc=original;}
});

test('offline, failed authorization and disposal clear memory; no confidential data persistence',async()=>{
 f.switchUser(ids.worker);let online=true,handler,removed=false;
 const a=create({client:f.client,enabled:()=>true,online:()=>online,events:{addEventListener:(event,fn)=>handler=fn,removeEventListener:()=>removed=true}});
 try{
  await a.ownContacts();assert.equal(a.getState().contacts.synthetic_contact,'own');online=false;handler();await assert.rejects(a.ownContacts());assert.equal(a.getState().contacts,null);
  online=true;await a.refreshIdentity();await f.db.query('update humbel_security.principals set enabled=false where user_id=$1',[ids.worker]);await assert.rejects(a.ownContacts());assert.equal(a.getState().identity,null);
 }finally{await f.db.query('update humbel_security.principals set enabled=true where user_id=$1',[ids.worker]);a.dispose();}
 assert.equal(removed,true);assert.equal(f.listeners.size,0);
 const source=require('fs').readFileSync('js/secure-access.js','utf8');assert(!/localStorage|indexedDB|sessionStorage|\.from\(['"]employees/.test(source.replace(/\/\*[\s\S]*?\*\//g,'')));
});

test('old pending response cannot overwrite or clear the new actor cache',async()=>{
 const original=f.client.rpc;let release,started,first=true;
 const waiting=new Promise(resolve=>started=resolve);f.switchUser(ids.worker);
 f.client.rpc=async(name,args)=>{const result=await original(name,args);if(name==='security_own_contacts'&&first){first=false;started();await new Promise(resolve=>release=resolve);}return result;};
 const a=create(options());
 try{const old=a.ownContacts();await waiting;f.switchUser(ids.other);assert.deepEqual(await a.ownContacts(),{});release();await assert.rejects(old);assert.equal(a.getState().identity.user_id,ids.other);assert.deepEqual(a.getState().contacts,{});}
 finally{f.client.rpc=original;a.dispose();}
});

test('context identity mismatch and anonymous SQL context access are rejected',async()=>{
 await assert.rejects(f.asUser(null,d=>d.query('select public.security_actor_context()'),{anonymous:true}),e=>e.code==='42501');
 f.switchUser(ids.worker);const original=f.client.rpc;
 f.client.rpc=async(name,args)=>{const result=await original(name,args);return name==='security_actor_context'?{...result,data:{...result.data,user_id:ids.other}}:result;};
 const a=create(options());try{await assert.rejects(a.refreshIdentity());assert.equal(a.getState().identity,null);}finally{f.client.rpc=original;a.dispose();}
});
