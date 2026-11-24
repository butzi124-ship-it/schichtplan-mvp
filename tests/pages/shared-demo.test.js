'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const os=require('node:os');
const crypto=require('node:crypto');
const {createSharedDemoStore}=require('../../scripts/pages/shared-demo-store');
const {createSharedDemoServer}=require('../../scripts/pages/shared-demo-server');

test('file-backed FABOS seed is scoped, persists allowed actions and rejects role bypasses',()=>{
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'fabos-shared-store-'));
 try{
  const file=path.join(dir,'workspace.json'),store=createSharedDemoStore(file);
  const master=store.snapshot('admin'),leader=store.snapshot('leader'),worker=store.snapshot('employee');
  assert.equal(master.data.departments.length,8);assert.equal(master.data.employees.length,50);
  assert.equal(master.data.machines.length,48);assert.equal(master.data.orders.length,50);
  assert.equal(master.data.operations.length,277);
  assert.equal(leader.data.departments.length,1);assert.equal(leader.data.machines.length,6);
  assert.equal(worker.data.departments.length,1);assert.equal(worker.data.machines.length,6);
  assert.equal(worker.actor.role,'employee');
  assert.throws(()=>store.change('employee','reset',{}),e=>e.status===403);
  assert.throws(()=>store.change('employee','create_employee',{first:'Unauthorized',last:'Person',department_id:master.data.departments[0].id}),e=>e.status===403);
  assert.throws(()=>store.change('leader','create_employee',{first:'Unauthorized',last:'Person',department_id:leader.data.departments[0].id}),e=>e.status===403);
  const other=master.data.operations.find(o=>o.department_id!==leader.actor.department_id&&o.state!=='completed');
  assert(other);assert.throws(()=>store.change('leader','approve',{op_id:other.id}),e=>e.status===403);
  const local=leader.data.operations.find(o=>o.state==='running'||o.state==='queued');
  assert(local);store.change('leader','approve',{op_id:local.id});
  assert(store.snapshot('leader').approved.includes(local.id));
  const running=master.data.operations.find(x=>x.state==='running'&&x.good_qty<master.data.orders.find(o=>o.id===x.order_id).ba_quantity);
  assert(running);
  const before=running.good_qty;
  const opOfOtherPerson=worker.data.operations.find(o=>o.state==='running'&&o.assigned_employee_id!==worker.actor.employee_id);
  if(opOfOtherPerson)assert.throws(()=>store.change('employee','book_qty',{op_id:opOfOtherPerson.id,qty:1}),e=>e.status===403);
  store.change('admin','book_qty',{op_id:running.id,qty:1});
  const fr=master.data.departments.find(d=>d.code==='FR'),m=master.data.machines.find(m=>m.department_id===fr.id);
  store.change('admin','create_order',{department_id:fr.id,machine_id:m.id,employee_id:worker.actor.employee_id,article:'Schaltwelle Test',article_no:'ART-999',material:'16MnCr5',instruction:'Aufspannen und fertigen, Erststück kontrollieren.',qty:12});
  store.change('admin','create_employee',{first:'Carla',last:'Beispiel',department_id:fr.id});
  const reopened=createSharedDemoStore(file),snapshot=reopened.snapshot('admin');
  assert.equal(snapshot.data.orders.length,51);assert.equal(snapshot.data.employees.length,51);
  assert.equal(snapshot.data.operations.find(o=>o.id===running.id).good_qty,before+1);
  assert(snapshot.approved.includes(local.id));
  const created=snapshot.data.employees.find(x=>x.display_name==='Carla Beispiel');
  const dr=snapshot.data.departments.find(x=>x.code==='DR');
  reopened.change('admin','move_employee',{employee_id:created.id,department_id:dr.id});
  assert.equal(reopened.snapshot('admin').data.employees.find(x=>x.id===created.id).department_id,dr.id);
  reopened.change('admin','reset',{});
  assert.equal(createSharedDemoStore(file).snapshot('admin').data.orders.length,50);
  assert.equal(createSharedDemoStore(file).snapshot('admin').data.employees.length,50);
 }finally{fs.rmSync(dir,{recursive:true,force:true});}
});

test('HTTP preview gates all data behind password and enforces role permissions across clients and restart',{timeout:90000},async()=>{
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'fabos-shared-http-'));
 const file=path.join(dir,'workspace.json');
 const passwords={admin:crypto.randomBytes(32).toString('hex'),leader:crypto.randomBytes(32).toString('hex'),employee:crypto.randomBytes(32).toString('hex')};
 let server;
 async function request(url,cookie,method='GET',body){
  const headers={};if(cookie)headers.Cookie=cookie;if(method==='POST'){headers.Origin=server.url;headers['Content-Type']='application/json';}
  return fetch(server.url+url,{method,headers,body:body?JSON.stringify(body):undefined,redirect:'manual'});
 }
 async function signIn(account){
  const response=await request('/login',null,'POST',{account,password:passwords[account]});
  assert.equal(response.status,200);
  const cookie=response.headers.get('set-cookie');
  assert.match(cookie,/HttpOnly/);assert.match(cookie,/SameSite=Strict/);
  return cookie.split(';')[0];
 }
 try{
  server=await createSharedDemoServer({passwords,filePath:file,online:false});
  assert.equal((await request('/api/state')).status,401);
  assert.equal((await request('/fertigung.html')).status,401);
  assert.equal((await request('/')).status,200);
  const master=await signIn('admin'),leader=await signIn('leader'),employee=await signIn('employee');
  const adminState=await (await request('/api/state',master)).json();
  assert.equal(adminState.data.machines.length,48);assert.equal(adminState.data.orders.length,50);
  assert.equal((await (await request('/api/state',employee)).json()).data.machines.length,6);
  assert.equal((await request('/app.html',employee)).status,403);
  const staffDenied=await request('/api/action',employee,'POST',{action:'create_employee',payload:{first:'No',last:'Access',department_id:adminState.data.departments[0].id}});
  assert.equal(staffDenied.status,403);
  const wrongDept=adminState.data.operations.find(o=>o.department_id!==adminState.data.departments.find(d=>d.code==='FR').id&&o.state!=='completed');
  const badApproval=await request('/api/action',leader,'POST',{action:'approve',payload:{op_id:wrongDept.id}});
  assert.equal(badApproval.status,403);
  const fr=adminState.data.departments.find(d=>d.code==='FR');
  const assigned=adminState.data.employees.find(e=>e.first_name==='Mila'&&e.last_name==='Fischer');
  const machine=adminState.data.machines.find(m=>m.department_id===fr.id);
  const mutation=await request('/api/action',master,'POST',{action:'create_order',payload:{department_id:fr.id,machine_id:machine.id,employee_id:assigned.id,article:'Prüfteil',article_no:'TEST-42',material:'C45',instruction:'Rüsten, fertigen und prüfen.',qty:9}});
  assert.equal(mutation.status,200,await mutation.text());
  assert.equal((await (await request('/api/state',master)).json()).data.orders.length,51);
  const page=await request('/fertigung.html?role=admin',employee);
  assert.equal(page.status,200);
  assert.match(await page.text(),/window.__FABOS_SERVER_ROLE="employee"/);
  const page2=await request('/fertigung.html',master);
  assert.match(await page2.text(),/window.__FABOS_SERVER_ROLE="master"/);
  const logout=await request('/logout',master,'POST',{});
  assert.equal(logout.status,200);
  assert.equal((await request('/api/state',master)).status,401);
  await server.close();server=null;
  server=await createSharedDemoServer({passwords,filePath:file,online:false});
  const again=await signIn('admin');
  assert.equal((await (await request('/api/state',again)).json()).data.orders.length,51);
 }finally{if(server)await server.close();fs.rmSync(dir,{recursive:true,force:true});}
});
