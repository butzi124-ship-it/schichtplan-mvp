'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const crypto=require('node:crypto');
const {rowsByTable}=require('../helpers/supabase-mock');
const {makeDemoData,validate}=require('../../scripts/preview/fabos-demo-data');
const {buildDemoPreview,renderDemoPlan}=require('../../scripts/preview/fabos-demo-adapter');
const {createPreview}=require('../../scripts/preview/server');

test('50 invented workers in 8 named departments, 6 machines each; plausible route and machine reservations',()=>{
 const d=makeDemoData();
 assert.doesNotThrow(()=>validate(d));
 assert.equal(d.departments.length,8);
 assert.equal(d.employees.length,50);
 assert.equal(d.machines.length,48);
 assert.equal(d.orders.length,50);
 assert.equal(d.operations.length,277);
 assert(d.operations.filter(o=>o.state==='running').length>=8);
 assert(d.operations.filter(o=>o.state==='queued').length>50);
 assert(d.operations.every(o=>o.description.length>60&&o.setup_minutes>=15&&o.cycle_seconds>=60));
 for(const dep of d.departments){assert(d.employees.filter(e=>e.department_id===dep.id).length>=6);assert.equal(d.machines.filter(m=>m.department_id===dep.id).length,6);}
 assert.deepEqual(d.departments.map(x=>x.name),['Dreherei','Fräserei','Verzahnerei','Entgraterei','Härterei','Schleiferei','Qualitätssicherung','Montage']);
 const serialized=JSON.stringify(d);
 assert(!serialized.includes('Lavdrim')&&!serialized.includes('Roger')&&!serialized.includes('Dashmir'));
 assert(!serialized.includes('https://')&&!serialized.includes('sb_publishable_'));
});

test('legacy preview exposes all 50 seeded BA including planned OP in an isolated fixture',()=>{
 const original=JSON.stringify(rowsByTable);
 const {data,rows}=buildDemoPreview(rowsByTable);
 assert.equal(JSON.stringify(rowsByTable),original);
 assert.equal(rows.departments.length,8);
 assert.equal(rows.employees.length,51); // separate isolated admin plus 50 invented personnel
 assert.equal(rows.production_machines.length,48);
 assert.equal(rows.tools.length,20);
 assert(rows.tools.some(t=>t.label==='Schaftfräser'));
 assert(rows.tools.some(t=>t.label==='Drehmeißel Schruppen'));
 assert(rows.tools.every(t=>t.t_number.startsWith('DEMO-')));
 assert.equal(rows.production_orders.length,50);
 assert(rows.production_orders.every(o=>['queued','running','completed'].includes(o.status)));
 assert(rows.production_orders.some(o=>o.status==='queued'));
 assert(rows.production_order_stations.every(s=>data.operations.some(o=>o.id===s.id)));
 assert(rows.production_station_counts.every(c=>rows.production_order_stations.some(s=>s.id===c.station_id)));
 const html=renderDemoPlan(data);
 assert(html.includes('Nächster Auftrag'));
 assert(html.includes('DEMO-BA-810001'));
 assert(html.includes('keine echte PlanOS'));
 assert(!html.includes('production database'));
});

test('demo queue requires preview login and is opt-in, never modifies default server', {timeout:20000},async()=>{
 const old=process.env.HUMBEL_PREVIEW_DEMO;
 const password=crypto.randomBytes(32).toString('hex');
 const backend={actors:{employee:{uid:'synthetic',role:'employee',name:'Test Mitarbeiter'}},rpc:async()=>({}),close:async()=>{}};
 try{
  delete process.env.HUMBEL_PREVIEW_DEMO;
  const plain=await createPreview({passwords:{employee:password},backend});
  try{const resp=await fetch(plain.url+'/demo-plan');assert.equal(resp.status,401);}finally{await plain.close();}
  process.env.HUMBEL_PREVIEW_DEMO='1';
  const testBackend={...backend,close:async()=>{}};
  const preview=await createPreview({passwords:{employee:password},backend:testBackend});
  try{
   let r=await fetch(preview.url+'/demo-plan');assert.equal(r.status,401);
   r=await fetch(preview.url+'/login',{method:'POST',headers:{Origin:preview.url,'Content-Type':'application/json'},body:JSON.stringify({account:'employee',password})});
   assert.equal(r.status,200);
   const cookie=r.headers.get('set-cookie').split(';')[0];
   r=await fetch(preview.url+'/demo-plan',{headers:{Cookie:cookie}});
   assert.equal(r.status,200);
   const html=await r.text();
   assert(html.includes('Maschinenbelegung'));
   assert(html.includes('DEMO-BA-810001'));
   r=await fetch(preview.url+'/preview-bootstrap.js',{headers:{Cookie:cookie}});
   const script=await r.text();
   assert(script.includes('window.__PREVIEW_DEMO=true'));
   assert(script.includes('DEMO-BA-810001'));
   r=await fetch(preview.url+'/logout',{method:'POST',headers:{Origin:preview.url,Cookie:cookie,'Content-Type':'application/json'},body:'{}'});
   assert.equal(r.status,200);
   r=await fetch(preview.url+'/demo-plan',{headers:{Cookie:cookie}});
   assert.equal(r.status,401);
  }finally{await preview.close();}
 }finally{if(old===undefined)delete process.env.HUMBEL_PREVIEW_DEMO;else process.env.HUMBEL_PREVIEW_DEMO=old;}
});
