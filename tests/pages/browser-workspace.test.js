'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const {makeDemoData}=require('../../scripts/preview/fabos-demo-data');
const script=fs.readFileSync('scripts/pages/browser-workspace.js','utf8');
function store(backing){
 const map=backing||new Map();
 return {getItem:key=>map.has(key)?map.get(key):null,setItem:(key,val)=>map.set(key,String(val)),
  removeItem:key=>map.delete(key)};
}
function openWorkspace(local,session){
 const window={localStorage:local,sessionStorage:session,location:{pathname:'/fertigung.html',reload(){}},addEventListener(){}};
 vm.runInNewContext(script,{window,console},{filename:'browser-workspace.js'});
 return window.FABOS_BROWSER_WORKSPACE;
}
test('Github-only seeded data and local persistence are fictional and structurally complete',()=>{
 const d=makeDemoData();
 const local=store(),first=openWorkspace(local,store()),snapshot={data:d,approved:[]};
 assert(first.valid(snapshot));
 assert.equal(first.save(snapshot),true);
 const second=openWorkspace(local,store());
 assert.equal(second.load().data.employees.length,50);
 assert.equal(second.load().data.machines.length,48);
 assert.equal(second.load().data.orders.length,50);
 assert.equal(second.load().data.operations.length,277);
 const running=d.operations.find(o=>o.state==='running'&&o.good_qty<d.orders.find(order=>order.id===o.order_id).ba_quantity);
 assert(running);
 snapshot.data.operations.find(o=>o.id===running.id).good_qty+=1;
 snapshot.approved.push(running.id);
 first.save(snapshot);
 assert.equal(second.load().data.operations.find(o=>o.id===running.id).good_qty,running.good_qty);
 assert(second.load().approved.includes(running.id));
 const copy=second.decode(second.encode(snapshot));
 assert.equal(copy.approved[0],running.id);
 assert.equal(copy.data.orders.length,50);
 first.reset();
 assert.equal(second.load(),null);
});
test('only recognized, validated FABOS fictional backups can be imported',()=>{
 const w=openWorkspace(store(),store()),snapshot={data:makeDemoData(),approved:[]};
 assert.throws(()=>w.decode('{"password":"fake"}'),/Keine gültige FABOS/);
 assert.throws(()=>w.decode('{'),/JSON/);
 assert.throws(()=>w.decode('x'.repeat(3*1024*1024+1)),/groß/);
 const parsed=JSON.parse(w.encode(snapshot));
 parsed.snapshot.data.machines[0].department_id='not-a-department';
 assert.throws(()=>w.decode(JSON.stringify(parsed)),/Fantasiedaten/);
 assert.throws(()=>w.save({data:makeDemoData(),approved:['unknown-id']}),/Ungültiger/);
});
test('browser storage blocked falls back only to current tab, never to cloud',()=>{
 const failed={getItem(){throw Error('unavailable')},setItem(){throw Error('unavailable')},removeItem(){throw Error('unavailable')}};
 const session=store(),w=openWorkspace(failed,session);
 const snapshot={data:makeDemoData(),approved:[]};
 assert.equal(w.save(snapshot),false);
 assert.equal(w.load().data.orders.length,50);
 w.reset();
 assert.equal(w.load(),null);
});
