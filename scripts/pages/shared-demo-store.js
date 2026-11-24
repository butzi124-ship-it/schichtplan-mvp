'use strict';
// Isolated fictional FABOS workspace. No production Supabase/DB/Auth dependency.
const fs=require('node:fs');
const path=require('node:path');
const crypto=require('node:crypto');
const {makeDemoData,validate}=require('../preview/fabos-demo-data');
const FORMAT=1;
function deny(status,message){const e=new Error(message);e.status=status;throw e;}
function required(value,max,label){
 if(typeof value!=='string'||!value.trim()||value.trim().length>max)deny(422,label+' ungültig');
 return value.trim();
}
function integer(value,min,max,label){
 if(!Number.isSafeInteger(value)||value<min||value>max)deny(422,label+' ungültig');
 return value;
}
function createSharedDemoStore(filePath){
 if(typeof filePath!=='string'||!path.isAbsolute(filePath))throw Error('Absolute isolated demo state file required');
 const parent=path.dirname(filePath);
 fs.mkdirSync(parent,{recursive:true,mode:0o700});
 const fresh=()=>({format:FORMAT,revision:0,data:makeDemoData(),approved:[],audit:[]});
 function write(state){
  const tmp=filePath+'.'+process.pid+'.'+crypto.randomBytes(6).toString('hex')+'.tmp';
  try{fs.writeFileSync(tmp,JSON.stringify(state),{flag:'wx',mode:0o600});fs.renameSync(tmp,filePath);}
  finally{if(fs.existsSync(tmp))fs.unlinkSync(tmp);}
 }
 let state;
 if(fs.existsSync(filePath)){
  state=JSON.parse(fs.readFileSync(filePath,'utf8'));
  if(state.format!==FORMAT||!Number.isSafeInteger(state.revision)||!Array.isArray(state.data?.departments)||
      state.data.departments.length!==8||state.data.machines.length!==48||
      state.data.employees.length<50||state.data.orders.length<50||
      !Array.isArray(state.data.operations)||!Array.isArray(state.approved)||!Array.isArray(state.audit))throw Error('Persisted demo schema incompatible: migration/review required');
 }else{state=fresh();validate(state.data);write(state);}
 const find=(a,id)=>a.find(v=>v.id===id);
 function actor(account){
  const d=state.data;const fr=d.departments.find(x=>x.code==='FR');
  if(!fr)throw Error('Fräserei missing');
  if(account==='admin')return {role:'master',department_id:null,employee_id:null,display_name:'FABOS Master'};
  if(account==='leader')return {role:'lead',department_id:fr.id,employee_id:fr.leader_employee_id,display_name:find(d.employees,fr.leader_employee_id)?.display_name||'Abteilungsleitung'};
  if(account==='employee'){
   const person=d.employees.find(x=>x.first_name==='Mila'&&x.last_name==='Fischer'&&x.department_id===fr.id);
   if(!person)throw Error('Isolated demo worker account missing');
   return {role:'employee',department_id:fr.id,employee_id:person.id,display_name:person.display_name};
  }
  deny(401,'Anmeldung erforderlich');
 }
 function snapshot(account){
  const a=actor(account),d=state.data,all=a.role==='master';
  const departments=d.departments.filter(x=>all||x.id===a.department_id);
  const employees=d.employees.filter(x=>all||x.department_id===a.department_id);
  const machines=d.machines.filter(x=>all||x.department_id===a.department_id);
  const ops=d.operations.filter(x=>all||x.department_id===a.department_id);
  const opIds=new Set(ops.map(x=>x.id)),orderIds=new Set(ops.map(x=>x.order_id));
  const orders=d.orders.filter(x=>orderIds.has(x.id)).map(x=>({...x,operations:x.operations.filter(id=>opIds.has(id))}));
  const queue=d.machine_queue.filter(x=>all||x.department_id===a.department_id).map(x=>({...x,operations:x.operations.filter(id=>opIds.has(id))}));
  return {revision:state.revision,actor:a,approved:state.approved.filter(id=>opIds.has(id)),
    data:{...d,departments,employees,machines,orders,operations:ops,machine_queue:queue}};
 }
 function departmentPermission(a,id){
  if(a.role!=='master'&&a.department_id!==id)deny(403,'Keine Rechte für diese Abteilung');
 }
 function canOperate(a,op){
  departmentPermission(a,op.department_id);
  if(a.role==='employee'&&a.employee_id!==op.assigned_employee_id)deny(403,'Arbeitsgang nicht zugeordnet');
 }
 function change(account,action,payload={}){
  if(!payload||typeof payload!=='object'||Array.isArray(payload))deny(400,'Ungültige Aktion');
  const a=actor(account),d=state.data;
  const op=id=>{const found=find(d.operations,id);if(!found)deny(404,'Arbeitsgang nicht gefunden');return found;};
  const person=id=>{const found=find(d.employees,id);if(!found)deny(404,'Mitarbeiter nicht gefunden');return found;};
  const dep=id=>{const found=find(d.departments,id);if(!found)deny(404,'Abteilung nicht gefunden');return found;};
  const machine=id=>{const found=find(d.machines,id);if(!found)deny(404,'Maschine nicht gefunden');return found;};
  const order=id=>{const found=find(d.orders,id);if(!found)deny(404,'Auftrag nicht gefunden');return found;};
  let audit={action,actor:a.display_name,at:new Date().toISOString()};
  if(action==='approve'){
   if(a.role==='employee')deny(403,'Freigabe nur durch Leitung');
   const v=op(required(payload.op_id,128,'Arbeitsgang'));
   departmentPermission(a,v.department_id);
   if(v.state==='completed')deny(409,'Arbeitsgang bereits abgeschlossen');
   if(!state.approved.includes(v.id))state.approved.push(v.id);
   audit.op_id=v.id;
  }else if(action==='create_employee'){
   if(a.role!=='master')deny(403,'Nur Master darf Personal anlegen');
   if(d.employees.length>=300)deny(422,'Maximale Testmitarbeiterzahl erreicht');
   const first=required(payload.first,50,'Vorname'),last=required(payload.last,50,'Nachname');
   const department=dep(required(payload.department_id,128,'Abteilung'));
   const id=crypto.randomUUID(),n=d.employees.length+1;
   const e={id,first_name:first,last_name:last,display_name:first+' '+last,personnel_no:'TEST-'+String(n).padStart(3,'0'),
     department_id:department.id,department:department.name,employee_type:'core',role:'employee',is_active:true,auth_user_id:null};
   d.employees.push(e);audit.employee_id=id;
  }else if(action==='move_employee'){
   if(a.role!=='master')deny(403,'Nur Master darf Personal umteilen');
   const e=person(required(payload.employee_id,128,'Mitarbeiter')),department=dep(required(payload.department_id,128,'Abteilung'));
   if(e.role==='department_admin')deny(409,'Abteilungsleitung kann nicht verschoben werden');
   if(d.operations.some(x=>x.assigned_employee_id===e.id&&x.state!=='completed'&&x.department_id!==department.id))deny(409,'Offene Arbeitsgänge zuerst anderen Mitarbeitern zuweisen');
   e.department_id=department.id;e.department=department.name;audit.employee_id=e.id;
  }else if(action==='assign_op'){
   if(a.role==='employee')deny(403,'Mitarbeiter dürfen nicht umteilen');
   const v=op(required(payload.op_id,128,'Arbeitsgang')),e=person(required(payload.employee_id,128,'Mitarbeiter'));
   departmentPermission(a,v.department_id);
   if(v.state==='completed'||e.department_id!==v.department_id||e.is_active===false)deny(409,'Ungültige OP-Zuordnung');
   v.assigned_employee_id=e.id;audit.op_id=v.id;audit.employee_id=e.id;
  }else if(action==='create_order'){
   if(a.role==='employee')deny(403,'Mitarbeiter dürfen keine BA anlegen');
   if(d.orders.length>=1000)deny(422,'Testauftragslimit erreicht');
   const department=dep(required(payload.department_id,128,'Abteilung'));
   departmentPermission(a,department.id);
   const m=machine(required(payload.machine_id,128,'Maschine')),e=person(required(payload.employee_id,128,'Mitarbeiter'));
   if(m.department_id!==department.id||e.department_id!==department.id||e.is_active===false)deny(409,'Maschine/Mitarbeiter außerhalb der Abteilung');
   const article=required(payload.article,90,'Artikel'),articleNo=required(payload.article_no,50,'Artikelnummer');
   const material=required(payload.material,50,'Material'),instruction=required(payload.instruction,250,'Arbeitsanweisung');
   const qty=integer(payload.qty,1,100000,'Sollmenge');
   const n=900001+d.orders.filter(x=>/^BA-9\d+$/.test(x.ba_number)).length;
   const id=crypto.randomUUID(),opId=crypto.randomUUID(),baNumber='BA-'+n;
   const latest=Math.max(Date.now(),...d.operations.filter(x=>x.machine_id===m.id).map(x=>Date.parse(x.planned_end)||0));
   const start=new Date(latest).toISOString(),duration=45+qty*3,end=new Date(latest+duration*60000).toISOString();
   const o={id,ba_number:baNumber,article_number:articleNo,article_name:article,material,ba_quantity:qty,
     release_at:start,operations:[opId],status:'queued',next_operation:opId,promised_date:end};
   const v={id:opId,order_id:id,op_number:'010',sequence:1,department_id:department.id,department:department.name,
     machine_id:m.id,machine:m.name,machine_code:m.machine_code,description:instruction,setup_minutes:45,
     cycle_seconds:180,planned_minutes:duration,planned_start:start,planned_end:end,state:'queued',
     good_qty:0,scrap_qty:0,assigned_employee_id:e.id,predecessor_op_id:null};
   d.orders.push(o);d.operations.push(v);
   const queue=d.machine_queue.find(x=>x.machine_id===m.id);
   if(queue)queue.operations.push(v.id);
   audit.order_id=id;audit.op_id=opId;
  }else if(['start_op','book_qty','finish_op'].includes(action)){
   const v=op(required(payload.op_id,128,'Arbeitsgang')),o=order(v.order_id);
   canOperate(a,v);audit.op_id=v.id;
   if(action==='start_op'){
    if(v.state!=='queued')deny(409,'Nur geplante Arbeitsgänge können gestartet werden');
    const predecessor=v.predecessor_op_id&&op(v.predecessor_op_id);
    if(predecessor&&predecessor.state!=='completed')deny(409,'Vorgänger-OP noch offen');
    if(d.operations.some(x=>x.machine_id===v.machine_id&&x.id!==v.id&&x.state==='running'))deny(409,'Maschine bereits belegt');
    v.state='running';o.status='running';
   }else if(action==='book_qty'){
    if(v.state!=='running')deny(409,'Arbeitsgang läuft nicht');
    const qty=integer(payload.qty,1,100000,'Gutmenge');
    if(v.good_qty+qty>o.ba_quantity)deny(409,'Menge größer als Restmenge');
    v.good_qty+=qty;audit.qty=qty;
   }else{
    if(v.state!=='running'||v.good_qty!==o.ba_quantity)deny(409,'Auftrag noch nicht fertiggezählt');
    v.state='completed';
    o.status=d.operations.filter(x=>x.order_id===o.id).every(x=>x.state==='completed')?'completed':'queued';
    o.next_operation=d.operations.find(x=>x.order_id===o.id&&x.state!=='completed')?.id||null;
   }
  }else if(action==='reset'){
   if(a.role!=='master')deny(403,'Nur Master kann Testdaten zurücksetzen');
   const next=fresh();next.revision=state.revision+1;next.audit=[audit];
   write(next);state=next;return snapshot(account);
  }else deny(400,'Aktion nicht unterstützt');
  state.revision++;
  state.audit.push(audit);
  if(state.audit.length>500)state.audit=state.audit.slice(-500);
  write(state);return snapshot(account);
 }
 return {snapshot,change,actor,filePath,revision:()=>state.revision};
}
module.exports={createSharedDemoStore};
