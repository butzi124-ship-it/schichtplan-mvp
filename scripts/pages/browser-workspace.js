'use strict';
/**
 * FABOS GitHub Pages fictional workspace, with browser-only persistence.
 * IMPORTANT: no authentication, central sync, or real company data.
 */
(function attachBrowserWorkspace(root){
 const KEY='fabos-github-pages-fictional-workspace-v1';
 const LEGACY='fabos-fictitious-workspace-v02019';
 const KIND='FABOS_GITHUB_PAGES_FICTITIOUS_BACKUP';
 const FORMAT=1;
 const maxBytes=3*1024*1024;
 const unique=a=>Array.isArray(a)&&new Set(a.map(x=>x.id)).size===a.length;
 function valid(snapshot){
  if(!snapshot||typeof snapshot!=='object'||Array.isArray(snapshot))return false;
  const d=snapshot.data;
  if(!d||typeof d!=='object'||!Array.isArray(snapshot.approved)||!Number.isFinite(Date.parse(d.as_of)))return false;
  if(!unique(d.departments)||d.departments.length!==8||
     !unique(d.employees)||d.employees.length<50||d.employees.length>500||
     !unique(d.machines)||d.machines.length!==48||
     !unique(d.orders)||d.orders.length<50||d.orders.length>2000||
     !unique(d.operations)||d.operations.length<277||d.operations.length>10000||
     snapshot.approved.length>10000)return false;
  const depts=new Set(d.departments.map(x=>x.id)),
   machines=new Map(d.machines.map(x=>[x.id,x])),
   employees=new Map(d.employees.map(x=>[x.id,x])),
   orders=new Map(d.orders.map(x=>[x.id,x])),
   operations=new Set(d.operations.map(x=>x.id));
  if([...depts].some(id=>typeof id!=='string'||!id))return false;
  if(d.departments.some(dep=>d.machines.filter(m=>m.department_id===dep.id).length!==6))return false;
  if(d.machines.some(m=>!depts.has(m.department_id)||typeof m.name!=='string'||m.name.length>150))return false;
  if(d.employees.some(e=>!depts.has(e.department_id)||typeof e.display_name!=='string'||e.display_name.length>150))return false;
  if(d.orders.some(o=>!Array.isArray(o.operations)||!o.operations.every(id=>operations.has(id))||
     typeof o.ba_number!=='string'||o.ba_number.length>40||
     !Number.isSafeInteger(o.ba_quantity)||o.ba_quantity<1||o.ba_quantity>100000))return false;
  if(d.operations.some(op=>{
   const m=machines.get(op.machine_id),e=employees.get(op.assigned_employee_id),o=orders.get(op.order_id);
   return !m||!e||!o||!depts.has(op.department_id)||m.department_id!==op.department_id||
    e.department_id!==op.department_id||
    !['running','queued','completed'].includes(op.state)||
    !Number.isSafeInteger(op.good_qty)||op.good_qty<0||op.good_qty>o.ba_quantity||
    (op.predecessor_op_id!==null&&op.predecessor_op_id!==undefined&&!operations.has(op.predecessor_op_id));
  }))return false;
  if(snapshot.approved.some(id=>!operations.has(id)))return false;
  return true;
 }
 function parse(value){try{const result=JSON.parse(value);return valid(result)?result:null;}catch{return null;}}
 function load(){
  let raw=null;
  try{raw=root.localStorage?.getItem(KEY)}catch{}
  const preferred=parse(raw);
  if(preferred)return preferred;
  try{raw=root.sessionStorage?.getItem(LEGACY)}catch{raw=null;}
  return parse(raw);
 }
 function save(snapshot){
  if(!valid(snapshot))throw Error('Ungültiger fiktiver Testdatenstand');
  const data=JSON.stringify(snapshot);
  let persistent=false;
  try{root.localStorage?.setItem(KEY,data);persistent=true;}catch{}
  try{root.sessionStorage?.setItem(LEGACY,data);}catch{}
  return persistent;
 }
 function reset(){
  try{root.localStorage?.removeItem(KEY);}catch{}
  try{root.sessionStorage?.removeItem(LEGACY);}catch{}
 }
 function encode(snapshot){
  if(!valid(snapshot))throw Error('Ungültiger Testdatenstand');
  return JSON.stringify({kind:KIND,format:FORMAT,exported_at:new Date().toISOString(),snapshot},null,2);
 }
 function decode(contents){
  if(typeof contents!=='string'||contents.length>maxBytes)throw Error('Datei zu groß oder ungültig');
  let file;
  try{file=JSON.parse(contents);}catch{throw Error('Keine gültige JSON-Datei');}
  if(!file||file.kind!==KIND||file.format!==FORMAT||!valid(file.snapshot))
   throw Error('Keine gültige FABOS-Fantasiedaten-Sicherung');
  return file.snapshot;
 }
 function download(snapshot){
  const data=encode(snapshot);
  const url=URL.createObjectURL(new Blob([data],{type:'application/json;charset=utf-8'}));
  const a=document.createElement('a');
  a.href=url;a.download='fabos-fantasiedaten-'+new Date().toISOString().slice(0,10)+'.json';
  a.style.display='none';document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),0);
 }
 if(root.addEventListener)root.addEventListener('storage',event=>{
  if(event.key===KEY&&root.location?.pathname?.endsWith('/fertigung.html'))root.location.reload();
 });
 root.FABOS_BROWSER_WORKSPACE={KEY,LEGACY,KIND,valid,load,save,reset,encode,decode,download};
 if(typeof module!=='undefined'&&module.exports)module.exports=root.FABOS_BROWSER_WORKSPACE;
})(typeof window!=='undefined'?window:globalThis);
