// Disposable synthetic contract. Real production columns sourced from the catalog inventory, not a schema restore.
const {articleDatabase,...ids}=require('./article-db');
const fs=require('node:fs');
const path=require('node:path');
const machine='30000000-0000-0000-0000-000000000001';
const foreignMachine='30000000-0000-0000-0000-000000000002';
const station='40000000-0000-0000-0000-000000000001';
const migration=path.resolve(__dirname,'../../supabase/migrations/20261106053328_operations_phase2_review.sql');
async function operationsDatabase(){
 const api=await articleDatabase();
 await api.db.exec(`
  create table public.production_machines(id uuid primary key,name text not null,machine_code text not null,department_id uuid references public.departments(id),active boolean not null default true);
  alter table public.production_orders add column machine_id uuid references public.production_machines(id),add column department_id uuid references public.departments(id),add column article_number text,add column target_quantity integer not null default 0,add column pallet_count integer not null default 0,add column pieces_per_pallet integer not null default 0,add column use_chain_logic boolean not null default true,add column status text not null default 'running',add column started_at timestamptz not null default now(),add column completed_at timestamptz,add column created_at timestamptz not null default now(),add column updated_at timestamptz not null default now();
  create table public.production_order_stations(id uuid primary key default gen_random_uuid(),order_id uuid not null references public.production_orders(id),station_no integer not null,op_number text,scrap_total integer not null default 0);
  create table public.production_station_counts(id uuid primary key default gen_random_uuid(),order_id uuid not null references public.production_orders(id),station_id uuid references public.production_order_stations(id),employee_id uuid,good integer not null default 0);
  create table public.production_order_checklist(order_id uuid references public.production_orders(id),done boolean not null default false);
  insert into public.production_machines values('${machine}','Testmaschine','TEST-52','${ids.department}',true),('${foreignMachine}','Fremde Maschine','OTHER','${ids.foreignDepartment}',true);
  update public.production_orders set ba_number='LEGACY-812150',machine_id='${machine}',department_id='${ids.department}',target_quantity=48 where id='${ids.order}';
  insert into public.production_order_stations(id,order_id,station_no,op_number,scrap_total) values('${station}','${ids.order}',1,'100',3);
  insert into public.production_station_counts(order_id,station_id,employee_id,good) values('${ids.order}','${station}','${ids.worker}',17);
  insert into public.production_order_checklist values('${ids.order}',true);
 `);
 await api.db.exec(fs.readFileSync(migration,'utf8'));
 const signatures={operation_plan:['p_revision_id','p_operations','p_expected_version'],operations_article_command:['p_action','p_revision_id','p_department_id','p_article_number','p_data','p_expected_version'],order_from_article:['p_revision_id','p_ba_number','p_quantity','p_due_date','p_machine_id'],operation_order_read:['p_order_id'],operation_scan:['p_barcode']};
 let pending=Promise.resolve();
 function rpc(uid,name,params={}){
  const task=async()=>{
   if(!Object.hasOwn(signatures,name))return api.rpc(uid,name,params);
   const keys=signatures[name];
   const r=await api.asUser(uid,()=>api.db.query(`select public.${name}(${keys.map((_,i)=>'$'+(i+1)).join(',')}) result`,keys.map(k=>params[k]??null)));
   return r.rows[0].result;
  };
  const result=pending.then(task);pending=result.catch(()=>{});return result;
 }
 return{...api,rpc};
}
module.exports={operationsDatabase,...ids,machine,foreignMachine,station};
