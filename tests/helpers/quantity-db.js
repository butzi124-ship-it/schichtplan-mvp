// Synthetic catalog-shaped counter fixtures; never a production restore.
const {setupDatabase,...ids}=require('./setup-db');
const fs=require('node:fs');
async function quantityDatabase({enabled=true}={}){
 const api=await setupDatabase();
 // Earlier embedded OP fixtures intentionally model only a subset of the catalog.
 // Extend only this new fixture, without weakening or replacing earlier tests.
 const inventory=JSON.parse(fs.readFileSync('docs/qa/schema-inventory.json'));
 for(const table of ['production_order_employees','production_qa_causes','production_station_events']){
  const cols=inventory.columns.filter(c=>c.table===table);
  await api.db.exec(`create table public.${table}(${cols.map(c=>`${c.name} ${c.type}${c.default?' default '+c.default:''}${c.nullable==='NO'?' not null':''}`).join(',')},primary key(id))`);
 }
 await api.db.exec(`alter table public.production_order_stations add column name text,add column lock_name boolean not null default false,add column clarify_total integer not null default 0,add column scrap_lifetime integer not null default 0,add column clarify_lifetime integer not null default 0,add column created_at timestamptz not null default now(),add column updated_at timestamptz not null default now();
 alter table public.production_station_counts add column order_employee_id uuid,add column good_qty integer not null default 0,add column created_at timestamptz not null default now(),add column updated_at timestamptz not null default now();
 update public.production_station_counts set good_qty=good;
 create unique index quantity_fixture_employee_station on public.production_station_counts(station_id,order_employee_id);`);
 await api.db.exec('alter table public.production_order_employees add unique(order_id,employee_id)');
 await api.db.exec(fs.readFileSync('supabase/review/quantity-phase4.sql','utf8'));
 if(enabled)await api.db.exec('update public.production_quantity_settings set enabled=true');
 const keys=['p_operation_id','p_action','p_request_id','p_station_id','p_order_employee_id','p_kind','p_delta','p_section_id','p_qa_cause_id','p_note','p_station_ids','p_new_count'];let pending=Promise.resolve();
 function rpc(uid,name,params={}){const task=async()=>{if(!['op_quantity_view','op_quantity_command'].includes(name))return api.rpc(uid,name,params);const k=name==='op_quantity_view'?[]:keys;return api.asUser(uid,async()=>(await api.db.query(`select public.${name}(${k.map((_,i)=>'$'+(i+1)).join(',')}) result`,k.map(k=>params[k]??null))).rows[0].result);};const p=pending.then(task);pending=p.catch(()=>{});return p;}
 return{...api,rpc};
}
module.exports={quantityDatabase,...ids};
