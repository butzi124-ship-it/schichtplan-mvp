const fs=require('node:fs');const {readinessDatabase,...ids}=require('./readiness-db');
async function forecastDatabase(){const api=await readinessDatabase();await api.db.query("select set_config('TimeZone',$1,false)",[require('../fixtures/forecast-clock.json').zone]);await api.db.exec(`
 create table public.employees(id uuid primary key,auth_user_id uuid,display_name text,is_active boolean,department_id uuid);
 create table public.shift_definitions(id uuid primary key,name text,shift_key text,start_time time,end_time time,requires_time boolean,active boolean);
 create table public.planner_assignments(id uuid primary key default gen_random_uuid(),shift_id text,shift_date date,assigned_user text);
 create table public.planner_absence_replacements(id uuid primary key default gen_random_uuid(),shift_id text,shift_date date,mode text,replacement_user text);
 create table public.planner_shift_cancellations(id uuid primary key default gen_random_uuid(),shift_id text,shift_date date,reason text);
 create table public.planner_availability(id uuid primary key default gen_random_uuid(),shift_id text,shift_date date,user_name text,status text);
 create table public.planner_absences(id uuid primary key default gen_random_uuid(),absence_date date,user_name text,absence_type text);
 create table public.planner_vacations(id uuid primary key default gen_random_uuid(),user_name text,from_date date,to_date date);
 create table public.planner_sick_leaves(id uuid primary key default gen_random_uuid(),user_name text,from_date date,to_date date);
 create table public.planner_swaps(id uuid primary key default gen_random_uuid(),user_a text,user_b text,start_date date,end_date date);
 create table public.planner_special_days(id uuid primary key default gen_random_uuid(),day_date date,type text,label text);
 alter table public.production_order_stations add actual_time_minutes integer,add clarify_total integer default 0;
 alter table public.production_station_counts add good_qty integer default 0;
 `);await api.db.exec(fs.readFileSync('supabase/review/issue-10-production-forecast.sql','utf8'));
 const signatures={production_forecast_input:['p_department_id','p_request_id'],production_forecast_view:[],production_forecast_report:['p_operation_id','p_request_id','p_data'],production_forecast_publish:['p_department_id','p_actor_id','p_run_id','p_result'],production_forecast_pending:[],production_forecast_worker_input:['p_department_id','p_actor_id']};
 async function forecastRpc(uid,name,p={},role='authenticated'){if(!signatures[name])return api.rpc(uid,name,p);const k=signatures[name];return(await api.asUser(uid,()=>api.db.query(`select public.${name}(${k.map((_,i)=>'$'+(i+1)).join(',')}) result`,k.map(x=>p[x]??null)),role)).rows[0].result;}
 return{...api,rpc:forecastRpc};}
module.exports={forecastDatabase,...ids};
// Synthetic roster/config only. Existing order snapshots are created through the release RPCs.
async function seedForecast(api){
 const {blank}=require('../../js/operations');
 await api.db.query("update public.production_orders set status='completed' where id=$1",[ids.order]);
 const today=(await api.db.query("select (now() at time zone 'Europe/Berlin')::date::text d")).rows[0].d;
 await api.db.query(`insert into public.production_forecast_settings(department_id,enabled,zone,start_date,end_date,valid_until,green_buffer_seconds,orange_buffer_seconds,max_age_seconds,feedback_max_age_seconds,dst_disambiguation,calculation_actor_id,roster_complete) values($1,true,'Europe/Berlin',$2::date,$2::date+3,now()+interval '1 day',7200,1800,300,86400,'reject',$3,true)`,[ids.department,today,ids.manager]);
 await api.db.query("insert into public.employees values($1,$1,'Synthetic Operator',true,$2)",[ids.worker,ids.department]);await api.db.exec("insert into public.shift_definitions values('50000000-0000-0000-0000-000000000001','Synthetic configured shift','TEST','00:00','23:59',false,true)");
 await api.db.query('insert into public.production_forecast_staff values($1,$2,$3,$4)',[ids.worker,'Synthetic Operator',['Q'],[ids.machine,ids.machineB,ids.machineC]]);
 for(const m of[ids.machine,ids.machineB,ids.machineC]){await api.db.query('insert into public.production_forecast_machines values($1,1,false)',[m]);await api.db.query("insert into public.production_forecast_shifts(machine_id,shift_definition_id,planner_suffix,weekdays,breaks,unmanned) values($1,'50000000-0000-0000-0000-000000000001','-mf-0',array[1,2,3,4,5,6,7],'[]',false)",[m]);}
 await api.db.query("insert into public.planner_assignments(shift_id,shift_date,assigned_user) select d::date::text||'-mf-0',d::date,'Synthetic Operator' from generate_series($1::date,$1::date+3,interval '1 day') d",[today]);
 let r=await api.rpc(ids.manager,'operations_article_command',{p_action:'create',p_department_id:ids.department,p_article_number:'SYNTH-FORECAST',p_data:{designation:'Synthetischer Prognoseauftrag',customer:'Testkunde'}});
 const ops=[ids.machine,ids.machineB,ids.machineC].map((machine_id,i)=>({...blank((i+1)*10,i?i*10:null),machine_id,machine_group:'TEST',activity:'Synthetic forecast OP',tr:{value:5,unit:'min'},te:{value:1,unit:'min'}}));
 const p=await api.rpc(ids.manager,'operation_plan',{p_revision_id:r.id,p_operations:ops,p_expected_version:r.row_version});r=await api.rpc(ids.reviewer,'operations_article_command',{p_action:'release',p_revision_id:r.id,p_expected_version:p.row_version});
 const out=await api.rpc(ids.manager,'order_from_article',{p_revision_id:r.id,p_ba_number:'812150',p_quantity:48,p_due_date:(await api.db.query('select ($1::date+1)::text d',[today])).rows[0].d,p_machine_id:ids.machine});
 for(const op of out.op_rows){await api.db.query("insert into public.production_forecast_rules(operation_id,priority,qualification,people,unmanned_allowed,transfer_seconds,quantity_source) values($1,$2,'Q',1,false,0,'reported')",[op.id,op.op_number]);await api.rpc(ids.manager,'production_forecast_report',{p_operation_id:op.id,p_request_id:require('crypto').randomUUID(),p_data:{setup_complete:false,setup_elapsed_s:0,processing_elapsed_s:0,current_piece_elapsed_s:0,good:0,scrap:0,clarify:0,replacement_qty:0,material_available_at:new Date(Date.now()-10000).toISOString(),evidence:'Isolated synthetic observation'}});}
 return out;
}
module.exports.seedForecast=seedForecast;
