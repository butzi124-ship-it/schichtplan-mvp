// Synthetic embedded contract only; native JWT/concurrency evidence comes from isolated CI.
const {baEntryDatabase,...ids}=require('./ba-entry-db');
const fs=require('node:fs');
const leaderB='00000000-0000-0000-0000-000000000005',deputy='00000000-0000-0000-0000-000000000006';
const machineA2='30000000-0000-0000-0000-000000000003';
async function routingDatabase({enabled=true}={}){
 const api=await baEntryDatabase();
 await api.db.query("update public.employees set role='employee' where auth_user_id=$1",[ids.reviewer]);
 await api.db.exec("alter table public.employees add constraint employees_role_check check(role in('admin','employee','tool_scanner'))");
 await api.db.exec(`alter table public.departments add column leader_employee_id uuid references public.employees(id);
 alter table public.production_order_stations add column actual_time_minutes integer;
 update public.production_order_stations set actual_time_minutes=42;
 insert into auth.users values('${leaderB}'),('${deputy}');
 insert into public.employees(auth_user_id,department_id,role) values('${leaderB}','${ids.foreignDepartment}','employee'),('${deputy}','${ids.foreignDepartment}','employee');
 update public.departments set leader_employee_id=(select id from public.employees where auth_user_id=case when departments.id='${ids.department}' then '${ids.reviewer}'::uuid else '${leaderB}'::uuid end);
 insert into public.production_machines values('${machineA2}','Alternative A','TEST-A2','${ids.department}',true);`);
 await api.db.exec(fs.readFileSync('supabase/review/issue-9-machine-readiness.sql','utf8'));
 await api.db.exec('update public.production_readiness_settings set enabled=true');
 await api.db.exec(fs.readFileSync('supabase/review/op-routing-phase2.sql','utf8'));
 await api.db.query("update public.employees set role='department_admin' where auth_user_id=any($1::uuid[])",[[ids.reviewer,leaderB]]);
 if(enabled)await api.db.exec('update public.production_routing_settings set enabled=true');
 await api.db.query("insert into public.production_routing_authorities values($1,$2,'admin',null,now()-interval '1 day',now()+interval '1 day',false,$1),($1,$3,'admin',null,now()-interval '1 day',now()+interval '1 day',false,$1),($4,$2,'leader',null,now()-interval '1 day',now()+interval '1 day',false,$1),($5,$3,'leader',null,now()-interval '1 day',now()+interval '1 day',false,$1),($6,$3,'deputy',$5,now()-interval '1 day',now()+interval '1 day',false,$1)",[ids.manager,ids.department,ids.foreignDepartment,ids.reviewer,leaderB,deputy]);
 await api.db.query("insert into public.production_routing_capabilities values($1,'TEST'),($2,'TEST'),($3,'TEST')",[ids.machine,machineA2,ids.foreignMachine]);
 await api.db.query("insert into public.production_routing_abort_reasons values('synthetic-machine-change','Synthetic configured machine change',false,true),('synthetic-other','Synthetic other',true,true)");
 for(const [uid,machine]of [[ids.worker,ids.machine],[ids.worker,machineA2],[ids.outsider,ids.foreignMachine]])await api.db.query("insert into public.production_readiness_grants values($1,$2,true,true,now()+interval '1 day',false,$3)",[uid,machine,ids.manager]);
 let pending=Promise.resolve();
 const signatures={op_routing_view:[],op_routing_command:['p_operation_id','p_action','p_request_id','p_expected_version','p_machine_id','p_reason','p_abort_code'],machine_readiness_view:[],machine_readiness_initialize:['p_order_id'],machine_readiness_command:['p_operation_id','p_action','p_request_id','p_expected_version','p_quantity','p_machine_id','p_planned_start','p_reason']};
 function rpc(uid,name,params={}){
  const task=async()=>{
   if(!Object.hasOwn(signatures,name))return api.rpc(uid,name,params);
   const keys=signatures[name];return api.asUser(uid,async()=> (await api.db.query(`select public.${name}(${keys.map((_,i)=>'$'+(i+1)).join(',')}) result`,keys.map(k=>params[k]??null))).rows[0].result);
  };const p=pending.then(task);pending=p.catch(()=>{});return p;
 }
 let revision;
 async function seed(ba='812200'){
  if(!revision){const {blank}=require('../../js/operations');let r=await rpc(ids.manager,'operations_article_command',{p_action:'create',p_department_id:ids.department,p_article_number:'32505',p_data:{designation:'Synthetic routing article'}});
   const p=await rpc(ids.manager,'operation_plan',{p_revision_id:r.id,p_operations:[{...blank(10),machine_id:ids.machine,machine_group:'TEST',activity:'Synthetic routing OP'}],p_expected_version:r.row_version});
   revision=await rpc(ids.reviewer,'operations_article_command',{p_action:'release',p_revision_id:r.id,p_expected_version:p.row_version});}
  return rpc(ids.worker,'ba_entry_command',{p_action:'create',p_ba_number:ba,p_article_number:'32505',p_quantity:48});
 }
 return{...api,rpc,seed};
}
module.exports={routingDatabase,...ids,leaderB,deputy,machineA2};
