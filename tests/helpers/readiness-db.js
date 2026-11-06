const {operationsDatabase,...ids}=require('./operations-db');
const fs=require('node:fs');const path=require('node:path');
const machineB='30000000-0000-0000-0000-000000000003',machineC='30000000-0000-0000-0000-000000000004';
async function readinessDatabase({enabled=true}={}){
 const api=await operationsDatabase();
 await api.db.query('insert into public.production_machines values($1,$2,$3,$4,true),($5,$6,$7,$4,true)',[machineB,'Testmaschine B','TEST-B',ids.department,machineC,'Testmaschine C','TEST-C']);
 await api.db.exec(fs.readFileSync(path.resolve(__dirname,'../../supabase/review/issue-9-machine-readiness.sql'),'utf8'));
 if(enabled)await api.db.exec('update public.production_readiness_settings set enabled=true');
 await api.db.query("insert into public.production_readiness_grants values($1,$2,true,true,now()+interval '1 hour',false,$3)",[ids.worker,ids.machine,ids.manager]);
 let pending=Promise.resolve();
 function rpc(uid,name,params={}){
  const task=async()=>{
   const signatures={machine_readiness_view:[],machine_readiness_initialize:['p_order_id'],machine_readiness_command:['p_operation_id','p_action','p_request_id','p_expected_version','p_quantity','p_machine_id','p_planned_start','p_reason']};
   if(!Object.hasOwn(signatures,name))return api.rpc(uid,name,params);
   const keys=signatures[name];const r=await api.asUser(uid,()=>api.db.query(`select public.${name}(${keys.map((_,i)=>'$'+(i+1)).join(',')}) result`,keys.map(k=>params[k]??null)));return r.rows[0].result;
  };const r=pending.then(task);pending=r.catch(()=>{});return r;
 }
 async function seed(ba='912150',quantity=48){
  const {blank}=require('../../js/operations');
  let r=await rpc(ids.manager,'operations_article_command',{p_action:'create',p_department_id:ids.department,p_article_number:'SYNTH-'+ba,p_data:{designation:'Synthetischer Folgeauftrag',customer:'Testkunde'}});
  const operations=[ids.machine,machineB,machineC].map((machine_id,i)=>({...blank((i+1)*10,i?i*10:null),machine_id,machine_group:'TEST',activity:'Synthetische OP '+(i+1)}));
  const p=await rpc(ids.manager,'operation_plan',{p_revision_id:r.id,p_operations:operations,p_expected_version:r.row_version});
  r=await rpc(ids.reviewer,'operations_article_command',{p_action:'release',p_revision_id:r.id,p_expected_version:p.row_version});
  return rpc(ids.manager,'order_from_article',{p_revision_id:r.id,p_ba_number:ba,p_quantity:quantity,p_due_date:'2099-12-31',p_machine_id:ids.machine});
 }
 async function forecast(op,quantity=48,extra={}){
  await api.db.query("insert into public.production_operation_arrivals(operation_id,machine_id,expected_on,expected_quantity,source_record,evidence,verified_by,verified_at,valid_until) values($1,$2,coalesce($3::date,(now() at time zone 'Europe/Berlin')::date),$4,'synthetic:existing-inbound-record','Synthetic verified source, not a computed ETA',$5,now(),now()+interval '1 day') on conflict(operation_id) do update set expected_on=excluded.expected_on,expected_quantity=excluded.expected_quantity,valid_until=excluded.valid_until",[op.operation_id||op.id,extra.machine||op.machine_id||op.data?.machine_id,extra.date||null,quantity,ids.manager]);
 }
 return{...api,rpc,seed,forecast};
}
module.exports={readinessDatabase,...ids,machineB,machineC};
