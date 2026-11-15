const {quantityDatabase,...ids}=require('./quantity-db');const fs=require('node:fs');
async function lifecycleDatabase({enabled=true}={}){
 const api=await quantityDatabase();await api.db.exec(fs.readFileSync('supabase/review/lifecycle-phase5.sql','utf8'));
 if(enabled)await api.db.exec('update public.production_lifecycle_settings set enabled=true');
 let pending=Promise.resolve();
 const keys=['p_order_id','p_action','p_request_id','p_expected_version','p_operation_id','p_data'];
 function rpc(uid,name,p={}){if(!['ba_lifecycle_command','ba_lifecycle_view'].includes(name))return api.rpc(uid,name,p);const task=()=>api.asUser(uid,async()=> (await api.db.query(`select public.${name}(${name==='ba_lifecycle_view'?'':keys.map((_,i)=>'$'+(i+1)).join(',')}) result`,name==='ba_lifecycle_view'?[]:keys.map(k=>p[k]??null))).rows[0].result);const r=pending.then(task);pending=r.catch(()=>{});return r;}
 return{...api,rpc};
}
module.exports={lifecycleDatabase,...ids};
