const {routingDatabase,...ids}=require('./routing-db');const fs=require('node:fs');
async function setupDatabase({enabled=true}={}){
 const api=await routingDatabase();await api.db.exec(fs.readFileSync('supabase/review/setup-phase3.sql','utf8'));
 if(enabled)await api.db.exec('update public.production_setup_settings set enabled=true');
 const keys={setup_time_view:[],setup_time_command:['p_action','p_request_id','p_operation_id','p_run_id','p_mode','p_segment_id','p_expected','p_change','p_reason','p_previous_user']};let pending=Promise.resolve();
 function rpc(uid,name,params={}){const task=async()=>{if(!Object.hasOwn(keys,name))return api.rpc(uid,name,params);return api.asUser(uid,async()=> (await api.db.query(`select public.${name}(${keys[name].map((_,i)=>'$'+(i+1)).join(',')}) result`,keys[name].map(k=>params[k]??null))).rows[0].result);};const p=pending.then(task);pending=p.catch(()=>{});return p;}
 return{...api,rpc};
}
module.exports={setupDatabase,...ids};
