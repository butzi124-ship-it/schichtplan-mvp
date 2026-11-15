// Embedded contract only: synthetic rows, no URL/network and no live schema claim.
const {operationsDatabase,...ids}=require('./operations-db');
const fs=require('node:fs');
async function baEntryDatabase({enabled=true}={}) {
  const api=await operationsDatabase();
  await api.db.exec(`create table public.employees(id uuid primary key default gen_random_uuid(),auth_user_id uuid references auth.users(id),department_id uuid references public.departments(id),role text not null,is_active boolean not null default true);
    insert into public.employees(auth_user_id,department_id,role) values('${ids.manager}','${ids.department}','admin'),('${ids.reviewer}','${ids.department}','department_admin'),('${ids.worker}','${ids.department}','employee'),('${ids.outsider}','${ids.foreignDepartment}','employee');`);
  await api.db.exec(fs.readFileSync('supabase/review/ba-entry-phase1.sql','utf8'));
  if(enabled)await api.db.exec('update public.production_ba_entry_settings set enabled=true');
  let pending=Promise.resolve();
  function rpc(uid,name,params={}) {
    if(name!=='ba_entry_command')return api.rpc(uid,name,params);
    const task=async()=>api.asUser(uid,async()=> (await api.db.query('select public.ba_entry_command($1,$2,$3,$4,$5) result', ['p_action','p_ba_number','p_article_number','p_quantity','p_op_number'].map(k=>params[k]??null))).rows[0].result);
    const result=pending.then(task);pending=result.catch(()=>{});return result;
  }
  return {...api,rpc};
}
module.exports={baEntryDatabase,...ids};
