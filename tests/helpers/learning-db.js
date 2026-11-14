const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const {randomUUID}=require('node:crypto');
const {fixture,ids}=require('../security/permissions-8b1-fixture');
async function setup(){
 const f=await fixture();await f.db.exec(fs.readFileSync('supabase/review/learning-issue39.sql','utf8'));
 for(const who of ['leader','worker','other','deputy'])await f.db.query("insert into humbel_learning.people(employee_id,user_id,department_id,access_enabled,evidence) values($1,$1,$2,true,'Independent synthetic learning enrollment')",[ids[who],who==='other'?ids.d2:ids.d1]);
 await f.db.query("insert into humbel_learning.scopes values($1,$2,true,true,false,true,now()+interval '1 day',false,'Synthetic own-department grant')",[ids.leader,ids.d1]);
 const call=(who,action,body={},request=randomUUID())=>f.asUser(ids[who],d=>d.query('select public.learning_command($1,$2,$3) result',[action,body,request])).then(r=>r.rows[0].result);
 const view=who=>f.asUser(ids[who],d=>d.query('select public.learning_view() result')).then(r=>r.rows[0].result);
 const safe=who=>f.db.query("insert into humbel_learning.context_windows(employee_id,starts_at,ends_at,eligible,safe_to_offer,evidence) values($1,now()-interval '1 hour',now()+interval '8 hours',true,true,'Synthetic trusted eligibility only')",[ids[who]]);
 const question=async ({kind='learning',difficulty=1,concept='concept-a',prompt='Synthetic theoretical question?',curriculum='general',previous_id=null}={})=>{
  await call('leader','question_save',{department:ids.d1,domain:'milling',topic:'topic-a',concept,difficulty,prompt,options:[{id:'a',label:'A'},{id:'b',label:'B'},{id:'c',label:'C'}],correct:['a','c'],explanation:'A and C are correct. Synthetic explanation.',source:'Synthetic source, not a production safety instruction',curriculum,kind,released:true,previous_id});
  return (await f.db.query('select id from humbel_learning.questions order by created_at desc,id desc limit 1')).rows[0].id;
 };
 const ready=async()=>f.db.exec('update humbel_learning.settings set enabled=true,safety_enabled=true');
 return {...f,call,view,safe,question,ready};
}

module.exports={setup,ids};
