// Synthetic test adapters only. Never accepts a hosted DB URL or production data.
const production=require('../../tests/helpers/lifecycle-db');
const learning=require('../../tests/helpers/learning-db');
async function createBackend(){
 const p=await production.lifecycleDatabase(),l=await learning.setup();await l.ready();await l.safe('worker');await l.question({prompt:'Synthetische Lernfrage: Welche Antworten sind A und C?'});await l.safe('leader');await l.call('leader','offer');
 await p.seed('800001');
 await l.db.query("insert into humbel_learning.people(employee_id,user_id,department_id,access_enabled,evidence) values($1,$1,$2,true,'Explicit independent synthetic preview admin enrollment')",[learning.ids.admin,learning.ids.d1]);
 const actors={admin:{role:'admin',uid:learning.ids.admin,puid:production.manager,who:'admin',name:'Test Administration'},leader:{role:'department_admin',uid:learning.ids.leader,puid:production.reviewer,who:'leader',name:'Test Abteilungsleitung'},employee:{role:'employee',uid:learning.ids.worker,puid:production.worker,who:'worker',name:'Test Mitarbeiter'}};
 // Each preview account has an independent synthetic learning principal.
 // Explicit mappings across separate fixtures are not native Auth acceptance.
 const allow=new Set(['article_master_capabilities','article_master_list','article_master_get','article_master_command','operations_article_command','operation_plan','order_from_article','operation_order_read','operation_scan','ba_entry_command','op_routing_view','op_routing_command','machine_readiness_view','machine_readiness_initialize','machine_readiness_command','setup_time_view','setup_time_command','op_quantity_view','op_quantity_command','ba_lifecycle_view','ba_lifecycle_command']);
 let chain=Promise.resolve();
 function rpc(account,name,body={}){const run=async()=>{const a=actors[account];if(!a)throw Error('No server identity');
  if(name==='learning_view')return l.view(a.who);
  if(name==='learning_command')return l.call(a.who,body.p_action,body.p_body,body.p_request);
  if(name==='learning_admin_view')return (await l.asUser(a.uid,d=>d.query('select public.learning_admin_view($1) result',[body.p_department]))).rows[0].result;
  if(!allow.has(name))throw Object.assign(Error('Unsupported test adapter'),{code:'42501'});
  return p.rpc(a.puid,name,body);
 };const r=chain.then(run);chain=r.catch(()=>{});return r;}
 return{actors,rpc,close:async()=>{await p.db.close();await l.db.close();}};
}
module.exports={createBackend};
