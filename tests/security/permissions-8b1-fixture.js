// Synthetic identity/session shims only; NOT cryptographic JWT or API acceptance.
const {randomUUID}=require('node:crypto');
const {originalFixture,enroll,ids}=require('./step7-fixture');
const baseline=require('../db-integration/security-step7-review');
const review=require('../db-integration/permissions-8b1-review');
const {PGlite}=require('@electric-sql/pglite');
// Cache ONLY a freshly built synthetic fixture image. Each test gets its own DB.
// This is not a production restore; it avoids reinstalling the unchanged Step7
// package for every new contract case while preserving complete test isolation.
let prepared;
async function preparedImage(){
 if(!prepared)prepared=(async()=>{const base=await originalFixture();try{
  await baseline.apply(base.db);await enroll(base);await review.apply(base.db);
  return await base.db.dumpDataDir();
 }finally{await base.db.close();}})();
 return prepared;
}
async function fixture(){
 const db=new PGlite({loadDataDir:await preparedImage()});await db.waitReady;
 await db.query("select set_config('humbel.isolated_sentinel','synthetic-8b1-cloned-fixture-sentinel-over-32',false)");
 async function asUser(who,fn,{anonymous=false,claims={}}={}){
  await db.exec('begin');try{await db.exec('set local role '+(anonymous?'anon':'authenticated'));
   await db.query("select set_config('request.jwt.claims',$1,true)",[JSON.stringify({sub:who,role:anonymous?'anon':'authenticated',...claims})]);
   const value=await fn(db);await db.exec('commit');return value;
  }catch(e){await db.exec('rollback');throw e;}
 }
 const f={db,asUser};
 const sessions={};
 async function identity(who){
  const session=randomUUID();sessions[who]=session;
  await f.db.query("insert into humbel_permissions.identities(user_id,employee_id,enabled,primary_department_id,valid_from,evidence) values($1,$1,true,$2,now()-interval '1 day','Independent SYNTHETIC identity approval')",[ids[who],who==='other'?ids.d2:ids.d1]);
  await f.db.query("insert into humbel_permissions.sessions values($1,$2,1,1,now()-interval '1 day',now()+interval '1 day',false,'Synthetic session allowlist, not GoTrue')",[session,ids[who]]);
 }
 async function role(who,code='employee',dept=null){
  const id=randomUUID();await f.db.query("insert into humbel_permissions.organization_roles(id,user_id,role_code,department_id,enabled,valid_from,evidence) values($1,$2,$3,$4,true,now()-interval '1 day','Synthetic role, never auto-mapped')",[id,ids[who],code,dept]);return id;
 }
 async function membership(who,dept){await f.db.query("insert into humbel_permissions.memberships(employee_id,department_id,enabled,valid_from,evidence) values($1,$2,true,now()-interval '1 day','Explicit synthetic membership')",[ids[who],dept]);}
 async function grant(who,roleId,dept=ids.d1,{scope='membership',action='read',until=null}={}){
  await f.db.query("insert into humbel_permissions.capabilities values('synthetic',$1,$2,true,'SYNTHETIC capability only') on conflict do nothing",[action,scope]);
  const id=randomUUID();await f.db.query("insert into humbel_permissions.module_grants(id,user_id,role_assignment_id,module,action,department_id,enabled,primary_epoch,valid_from,valid_until,evidence) values($1,$2,$3,'synthetic',$4,$5,true,1,now()-interval '1 day',$6,'SYNTHETIC explicit grant')",[id,ids[who],roleId,action,dept,until]);return id;
 }
 // TEST-ONLY exposure, never part of SQL candidate/application. Simulates a proposed
 // API surface to attack the real read evaluators as authenticated/anon roles.
 async function expose(){await f.db.exec('grant usage on schema humbel_permissions to authenticated,anon;grant execute on function humbel_permissions.actor_context(),humbel_permissions.can_perform(text,text,uuid) to authenticated,anon');}
 async function enabled(){await f.db.exec('update humbel_permissions.runtime set enabled=true');}
 const call=(who,dept=ids.d1,claims={})=>f.asUser(ids[who]||null,d=>d.query("select humbel_permissions.can_perform('synthetic','read',$1) allowed",[dept]),{claims:{session_id:sessions[who],...claims}}).then(r=>r.rows[0].allowed);
 const context=(who,claims={})=>f.asUser(ids[who]||null,d=>d.query('select * from humbel_permissions.actor_context()'),{claims:{session_id:sessions[who],...claims}}).then(r=>r.rows);
 return {...f,identity,role,membership,grant,expose,enabled,call,context,sessions};
}
module.exports={fixture,ids};
