// Synthetic SQL logins / session attestations. Not actual GoTrue or JWT evidence.
const {PGlite}=require('@electric-sql/pglite');
const {randomUUID}=require('node:crypto');
const prior=require('./permissions-8b1-fixture'),review=require('../db-integration/permissions-8b2-review');
const {ids}=prior;let image;
async function prepared(){
 if(!image)image=(async()=>{const f=await prior.fixture();try{await review.apply(f.db);return await f.db.dumpDataDir();}finally{await f.db.close();}})();
 return image;
}
async function fixture(){
 const db=new PGlite({loadDataDir:await prepared()});await db.waitReady;
 await db.query("select set_config('humbel.isolated_sentinel','8b2-synthetic-clone-approved-sentinel-over32',false)");
 async function identity(who,active=false){await db.query("insert into humbel_permissions.identities(user_id,employee_id,enabled,primary_department_id,valid_from,evidence) values($1,$1,$2,$3,now()-interval '1 day','SYNTHETIC prepared identity')",[ids[who],active,who==='other'?ids.d2:ids.d1]);}
 async function login(role,fn,{claims={}}={}){
  if(!['op_a','op_b','outsider'].includes(role))throw Error('Synthetic login allowlist');
  await db.exec('set session authorization '+role);await db.exec('begin');
  try{await db.query("select set_config('request.jwt.claims',$1,true)",[JSON.stringify(claims)]);const r=await fn(db);await db.exec('commit');return r;}
  catch(e){await db.exec('rollback');throw e;}finally{await db.exec('set session authorization postgres;reset role');}
 }
 async function provision(){
  await identity('admin',true);await identity('worker');await identity('other');
  await db.exec('create role op_a login nosuperuser nocreaterole nocreatedb nobypassrls;create role op_b login nosuperuser nocreaterole nocreatedb nobypassrls;create role outsider login nosuperuser nocreaterole nocreatedb nobypassrls');
  await db.query("insert into humbel_identity_admin.operators select oid,rolname,$1,1,1,true,false,now()-interval '1 day',now()+interval '1 day',$2 from pg_roles where rolname='op_a'",[ids.admin,randomUUID()]);
  const oid=(await db.query("select oid from pg_roles where rolname='op_a'")).rows[0].oid;
  for(const action of ['identity_activate','identity_deactivate','epoch_bump','session_allow','session_revoke','context_check'])await db.query("insert into humbel_identity_admin.scopes values($1,$2,$3,true,false,now()-interval '1 day',now()+interval '1 day',$4)",[oid,ids.worker,action,randomUUID()]);
  await db.exec('update humbel_identity_admin.runtime set enabled=true');
  // TEST-ONLY entry grants. No table grants, no helper grant, no production role.
  await db.exec('grant usage on schema humbel_identity_admin to op_a,op_b,outsider;grant execute on function humbel_identity_admin.command(text,uuid,uuid,bigint,uuid,timestamptz),humbel_identity_admin.server_context(uuid,uuid) to op_a,op_b,outsider');
 }
 async function command(action,revision,{target=ids.worker,session=null,request=randomUUID(),until=null,as='op_a',claims={}}={}){
  return login(as,d=>d.query('select humbel_identity_admin.command($1,$2,$3,$4,$5,$6) result',[action,target,session,revision,request,until]),{claims}).then(r=>r.rows[0].result);
 }
 const context=(sid,{as='op_a',target=ids.worker}={})=>login(as,d=>d.query('select humbel_identity_admin.server_context($1,$2) result',[target,sid])).then(r=>r.rows[0].result);
 async function attest(sid=randomUUID(),who='worker'){
  const until=(await db.query("select now()+interval '1 hour' t")).rows[0].t;
  await db.query("insert into humbel_identity_admin.session_attestations select $1,$2,i.binding_version,i.authorization_epoch,now()-interval '1 minute',$3,$4 from humbel_permissions.identities i where i.user_id=$2",[sid,ids[who],until,randomUUID()]);return {sid,until};
 }
 return{db,identity,login,provision,command,context,attest};
}
module.exports={fixture,ids};
