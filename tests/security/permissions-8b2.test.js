const {test}=require('node:test'),assert=require('node:assert/strict'),{randomUUID}=require('node:crypto');
const {fixture,ids}=require('./permissions-8b2-fixture'),review=require('../db-integration/permissions-8b2-review');
const deny=p=>assert.rejects(p,e=>e.code==='42501');
async function use(fn){const f=await fixture();try{await fn(f);}finally{await f.db.close();}}

test('8B2 DEFAULT: verified additive package remains off, empty, private, repeatable and fails on drift', {timeout:300000},async()=>use(async f=>{
 await review.postconditions(f.db);assert.equal((await review.apply(f.db)).status,'ISOLATED_8B2_VERIFIED_NOOP');
 await f.db.exec('alter table humbel_permissions.installation rename to synthetic_missing_8b1');await assert.rejects(review.apply(f.db),/8B1 prerequisite not installed/);await f.db.exec('alter table humbel_permissions.synthetic_missing_8b1 rename to installation');
 for(const t of ['operators','scopes','session_attestations','audit','requests'])assert.equal((await f.db.query('select count(*)::int n from humbel_identity_admin.'+t)).rows[0].n,0);
 await deny(f.db.query('select humbel_identity_admin.command(\'identity_activate\',$1,null,0,$2,null)',[ids.worker,randomUUID()]));
 await f.db.exec('begin;set local role authenticated');await assert.rejects(f.db.query('select * from humbel_identity_admin.operators'),e=>e.code==='42501');await f.db.exec('rollback');
 await f.db.exec('alter table humbel_identity_admin.scopes add column unexpected text');await assert.rejects(review.apply(f.db),/8B2 drift/);
}));

test('8B2 AUTHORITY: JWT/target spoofing, missing scopes, inactive operators, role switches and self approval cannot authorize', {timeout:300000},async()=>use(async f=>{
 await f.provision();
 for(const who of ['outsider','op_b'])await deny(f.command('identity_activate',0,{as:who,claims:{sub:ids.admin,role:'admin',session_id:randomUUID(),user_metadata:{role:'executive'}}}));
 await deny(f.command('identity_activate',0,{target:ids.other}));
 await f.db.exec("update humbel_identity_admin.scopes set valid_until=now()-interval '1 second'");await deny(f.command('identity_activate',0));await f.db.exec("update humbel_identity_admin.scopes set valid_until=now()+interval '1 day'");
 const oid=(await f.db.query("select oid from pg_roles where rolname='op_a'")).rows[0].oid;
 await f.db.query("insert into humbel_identity_admin.scopes values($1,$2,'identity_activate',true,false,now()-interval '1 day',now()+interval '1 day',$3)",[oid,ids.admin,randomUUID()]);
 await deny(f.command('identity_activate',0,{target:ids.admin}));
 await f.db.exec('update humbel_identity_admin.operators set revoked=true');await deny(f.command('identity_activate',0));await f.db.exec('update humbel_identity_admin.operators set revoked=false');
 await f.db.exec('update humbel_identity_admin.operators set authorization_epoch=99');await deny(f.command('identity_activate',0));await f.db.exec('update humbel_identity_admin.operators set authorization_epoch=1');
 await f.db.query('update humbel_permissions.identities set enabled=false where user_id=$1',[ids.admin]);await deny(f.command('identity_activate',0));await f.db.query('update humbel_permissions.identities set enabled=true where user_id=$1',[ids.admin]);
 await deny(f.login('op_a',d=>d.exec('insert into humbel_identity_admin.session_attestations values(gen_random_uuid(),\''+ids.worker+'\',1,1,now(),now()+interval \'1 day\',gen_random_uuid())')));
 await f.db.exec('grant authenticated to op_a;grant usage on schema humbel_identity_admin to authenticated;grant execute on function humbel_identity_admin.command(text,uuid,uuid,bigint,uuid,timestamptz) to authenticated');
 await deny(f.login('op_a',async d=>{await d.exec('set local role authenticated');return d.query("select humbel_identity_admin.command('identity_activate',$1,null,0,$2,null)",[ids.worker,randomUUID()]);}));
 assert.equal((await f.db.query('select count(*)::int n from humbel_identity_admin.audit')).rows[0].n,0);
}));

test('8B2 LIFECYCLE: activation/revocation is versioned, audited and never reenables old roles or session IDs', {timeout:300000},async()=>use(async f=>{
 await f.provision();const r=randomUUID(),g=randomUUID();
 await f.db.query("insert into humbel_permissions.organization_roles(id,user_id,role_code,enabled,valid_from,evidence) values($1,$2,'employee',true,now()-interval '1 day','Synthetic prior role')",[r,ids.worker]);
 await f.db.exec("insert into humbel_permissions.capabilities values('synthetic','read','explicit_department',true,'Synthetic only')");
 await f.db.query("insert into humbel_permissions.module_grants values($1,$2,$3,'synthetic','read',$4,true,false,1,now()-interval '1 day',null,'Synthetic prior grant')",[g,ids.worker,r,ids.d1]);
 const activation=await f.command('identity_activate',0,{claims:{sub:ids.other,role:'executive'}});
 assert.equal((await f.db.query('select actor_user_id from humbel_identity_admin.audit')).rows[0].actor_user_id,ids.admin);assert.equal(activation.revision,1);assert.equal(activation.authorization_epoch,2);
 assert.equal((await f.db.query('select enabled from humbel_permissions.organization_roles where id=$1',[r])).rows[0].enabled,false);
 const proof=await f.attest();await f.command('session_allow',1,{session:proof.sid,until:proof.until});assert.equal((await f.context(proof.sid)).authorization_epoch,2);
 const unusedProof=await f.attest();await f.command('identity_deactivate',2);await deny(f.context(proof.sid));
 await f.command('identity_activate',3);await deny(f.command('session_allow',4,{session:proof.sid,until:proof.until}));await deny(f.command('session_allow',4,{session:unusedProof.sid,until:unusedProof.until}));
 assert.equal((await f.db.query('select revoked,enabled from humbel_permissions.module_grants where id=$1',[g])).rows[0].revoked,true);
 assert.equal((await f.db.query('select count(*)::int n from humbel_identity_admin.audit')).rows[0].n,4);
 assert.equal((await f.db.query('select count(*)::int n from humbel_identity_admin.session_tombstones where session_id=$1',[proof.sid])).rows[0].n,1);
}));

test('8B2 SESSIONS: independent evidence, expiry, binding/epoch mismatch and irreversible revocation defeat forged context and retries', {timeout:300000},async()=>use(async f=>{
 await f.provision();await f.command('identity_activate',0);
 await deny(f.command('session_allow',1,{session:randomUUID(),until:new Date(Date.now()+60000).toISOString()}));
 const foreign=await f.attest(randomUUID(),'other');await deny(f.command('session_allow',1,{session:foreign.sid,until:foreign.until}));
 const proof=await f.attest();const request=randomUUID();await f.command('session_allow',1,{session:proof.sid,until:proof.until,request});
 const replay=await f.command('session_allow',1,{session:proof.sid,until:proof.until,request});assert.equal(replay.replayed,true);
 await f.db.exec('update humbel_identity_admin.scopes set revoked=true');await deny(f.command('session_allow',1,{session:proof.sid,until:proof.until,request}));await f.db.exec('update humbel_identity_admin.scopes set revoked=false');
 const valid=(await f.context(proof.sid));assert.equal(valid.user_id,ids.worker);
 await f.db.query('update humbel_permissions.sessions set binding_version=99 where session_id=$1',[proof.sid]);await deny(f.context(proof.sid));await f.db.query('update humbel_permissions.sessions set binding_version=1 where session_id=$1',[proof.sid]);
 await f.command('session_revoke',2,{session:proof.sid});await deny(f.context(proof.sid));
 await deny(f.command('session_allow',1,{session:proof.sid,until:proof.until,request}));await deny(f.command('session_allow',3,{session:proof.sid,until:proof.until}));
 // Even corrupted legacy allowlist flags cannot bypass the permanent tombstone.
 await f.db.query('update humbel_permissions.sessions set revoked=false where session_id=$1',[proof.sid]);await deny(f.context(proof.sid));
 const expired=randomUUID();await f.db.query("insert into humbel_identity_admin.session_attestations select $1,$2,i.binding_version,i.authorization_epoch,now()-interval '2 hours',now()-interval '1 hour',$3 from humbel_permissions.identities i where i.user_id=$2",[expired,ids.worker,randomUUID()]);await deny(f.command('session_allow',3,{session:expired,until:new Date(Date.now()+60000).toISOString()}));
 const p2=await f.attest();await f.command('session_allow',3,{session:p2.sid,until:p2.until});
 await f.db.query("update humbel_permissions.sessions set valid_from=now()-interval '1 hour',valid_until=now()-interval '1 second' where session_id=$1",[p2.sid]);await deny(f.context(p2.sid));await f.db.query("update humbel_permissions.sessions set valid_from=now()-interval '1 minute',valid_until=$2 where session_id=$1",[p2.sid,p2.until]);assert.equal((await f.context(p2.sid)).user_id,ids.worker);
 await f.command('epoch_bump',4);await deny(f.context(p2.sid));
 assert.equal((await f.db.query('select count(*)::int n from humbel_identity_admin.audit')).rows[0].n,5);
}));

test('8B2 ATOMICITY: conflicting versions/request IDs and audit failure roll back all changes; append-only history resists maintenance writes', {timeout:300000},async()=>use(async f=>{
 await f.provision();const req=randomUUID();const result=await f.command('identity_activate',0,{request:req});
 await assert.rejects(f.command('epoch_bump',0),e=>e.code==='40001');
 await assert.rejects(f.command('identity_deactivate',1,{request:req}),e=>e.code==='22023');
 assert.equal((await f.command('identity_activate',0,{request:req})).revision,result.revision);
 const before=(await f.db.query('select enabled,blocked,authorization_epoch from humbel_permissions.identities where user_id=$1',[ids.worker])).rows;
 await f.db.exec("create function humbel_identity_admin.synthetic_failure() returns trigger language plpgsql as $$begin raise exception 'Synthetic audit unavailable';end$$;create trigger synthetic_fail before insert on humbel_identity_admin.audit for each row execute function humbel_identity_admin.synthetic_failure()");
 await assert.rejects(f.command('identity_deactivate',1),/Synthetic audit unavailable/);
 assert.deepEqual((await f.db.query('select enabled,blocked,authorization_epoch from humbel_permissions.identities where user_id=$1',[ids.worker])).rows,before);
 assert.equal((await f.db.query('select count(*)::int n from humbel_identity_admin.requests')).rows[0].n,1);
 await f.db.exec('drop trigger synthetic_fail on humbel_identity_admin.audit');
 for(const q of ['update humbel_identity_admin.audit set outcome=\'applied\'','delete from humbel_identity_admin.audit','truncate humbel_identity_admin.audit','delete from humbel_identity_admin.requests'])await deny(f.db.exec(q));
 assert.equal((await f.db.query('select count(*)::int n from humbel_identity_admin.audit')).rows[0].n,1);
 await f.db.exec('update humbel_permissions.runtime set enabled=true');await deny(f.command('identity_deactivate',1));
}));
