const {test}=require('node:test'),assert=require('node:assert/strict'),{randomUUID,createHash}=require('node:crypto'),fs=require('node:fs');
const {originalFixture,enroll,catalog,ids}=require('./step7-fixture'),review=require('../db-integration/security-step7-review'),prior=require('../db-integration/security-step6-review');
const deny=p=>assert.rejects(p,e=>e.code==='42501');
async function use(fn){const f=await originalFixture();try{await fn(f);}finally{await f.db.close();}}

test('STEP7 OBSERVATION: catalog covers 31/308/127 plus exact provenance and independent role-right observations',()=>{
 const c=catalog(),m=review.manifest(),extra=JSON.parse(fs.readFileSync('docs/qa/security-step7-effective-rights.json','utf8'));
 assert.equal(c.provenance.authoritativePgDump,false);assert.equal(c.provenance.productionWrites,false);assert.equal(c.relations.filter(t=>t.schema==='public').length,31);assert.equal(c.columns.filter(t=>t.schema==='public').length,308);assert.equal(c.policies.length,127);
 assert.equal(createHash('sha256').update(fs.readFileSync('docs/qa/security-step7-original-catalog.json.gz')).digest('hex'),m.catalogGzipSha256);
 assert.equal(extra.table_rights.length,496);assert(extra.table_rights.every(x=>x.allowed));assert.equal(extra.reported_api_schemas,null);
 assert(c.provenance.redactions.some(x=>x.schema==='storage'));assert.equal(c.functions.find(x=>x.schema==='storage'&&x.name==='list_objects_with_delimiter').definition,'REDACTED_PENDING_ADMIN_REVIEW');
});

test('STEP7 BASELINE ONLY: original role contradiction and mutable-employee helper escalation reproduced synthetically',async()=>use(async f=>{
 await assert.rejects(f.db.query("update public.employees set role='department_admin' where id=$1",[ids.leader]),e=>e.code==='23514');
 assert.equal((await f.asUser(ids.worker,d=>d.query('select public.is_admin() v'))).rows[0].v,false);
 await f.asUser(ids.worker,d=>d.query("update public.employees set role='admin' where id=$1",[ids.worker]));
 assert.equal((await f.asUser(ids.worker,d=>d.query('select public.is_admin() v'))).rows[0].v,true);
 assert.equal((await f.db.query("select has_table_privilege('anon','public.order_events','TRUNCATE') v")).rows[0].v,true);
}));

test('STEP7 GAP: unchanged PR31 does not close maintenance privileges on the other sixteen tables',async()=>use(async f=>{
 await prior.apply(f.db);assert.equal((await f.db.query("select has_table_privilege('authenticated','public.tools','TRUNCATE') v")).rows[0].v,true);
}));

test('STEP7 CORRECTED: all31 maintenance/column/default rights denied and role transition preserves profiles without enrollment',async()=>use(async f=>{
 const before=(await f.db.query('select * from public.employees order by id')).rows;
 await review.apply(f.db);await review.postconditions(f.db);assert.deepEqual((await f.db.query('select * from public.employees order by id')).rows,before);
 assert.equal((await f.db.query('select count(*)::int n from humbel_security.principals')).rows[0].n,0);
 await f.db.query("update public.employees set role='department_admin' where id=$1",[ids.leader]);
 await deny(f.assign(ids.leader)); // role string alone still grants no authority
 await enroll(f);for(const role of ['anon','authenticated'])await deny(f.asUser(null,d=>d.query('truncate public.order_events'),{anonymous:role==='anon'}));
 assert.equal((await review.apply(f.db)).status,'ISOLATED_STEP7_VERIFIED_NOOP');
}));

test('STEP7 CORRECTED: staff, leader, deputy and scanner cannot alter any of five own or foreign protected fields',async()=>use(async f=>{
 await review.apply(f.db);await enroll(f);const before=(await f.db.query('select * from public.employees order by id')).rows;
 for(const who of [ids.worker,ids.leader,ids.deputy,ids.scanner])for(const target of [who,ids.other])for(const change of ["role='admin'",`auth_user_id='${ids.orphan}'`,`department_id='${ids.d2}'`,"is_active=false",'id=gen_random_uuid()']){
  const r=await f.asUser(who,d=>d.query('update public.employees set '+change+' where id=$1 returning id',[target]));assert.deepEqual(r.rows,[]);
 }
 for(const who of [ids.worker,ids.leader,ids.deputy,ids.scanner]){
  await deny(f.asUser(who,d=>d.query("insert into public.employees(display_name,role) values('Denied','admin')")));
  assert.deepEqual((await f.asUser(who,d=>d.query('delete from public.employees where id=$1 returning id',[ids.other]))).rows,[]);
 }
 assert.deepEqual((await f.db.query('select * from public.employees order by id')).rows,before);
 for(const role of ['anon','authenticated'])await deny(f.asUser(null,d=>d.query('truncate public.employees cascade'),{anonymous:role==='anon'}));
}));

test('STEP7 CORRECTED: verified admin keeps CRUD; own approved test contact and timestamps work with trigger EXECUTE revoked',async()=>use(async f=>{
 await review.apply(f.db);await enroll(f);
 const row=(await f.asUser(ids.admin,d=>d.query("insert into public.employees(display_name,role) values('Synthetic managed','employee') returning id"))).rows[0];
 await f.asUser(ids.admin,d=>d.query("update public.employees set first_name='Synthetic admin update' where id=$1 returning id",[row.id]));
 assert.equal((await f.asUser(ids.admin,d=>d.query('delete from public.employees where id=$1 returning id',[row.id]))).rows[0].id,row.id);
 await f.db.query("insert into humbel_security.contact_fields values('synthetic_contact','Test-only field decision')");
 assert.deepEqual((await f.asUser(ids.worker,d=>d.query('select public.security_own_contacts($1) v',[{synthetic_contact:'own'}]))).rows[0].v,{synthetic_contact:'own'});
 await deny(f.asUser(ids.worker,d=>d.query('select public.security_own_contacts($1)',[{role:'admin'}])));
 assert.equal((await f.asUser(ids.worker,d=>d.query('select public.is_admin() v'))).rows[0].v,false);
 for(const fn of ['is_admin()','current_app_employee_role()','current_employee_id()'])await deny(f.asUser(null,d=>d.query('select public.'+fn),{anonymous:true}));
}));

test('STEP7 CORRECTED: own AL plan, foreign denial, delegated expiry, separate publication and privacy survive',async()=>use(async f=>{
 await review.apply(f.db);await enroll(f);assert.equal((await f.assign(ids.leader)).rows[0].result.version,1);await deny(f.assign(ids.leader,randomUUID(),0,ids.d2,ids.other));
 await f.db.query("insert into humbel_security.delegations values($1,$2,$3,now()-interval '1 day',now()+interval '1 day',false,'Synthetic delegation')",[ids.deputy,ids.d1,ids.leader]);assert.equal((await f.assign(ids.deputy,randomUUID())).rows[0].result.version,1);
 for(const update of ["revoked=true","revoked=false,valid_until=now()-interval '1 second'"]){await f.db.query('update humbel_security.delegations set '+update+' where user_id=$1',[ids.deputy]);await deny(f.assign(ids.deputy,randomUUID()));}
 await deny(f.asUser(ids.leader,d=>d.query('select public.security_publish_assignment($1,1)',[ids.assignment])));
 for(const [who,read,publish] of [[ids.worker,true,false],[ids.admin,false,true]])await f.db.query("insert into humbel_security.schedule_grants values($1,$2,$3,$4,now()-interval '1 day',now()+interval '1 day',false,'Synthetic schedule scope')",[who,ids.d1,read,publish]);
 await f.asUser(ids.admin,d=>d.query('select public.security_publish_assignment($1,1)',[ids.assignment]));const rows=(await f.asUser(ids.worker,d=>d.query('select * from public.security_published_schedule()'))).rows;
 assert.equal(rows.length,1);assert.deepEqual(Object.keys(rows[0]).sort(),['assignment_id','employee_id','shift_date','shift_id']);
 assert.deepEqual((await f.asUser(ids.worker,d=>d.query('select * from public.planner_sick_leaves'))).rows,[]);
}));

test('STEP7: unrecognized role constraint rolls back all stages and preserves original data/rights',async()=>use(async f=>{
 await f.db.exec('alter table public.employees drop constraint employees_role_check;alter table public.employees add constraint employees_role_check check(length(role)>0)');
 const before=await prior.snapshot(f.db),rows=(await f.db.query('select * from public.employees order by id')).rows;
 await assert.rejects(review.apply(f.db),e=>e.code==='55000');assert.equal(await prior.snapshot(f.db),before);assert.deepEqual((await f.db.query('select * from public.employees order by id')).rows,rows);
}));

test('STEP7 COMPATIBILITY ONLY: existing tool stock and journal payloads still execute; no full scanner/atomicity claim',async()=>use(async f=>{
 await review.apply(f.db);await enroll(f);
 const tool=(await f.asUser(ids.admin,d=>d.query("insert into public.tools(t_number,label,diameter,shelf,article_no,holder,stock) values('SYN7','Synthetic tool','10','SYN','SYN7','HSK 63',5) returning id"))).rows[0];
 const result=await f.asUser(ids.worker,async d=>{
  const row=(await d.query('update public.tools set stock=$1 where id=$2 returning *',[4,tool.id])).rows[0];assert.equal(row.stock,4);
  return d.query("insert into public.tool_journal(tool_id,tool_t_number,tool_label,action,qty,stock_before,stock_after,user_name) values($1,'SYN7','Synthetic tool','Synthetic withdrawal',1,5,4,'worker') returning id",[tool.id]);
 });assert.equal(result.rows.length,1);
 assert.equal((await f.db.query('select count(*)::int n from public.tool_journal where tool_id=$1',[tool.id])).rows[0].n,1);
}));

test('STEP7 CORRECTED: temp/public name spoofing cannot create admin authority; native unknown runner rejected',async()=>use(async f=>{
 await review.apply(f.db);await enroll(f);
 await deny(f.asUser(ids.worker,d=>d.exec('create table public.synthetic_attack(id int)')));
 await f.asUser(ids.worker,async d=>{await d.exec("create temporary table employees(role text);insert into employees values('admin')");assert.equal((await d.query('select public.is_admin() v')).rows[0].v,false);});
 const os=require('node:os'),path=require('node:path'),{prepare}=require('../db-integration/security-step7-runner');const dir=fs.mkdtempSync(path.join(os.tmpdir(),'step7-native-'));
 try{fs.mkdirSync(path.join(dir,'scripts/db-integration'),{recursive:true});const file=path.join(dir,'scripts/db-integration/run.js');fs.writeFileSync(file,'unknown');assert.throws(()=>prepare(dir),/Exact Step6/);assert.equal(fs.readFileSync(file,'utf8'),'unknown');}finally{fs.rmSync(dir,{recursive:true,force:true});}
}));
