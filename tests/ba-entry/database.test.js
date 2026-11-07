const {test,before,after}=require('node:test');
const assert=require('node:assert/strict');
const {baEntryDatabase,manager,worker,reviewer,outsider,department,order}=require('../helpers/ba-entry-db');
const {blank}=require('../../js/operations');
const article=require('../fixtures/article-reference.json');
let api,revision;
const create=(uid,ba,quantity=48,number='32505')=>api.rpc(uid,'ba_entry_command',{p_action:'create',p_ba_number:ba,p_article_number:number,p_quantity:quantity});
const open=(uid,ba,op=null)=>api.rpc(uid,'ba_entry_command',{p_action:'open',p_ba_number:ba,p_op_number:op});
before(async()=>{
  api=await baEntryDatabase();
  let r=await api.rpc(manager,'operations_article_command',{p_action:'create',p_department_id:department,p_article_number:'32505',p_data:article.data});
  const ops=[10,20,100].map((n,i,a)=>({...blank(n,i?a[i-1]:null),activity:'Synthetic OP '+n,machine_group:'52'}));
  const plan=await api.rpc(manager,'operation_plan',{p_revision_id:r.id,p_operations:ops,p_expected_version:r.row_version});
  revision=await api.rpc(reviewer,'operations_article_command',{p_action:'release',p_revision_id:r.id,p_expected_version:plan.row_version});
});
after(async()=>api?.db.close());
test('assigned first-OP employee creates from exactly three fields; no machine/deadline/AL invented; no legacy counters copied',async()=>{
  const before=(await api.db.query('select * from public.production_station_counts')).rows;
  const s=await create(worker,'000812160');
  assert.equal(s.quantity,48);assert.equal(s.due_date,null);assert.equal(s.selected_op,10);assert.equal(s.responsibility_status,'pending_rule');
  assert.equal(s.op_rows.length,3);assert.equal((await api.db.query('select machine_id from public.production_orders where id=$1',[s.order_id])).rows[0].machine_id,null);
  assert.deepEqual((await api.db.query('select * from public.production_station_counts')).rows,before);
  assert.equal((await api.db.query('select count(*)::int n from public.production_order_stations where order_id=$1',[s.order_id])).rows[0].n,0);
  const audit=(await api.db.query('select * from public.production_ba_entry_records where order_id=$1',[s.order_id])).rows[0];
  assert.equal(audit.authorization_source,'first_op_assignment');assert.equal(audit.actor_id,worker);assert.equal(audit.before_value,null);assert.equal(audit.after_value.ba_number,'000812160');
});
test('duplicate submissions open existing immutable BA and do not change quantities or add creation audit',async()=>{
  const s=await create(worker,'000812160',999);assert.equal(s.existing,true);assert.equal(s.quantity,48);
  assert.equal((await api.db.query("select count(*)::int n from public.production_orders where ba_number='000812160'")).rows[0].n,1);
  assert.equal((await api.db.query('select count(*)::int n from public.production_ba_entry_records where order_id=$1',[s.order_id])).rows[0].n,1);
});
test('existing barcode and manual OP lookup preserve identity/leading zeros without writes; unknown OP denied',async()=>{
  for(const n of [10,20,100])assert.equal((await open(worker,'000812160',n)).selected_op,n);
  await assert.rejects(open(worker,'000812160',110),e=>e.code==='22023');
  for(const ba of ['812160.100','1e3',' 000812160'])await assert.rejects(open(worker,ba),e=>e.code==='22023');
  await assert.rejects(open(outsider,'000812160'),e=>e.code==='42501');
});
test('wrong first OP, expiry, revocation, scanner, inactive and role escalation without trusted scope fail closed',async()=>{
  await api.db.query('update public.article_master_assignments set first_op_number=20 where user_id=$1',[worker]);
  await assert.rejects(create(worker,'812161'));await api.db.query('update public.article_master_assignments set first_op_number=10,revoked=true where user_id=$1',[worker]);
  await assert.rejects(create(worker,'812161'));await api.db.query("update public.article_master_assignments set revoked=false,expires_at=now()-interval '1 second' where user_id=$1",[worker]);
  await assert.rejects(create(worker,'812161'));await api.db.query("update public.article_master_assignments set expires_at=now()+interval '1 hour' where user_id=$1",[worker]);
  await api.db.query("update public.employees set role='tool_scanner' where auth_user_id=$1",[worker]);await assert.rejects(create(worker,'812161'));
  await api.db.query("update public.employees set role='employee',is_active=false where auth_user_id=$1",[worker]);await assert.rejects(create(worker,'812161'));
  await api.db.query('update public.employees set is_active=true where auth_user_id=$1',[worker]);
  await api.db.query("update public.employees set role='admin' where auth_user_id=$1",[outsider]);await assert.rejects(create(outsider,'812161'));
});
test('existing legacy BA can be opened by scoped manager without conversion; worker cannot infer old OP authority',async()=>{
  await api.db.query("update public.production_orders set ba_number='812170' where id=$1",[order]);
  const s=await create(manager,'812170');assert.equal(s.existing,true);assert.equal(s.legacy,true);assert.equal(s.order_id,order);
  assert.equal((await api.db.query('select count(*)::int n from public.production_order_article_snapshots where order_id=$1',[order])).rows[0].n,0);
  await assert.rejects(open(worker,'812170'),e=>e.code==='42501');
});
test('audit failure aborts the whole BA transaction; creation audit and snapshots cannot be edited',async()=>{
  await api.db.exec("create function public.synthetic_audit_failure() returns trigger language plpgsql as $$begin raise exception 'audit unavailable';end$$;create trigger synthetic_audit_failure before insert on public.production_ba_entry_records for each row execute function public.synthetic_audit_failure()");
  try{await assert.rejects(create(worker,'812171'));assert.equal((await api.db.query("select count(*)::int n from public.production_orders where ba_number='812171'")).rows[0].n,0);}
  finally{await api.db.exec('drop trigger synthetic_audit_failure on public.production_ba_entry_records;drop function public.synthetic_audit_failure()');}
  await assert.rejects(api.db.query("update public.production_ba_entry_records set after_value='{}'"),e=>e.code==='42501');
  await assert.rejects(api.asUser(worker,()=>api.db.query('select * from public.production_ba_entry_records')),e=>e.code==='42501');
  await assert.rejects(api.asUser(worker,()=>api.db.query('select article_master_private.ba_entry_read(null,null)')),e=>e.code==='42501');
});
test('new released article revision affects only future BA; existing OP snapshot stays frozen',async()=>{
  const old=await open(worker,'000812160');
  const draft=await api.rpc(manager,'operations_article_command',{p_action:'revise',p_revision_id:revision.id,p_expected_version:revision.row_version});
  const plan=await api.rpc(manager,'operation_plan',{p_revision_id:draft.id,p_operations:old.operations.map(o=>({...o,description:'new revision only'})),p_expected_version:draft.row_version});
  await api.rpc(reviewer,'operations_article_command',{p_action:'release',p_revision_id:draft.id,p_expected_version:plan.row_version});
  const next=await create(worker,'812172',12);assert.equal(next.article_revision,2);assert.equal(next.op_rows[0].data.description,'new revision only');
  assert.deepEqual((await open(worker,'000812160')).op_rows,old.op_rows);
});
test('server defaults off and anon is denied; existing five-field API validation is unchanged',async()=>{
  await api.db.exec('update public.production_ba_entry_settings set enabled=false');
  try{await assert.rejects(create(manager,'812173'));await assert.rejects(open(worker,'000812160'));}
  finally{await api.db.exec('update public.production_ba_entry_settings set enabled=true');}
  await assert.rejects(api.asUser(null,()=>api.db.query("select public.ba_entry_command('create','812173','32505',1)"),'anon'),e=>e.code==='42501');
  await assert.rejects(api.rpc(manager,'order_from_article',{p_revision_id:revision.id,p_ba_number:'812173',p_quantity:1}),e=>e.code==='22023');
});
