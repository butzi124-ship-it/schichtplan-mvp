// Executed only inside #15's pinned, sentinel-verified native isolated stack.
const fs=require('node:fs');
const assert=require('node:assert/strict');
const {rpc,http,connect,sqlAs}=require('../../scripts/db-integration/client');
async function run({db,s,check,results}) {
  const {people:p}=s;
  await db.query(fs.readFileSync('supabase/review/ba-entry-phase1.sql','utf8'));
  await db.query("notify pgrst,'reload schema'");await new Promise(resolve=>setTimeout(resolve,1200));
  const create=(ba,quantity=7,user=p.worker)=>rpc(user,'ba_entry_command',{p_action:'create',p_ba_number:ba,p_article_number:'SYNTH-ACCEPT',p_quantity:quantity});
  const open=(ba,op=null,user=p.worker)=>rpc(user,'ba_entry_command',{p_action:'open',p_ba_number:ba,p_op_number:op});
  let made;
  await check('BA entry defaults off; real JWT cannot create before isolated-only enablement',async()=>{
    await assert.rejects(create('812180'));
    await db.query('update public.production_ba_entry_settings set enabled=true');
    await db.query("insert into public.article_master_assignments(user_id,department_id,article_number,first_op_number,first_op_evidence,expires_at,issued_by) values($1,$2,'SYNTH-ACCEPT',10,'Synthetic independent first-OP grant',now()+interval '1 hour',$3)",[p.worker.id,s.dept,p.manager.id]);
  });
  await check('Real worker JWT creates BA using three fields; immutable plan, optional planning, pending AL and legacy counts preserved',async()=>{
    made=await create('000812180');assert.equal(made.existing,false);assert.equal(made.due_date,null);assert.equal(made.quantity,7);assert.equal(made.selected_op,10);assert.equal(made.op_rows.length,3);
    assert.equal((await db.query('select machine_id from public.production_orders where id=$1',[made.order_id])).rows[0].machine_id,null);
    const audit=(await db.query('select * from public.production_ba_entry_records where order_id=$1',[made.order_id])).rows[0];assert.equal(audit.actor_id,p.worker.id);assert.equal(audit.responsibility_status,'pending_rule');
    assert.equal((await db.query('select good_qty from public.production_station_counts where order_id=$1',[s.legacy])).rows[0].good_qty,17);
    assert.equal(Number((await db.query('select count(*) n from public.production_order_stations where order_id=$1',[made.order_id])).rows[0].n),0);
  });
  await check('Exact BA/OP reads and duplicate submissions do not modify data; foreign/anon/raw-table access denied',async()=>{
    assert.equal((await open('000812180',20)).selected_op,20);assert.equal((await create('000812180',99)).quantity,7);
    await assert.rejects(open('000812180',999));await assert.rejects(open('000812180',null,p.outsider));await assert.rejects(create('812181',1,p.outsider));
    const anon=await http(null,'/rpc/ba_entry_command',{method:'POST',body:{p_action:'open',p_ba_number:'000812180'}});assert.ok([401,403].includes(anon.status));
    for(const table of ['production_ba_entry_records','production_ba_entry_settings'])assert.ok([401,403].includes((await http(p.worker,'/'+table+'?limit=1')).status));
  });
  await check('Native first-OP wrong/expired/revoked grants and scanner/inactive roles fail closed',async()=>{
    for(const patch of ["first_op_number=20","first_op_number=10,revoked=true","revoked=false,expires_at=now()-interval '1 second'"]){await db.query('update public.article_master_assignments set '+patch+' where user_id=$1',[p.worker.id]);await assert.rejects(create('812181'));}
    await db.query("update public.article_master_assignments set expires_at=now()+interval '1 hour' where user_id=$1",[p.worker.id]);
    await db.query("update public.employees set role='tool_scanner' where id=$1",[p.worker.employee_id]);await assert.rejects(create('812181'));
    await db.query("update public.employees set role='employee',is_active=false where id=$1",[p.worker.employee_id]);await assert.rejects(create('812181'));
    await db.query('update public.employees set is_active=true where id=$1',[p.worker.employee_id]);
    await assert.rejects(create('812181',1,p.scanner));
  });
  await check('Independent concurrent BA requests resolve to one order and one creation audit',async()=>{
    const c1=await connect(),c2=await connect();
    try {
      const rows=await Promise.all([c1,c2].map(c=>sqlAs(p.worker,c=>c.query("select public.ba_entry_command('create','812182','SYNTH-ACCEPT',7,null) result"),{connection:c})));
      assert.equal(rows[0].rows[0].result.order_id,rows[1].rows[0].result.order_id);
      assert.deepEqual(rows.map(r=>r.rows[0].result.existing).sort(),[false,true]);
      assert.equal(Number((await db.query("select count(*) n from public.production_orders where ba_number='812182'")).rows[0].n),1);
      assert.equal(Number((await db.query("select count(*) n from public.production_ba_entry_records a join public.production_orders o on o.id=a.order_id where o.ba_number='812182'")).rows[0].n),1);
    } finally {await c1.end();await c2.end();}
  });
  await check('Native audit failure rolls back whole BA; records and prior OP snapshots are immutable',async()=>{
    await db.query("create function public.synthetic_ba_audit_abort() returns trigger language plpgsql as $$begin raise exception 'Synthetic audit failure';end$$;create trigger synthetic_ba_audit_abort before insert on public.production_ba_entry_records for each row execute function public.synthetic_ba_audit_abort()");
    try{await assert.rejects(create('812183'));assert.equal(Number((await db.query("select count(*) n from public.production_orders where ba_number='812183'")).rows[0].n),0);}
    finally{await db.query('drop trigger synthetic_ba_audit_abort on public.production_ba_entry_records;drop function public.synthetic_ba_audit_abort()');}
    await assert.rejects(db.query("update public.production_ba_entry_records set after_value='{}' where order_id=$1",[made.order_id]),e=>e.code==='42501');
    await assert.rejects(db.query("update public.production_order_operations set op_number=777 where order_id=$1",[made.order_id]),e=>e.code==='42501');
    assert.deepEqual((await open('000812180')).operations,made.operations);
  });
  results.limitations.push('BA Phase 1 tested against #15 structural mirror only; no authoritative restore/ACL approval, AL rule unresolved, no existing counter conversion or Phase 2-6 activation.');
}
module.exports={run};
