const {test}=require('node:test');
const a=require('node:assert/strict');
const fs=require('node:fs');
const {randomUUID}=require('node:crypto');
const {quantityDatabase,worker,machine}=require('../helpers/quantity-db');

test('populated Phase 4 upgrades atomically: backfill commits after DDL, with RLS and deferred FK integrity intact',async()=>{
 const api=await quantityDatabase();
 try{
  // Match native CI: Phase 4 already ran and detected a real tension discrepancy.
  const source=await api.seed('816100'),op=source.op_rows[0].id;
  await api.rpc(worker,'op_routing_command',{p_operation_id:op,p_action:'select',p_request_id:randomUUID(),p_expected_version:0,p_machine_id:machine,p_reason:'Synthetic pre-Phase-5 machine'});
  await api.rpc(worker,'op_quantity_command',{p_operation_id:op,p_action:'configure',p_request_id:randomUUID(),p_new_count:2,p_note:'Synthetic existing tension plan'});
  const row=(await api.rpc(worker,'op_quantity_view')).rows.find(r=>r.operation_id===op);
  await api.rpc(worker,'op_quantity_command',{p_operation_id:op,p_action:'count',p_request_id:randomUUID(),p_station_id:row.stations[1].id,p_order_employee_id:row.own_person_id,p_section_id:row.section_id,p_kind:'good',p_delta:1});
  const conflict=(await api.db.query('select * from public.production_quantity_conflicts where operation_id=$1',[op])).rows[0];
  a.ok(conflict);
  const oldTables=['production_quantity_conflicts','production_quantity_notifications','production_quantity_events','production_station_counts','production_station_events','production_operation_machine_sections','production_operation_activity_segments'];
  async function snapshot(){const out={};for(const table of oldTables)out[table]=(await api.db.query(`select to_jsonb(t) value from public.${table} t order by to_jsonb(t)::text`)).rows;return out;}
  const before=await snapshot();
  // v0.14.0 raises 55006 here: backfilled rows queue deferred triggers before ALTER TABLE.
  await api.db.exec(fs.readFileSync('supabase/review/lifecycle-phase5.sql','utf8'));
  a.deepEqual(await snapshot(),before);
  const adopted=(await api.db.query('select * from public.production_lifecycle_cases where conflict_id=$1',[conflict.id])).rows;
  a.equal(adopted.length,1);a.equal(adopted[0].quantity_request,conflict.first_request);
  a.equal(adopted[0].detected_at.toISOString(),conflict.first_detected.toISOString());a.equal(adopted[0].after_completion,null);
  a.equal((await api.db.query('select enabled from public.production_lifecycle_settings')).rows[0].enabled,false);
  const readLifecycle=async()=> (await api.asUser(worker,()=>api.db.query('select public.ba_lifecycle_view() result'))).rows[0].result;
  a.deepEqual((await readLifecycle()).rows,[]);
  const protectedTables=(await api.db.query("select c.relname,c.relrowsecurity,has_table_privilege('authenticated',c.oid,'SELECT,INSERT,UPDATE,DELETE') auth_raw,has_table_privilege('anon',c.oid,'SELECT,INSERT,UPDATE,DELETE') anon_raw from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname='public' and c.relkind='r' and c.relname like 'production_lifecycle_%' order by c.relname")).rows;
  a.equal(protectedTables.length,10);a.ok(protectedTables.every(t=>t.relrowsecurity&&!t.auth_raw&&!t.anon_raw));
  await a.rejects(api.asUser(worker,()=>api.db.query('select * from public.production_lifecycle_cases')),e=>e.code==='42501');
  await a.rejects(api.db.query('update public.production_lifecycle_cases set after_completion=true where id=$1',[adopted[0].id]),e=>e.code==='42501');
  const deferred=(await api.db.query("select condeferrable,condeferred from pg_constraint where contype='f' and conrelid in('public.production_lifecycle_closures'::regclass,'public.production_lifecycle_rework'::regclass,'public.production_lifecycle_cases'::regclass,'public.production_lifecycle_case_entries'::regclass) and condeferrable")).rows;
  a.equal(deferred.length,4);a.ok(deferred.every(c=>c.condeferrable&&c.condeferred));
  // The repair must not turn deferred checks off or silently accept an orphan.
  await api.db.exec('begin');
  await api.db.query('insert into public.production_lifecycle_cases(conflict_id,quantity_request,after_completion) values($1,$2,false)',[conflict.id,randomUUID()]);
  a.equal((await api.db.query('select count(*)::int n from public.production_lifecycle_cases')).rows[0].n,2);
  await a.rejects(api.db.exec('commit'),e=>e.code==='23503');
  await api.db.exec('rollback');
  a.equal((await api.db.query('select count(*)::int n from public.production_lifecycle_cases')).rows[0].n,1);
  a.deepEqual(await snapshot(),before);
  // The first native Phase-5 test can now enable the synthetic module and use RPCs.
  await api.db.exec('update public.production_lifecycle_settings set enabled=true');
  const active=(await readLifecycle()).rows.find(r=>r.order_id===source.order_id);
  a.equal(active.cases.length,1);a.equal(active.cases[0].id,adopted[0].id);
  a.equal(active.cases[0].open,true);a.equal(active.cases[0].remaining,1);
  a.equal(active.operations[0].finished,1);a.deepEqual(await snapshot(),before);
 }finally{await api.db.close();}
});
