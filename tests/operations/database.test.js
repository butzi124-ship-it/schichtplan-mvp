const {test,before,after}=require('node:test');
const assert=require('node:assert/strict');
const {operationsDatabase,manager,worker,outsider,reviewer,department,foreignDepartment,machine,foreignMachine,order,station}=require('../helpers/operations-db');
const reference=require('../fixtures/article-reference.json');
const fixture=require('../fixtures/operations-reference.json');
const {blank}=require('../../js/operations');
let api,first,second;
before(async()=>{api=await operationsDatabase();});after(async()=>{await api?.db.close();});
const denied=(p,code)=>assert.rejects(p,e=>!code||e.code===code);
const command=(uid,action,r=null,extra={})=>api.rpc(uid,'operations_article_command',{p_action:action,p_department_id:department,p_article_number:'32505',p_data:reference.data,p_revision_id:r?.id,p_expected_version:r?.row_version,...extra});
const plan=(uid,r,operations)=>api.rpc(uid,'operation_plan',{p_revision_id:r.id,p_operations:operations,p_expected_version:r.row_version});
const createBA=(r,ba='812150',quantity=48,extra={},uid=manager)=>api.rpc(uid,'order_from_article',{p_revision_id:r.id,p_ba_number:ba,p_quantity:quantity,p_due_date:'2099-12-31',p_machine_id:machine,...extra});
async function released(number,operations=fixture.operations){let r=await command(manager,'create',null,{p_article_number:number});const p=await plan(manager,r,operations);r.row_version=p.row_version;return command(reviewer,'release',r);}
async function oldData(){return (await api.db.query(`select jsonb_build_object('orders',(select jsonb_agg(to_jsonb(o)) from public.production_orders o where id='${order}'),'stations',(select jsonb_agg(to_jsonb(s)) from public.production_order_stations s),'counts',(select jsonb_agg(to_jsonb(c)) from public.production_station_counts c),'checklist',(select jsonb_agg(to_jsonb(c)) from public.production_order_checklist c)) value`)).rows[0].value;}
test('only assigned author may edit plan; first OP cannot release or create BA; direct tables and anonymous denied',async()=>{
 first=await command(worker,'create');
 await denied(plan(outsider,first,fixture.operations),'42501');
 const p=await plan(worker,first,fixture.operations);first.row_version=p.row_version;
 await denied(command(worker,'release',first),'42501');
 await denied(createBA(first,'NOT-ALLOWED',1,{},worker),'42501');
 await denied(api.asUser(null,()=>api.db.query('select public.operation_scan($1)',['812150']),'anon'),'42501');
 for(const table of ['article_operation_plans','production_order_article_snapshots','production_order_operations'])await denied(api.asUser(manager,()=>api.db.query(`select * from public.${table}`)),'42501');
 const flags=await api.db.query("select count(*)::int n from pg_class where relname in ('article_operation_plans','production_order_article_snapshots','production_order_operations') and relrowsecurity");assert.equal(flags.rows[0].n,3);
});
test('SQL validates unique OPs, ordering/predecessors, exact fields and Tr/Te pairs without inventing unknown values',async()=>{
 const op=fixture.operations[0];
 const invalid=[[{...op,number:0}],[{...op,number:1000}],[op,op],[{...op,predecessor:100}],[{...op,tr:{value:-1,unit:'min'}}],[{...op,te:{value:1,unit:null}}],[{...op,tr:{value:null,unit:'h'}}],[{...op,te:{value:'2',unit:'min'}}],[{...op,tr:{value:1,unit:'days'}}],[{...op,ba_quantity:48}],[{...op,activity:''}],[{...op,number:'010'}],[{...op,number:1.5}]];
 for(const ops of invalid)await denied(plan(worker,first,ops),'22023');
 const result=await plan(worker,first,[{...op,tr:{value:1.5,unit:'h'},te:{value:0,unit:'s'}}]);first.row_version=result.row_version;
 const unknown=await plan(worker,first,fixture.operations);first.row_version=unknown.row_version;
 assert.deepEqual(unknown.operations[0].tr,{value:null,unit:null});
});
test('plan saves increment article version; stale requests, foreign/inactive machine and empty release roll back',async()=>{
 await denied(plan(worker,{...first,row_version:first.row_version-1},fixture.operations),'40001');
 await denied(plan(worker,first,[{...fixture.operations[0],machine_id:foreignMachine}]),'42501');
 const empty=await command(manager,'create',null,{p_article_number:'empty'});
 await denied(command(manager,'release',empty),'22023');
 assert.equal((await api.rpc(manager,'article_master_list',{p_query:'empty'}))[0].status,'draft');
 first=await command(manager,'release',first);
 await denied(plan(manager,first,fixture.operations),'40001');
 await denied(api.db.query('update public.article_operation_plans set operations=$1 where revision_id=$2',[[],first.id]),'42501');
});
test('two BAs use one released plan but only new quantities/numbers/due dates, never copy legacy counters/checklists',async()=>{
 const old=await oldData();
 const a=await createBA(first),b=await createBA(first,'812151',12);
 assert.equal(a.article_revision,first.revision);assert.equal(a.release_status,'released');assert.equal(a.released_by,first.released_by);assert.equal(a.released_at,first.released_at);assert.equal(a.ba_number,'812150');assert.equal(a.quantity,48);assert.equal(b.quantity,12);assert.notEqual(a.order_id,b.order_id);
 assert.deepEqual(a.operations,b.operations);assert.equal(a.op_rows[0].op_number,100);assert.equal(a.op_rows[0].data.machine_group,'52');
 assert.deepEqual(await oldData(),old);
 for(const id of [a.order_id,b.order_id]){
  const counts=await api.db.query('select (select count(*) from public.production_order_stations where order_id=$1)::int stations,(select count(*) from public.production_station_counts where order_id=$1)::int counters,(select count(*) from public.production_order_checklist where order_id=$1)::int checks',[id]);assert.deepEqual(counts.rows[0],{stations:0,counters:0,checks:0});
  const o=(await api.db.query('select ba_quantity,target_quantity,pallet_count,pieces_per_pallet,use_chain_logic,completed_at from public.production_orders where id=$1',[id])).rows[0];assert.equal(o.ba_quantity,o.target_quantity);assert.equal(o.pallet_count,0);assert.equal(o.pieces_per_pallet,0);assert.equal(o.use_chain_logic,false);assert.equal(o.completed_at,null);
 }
 second=b;
});
test('duplicate/invalid BA fields, draft/archived/unassigned sources reject without partial order',async()=>{
 const before=(await api.db.query('select count(*)::int n from public.production_orders')).rows[0].n;
 await denied(createBA(first),'23505');
 for(const ba of ['',' 812152','812152.010','ABC','1e3'])await denied(createBA(first,ba),'22023');
 for(const quantity of [0,-1,null])await denied(createBA(first,'812152',quantity),'22023');
 await denied(createBA(first,'812152',1,{p_due_date:'2020-01-01'}),'22023');
 await denied(createBA(first,'812152',1,{p_machine_id:foreignMachine}),'42501');
 await denied(createBA(first,'812152',1,{},outsider),'42501');
 const draft=await command(manager,'revise',first);await denied(createBA(draft,'812152'),'22023');
 const p=await plan(manager,draft);assert.deepEqual(p.operations,fixture.operations);
 assert.equal((await api.db.query('select count(*)::int n from public.production_orders')).rows[0].n,before);
});
test('injected OP insert failure rolls back order, link and snapshot atomically',async()=>{
 await api.db.exec("create function public.test_abort_op() returns trigger language plpgsql as $$begin raise exception 'synthetic failure'; end$$; create trigger test_abort before insert on public.production_order_operations for each row execute function public.test_abort_op()");
 await denied(createBA(first,'812152'));
 await api.db.exec('drop trigger test_abort on public.production_order_operations;drop function public.test_abort_op()');
 assert.equal((await api.db.query("select count(*)::int n from public.production_orders where ba_number='812152'")).rows[0].n,0);
 assert.equal((await api.db.query("select count(*)::int n from public.production_order_article_snapshots where ba_number='812152'")).rows[0].n,0);
 const valid=await createBA(first,'812152');assert.equal(valid.op_rows.length,1);
});
test('released OP snapshots cannot change; article revision edits preserve old BAs; optional second reviewer retained',async()=>{
 const before=await api.rpc(manager,'operation_order_read',{p_order_id:second.order_id});
 for(const sql of ['update public.production_order_operations set data=\'{}\' where order_id=$1','delete from public.production_order_operations where order_id=$1','update public.production_order_article_snapshots set quantity=999 where order_id=$1','delete from public.production_order_article_snapshots where order_id=$1',"update public.production_orders set ba_quantity=999 where id=$1"])await denied(api.db.query(sql,[second.order_id]),'42501');
 let draft=(await api.rpc(manager,'article_master_list',{p_query:'32505'})).find(r=>r.status==='draft');
 const ops=[{...fixture.operations[0],activity:'Synthetic revised activity'}];const p=await plan(manager,draft,ops);draft.row_version=p.row_version;
 await api.db.query('insert into public.article_master_settings values($1,true)',[department]);
 await denied(command(manager,'release',draft),'42501');
 draft=await command(reviewer,'release',draft);
 const fresh=await createBA(draft,'812153',7);assert.equal(fresh.operations[0].activity,'Synthetic revised activity');
 assert.deepEqual(await api.rpc(manager,'operation_order_read',{p_order_id:second.order_id}),before);
});
test('scans preserve leading BA zeros, enforce exactly 3 OP digits and remain read-only on repeats',async()=>{
 const source=await released('barcode-test',[{...blank(10),activity:'Synthetic OP 10',machine_group:'TEST'},{...blank(20,10),activity:'Synthetic OP 20',machine_group:'TEST'},{...blank(100,20),activity:'Synthetic OP 100',machine_group:'TEST'},{...blank(110,100),activity:'Synthetic OP 110',machine_group:'TEST'}]);
 const a=await createBA(source,'000812150',5);
 for(const op of ['010','020','100','110']){const first=await api.rpc(manager,'operation_scan',{p_barcode:'000812150.'+op});const again=await api.rpc(manager,'operation_scan',{p_barcode:'000812150.'+op});assert.deepEqual(first,again);assert.equal(first.selected_op,Number(op));assert.equal(first.ba_number,'000812150');}
 for(const s of ['000812150.10','000812150.000','000812150.999','000812150.010x',' 000812150','000812150.001.010'])await denied(api.rpc(manager,'operation_scan',{p_barcode:s}),'22023');
 await denied(api.rpc(outsider,'operation_scan',{p_barcode:'000812150.010'}),'42501');
 assert.equal((await api.rpc(manager,'operation_scan',{p_barcode:'000812150'})).selected_op,null);
 assert.equal(a.op_rows.length,4);
});
test('multiple tensions/employees can share one OP without modifying the snapshot; legacy orders remain editable',async()=>{
 const snapshot=await api.rpc(manager,'operation_order_read',{p_order_id:second.order_id});
 await api.db.query("insert into public.production_order_stations(order_id,station_no,op_number) values($1,1,'100'),($1,2,'100')",[second.order_id]);
 const stations=(await api.db.query('select * from public.production_order_stations where order_id=$1',[second.order_id])).rows;
 for(const [i,s] of stations.entries())await api.db.query('insert into public.production_station_counts(order_id,station_id,employee_id,good) values($1,$2,$3,$4)',[second.order_id,s.id,i?worker:manager,i+1]);
 assert.deepEqual(await api.rpc(manager,'operation_order_read',{p_order_id:second.order_id}),snapshot);
 await api.db.query('update public.production_orders set target_quantity=49 where id=$1',[order]);
 assert.equal((await api.db.query('select target_quantity from public.production_orders where id=$1',[order])).rows[0].target_quantity,49);
 await denied(api.db.query("insert into public.production_orders(id,ba_number,ba_quantity) values(gen_random_uuid(),'812150',2)"),'23505');
});

test('archived source, inactive scope/machine and revoked first OP rights fail on the next server request',async()=>{
 const r=await released('archived-source');const a=await command(manager,'archive',r);await denied(createBA(a,'900001'),'22023');
 await api.db.query('update public.production_machines set active=false where id=$1',[machine]);await denied(createBA(first,'900001'),'42501');await api.db.query('update public.production_machines set active=true where id=$1',[machine]);
 await api.db.query('update public.article_master_assignments set revoked=true where user_id=$1',[worker]);await denied(plan(worker,first),'42501');
 await api.db.query('update public.departments set active=false where id=$1',[department]);await denied(api.rpc(manager,'operation_order_read',{p_order_id:second.order_id}),'42501');await api.db.query('update public.departments set active=true where id=$1',[department]);
});
