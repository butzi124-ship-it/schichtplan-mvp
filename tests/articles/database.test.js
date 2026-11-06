const {test,before,after} = require('node:test');
const assert = require('node:assert/strict');
const {articleDatabase,manager,worker,outsider,reviewer,department,foreignDepartment,order} = require('../helpers/article-db');
const reference = require('../fixtures/article-reference.json');
let api, row;
before(async()=>{api=await articleDatabase();});
after(async()=>{await api?.db.close();});
const command=(uid,action,extra={})=>api.rpc(uid,'article_master_command',{p_action:action,p_department_id:department,p_article_number:'32505',p_data:reference.data,...extra});
const denied=(promise,code='42501')=>assert.rejects(promise,e=>e.code===code);

test('anonymous and unassigned users cannot execute privileged mutations or read data',async()=>{
 await denied(api.asUser(null,()=>api.db.query('select public.article_master_capabilities()'),'anon'));
 await denied(command(outsider,'create'));
 assert.deepEqual(await api.rpc(outsider,'article_master_list',{p_query:''}),[]);
 for(const table of ['scopes','assignments','items','revisions','audit','ba_links']) {
  await denied(api.asUser(manager,()=>api.db.query(`select * from public.article_master_${table}`)));
 }
 await denied(api.asUser(worker,()=>api.db.query(`insert into public.article_master_scopes values('${worker}','${foreignDepartment}')`)));
});
test('first OP assignment is exact, scoped and allows an initial draft only',async()=>{
 await denied(command(worker,'create',{p_article_number:'32506'}));
 await denied(command(worker,'create',{p_department_id:foreignDepartment}));
 row=await command(worker,'create');
 assert.equal(row.article_number,'32505'); assert.equal(row.revision,1); assert.equal(row.created_by,worker);
 assert.equal(row.data.material,'90MnCrV8');
 assert(!JSON.stringify(row.data).includes('812150')); assert(!JSON.stringify(row.data).includes('48'));
 for(const action of ['release','archive','revise']) await denied(command(worker,action,{p_revision_id:row.id,p_expected_version:row.row_version}));
});
test('duplicate identity normalization and hostile BA fields fail transactionally',async()=>{
 await denied(command(manager,'create',{p_article_number:' 32505 '}),'23505');
 for(const key of ['ba_number','ba_quantity','target_quantity','delivery_date','completed_at','unknown']) {
  await denied(command(manager,'create',{p_article_number:'bad-'+key,p_data:{...reference.data,[key]:'812150'}}),'22023');
 }
 await denied(command(manager,'create',{p_article_number:'bad-array',p_data:[]}), '22023');
 await denied(command(manager,'create',{p_article_number:'bad-type',p_data:{material:48}}),'22023');
 await denied(command(manager,'create',{p_article_number:'bad-long',p_data:{technical_notes:'x'.repeat(12001)}}),'22023');
 const items=await api.db.query('select count(*)::int as n from public.article_master_items'); assert.equal(items.rows[0].n,1);
});
test('edits use optimistic version, preserve content on stale request, and audit only successes',async()=>{
 const first=await command(worker,'update',{p_revision_id:row.id,p_expected_version:row.row_version,p_data:{...reference.data,drawing_number:'D-32505',drawing_revision:'A'}});
 await denied(command(worker,'update',{p_revision_id:row.id,p_expected_version:row.row_version}), '40001'); row=first;
 const audit=await api.db.query('select action from public.article_master_audit order by id'); assert.deepEqual(audit.rows.map(r=>r.action),['create','update']);
});
test('optional independent release, frozen linked revision, revision copy and archival retain old BA',async()=>{
 await api.db.query('insert into public.article_master_settings values($1,true)',[department]);
 // Worker-created draft can be released by a separately provisioned responsible person.
 row=await command(manager,'release',{p_revision_id:row.id,p_expected_version:row.row_version});
 assert.equal(row.released_by,manager); assert(row.released_at);
 await api.db.query('insert into public.article_master_ba_links values($1,$2)',[order,row.id]);
 await denied(command(manager,'update',{p_revision_id:row.id,p_expected_version:row.row_version}), '22023');
 await denied(api.db.query('update public.article_master_revisions set data=$1 where id=$2',[{material:'changed'},row.id]));
 await denied(api.db.query('delete from public.article_master_revisions where id=$1',[row.id]));
 let next=await command(manager,'revise',{p_revision_id:row.id,p_expected_version:row.row_version});
 assert.equal(next.revision,2); assert.deepEqual(next.data,row.data);
 await denied(command(manager,'revise',{p_revision_id:row.id,p_expected_version:row.row_version}),'23505');
 await denied(command(manager,'release',{p_revision_id:next.id,p_expected_version:next.row_version}));
 next=await command(reviewer,'release',{p_revision_id:next.id,p_expected_version:next.row_version}); assert.equal(next.status,'released');
 const archived=await command(manager,'archive',{p_revision_id:row.id,p_expected_version:row.row_version}); assert.equal(archived.status,'archived');
 const ba=await api.db.query('select o.ba_number,o.ba_quantity,l.revision_id from public.production_orders o join public.article_master_ba_links l on l.order_id=o.id');
 assert.deepEqual(ba.rows,[{ba_number:'812150',ba_quantity:48,revision_id:row.id}]);
 const old=await api.db.query('select data from public.article_master_revisions where id=$1',[row.id]); assert.deepEqual(old.rows[0].data,row.data);
});
test('expired/revoked first OP assignments and foreign department access are denied',async()=>{
 await api.db.query('update public.article_master_assignments set revoked=true where user_id=$1',[worker]);
 assert.deepEqual(await api.rpc(worker,'article_master_list',{p_query:''}),[]);
 assert.deepEqual((await api.rpc(worker,'article_master_capabilities')).assignments,[]);
 await denied(command(worker,'create'));
 await denied(command(manager,'create',{p_department_id:foreignDepartment,p_article_number:'foreign'}));
 await api.db.query("update public.article_master_assignments set revoked=false,expires_at=now()-interval '1 second' where user_id=$1",[worker]);
 await denied(command(worker,'create'));
});
test('scope uniqueness permits an independently authorized second department; inactive scopes and direct table reads fail closed',async()=>{
 await api.db.query('insert into public.article_master_scopes values($1,$2)',[reviewer,foreignDepartment]);
 const second=await command(reviewer,'create',{p_department_id:foreignDepartment});
 assert.equal(second.article_number,'32505'); assert.notEqual(second.article_id,row.article_id);
 await api.db.query('update public.departments set active=false where id=$1',[foreignDepartment]);
 assert((await api.rpc(reviewer,'article_master_list',{p_query:'32505'})).every(r=>r.department_id===department));
 await denied(command(reviewer,'update',{p_revision_id:second.id,p_expected_version:second.row_version}));
 const rls=await api.db.query("select count(*)::int n from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname='public' and c.relname like 'article_master_%' and c.relkind='r' and c.relrowsecurity");
 assert.equal(rls.rows[0].n,7);
 await api.db.exec('grant select on public.article_master_items,public.article_master_revisions to authenticated');
 const direct=await api.asUser(outsider,()=>api.db.query('select * from public.article_master_revisions'));
 assert.deepEqual(direct.rows,[]);
 await api.db.exec('revoke select on public.article_master_items,public.article_master_revisions from authenticated');
});

test('article search escapes wildcard input and case-folded duplicate numbers are constrained',async()=>{
 await command(manager,'create',{p_article_number:'N%_123'});
 await command(manager,'create',{p_article_number:'NXX123'});
 const literal=await api.rpc(manager,'article_master_list',{p_query:'n%_'});
 assert.deepEqual(literal.map(r=>r.article_number),['N%_123']);
 await command(manager,'create',{p_article_number:'AB-10'});
 await denied(command(manager,'create',{p_article_number:' ab-10 '}),'23505');
});
