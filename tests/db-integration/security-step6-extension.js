// Separate additive acceptance. All earlier baseline/BA-OP/Step3/4 proofs retained.
const assert=require('node:assert/strict');
const {http}=require('../../scripts/db-integration/client');
const {sources,postconditions,ten}=require('./security-step6-review');
async function run(context){
 await require('./security-step4-extension').run(context);
 const {db,s,check,results}=context,p=s.people;assert.equal(results.tests.filter(t=>t.status==='failed').length,0);
 const start=results.tests.length;
 const contents=async()=>{const rows={};for(const t of [...ten,'tools','tool_journal'])rows[t]=(await db.query('select * from public.'+t+' order by id')).rows;return rows;};
 const before=await contents();
 await check('Step6 native additive ACL transition preserves all ten tables, tools and journal rows',async()=>{
  const body=sources()[2];await db.query('begin');
  try{await db.query('lock table public.tool_materials in access exclusive mode');await db.query(body);await postconditions(db);await db.query('commit');}
  catch(e){await db.query('rollback');throw e;}
  assert.deepEqual(await contents(),before);
 });
 await check('Step6 real REST anon cannot read any of ten formerly unprotected tables',async()=>{
  for(const t of ten)assert.equal((await http(null,'/'+t+'?select=id&limit=1')).ok,false,t);
 });
 let id;
 await check('Step6 real REST admin creates, renames and deactivates material with existing application payloads',async()=>{
  const r=await http(p.manager,'/tool_materials',{method:'POST',headers:{prefer:'return=representation'},body:{name:'Synthetic Step6 native material',sort_order:900,is_active:true}});assert.equal(r.ok,true);id=r.data[0].id;
  assert.equal((await http(p.manager,'/tool_materials?id=eq.'+id,{method:'PATCH',body:{name:'Synthetic Step6 renamed',updated_at:new Date().toISOString()}})).ok,true);
  assert.equal((await http(p.manager,'/tool_materials?id=eq.'+id,{method:'PATCH',body:{is_active:false,updated_at:new Date().toISOString()}})).ok,true);
  assert.equal((await db.query('select is_active from public.tool_materials where id=$1',[id])).rows[0].is_active,false);
 });
 await check('Step6 real REST staff/scanner material read retained, create and protected writes/delete denied',async()=>{
  assert(id,'Admin reference setup must have succeeded');
  for(const user of [p.worker,p.scanner]){
   assert.equal((await http(user,'/tool_materials?select=id,name,sort_order,is_active')).ok,true);
   assert.equal((await http(user,'/tool_materials',{method:'POST',body:{name:'Denied native material'}})).ok,false);
   const r=await http(user,'/tool_materials?id=eq.'+id,{method:'PATCH',headers:{prefer:'return=representation'},body:{name:'Denied rename'}});assert(!r.ok||r.data.length===0);
  }
  assert.equal((await http(p.manager,'/tool_materials?id=eq.'+id,{method:'PATCH',body:{created_at:new Date().toISOString()}})).ok,false);
  assert.equal((await http(p.manager,'/tool_materials?id=eq.'+id,{method:'DELETE'})).ok,false);
  assert.equal((await db.query('select name from public.tool_materials where id=$1',[id])).rows[0].name,'Synthetic Step6 renamed');
 });
 results.step6ReviewStatus=results.tests.slice(start).every(t=>t.status==='passed')?'NATIVE_ADDITIVE_REVIEW_VERIFIED__NO_PRODUCTION_CLEARANCE':'FAILED';
 results.limitations.push('Step6 native continuation tests additive material ACLs after original Step3/4+BA/OP proofs; complete fresh-install/NOOP/rollback package tested separately in PGlite, not as a native original-schema restore. Full planner/private/tool workflow compatibility remains blocked.');
}
module.exports={run};
