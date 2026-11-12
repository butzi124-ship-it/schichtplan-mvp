const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),os=require('node:os'),path=require('node:path'),crypto=require('node:crypto');
const YAML=require('yaml');
const {fixture}=require('./fixture');
const {prepare}=require('../db-integration/security-step3-runner');
const hash=file=>crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');

test('test catalog generator and original metadata are byte-identical to pinned PR15 sources',()=>{
 assert.equal(hash('tests/security/catalog-baseline.js'),'89ada9a20c0a8065800e3a0ac38b1fd905ef7ced4bf1bec6df5462e8541f368c');
 assert.equal(hash('tests/security/source-catalog.js'),'0aef5abec9cd227985c6334dbdd06eb10e21e85b346c4ca3f10436f0c492dfd7');
});

test('new native workflow preserves all three original acceptance modes and adds guarded independent candidate job',()=>{
 const original=YAML.parse(fs.readFileSync('.github/workflows/ba-entry-isolated.yml','utf8'));
 assert.deepEqual(original.jobs['ba-entry-native'].strategy.matrix.acceptance,['diagnostic','hardening','ba-op']);
 const next=YAML.parse(fs.readFileSync('.github/workflows/security-step3-isolated.yml','utf8'));
 assert.equal(next.jobs['security-step3-native'].needs,'security-contract');
 const commands=next.jobs['security-step3-native'].steps.map(s=>s.run||'');
 assert(commands.some(c=>c.includes('integration-source.js assemble implementation isolated')));
 assert(commands.findIndex(c=>c.includes('prepare-runner'))<commands.findIndex(c=>c.includes('security-step3-runner.js')));
 assert(commands.some(c=>c==='docker exec humbel-gateway node scripts/db-integration/run.js ba-op'));
 assert(next.jobs['security-step3-native'].steps.some(s=>s.if==='always()'&&s.run?.includes('runtime.js stop')));
 const extension=fs.readFileSync('tests/db-integration/security-step3-extension.js','utf8');
 assert(extension.indexOf("run(context,'ba-op')")<extension.indexOf("security-step3-isolated.sql"));
 assert(extension.includes('results.securityStatus=baseline'));
 assert(extension.includes('Promise.allSettled(connections.map'));
});

test('unknown native runner is rejected without modifying its source',()=>{
 const root=fs.mkdtempSync(path.join(os.tmpdir(),'step3-guard-'));
 try{fs.mkdirSync(path.join(root,'scripts/db-integration'),{recursive:true});const file=path.join(root,'scripts/db-integration/run.js');fs.writeFileSync(file,'unreviewed runner');assert.throws(()=>prepare(root),/Unknown runner/);assert.equal(fs.readFileSync(file,'utf8'),'unreviewed runner');}
 finally{fs.rmSync(root,{recursive:true,force:true});}
});

test('candidate SQL aborts before creating objects without the isolated sentinel',async()=>{
 const f=await fixture();
 try{
  await assert.rejects(f.db.exec(fs.readFileSync('supabase/review/security-step3-isolated.sql','utf8')),e=>e.code==='42501');
  await f.db.exec('rollback');
  assert.equal((await f.db.query("select count(*)::int n from pg_namespace where nspname='humbel_security'")).rows[0].n,0);
 }finally{await f.db.close();}
});
