const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),os=require('node:os'),path=require('node:path'),YAML=require('yaml');
const {prepare}=require('../db-integration/security-step4-runner');
test('step4 native runner rejects unknown source before writing; previous acceptance workflows remain required',()=>{
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'step4-native-'));
 try{fs.mkdirSync(path.join(dir,'scripts/db-integration'),{recursive:true});const file=path.join(dir,'scripts/db-integration/run.js');fs.writeFileSync(file,'unreviewed');assert.throws(()=>prepare(dir),/Exact step3 native runner/);assert.equal(fs.readFileSync(file,'utf8'),'unreviewed');}
 finally{fs.rmSync(dir,{recursive:true,force:true});}
 const w=YAML.parse(fs.readFileSync('.github/workflows/security-step4-isolated.yml','utf8'));
 const commands=w.jobs['security-step4-native'].steps.map(s=>s.run||'');
 assert.equal(w.jobs['security-step4-native'].needs,'security-contract');
 assert(commands.findIndex(s=>s.includes('security-step3-runner.js'))<commands.findIndex(s=>s.includes('security-step4-runner.js')));
 assert(commands.some(s=>s==='docker exec humbel-gateway node scripts/db-integration/run.js ba-op'));
 const source=fs.readFileSync('tests/db-integration/security-step4-extension.js','utf8');
 assert(source.indexOf("security-step3-extension").valueOf()<source.indexOf('security-step4-integration-isolated.sql'));
 const sql=fs.readFileSync('supabase/review/security-step4-integration-isolated.sql','utf8');
 assert(sql.includes('pg_catalog.current_database()'));assert(sql.includes('humbel.isolated_sentinel'));
 assert.doesNotMatch(sql,/alter table|drop policy|create or replace/i);
 assert(!fs.readFileSync('schichtplan_mvp.html','utf8').includes('secure-access.js'),'isolated adapter must not be enabled by production HTML');
});
