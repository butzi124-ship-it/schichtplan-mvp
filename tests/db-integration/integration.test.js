const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');
const {assemble, inventory, verify, prepareRunner} = require('./integration-source');
const {run} = require('./integration-dispatch');
const {forecastDatabase} = require('../helpers/forecast-db');

test('native paths are explicit, fail closed and preserve the baseline security finding', async () => {
  for (const [mode, expected] of [['diagnostic',[]],['hardening',['./hardening-extension']],['ba-op',['./initial-responsibility-extension']]]) {
    const calls=[], context={results:{securityStatus:'CRITICAL_BASELINE_FINDINGS_REPRODUCED_NOT_FIXED'}};
    await run(context, mode, name => ({run:async ctx => {calls.push(name);ctx.results.securityStatus='CANDIDATE_ONLY';}}));
    assert.deepEqual(calls,expected);
    assert.equal(context.results.securityStatus,'CRITICAL_BASELINE_FINDINGS_REPRODUCED_NOT_FIXED');
    if(mode==='hardening') assert.equal(context.results.hardeningCandidateStatus,'CANDIDATE_ONLY');
  }
  for(const mode of [undefined,'','unknown','toString']) await assert.rejects(run({results:{}},mode),/Explicit native acceptance mode required/);
  await assert.rejects(run({results:{}},'hardening',()=>{throw Error('Missing mandatory module');}),/Missing mandatory module/);
});

test('assembly replaces foreign SQL and hashes the concrete checkout; mutations fail before DB use', () => {
  const tmp=fs.mkdtempSync(path.join(os.tmpdir(),'humbel-integration-'));
  const source=path.join(tmp,'source'),target=path.join(tmp,'target');
  try {
    for(const root of [source,target]) for(const dir of ['supabase/review','js','tests/db-integration','tests/fixtures'])fs.mkdirSync(path.join(root,dir),{recursive:true});
    for(const [file,content] of Object.entries({'supabase/review/a.sql':'select 1;','js/a.js':'// integration','tests/db-integration/a.js':'// extension','tests/fixtures/forecast-clock.json':'{}'}))fs.writeFileSync(path.join(source,file),content);
    fs.writeFileSync(path.join(source,'.gitignore'),'supabase/.temp/\n');
    const git=args=>execFileSync('git',args,{cwd:source,stdio:'pipe'});
    git(['init']);git(['add','.']);git(['-c','user.name=Test','-c','user.email=test@example.invalid','commit','-m','synthetic integration']);
    fs.writeFileSync(path.join(target,'supabase/review/foreign.sql'),'select 2;');
    assemble(source,target);
    assert.equal(fs.existsSync(path.join(target,'supabase/review/foreign.sql')),false);
    assert.deepEqual(verify(target).files,inventory(source));
    assert.equal(verify(target).commit,git(['rev-parse','HEAD']).toString().trim());
    for(const change of ['replace','delete','add']) {
      const file=path.join(target,'supabase/review/a.sql');
      if(change==='replace')fs.writeFileSync(file,'select 3;');
      if(change==='delete')fs.unlinkSync(file);
      if(change==='add')fs.writeFileSync(path.join(target,'supabase/review/foreign.sql'),'select 4;');
      assert.throws(()=>verify(target),/Integration source mismatch/);
      fs.writeFileSync(file,'select 1;');fs.rmSync(path.join(target,'supabase/review/foreign.sql'),{force:true});
    }
    fs.writeFileSync(path.join(source,'supabase/review/untracked.sql'),'select 5;');
    assert.throws(()=>assemble(source,target),/Untracked acceptance inputs/);
    fs.unlinkSync(path.join(source,'supabase/review/untracked.sql'));
    fs.mkdirSync(path.join(source,'supabase/.temp'));
    fs.writeFileSync(path.join(source,'supabase/.temp/ignored'),'untracked input');
    assert.throws(()=>assemble(source,target),/Untracked acceptance inputs/);
    fs.rmSync(path.join(source,'supabase/.temp'),{recursive:true});
    fs.appendFileSync(path.join(source,'supabase/review/a.sql'),'select 6;');
    assert.throws(()=>assemble(source,target));
    fs.symlinkSync(path.join(source,'supabase/review/a.sql'),path.join(target,'supabase/review/link.sql'));
    assert.throws(()=>verify(target),/Symlink forbidden/);
  } finally {fs.rmSync(tmp,{recursive:true,force:true});}
});

test('unknown runner cannot acquire an extension adapter', () => {
  const tmp=fs.mkdtempSync(path.join(os.tmpdir(),'humbel-runner-'));
  try{fs.mkdirSync(path.join(tmp,'scripts/db-integration'),{recursive:true});const file=path.join(tmp,'scripts/db-integration/run.js');fs.writeFileSync(file,'// changed harness');assert.throws(()=>prepareRunner(tmp),/Unknown native runner/);assert.equal(fs.readFileSync(file,'utf8'),'// changed harness');}
  finally{fs.rmSync(tmp,{recursive:true,force:true});}
});

test('installed forecast source triggers are AFTER STATEMENT on every bound table', async () => {
  const api=await forecastDatabase();
  try {
    const rows=(await api.db.query("select t.tgtype::int kind,pg_get_triggerdef(t.oid) definition from pg_trigger t where t.tgname='forecast_source_dirty'")).rows;
    assert.equal(rows.length,23);
    for(const row of rows){assert.equal(row.kind & 1,0,'Statement trigger required');assert.equal(row.kind & 2,0,'BEFORE lock inversion forbidden');assert.equal(row.kind & 64,0,'INSTEAD trigger forbidden');assert.match(row.definition,/AFTER INSERT OR DELETE OR UPDATE[\s\S]*FOR EACH STATEMENT/);}
  } finally {await api.db.close();}
});

test('workflow runs all modes independently with current-tree assembly and no renamed security hook', () => {
  const workflow=require('yaml').parse(fs.readFileSync('.github/workflows/ba-entry-isolated.yml','utf8'));
  const job=workflow.jobs['ba-entry-native'];
  assert.equal(job.strategy['fail-fast'],false);
  assert.deepEqual(job.strategy.matrix.acceptance,['diagnostic','hardening','ba-op']);
  const steps=job.steps.map(s=>s.run||'').join('\n');
  assert.match(steps,/integration-source.js assemble implementation isolated/);
  assert.match(steps,/run.js \$\{\{ matrix.acceptance \}\}/);
  assert.doesNotMatch(steps,/cp .*security-extension/);
  for(const chain of ['initial-responsibility:lifecycle','lifecycle:quantity','quantity:setup','setup:routing','routing:ba-entry']) {
    const [current,previous]=chain.split(':');assert.match(fs.readFileSync(`tests/db-integration/${current}-extension.js`,'utf8'),new RegExp(`require\\('\\./${previous}-extension'\\)\\.run\\(context\\)`));
  }
});
