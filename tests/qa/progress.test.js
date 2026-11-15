const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs'),os=require('node:os'),path=require('node:path');
const Reporter=require('../../scripts/qa-progress-reporter');
const YAML=require('yaml');
test('functional suite gets bounded graceful timeout before the job deadline without relaxing test assertions',()=>{
 const cfg=require('../../playwright.config');const workflow=YAML.parse(fs.readFileSync('.github/workflows/e2e.yml','utf8'));
 assert.equal(workflow.jobs.functional['timeout-minutes'],40);assert.equal(cfg.globalTimeout,30*60*1000);
 assert(cfg.globalTimeout+5*60*1000<workflow.jobs.functional['timeout-minutes']*60*1000);
 assert.equal(cfg.timeout,30000);assert.equal(cfg.expect.timeout,5000);assert.equal(cfg.retries,process.env.CI?1:0);
 assert.equal(workflow.concurrency['cancel-in-progress'],true);
 assert.equal(workflow.jobs['quality-gate'].if,'always()');
});
test('progress survives missing onEnd and records all attempts, failures and interrupted execution honestly',()=>{
 const previous=process.env.QA_ARTIFACT_DIR,dir=fs.mkdtempSync(path.join(os.tmpdir(),'qa-progress-'));process.env.QA_ARTIFACT_DIR=dir;
 const a={id:'a',expectedStatus:'passed',parent:{project:()=>({name:'chromium'})},titlePath:()=>['suite','a']};
 const b={...a,id:'b',titlePath:()=>['suite','b']};
 try{
  const r=new Reporter();r.onBegin({}, {allTests:()=>[a,b]});r.onTestBegin(a,{retry:0});r.onTestEnd(a,{retry:0,status:'failed'});r.onTestBegin(a,{retry:1});r.onTestEnd(a,{retry:1,status:'passed'});r.onTestBegin(b,{retry:0});
  let d=JSON.parse(fs.readFileSync(path.join(dir,'progress.json')));assert.equal(d.complete,false);assert.equal(d.planned,2);assert.deepEqual(d.attempts.map(x=>x.status),['failed','passed']);assert.equal(d.inFlight[0].id,'b');
  r.onError();r.onTestEnd(b,{retry:0,status:'interrupted'});r.onEnd({status:'timedout'});
  d=JSON.parse(fs.readFileSync(path.join(dir,'progress.json')));assert.equal(d.complete,true);assert.equal(d.runnerStatus,'timedout');assert.equal(d.runnerError,true);assert.equal(d.attempts.at(-1).status,'interrupted');assert.equal(d.inFlight.length,0);
 }finally{if(previous===undefined)delete process.env.QA_ARTIFACT_DIR;else process.env.QA_ARTIFACT_DIR=previous;fs.rmSync(dir,{recursive:true,force:true});}
});
