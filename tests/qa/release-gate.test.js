const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {spawnSync} = require('node:child_process');
const YAML = require('yaml');
const workflow = YAML.parse(fs.readFileSync('.github/workflows/e2e.yml','utf8'));

test('actual Quality Gate script rejects every non-success dependency, including cancellations and skips',()=>{
 const file=path.join(fs.mkdtempSync(path.join(os.tmpdir(),'humbel-gate-')),'summary.md');
 const run = values => spawnSync('bash',['-e','-c',workflow.jobs['quality-gate'].steps[0].run],{env:{...process.env,GITHUB_STEP_SUMMARY:file,COMMIT:'isolated-test',...values}}).status;
 try {
  assert.equal(run({FUNCTIONAL:'success',INTEGRITY:'success',VISUAL:'success'}),0);
  for(const name of ['FUNCTIONAL','INTEGRITY','VISUAL']) for(const result of ['failure','cancelled','skipped','timed_out','']) {
   assert.notEqual(run({FUNCTIONAL:'success',INTEGRITY:'success',VISUAL:'success',[name]:result}),0,`${name}=${result}`);
  }
 } finally {fs.rmSync(path.dirname(file),{recursive:true});}
});
test('actual required e2e script fails closed; deployment needs separate manual owner approval',()=>{
 const job=workflow.jobs.e2e;
 assert.equal(job.name,'e2e');assert.equal(job.if,'always()');assert.equal(job.needs,'quality-gate');
 for(const result of ['success','failure','cancelled','skipped','']) {
  const status=spawnSync('bash',['-e','-c',job.steps[0].run],{env:{...process.env,GATE:result}}).status;
  assert.equal(status===0,result==='success');
 }
 assert.equal(workflow.jobs.deploy.needs,'e2e');
 assert.equal(workflow.jobs.deploy.if,"github.event_name == 'workflow_dispatch' && github.ref == 'refs/heads/main' && inputs.deploy_approved == true");
 assert.equal(workflow.on.workflow_dispatch.inputs.deploy_approved.default,false);
 assert.equal(workflow.on.workflow_dispatch.inputs.deploy_approved.type,'boolean');
 assert.deepEqual(workflow.permissions,{contents:'read'});
});

test('Git ownership exception is scoped to the checked-out workspace and SHA capture propagates errors',()=>{
 const steps=workflow.jobs.functional.steps;
 const syntax=steps.find(s=>s.name==='Syntax and whitespace checks').run;
 const report=steps.find(s=>s.name==='Preserve version and commit report').run;
 assert(syntax.includes('git -c safe.directory="$GITHUB_WORKSPACE" show --format= --check HEAD'));
 assert(report.includes('TESTED_SHA="$(git -c safe.directory="$GITHUB_WORKSPACE" rev-parse HEAD)"\nexport TESTED_SHA'));
 assert(!syntax.includes('--global'));assert(!report.includes('--global'));
});
