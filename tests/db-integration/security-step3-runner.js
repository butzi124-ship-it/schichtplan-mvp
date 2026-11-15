// Exact, additive adapter AFTER the existing #26 runner/provenance checks.
const fs=require('node:fs');
const assert=require('node:assert/strict');
const {dispatch}=require('./integration-source');
const replacement=" await require('../../tests/db-integration/security-step3-extension').run({db,s,check,results});";
function prepare(root='.'){
 const file=require('node:path').join(root,'scripts/db-integration/run.js');
 const before=fs.readFileSync(file,'utf8');
 assert.equal(require('node:crypto').createHash('sha256').update(before).digest('hex'),
  'bc7a1d28fcf1c79e326510f01a5b52f7c232114a9fa82aed37587442cf4055f4',
  'Unknown runner: exact pinned clock/provenance adapters required first');
 assert.equal(before.split(dispatch).length-1,1,'Exact existing dispatch required');
 const after=before.replace(dispatch,replacement);
 assert.equal(after.replace(replacement,dispatch),before,'All native baseline assertions must survive');
 fs.writeFileSync(file,after);
}
module.exports={prepare};
if(require.main===module)prepare();
