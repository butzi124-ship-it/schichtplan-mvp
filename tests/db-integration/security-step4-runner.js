// Add a continuation to the exact already-adapted step-3 runner, preserving every assertion.
const fs=require('node:fs'),assert=require('node:assert/strict'),crypto=require('node:crypto'),path=require('node:path');
const before=" await require('../../tests/db-integration/security-step3-extension').run({db,s,check,results});";
const after=" await require('../../tests/db-integration/security-step4-extension').run({db,s,check,results});";
function prepare(root='.'){
 const file=path.join(root,'scripts/db-integration/run.js'),source=fs.readFileSync(file,'utf8');
 assert.equal(crypto.createHash('sha256').update(source).digest('hex'),'9b929b596e1fb07ea2404d8640e1cc545c67d50e32d150cf677ff3d308688fff','Exact step3 native runner required');
 assert.equal(source.split(before).length-1,1);
 const changed=source.replace(before,after);assert.equal(changed.replace(after,before),source);fs.writeFileSync(file,changed);
}
module.exports={prepare};if(require.main===module)prepare();
