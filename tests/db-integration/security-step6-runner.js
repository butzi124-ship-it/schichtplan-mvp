// Exact additional continuation after the unchanged Step4 runner adapter.
const fs=require('node:fs'),assert=require('node:assert/strict'),crypto=require('node:crypto'),path=require('node:path');
function prepare(root='.'){
 const file=path.join(root,'scripts/db-integration/run.js'),source=fs.readFileSync(file,'utf8');
 assert.equal(crypto.createHash('sha256').update(source).digest('hex'),'0d53f4830a02fb9a6194c979c8badd24365269240f8c0e6a3ca5f20ffc20342f','Exact Step4 native runner required');
 const before=" await require('../../tests/db-integration/security-step4-extension').run({db,s,check,results});";
 const after=" await require('../../tests/db-integration/security-step6-extension').run({db,s,check,results});";
 assert.equal(source.split(before).length,2);const next=source.replace(before,after);assert.equal(next.replace(after,before),source);fs.writeFileSync(file,next);
}
module.exports={prepare};if(require.main===module)prepare();
