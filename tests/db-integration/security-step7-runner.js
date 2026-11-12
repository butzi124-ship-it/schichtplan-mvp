const fs=require('node:fs'),assert=require('node:assert/strict'),crypto=require('node:crypto'),path=require('node:path');
function prepare(root='.'){
 const file=path.join(root,'scripts/db-integration/run.js'),source=fs.readFileSync(file,'utf8');
 assert.equal(crypto.createHash('sha256').update(source).digest('hex'),'f43e7aaeafe5221aa2ff250158cebf884e6a36f2354ca9f28674024f7536cdd0','Exact Step6 native runner required');
 const before=" await require('../../tests/db-integration/security-step6-extension').run({db,s,check,results});",after=" await require('../../tests/db-integration/security-step7-extension').run({db,s,check,results});";
 assert.equal(source.split(before).length,2);const next=source.replace(before,after);assert.equal(next.replace(after,before),source);fs.writeFileSync(file,next);
}
module.exports={prepare};if(require.main===module)prepare();
