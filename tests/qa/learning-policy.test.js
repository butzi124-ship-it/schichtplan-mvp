const {test}=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs');
const {reverseFile}=require('../../scripts/learning-integration-policy');
const changes=require('./learning-integration-patches.json');
test('learning integration: exact assets reject missing/duplicate/changed hooks and preserve unrelated edits',()=>{
 const html=fs.readFileSync('schichtplan_mvp.html','utf8'),p=changes.filePatches[0];
 assert(reverseFile(p.file,html).includes(p.before));
 for(const altered of [html.replace(p.after,''),html.replace(p.after,p.after+p.after),html.replace(p.after,p.after.replace('false','true'))])assert.throws(()=>reverseFile(p.file,altered));
 assert(reverseFile(p.file,html+'UNREVIEWED').endsWith('UNREVIEWED'));
});
test('learning integration: every app hook exact/unique; protected source baseline remains unchanged',()=>{
 let s=fs.readFileSync('schichtplan_mvp.js','utf8');for(const p of [...changes.patches].reverse()){assert.equal(s.split(p.after).length-1,1);s=s.replace(p.after,p.before);}assert(!s.includes('const learning ='));
});
