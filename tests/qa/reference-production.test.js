const fs=require('node:fs'),cp=require('node:child_process');
const {test}=require('node:test'),assert=require('node:assert/strict');
const {reverseReferenceFile}=require('../../scripts/reference-production-policy');
const {appBusiness,hash}=require('../../scripts/qa-policy');
const manifest=require('./reference-production-patches.json');
test('reference design restores the exact PR44 app and HTML; business and flags remain unchanged',()=>{
 for(const file of new Set(manifest.patches.map(p=>p.file))){
  const source=fs.readFileSync(file,'utf8');
  assert.equal(hash(appBusiness(reverseReferenceFile(file,source))),manifest.baseHashes[file]);
 }
});
test('reference guards reject changed, missing and duplicated presentation hooks',()=>{
 for(const p of manifest.patches){const s=fs.readFileSync(p.file,'utf8');assert.throws(()=>reverseReferenceFile(p.file,s.replace(p.after,'')));assert.throws(()=>reverseReferenceFile(p.file,s.replace(p.after,p.after+p.after)));}
});
test('reference presentation has no backend, storage, role or feature flag access',()=>{
 const s=fs.readFileSync('js/production-design.js','utf8');
 assert(!/\b(?:fetch|XMLHttpRequest|supabaseClient|localStorage|sessionStorage|currentUser|currentSupabaseUser|state)\b/.test(s));
 assert(!s.includes('HUMBEL_'));assert(!s.includes('32505'));assert(!s.includes('52978'));
});
