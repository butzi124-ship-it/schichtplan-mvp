const assert=require('node:assert/strict');
const changes=require('../tests/qa/learning-integration-patches.json');
function reverseFile(file,source){
 for(const p of [...changes.filePatches].reverse())if(p.file===file){
  assert.equal(source.split(p.after).length-1,1,'Learning asset hook must remain exact and unique');
  source=source.replace(p.after,p.before);
 }
 return source;
}
module.exports={reverseFile};
