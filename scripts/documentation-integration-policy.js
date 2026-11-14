const assert=require('node:assert/strict'),patches=require('../tests/qa/documentation-integration-patches.json');
function reverseFile(file,source){for(const p of patches)if(p.file===file){assert.equal(source.split(p.after).length-1,1,'Reviewed help assets must be exact and unique');source=source.replace(p.after,p.before);}return source;}module.exports={reverseFile};
