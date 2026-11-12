// Crash-resilient evidence only; never substitutes for final reports or changes status.
const fs=require('node:fs');
const path=require('node:path');
class ProgressReporter {
 onBegin(config,suite){
  this.dir=process.env.QA_ARTIFACT_DIR||'artifacts/functional';fs.mkdirSync(this.dir,{recursive:true});
  this.file=path.join(this.dir,'progress.json');
  this.report={sha:process.env.GITHUB_SHA||'local',complete:false,runnerStatus:'running',planned:suite.allTests().length,attempts:[],inFlight:[]};
  this.flush();
 }
 identity(test){return{id:test.id,project:test.parent.project()?.name||'unknown',title:test.titlePath().slice(1).join(' > ')};}
 onTestBegin(test,result){const row={...this.identity(test),retry:result.retry};this.report.inFlight.push(row);this.flush();}
 onTestEnd(test,result){
  const row={...this.identity(test),retry:result.retry,status:result.status,expectedStatus:test.expectedStatus};
  this.report.attempts.push(row);this.report.inFlight=this.report.inFlight.filter(x=>x.id!==row.id||x.retry!==row.retry);this.flush();
 }
 onError(){if(!this.report){this.onBegin({}, {allTests:()=>[]});this.report.planned=null;}this.report.runnerError=true;this.flush();}
 onEnd(result){this.report.complete=true;this.report.runnerStatus=result.status;this.flush();}
 flush(){const tmp=this.file+'.tmp';fs.writeFileSync(tmp,JSON.stringify(this.report,null,2));fs.renameSync(tmp,this.file);}
}
module.exports=ProgressReporter;
