const fs=require('node:fs');const {spawnSync}=require('node:child_process');
for(const file of ['supabase/functions/_shared/forecast-runtime.ts','supabase/functions/production-forecast/index.ts','supabase/functions/production-forecast-worker/index.ts']){
 const r=spawnSync(process.execPath,['--input-type=module','--check'],{input:fs.readFileSync(file),encoding:'utf8'});
 if(r.status!==0){process.stderr.write(file+'\n'+r.stderr);process.exit(1);}
}
console.log('Edge entry files: ECMAScript syntax passed; deployed Deno/Supabase execution is NOT tested.');
