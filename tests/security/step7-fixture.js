// Threat fixture informed by current observed public ACLs, NOT original restore.
const {fixture,ids}=require('./fixture'),fs=require('node:fs'),zlib=require('node:zlib');
function catalog(){return JSON.parse(zlib.gunzipSync(fs.readFileSync('docs/qa/security-step7-original-catalog.json.gz')));}
async function originalFixture(){
 const f=await fixture(),c=catalog();
 await f.db.query("update public.employees set role='employee' where id=$1",[ids.leader]);
 const check=c.constraints.find(x=>x.schema==='public'&&x.table==='employees'&&x.name==='employees_role_check');
 await f.db.exec('alter table public.employees drop constraint employees_role_check;alter table public.employees add constraint employees_role_check '+check.definition);
 await f.db.exec('create role supabase_admin nologin');
 for(const t of c.relations.filter(t=>t.schema==='public'&&t.kind==='r'))await f.db.exec('grant all privileges on public.'+t.name+' to anon,authenticated');
 for(const fn of c.functions.filter(x=>x.schema==='public'))await f.db.exec('grant execute on function public.'+fn.name+'('+fn.identity_arguments+') to public,anon,authenticated');
 for(const owner of ['postgres','supabase_admin'])await f.db.exec('alter default privileges for role '+owner+' in schema public grant all on tables to anon,authenticated');
 await f.db.exec("select set_config('humbel.isolated_sentinel','synthetic-step7-only-sentinel-over-32-characters',false)");
 return f;
}
async function enroll(f){
 await f.db.query("update public.employees set role='department_admin' where id=$1",[ids.leader]);
 for(const who of ['admin','leader','worker','other','deputy','scanner']){
  const role=who==='admin'?'admin':who==='leader'?'department_admin':who==='scanner'?'tool_scanner':'employee';
  await f.db.query("insert into humbel_security.principals values($1,$1,$2,$3,true,now()+interval '1 day','Independent synthetic Step7 approval')",[ids[who],role,who==='other'?ids.d2:ids.d1]);
 }
}
module.exports={originalFixture,enroll,catalog,ids};
