// Embedded structural-catalog contract, never an original-schema restore or a JWT test.
const {PGlite}=require('@electric-sql/pglite');
const fs=require('node:fs');
const {sql}=require('./catalog-baseline');
const uid=n=>`10000000-0000-0000-0000-${String(n).padStart(12,'0')}`;
const ids={admin:uid(1),leader:uid(2),worker:uid(3),other:uid(4),deputy:uid(5),scanner:uid(6),orphan:uid(7),d1:uid(101),d2:uid(102),assignment:uid(201)};
async function fixture(){
 const db=new PGlite();
 await db.exec(`create role anon;create role authenticated;create role service_role;
 create schema auth;create table auth.users(id uuid primary key);
 `);
 // Explicit synthetic auth shims; real GoTrue/PostgREST coverage is a different runner.
 await db.exec(`create or replace function auth.uid() returns uuid language sql stable as $$select (nullif(current_setting('request.jwt.claims',true),'')::jsonb->>'sub')::uuid$$;
 create function auth.jwt() returns jsonb language sql stable as $$select coalesce(nullif(current_setting('request.jwt.claims',true),'')::jsonb,'{}')$$;
 grant usage on schema auth to anon,authenticated;grant execute on all functions in schema auth to anon,authenticated;`);
 await db.exec(sql());
 // Known phase-2 review transition, not claimed as the original constraint.
 await db.exec("alter table public.employees drop constraint employees_role_check;alter table public.employees add constraint employees_role_check check(role in('admin','employee','tool_scanner','department_admin'))");
 await db.query("insert into public.departments(id,name,code) values($1,'Synthetic A','A'),($2,'Synthetic B','B')",[ids.d1,ids.d2]);
 for(const [label,role,dept] of [['admin','admin','d1'],['leader','department_admin','d1'],['worker','employee','d1'],['other','employee','d2'],['deputy','employee','d1'],['scanner','tool_scanner','d1'],['orphan','admin','d1']]){
  await db.query('insert into auth.users values($1)',[ids[label]]);
  await db.query('insert into public.employees(id,auth_user_id,display_name,role,department_id) values($1,$1,$2,$3,$4)',[ids[label],label,role,ids[dept]]);
 }
 await db.query('update public.departments set leader_employee_id=$1 where id=$2',[ids.leader,ids.d1]);
 async function asUser(who,fn,{anonymous=false,claims={}}={}){
  await db.exec('begin');
  try{await db.exec('set local role '+(anonymous?'anon':'authenticated'));await db.query("select set_config('request.jwt.claims',$1,true)",[JSON.stringify({sub:who,role:anonymous?'anon':'authenticated',...claims})]);const r=await fn(db);await db.exec('commit');return r;}
  catch(e){await db.exec('rollback');throw e;}
 }
 async function install(){
  await db.query("select set_config('humbel.isolated_sentinel',$1,false)",['synthetic-embedded-only-sentinel-32-characters']);
  await db.exec(fs.readFileSync('supabase/review/security-step3-isolated.sql','utf8'));
  // Independently chosen synthetic fixture roles; the candidate itself inserts NO identities.
  for(const label of ['admin','leader','worker','other','deputy','scanner']){
   const role=label==='admin'?'admin':label==='leader'?'department_admin':label==='scanner'?'tool_scanner':'employee';
   await db.query("insert into humbel_security.principals values($1,$1,$2,$3,true,now()+interval '1 day','Independent synthetic enrollment')",[ids[label],role,label==='other'?ids.d2:ids.d1]);
  }
 }
 const assign=(who,id=ids.assignment,version=0,dept=ids.d1,subject=ids.worker)=>asUser(who,d=>d.query("select public.security_assignment_command($1,$2,$3,'synthetic-'||($1::uuid)::text,'2026-11-12',$4) result",[id,dept,subject,version]));
 return{db,asUser,install,assign};
}
module.exports={fixture,ids,uid};
