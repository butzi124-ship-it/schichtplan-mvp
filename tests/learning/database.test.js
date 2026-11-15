const {test}=require('node:test');
const assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto');
const {setup,ids}=require('../helpers/learning-db');
test('learning: closed deployment defaults, independent enrollment, private tables and exact guarded install', {timeout:180000},async()=>{
 const f=await setup();try{
  const v=await f.view('worker');assert.equal(v.game_ready,false);assert.equal(v.enabled,true);assert.equal(v.visible,false);
  await assert.rejects(()=>f.call('worker','offer'),{code:'42501'});
  await assert.rejects(()=>f.asUser(ids.worker,d=>d.query('select * from humbel_learning.questions')),{code:'42501'});
  await assert.rejects(()=>f.asUser(ids.scanner,d=>d.query('select public.learning_view()')),{code:'42501'});
  await assert.rejects(()=>f.asUser(null,d=>d.query('select public.learning_view()'),{anonymous:true}),{code:'42501'});
  await f.db.query('update public.employees set is_active=false where id=$1',[ids.worker]);await assert.rejects(()=>f.view('worker'),{code:'42501'});
 }finally{await f.db.close();}
});
test('learning: adaptive offer, multi-select exact scoring, learned-repeat losses, experience and immutable evidence', {timeout:180000},async()=>{
 const f=await setup();try{
  await f.ready();await f.safe('worker');await f.question({difficulty:3});
  let v=await f.call('worker','offer');assert(v.unit);assert(!JSON.stringify(v.unit).includes('explanation'));assert(!Object.hasOwn(v.unit,'correct'));
  const first=v.unit.id,req=randomUUID();v=await f.call('worker','answer',{unit:first,answers:['a','c']},req);assert.equal(v.competence,20);assert.equal(v.experience,20);assert.equal(v.participation,1);
  assert.equal((await f.call('worker','answer',{unit:first,answers:['a','c']},req)).competence,20);
  await assert.rejects(()=>f.call('worker','answer',{unit:first,answers:['a','c']}),{code:'40001'});
  await f.question({difficulty:3,prompt:'A differently worded synthetic variant?'});
  await f.db.exec("update humbel_learning.mastery set due_at=now()-interval '1 day';update humbel_learning.people set next_offer=now()-interval '1 day'");
  v=await f.call('worker','offer');assert.equal(v.unit.repeat,true);assert.notEqual(v.unit.id,first);
  v=await f.call('worker','answer',{unit:v.unit.id,answers:['a']});assert.equal(v.competence,8);assert.equal(v.experience,20);assert.equal(v.history[0].details.correct,false);assert.match(v.history[0].details.explanation,/correct/);
  await assert.rejects(()=>f.db.exec('update humbel_learning.events set competence=999'),{code:'42501'});
 }finally{await f.db.close();}
});
test('learning: first wrong answer has no penalty, foreign unit and forged authority denied', {timeout:180000},async()=>{
 const f=await setup();try{await f.ready();await f.safe('worker');await f.question();let v=await f.call('worker','offer');
 await assert.rejects(()=>f.call('other','answer',{unit:v.unit.id,answers:['a','c']}),{code:'40001'});
 await assert.rejects(()=>f.call('worker','answer',{unit:v.unit.id,answers:['a','a']}),{code:'22023'});
 await assert.rejects(()=>f.asUser(ids.worker,d=>d.query('select public.learning_admin_view($1)',[ids.d1]),{claims:{role:'admin'}}),{code:'42501'});
 v=await f.call('worker','answer',{unit:v.unit.id,answers:['b']});assert.equal(v.competence,0);assert.equal(v.experience,0);
 }finally{await f.db.close();}
});
test('learning: deferrals charge only after third, exempt/unknown contexts never charge or advance penalty count', {timeout:180000},async()=>{
 const f=await setup();try{await f.ready();await f.safe('worker');await f.question();let v=await f.call('worker','offer');const id=v.unit.id;
 for(let i=1;i<=4;i++){await f.db.exec("update humbel_learning.units set available_at=now()-interval '1 minute'");v=await f.call('worker','defer',{unit:id,until:'30m'});assert.equal(v.participation,i<3?0:-5*(i-2));assert.equal(v.competence,0);}
 await f.db.query("insert into humbel_learning.context_windows(employee_id,starts_at,ends_at,eligible,exempt,evidence) values($1,now()-interval '1 hour',now()+interval '1 day',false,true,'Synthetic absence, no private reason')",[ids.worker]);
 await f.db.exec("update humbel_learning.units set available_at=now()-interval '1 minute'");v=await f.call('worker','defer',{unit:id,until:'today'});assert.equal(v.participation,-10);assert.equal(v.unit.deferrals,4);
 await f.db.exec("delete from humbel_learning.context_windows;update humbel_learning.units set available_at=now()-interval '1 minute'");v=await f.call('worker','defer',{unit:id,until:'30m'});assert.equal(v.participation,-10);assert.equal(v.unit.deferrals,4);
 }finally{await f.db.close();}
});
test('learning: privacy consent is separate from scoped participation administration; individual activation wins', {timeout:180000},async()=>{
 const f=await setup();try{await f.ready();assert.equal((await f.view('worker')).leaderboard.length,0);
 await f.call('worker','visibility',{visible:true});assert.equal((await f.view('deputy')).leaderboard.length,1);assert.equal((await f.view('other')).leaderboard.length,0);
 await f.call('worker','visibility',{visible:false});assert.equal((await f.view('deputy')).leaderboard.length,0);
 const admin=await f.asUser(ids.leader,d=>d.query('select public.learning_admin_view($1) result',[ids.d1]));assert(admin.rows[0].result.people.some(x=>x.id===ids.worker));assert(!JSON.stringify(admin.rows[0].result.people).includes('answers'));
 await assert.rejects(()=>f.asUser(ids.leader,d=>d.query('select public.learning_admin_view($1)',[ids.d2])),{code:'42501'});
 await assert.rejects(()=>f.call('leader','person_enabled',{department:ids.d1,person:ids.other,enabled:false}),{code:'42501'});
 await f.call('leader','department_enabled',{department:ids.d1,enabled:false});assert.equal((await f.view('worker')).enabled,false);
 await f.call('leader','person_enabled',{department:ids.d1,person:ids.worker,enabled:true});assert.equal((await f.view('worker')).enabled,true);
 }finally{await f.db.close();}
});
test('learning: versions, global publication approval and apprentice curriculum/topic gates', {timeout:180000},async()=>{
 const f=await setup();try{await f.ready();await f.safe('worker');const id=await f.question();
 await assert.rejects(()=>f.db.query("update humbel_learning.questions set prompt='overwrite' where id=$1",[id]),{code:'42501'});
 await f.db.query("update humbel_learning.people set learner=true,curriculum='futuremem2026',education_start='2026-08-01' where employee_id=$1",[ids.worker]);
 assert.equal((await f.call('worker','offer')).unit,null);
 await f.question({curriculum:'futuremem2026'});assert.equal((await f.call('worker','offer')).unit,null);
 await f.call('leader','education_topic',{department:ids.d1,person:ids.worker,topic:'topic-a',until:new Date(Date.now()+86400000).toISOString(),evidence:'Synthetic independently taught topic'});
 assert((await f.call('worker','offer')).unit);
 await assert.rejects(()=>f.call('leader','question_save',{department:ids.d1,global_pool:true,released:true}),{code:'42501'});
 }finally{await f.db.close();}
});
test('learning: current badge descends while achievement history/experience survive', {timeout:180000},async()=>{
 const f=await setup();try{await f.ready();await f.safe('worker');await f.db.exec('update humbel_learning.settings set thresholds=array[5,10,15,20,25]');await f.question();let v=await f.call('worker','offer');
 v=await f.call('worker','answer',{unit:v.unit.id,answers:['a','c']});assert.equal(v.badges.find(b=>b.code==='milling').stage,1);
 await f.question({prompt:'Changed variant'});await f.db.exec("update humbel_learning.mastery set due_at=now()-interval '1 day';update humbel_learning.people set next_offer=now()-interval '1 day'");v=await f.call('worker','offer');v=await f.call('worker','answer',{unit:v.unit.id,answers:['b']});assert.equal(v.competence,2);assert.equal(v.experience,5);assert.equal(v.badges.find(b=>b.code==='milling').stage,0);assert(v.badge_history.some(b=>b.stage===1));
 }finally{await f.db.close();}
});
test('learning: safety is independent, theory is not practical certification, failures and periodic evidence retained', {timeout:180000},async()=>{
 const f=await setup();try{await f.ready();const q=await f.question({kind:'safety'});await f.call('leader','person_enabled',{department:ids.d1,person:ids.worker,enabled:false});
 await f.call('leader','instruction_assign',{department:ids.d1,person:ids.worker,question:q,reason:'Neue Maschine – synthetic'});let v=await f.view('worker');assert.equal(v.enabled,false);const id=v.instructions[0].id;
 await assert.rejects(()=>f.call('leader','instruction_confirm',{department:ids.d1,assignment:id,evidence:'not yet taught'}),{code:'22023'});
 v=await f.call('worker','instruction_answer',{assignment:id,answers:['b']});assert.equal(v.instructions[0].failures,1);assert.equal(v.competence,0);
 v=await f.call('worker','instruction_answer',{assignment:id,answers:['a','c']});assert(v.instructions[0].theory_passed_at);assert.equal(v.instructions[0].confirmed_at,null);
 await assert.rejects(()=>f.call('worker','instruction_confirm',{department:ids.d1,assignment:id,evidence:'self-approved'}),{code:'42501'});
 await f.call('leader','instruction_confirm',{department:ids.d1,assignment:id,evidence:'Synthetic actual instruction evidence, NOT a production certificate'});v=await f.view('worker');assert.equal(v.instructions.length,2);assert(v.instructions.some(a=>a.confirmed_at));assert.equal(v.participation,0);
 }finally{await f.db.close();}
});
test('learning: failed audit rolls back answer/score/receipt atomically; request payload conflict preserved', {timeout:180000},async()=>{
 const f=await setup();try{await f.ready();await f.safe('worker');await f.question();let v=await f.call('worker','offer');const id=v.unit.id;
 await f.db.exec("create function public.synthetic_learning_failure() returns trigger language plpgsql as $$begin raise exception 'synthetic audit failure';end$$;create trigger fail_learning before insert on humbel_learning.events for each row execute function public.synthetic_learning_failure()");
 await assert.rejects(()=>f.call('worker','answer',{unit:id,answers:['a','c']}),/synthetic audit failure/);v=await f.view('worker');assert.equal(v.competence,0);assert.equal(v.unit.id,id);
 await f.db.exec('drop trigger fail_learning on humbel_learning.events');const req=randomUUID();await f.call('worker','visibility',{visible:true},req);await assert.rejects(()=>f.call('worker','visibility',{visible:false},req),{code:'40001'});
 }finally{await f.db.close();}
});
test('learning: all difficulty rewards, persistent participation counts and revoke/expiry denial', {timeout:180000},async()=>{
 const f=await setup();try{await f.ready();await f.safe('worker');let total=0;
 for(const [difficulty,points] of [[1,5],[2,10],[3,20]]){await f.question({difficulty,concept:'difficulty-'+difficulty});await f.db.exec("update humbel_learning.people set next_offer=now()-interval '1 day'");let v=await f.call('worker','offer');v=await f.call('worker','answer',{unit:v.unit.id,answers:['a','c']});total+=points;assert.equal(v.competence,total);}
 assert.equal((await f.view('worker')).answered_count,3);
 await f.db.query("update humbel_learning.scopes set revoked=true where user_id=$1",[ids.leader]);await assert.rejects(()=>f.asUser(ids.leader,d=>d.query('select public.learning_admin_view($1)',[ids.d1])),{code:'42501'});
 await f.call('worker','visibility',{visible:true});await f.db.exec('update humbel_learning.settings set enabled=false');await f.call('worker','visibility',{visible:false});assert.equal((await f.view('worker')).visible,false);
 }finally{await f.db.close();}
});
test('learning: exact global editorial approval permits a shared pool; altered payload is refused', {timeout:180000},async()=>{
 const f=await setup();try{await f.ready();await f.safe('other');const previous=await f.question();
 const body={department:ids.d1,domain:'milling',topic:'topic-a',concept:'shared-concept',difficulty:1,prompt:'Synthetic shared theoretical variant?',options:[{id:'a',label:'A'},{id:'b',label:'B'}],correct:['a'],explanation:'Synthetic verified explanation',source:'Synthetic editorial source',global_pool:true,released:true,previous_id:previous};
 await f.db.query('update humbel_learning.scopes set can_global=true where user_id=$1',[ids.leader]);
 await assert.rejects(()=>f.call('leader','question_save',body),{code:'42501'});
 await f.db.query("insert into humbel_learning.global_approvals values($1,$2,now()+interval '1 day','Independent synthetic GL editorial approval, not production enrollment')",[previous,body]);
 await assert.rejects(()=>f.call('leader','question_save',{...body,prompt:'Unapproved alteration'}),{code:'42501'});
 await f.call('leader','question_save',body);assert((await f.call('other','offer')).unit);
 }finally{await f.db.close();}
});
