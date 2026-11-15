const {test}=require('node:test'),assert=require('node:assert/strict');
const {create,badge}=require('../../js/learning');
const tick=()=>new Promise(r=>setTimeout(r,0));
const model=()=>({user_id:'person',installed:true,game_ready:true,safety_ready:true,enabled:true,visible:false,learner:false,provisional:true,competence:5,participation:-5,experience:20,level:1,review_count:1,next_offer:'2099-01-01',badges:[{code:'milling',label:'Fräsen',motif:'cutter',stage:1,points:5}],unit:null,history:[],badge_history:[],leaderboard:[],instructions:[],departments:[],author_departments:[]});
test('learning UI: rollout off sends no RPC and never displays invented personal scores',async()=>{
 let calls=0;const ui=create({client:{rpc:async()=>{calls++;}},actor:()=> 'person',redraw:()=>{},enabled:()=>false});
 ui.ensureSession();await tick();assert.equal(calls,0);assert.match(ui.render(),/keine Beispielpunkte/);assert(!ui.render().includes('Fachkompetenzpunkte</span>'));
});
test('learning UI: private metrics, all five distinct tool tiers, safety separation and escaping',async()=>{
 const v=model();v.history=[{kind:'answered',competence:5,participation:1,created_at:'2026-01-01',details:{correct:false,explanation:'<img src=x onerror=bad()>',source:'<secret>'}}];
 v.instructions=[{id:'one',prompt:'Theory?',options:[],due_at:'2026-01-01',reason:'New entry',source:'Source',resource:'<machine>',failures:0,confirmed_at:null,theory_passed_at:null}];
 const ui=create({client:{rpc:async()=>({data:v,error:null})},actor:()=> 'person',redraw:()=>{},enabled:()=>true});ui.ensureSession();await tick();
 const html=ui.render();assert.match(html,/Fachkompetenzpunkte/);assert.match(html,/Teilnahmepunkte/);assert.match(html,/Privater Lernfortschritt/);assert(!html.includes('<img'));assert(html.includes('&lt;img'));assert.match(ui.render(true),/ersetzt keine praktische Unterweisung/);assert(!ui.render(true).includes('Fachkompetenzpunkte'));
 assert.equal(new Set([1,2,3,4,5].map(n=>badge('cutter',n))).size,5);assert(!badge('unknown',1).includes('undefined'));
});
test('learning UI: delayed response cannot cross a user switch or logout',async()=>{
 let actor='first',resolveFirst;const ui=create({client:{rpc:async()=>actor==='first'?new Promise(r=>resolveFirst=r):({data:{...model(),user_id:'second',competence:77},error:null})},actor:()=>actor,redraw:()=>{},enabled:()=>true});
 ui.ensureSession();actor='second';ui.ensureSession();await tick();resolveFirst({data:{...model(),user_id:'first',competence:999},error:null});await tick();assert(ui.render().includes('77'));assert(!ui.render().includes('999'));
 ui.reset();assert(!ui.render().includes('77'));
});
test('learning UI: failed transport reuses request id, never sends client score/role authority',async()=>{
 const calls=[],handlers={};let lose=true;
 const client={rpc:async(name,params)=>{if(name==='learning_command'){calls.push(params);if(lose){lose=false;return{error:{code:'08006'}};}}return{data:model(),error:null};}};
 const ui=create({client,actor:()=> 'person',redraw:()=>{},enabled:()=>true});ui.ensureSession();await tick();ui.bind({addEventListener:(name,fn)=>handlers[name]=fn});
 const click=()=>handlers.click({target:{closest:()=>({dataset:{learn:'offer'}})}});click();await tick();click();await tick();assert.equal(calls.length,2);assert.equal(calls[0].p_request,calls[1].p_request);assert.deepEqual(calls[0].p_body,{});assert(!JSON.stringify(calls).includes('role'));
});
test('learning UI: authorized support topics remain distinct from participation and are escaped',async()=>{
 const handlers={},v={...model(),departments:[{id:'own',name:'Own'}]};
 const admin={people:[{id:'peer',name:'Peer',offered:3,answered:1,deferrals:2,participation:0,enabled:true,support_topics:['<topic>']}],questions:[],instructions:[]};
 const ui=create({client:{rpc:async n=>({data:n==='learning_admin_view'?admin:v,error:null})},actor:()=> 'person',redraw:()=>{},enabled:()=>true});
 ui.ensureSession();await tick();ui.bind({addEventListener:(n,fn)=>handlers[n]=fn});
 await handlers.change({target:{matches:s=>s==='[data-learn-department]',value:'own'}});
 const html=ui.render();assert(html.includes('Wiederholungsbedarf'));assert(html.includes('&lt;topic&gt;'));assert(!html.includes('<topic>'));assert(html.includes('Teilnahmeübersicht'));
});
