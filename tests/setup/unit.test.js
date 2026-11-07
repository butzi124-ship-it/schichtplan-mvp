const {test}=require('node:test'),assert=require('node:assert/strict');const {create}=require('../../js/setup-times');
test('planned shift ending renders reminder only, never calls a time command',async()=>{
 const flags=['HUMBEL_SETUP_TIMES_ENABLED','HUMBEL_OP_ROUTING_ENABLED','HUMBEL_MACHINE_READINESS_ENABLED','HUMBEL_OPERATIONS_ENABLED','HUMBEL_ARTICLE_MASTER_ENABLED'];for(const f of flags)globalThis[f]=true;
 const calls=[],api=create({client:{rpc:async(name)=>{calls.push(name);return{data:{rows:[]}};}},actor:()=> 'synthetic',role:()=> 'employee',redraw:()=>{},shiftEnd:()=>new Date(Date.now()-60000)});
 try{await api.ensureSession();assert.match(api.renderGeneral(),/Geplantes Schichtende/);assert.deepEqual(calls,['setup_time_view']);}finally{api.reset();for(const f of flags)delete globalThis[f];}
});
