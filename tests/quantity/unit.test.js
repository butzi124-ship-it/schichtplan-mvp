const {test}=require('node:test'),assert=require('node:assert/strict');
const {create}=require('../../js/op-quantities');
const flags=['HUMBEL_OP_QUANTITIES_ENABLED','HUMBEL_OP_ROUTING_ENABLED','HUMBEL_MACHINE_READINESS_ENABLED','HUMBEL_OPERATIONS_ENABLED','HUMBEL_ARTICLE_MASTER_ENABLED'];
function enable(){for(const f of flags)globalThis[f]=true;}
test('adapter gates every RPC and clears data on actor/scanner changes',async()=>{
 enable();let actor='A',role='employee',calls=0;const app=create({actor:()=>actor,role:()=>role,redraw(){},client:{async rpc(){calls++;return{data:{rows:[{operation_id:'op',order_id:'ba',stations:[],finished_good:0}],notifications:[]}};}},renderEmployee(){},renderAmounts(){}});
 await app.ensureSession();assert.equal(calls,1);assert(app.hasAccess());actor='B';assert.equal(app.hasAccess(),false);assert.equal(app.render(),'');role='tool_scanner';await app.ensureSession();assert.equal(app.hasAccess(),false);assert.equal(calls,1);app.reset();
});
test('disabled counter adapter creates no network call or counter',async()=>{
 enable();globalThis.HUMBEL_OP_QUANTITIES_ENABLED=false;const app=create({actor:()=> 'A',role:()=> 'employee',redraw(){},client:{rpc(){throw Error('Unexpected RPC');}}});await app.ensureSession();assert.equal(app.render(),'');assert.equal(app.station('x'),null);app.reset();
});
test('no production persistence or independent quantity accumulator in adapter',()=>{
 const source=require('node:fs').readFileSync('js/op-quantities.js','utf8');assert(!source.includes('localStorage'));assert(!source.includes('.from('));assert(source.includes("rpc('op_quantity_command'"));assert(source.includes('renderEmployee(s,e,'));
});
