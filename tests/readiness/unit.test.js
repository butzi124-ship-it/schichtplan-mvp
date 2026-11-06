const {test}=require('node:test');const assert=require('node:assert/strict');const {material,may,terminal,deadline}=require('../../js/machine-readiness');
const base={initialized:true,order_status:'running',work_status:'planned',material_status:'expected',machine_active:true,can_confirm:true,can_work:true,can_plan:false,predecessor_released:false,received_quantity:0,quantity:48};
test('four material indications carry distinct symbols and explicit forecast-versus-release text',()=>{
 assert.equal(material('planned').css,'yellow');assert.equal(material('expected').css,'alternating');assert.equal(material('ready').css,'blinking');assert.equal(material('started').css,'green');assert.match(material('expected').label,/keine Bearbeitungsfreigabe/);assert.equal(new Set(['planned','expected','ready','started'].map(s=>material(s).symbol)).size,4);
});
test('forecast never enables work; full predecessor release and physical receipt are independent requirements',()=>{
 assert.equal(may(base,'start'),false);assert.equal(may(base,'handover'),false);assert.equal(may({...base,predecessor_released:true},'handover'),true);assert.equal(may({...base,work_status:'ready',received_quantity:47},'start'),false);assert.equal(may({...base,work_status:'ready',received_quantity:48},'start'),true);
});
test('paused/terminal/unavailable/uninitialized operations fail closed without opening planning to operators',()=>{
 assert.equal(may({...base,machine_active:false,work_status:'running'},'pause'),true);assert.equal(may(base,'assign'),false);assert.equal(may({...base,can_plan:true},'assign'),true);for(const action of ['start','handover','resume','complete'])assert.equal(may({...base,machine_active:false},action),false);for(const state of ['completed','cancelled']){assert(terminal({...base,work_status:state}));for(const action of ['start','assign','handover','resume'])assert.equal(may({...base,work_status:state,can_plan:true},action),false);}assert.equal(may({...base,initialized:false},'handover'),false);
});
test('deadline facts stay separate and never manufacture a completion forecast',()=>{
 assert.match(deadline({due_date:'2026-11-05'},'2026-11-06'),/überschritten/);assert.match(deadline({due_date:'2099-01-01'},'2026-11-06'),/Prognose fehlt|prognose fehlt/);assert.equal(material('ready').css,'blinking');
});
