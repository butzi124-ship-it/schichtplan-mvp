const {test,expect}=require('@playwright/test'),{randomUUID}=require('node:crypto');
const {prepareMockPage,appUrl}=require('./helpers/mock-page');
const {quantityDatabase,worker,machine,machineA2}=require('./helpers/quantity-db');
test.describe('Machine readiness Phase 4 existing quantity counter',()=>{
 let api,snapshot,calls,loseResponse;
 test.beforeEach(async({page})=>{
  api=await quantityDatabase();snapshot=await api.seed();calls=[];loseResponse=false;
  const op=snapshot.op_rows[0].id;await api.rpc(worker,'op_routing_command',{p_operation_id:op,p_action:'select',p_request_id:randomUUID(),p_expected_version:0,p_machine_id:machine,p_reason:'Synthetic phase 4 selection'});
  await api.rpc(worker,'op_quantity_command',{p_operation_id:op,p_action:'configure',p_request_id:randomUUID(),p_new_count:3,p_note:'Synthetic confirmed plan'});
  await page.exposeFunction('__quantityRpc',async(name,params)=>{calls.push({name,params});try{const data=await api.rpc(worker,name,params);if(loseResponse&&name==='op_quantity_command'){loseResponse=false;return{error:{code:'08006'}};}return{data,error:null};}catch(e){return{data:null,error:{code:e.code}};}});
  await prepareMockPage(page);await page.addInitScript(()=>{const make=window.supabase.createClient;window.supabase.createClient=(...args)=>({...make(...args),rpc:(name,params)=>window.__quantityRpc(name,params)});});
 });
 test.afterEach(async()=>api?.db.close());
 async function login(page,enabled=true){await page.goto(appUrl);await page.getByRole('button',{name:'Mido/Admin'}).click();await page.evaluate(({worker,enabled})=>{window.HUMBEL_ARTICLE_MASTER_ENABLED=true;window.HUMBEL_OPERATIONS_ENABLED=true;window.HUMBEL_MACHINE_READINESS_ENABLED=true;window.HUMBEL_OP_ROUTING_ENABLED=true;window.HUMBEL_OP_QUANTITIES_ENABLED=enabled;currentSupabaseUser={id:worker};currentUser.role='employee';render();},{worker,enabled});if(enabled){await expect(page.locator('#op-quantities')).toHaveCount(0);await page.locator('#tabs').getByRole('button',{name:'Produktion',exact:true}).click();await page.evaluate(()=>setProductionSubTab('counts'));await expect(page.locator('#op-quantities')).toBeVisible();}}
 const tension=(page,n)=>page.getByRole('region',{name:`OP Spannung ${n}`,exact:true});
 const plus=(page,n)=>tension(page,n).getByRole('button',{name:'Gutteile +1 Mitarbeiter',exact:true});
 test('uses original employee buttons, final tension only and responsive employee access',async({page})=>{
  await login(page);await plus(page,1).click();await expect(page.locator('#op-quantities [role=status]')).toContainText('bestätigt');await expect(page.locator('[data-op-good]')).toHaveText('0');await plus(page,3).click();await expect(page.locator('[data-op-good]')).toHaveText('1');await expect(tension(page,3)).toContainText('Offener Mengenwiderspruch');
  expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1)).toBe(true);expect(calls.some(c=>c.name==='op_quantity_command'&&c.params.p_kind==='good')).toBe(true);
 });
 test('lost response retry retains request identity and counts once',async({page})=>{
  await login(page);loseResponse=true;await plus(page,3).click();await expect(page.locator('#op-quantities [role=alert]').first()).toContainText('nicht bestätigt');await plus(page,3).click();await expect(page.locator('#op-quantities [role=status]')).toContainText('bestätigt');await expect(page.locator('[data-op-good]')).toHaveText('1');const bookings=calls.filter(c=>c.name==='op_quantity_command');expect(bookings).toHaveLength(2);expect(bookings[0].params.p_request_id).toBe(bookings[1].params.p_request_id);
 });
 test('machine switch preserves OP amount and separates original counter events by section',async({page})=>{
  await login(page);await plus(page,3).click();await expect(page.locator('[data-op-good]')).toHaveText('1');const r=(await api.rpc(worker,'op_routing_view')).rows.find(r=>r.operation_id===snapshot.op_rows[0].id);await api.rpc(worker,'op_routing_command',{p_operation_id:r.operation_id,p_action:'select',p_request_id:randomUUID(),p_expected_version:r.row_version,p_machine_id:machineA2,p_reason:'Synthetic browser machine change'});
  await page.getByRole('button',{name:'OP-Mengen neu laden',exact:true}).click();await expect(page.locator('#op-quantities')).toContainText('Alternative A');await plus(page,3).click();await expect(page.locator('[data-op-good]')).toHaveText('2');expect((await api.db.query('select count(distinct section_id)::int n from public.production_quantity_attributions')).rows[0].n).toBe(2);
 });
 test('6M modal and original scrap labels survive; flag off makes no quantity RPC',async({page})=>{
  const id=randomUUID();await api.db.query("insert into public.production_qa_causes(id,group_id,group_label,reason_id,reason_label) values($1,'M','Mensch','AA','1AA')",[id]);await login(page);await page.evaluate(id=>{state.productionQaCauses=[{id,group_id:'M',group_label:'Mensch',reason_id:'AA',reason_label:'1AA',active:true}];render();},id);
  await tension(page,1).getByRole('button',{name:'Ausschuss +1',exact:true}).click();await page.getByRole('button',{name:'1AA',exact:true}).click();await page.getByRole('button',{name:'Ursache speichern',exact:true}).click();await expect(page.locator('#op-quantities [role=status]')).toContainText('bestätigt');await expect(tension(page,1)).toContainText('Ausschuss: 1');await expect(page.locator('[data-op-good]')).toHaveText('0');
  calls=[];await login(page,false);expect(calls.some(c=>c.name.startsWith('op_quantity_'))).toBe(false);await expect(page.locator('#op-quantities')).toHaveCount(0);
 });
});
