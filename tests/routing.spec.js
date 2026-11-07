const {test,expect}=require('@playwright/test');const {prepareMockPage,appUrl}=require('./helpers/mock-page');
const {routingDatabase,worker,reviewer,leaderB,machine,machineA2,foreignMachine,department}=require('./helpers/routing-db');
test.describe('Machine readiness Phase 2 routing',()=>{
 let api,uid,snapshot,calls,failNext;
 test.beforeEach(async({page})=>{
  api=await routingDatabase();snapshot=await api.seed();uid=worker;calls=[];failNext=false;
  await page.exposeFunction('__routingRpc',async(name,params)=>{calls.push({name,params});if(failNext&&name==='op_routing_command'){failNext=false;return{error:{code:'08006'}};}try{return{data:await api.rpc(uid,name,params),error:null};}catch(e){return{data:null,error:{code:e.code}};}});
  await prepareMockPage(page);await page.addInitScript(()=>{const make=window.supabase.createClient;window.supabase.createClient=(...args)=>({...make(...args),rpc:(name,params)=>window.__routingRpc(name,params)});});
 });
 test.afterEach(async()=>api?.db.close());
 async function login(page,{user=worker,role='employee',enabled=true}={}){
  uid=user;await page.goto(appUrl);await page.getByRole('button',{name:'Mido/Admin'}).click();
  await page.evaluate(({user,role,enabled})=>{window.HUMBEL_ARTICLE_MASTER_ENABLED=true;window.HUMBEL_OPERATIONS_ENABLED=true;window.HUMBEL_BA_ENTRY_ENABLED=true;window.HUMBEL_MACHINE_READINESS_ENABLED=true;window.HUMBEL_OP_ROUTING_ENABLED=enabled;currentSupabaseUser={id:user};currentUser.role=role;render();},{user,role,enabled});
  await page.locator('#tabs').getByRole('button',{name:'Produktion',exact:true}).click();
  await page.getByRole('button',{name:'OP-Maschinenübersicht',exact:true}).click();
 }
 const card=page=>page.getByRole('article',{name:'Maschinenzuordnung BA 812200 OP 010'});
 async function choose(page,target,action='Maschine wählen'){
  const c=card(page);await c.getByLabel('Zielmaschine',{exact:true}).selectOption(target);await c.getByLabel('Begründung',{exact:true}).fill('Synthetic reviewed machine move');await c.getByRole('button',{name:action,exact:true}).click();
 }
 test('article default preselected; suitable own-department switch retains existing readiness controls and immutable snapshots',async({page})=>{
  await login(page);const c=card(page);await expect(c).toBeVisible();await expect(c.getByLabel('Zielmaschine',{exact:true})).toHaveValue(machine);
  await expect(c.getByLabel('Zielmaschine',{exact:true}).locator('option[value="'+foreignMachine+'"]')).toHaveCount(0);
  await choose(page,machineA2);await expect(page.locator('#op-routing [role=status]')).toContainText('Audit gespeichert');await expect(c.getByLabel('Zielmaschine',{exact:true})).toHaveValue(machineA2);
  await expect(page.locator('#machine-readiness')).toBeVisible();await expect(page.getByRole('button',{name:'Physische Übergabe bestätigen',exact:true})).toBeVisible();
  expect((await api.db.query('select operations from public.production_order_article_snapshots where order_id=$1',[snapshot.order_id])).rows[0].operations[0].machine_id).toBe(machine);
  expect((await api.db.query('select count(*)::int n from public.production_station_counts')).rows[0].n).toBe(1);
 });
 test('AL binds across departments; receiving AL gets in-app notice and can return directly without BA responsibility change',async({page})=>{
  await login(page,{user:reviewer,role:'department_admin'});await choose(page,foreignMachine,'Verbindlich festlegen (AL)');await expect(page.locator('#op-routing [role=status]')).toContainText('AL wurden innerhalb der App benachrichtigt');
  await expect(card(page)).toContainText('Verbindliche BA-OP-Maschine');await expect(card(page).getByRole('button',{name:'Verbindlich festlegen (AL)',exact:true})).toHaveCount(0);
  await page.evaluate(()=>logoutSupabase());await login(page,{user:leaderB,role:'department_admin'});await expect(page.getByRole('region',{name:'AL-Benachrichtigungen'})).toContainText('Operative Verantwortung übernommen');
  await choose(page,machine,'An ursprüngliche Abteilung zurückgeben');await expect(card(page)).toContainText('Ausführende Abteilung: Fertigung');
  expect((await api.db.query('select department_id from public.production_orders where id=$1',[snapshot.order_id])).rows[0].department_id).toBe(department);
 });
 test('configured setup abort mandatory; two personal segments close together and new machine starts no timer',async({page})=>{
  const op=snapshot.op_rows[0].id;
  await api.rpc(worker,'op_routing_command',{p_operation_id:op,p_action:'select',p_request_id:require('node:crypto').randomUUID(),p_expected_version:0,p_machine_id:machine,p_reason:'Synthetic initial adoption'});
  const section=(await api.db.query('select id from public.production_operation_machine_sections where operation_id=$1 and closed_at is null',[op])).rows[0].id;
  await api.db.query("insert into public.production_operation_activity_segments(section_id,user_id,activity,started_at) values($1,$2,'setup_internal',now()-interval '1 minute'),($1,$3,'setup_external',now()-interval '1 minute')",[section,worker,reviewer]);
  await login(page);await choose(page,machineA2);await expect(page.locator('#op-routing [role=alert]')).toContainText('Abbruchgrund');
  await card(page).getByLabel('Rüst-Abbruchgrund').selectOption('synthetic-machine-change');await card(page).getByRole('button',{name:'Maschine wählen',exact:true}).click();await expect(page.locator('#op-routing [role=status]')).toContainText('kein automatischer');
  const r=(await api.db.query('select ended_at from public.production_operation_activity_segments where section_id=$1',[section])).rows;expect(r[0].ended_at.getTime()).toBe(r[1].ended_at.getTime());
  expect((await api.db.query('select count(*)::int n from public.production_operation_activity_segments where ended_at is null')).rows[0].n).toBe(0);
 });
 test('transport failure retains inputs; reload/session switch clears sensitive state; mobile/tablet fit',async({page})=>{
  await login(page);failNext=true;await choose(page,machineA2);await expect(page.locator('#op-routing [role=alert]')).toContainText('Eingaben bleiben erhalten');await expect(card(page).getByLabel('Begründung')).toHaveValue('Synthetic reviewed machine move');
  expect(await page.locator('#op-routing').evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1)).toBe(true);
  await card(page).getByRole('button',{name:'Maschine wählen',exact:true}).click();await expect(page.locator('#op-routing [role=status]')).toContainText('Audit gespeichert');
  await page.evaluate(()=>logoutSupabase());await expect(page.locator('#op-routing')).toHaveCount(0);expect(await page.evaluate(()=>Object.values(localStorage).join(' '))).not.toContain('Synthetic reviewed machine move');
 });
 test('phase disabled preserves old readiness UI; snapshot button opens the exact OP using existing BA/OP route',async({page})=>{
  await login(page,{enabled:false});await expect(page.locator('#op-routing')).toHaveCount(0);expect(calls.some(c=>c.name==='op_routing_view')).toBe(false);
  await page.evaluate(()=>{window.HUMBEL_OP_ROUTING_ENABLED=true;render();});await expect(page.locator('#op-routing')).toBeVisible();
  await page.getByRole('button',{name:'Artikelstamm',exact:true}).first().click();await page.getByRole('button',{name:'Aufträge / BA',exact:true}).first().click();const p=page.getByRole('region',{name:'BA-Erstanlage'});
  await p.getByLabel('BA-/OP-Barcode',{exact:true}).fill('812200.010');await p.getByRole('button',{name:'Barcode öffnen',exact:true}).click();await expect(card(page)).toBeVisible();await expect(page.locator('#op-routing .routing-card')).toHaveCount(1);
 });
});
