const {test,expect}=require('@playwright/test'),{randomUUID}=require('node:crypto');const {prepareMockPage,appUrl}=require('./helpers/mock-page');
const {setupDatabase,worker,reviewer,machine,machineA2,foreignMachine}=require('./helpers/setup-db');
test.describe('Machine readiness Phase 3 setup times',()=>{
 let api,uid,snapshot,calls,failNext;
 test.beforeEach(async({page})=>{
  api=await setupDatabase();snapshot=await api.seed();uid=worker;calls=[];failNext=false;
  await api.rpc(worker,'op_routing_command',{p_operation_id:snapshot.op_rows[0].id,p_action:'select',p_request_id:randomUUID(),p_expected_version:0,p_machine_id:machine,p_reason:'Synthetic adoption'});
  await page.exposeFunction('__setupRpc',async(name,params)=>{calls.push({name,params});if(failNext&&name==='setup_time_command'){failNext=false;return{error:{code:'08006'}};}try{return{data:await api.rpc(uid,name,params),error:null};}catch(e){return{data:null,error:{code:e.code}};}});
  await prepareMockPage(page);await page.addInitScript(()=>{const make=window.supabase.createClient;window.supabase.createClient=(...args)=>({...make(...args),rpc:(name,params)=>window.__setupRpc(name,params)});});
 });
 test.afterEach(async()=>api?.db.close());
 async function login(page,{enabled=true}={}){await page.goto(appUrl);await page.getByRole('button',{name:'Mido/Admin'}).click();await page.evaluate(({uid,enabled})=>{window.HUMBEL_ARTICLE_MASTER_ENABLED=true;window.HUMBEL_OPERATIONS_ENABLED=true;window.HUMBEL_MACHINE_READINESS_ENABLED=true;window.HUMBEL_OP_ROUTING_ENABLED=true;window.HUMBEL_SETUP_TIMES_ENABLED=enabled;currentSupabaseUser={id:uid};currentUser.role='employee';render();},{uid,enabled});await page.locator('#tabs').getByRole('button',{name:'Produktion',exact:true}).click();await page.getByRole('button',{name:'OP-Maschinenübersicht',exact:true}).click();}
 const card=p=>p.getByRole('article',{name:'Rüstung BA 812200 OP 10',exact:true});
 const command=(action,run,extra={})=>api.rpc(uid,'setup_time_command',{p_action:action,p_request_id:randomUUID(),p_operation_id:snapshot.op_rows[0].id,p_run_id:run,...extra});
 async function begin(p,mode='Intern rüsten'){await card(p).getByRole('button',{name:'Rüsten beginnen',exact:true}).click();await card(p).getByRole('button',{name:mode,exact:true}).click();await expect(p.locator('#setup-personal [role=status]')).toContainText('bestätigt');}
 test('explicit start then mode switch, personal stop and shared close; responsive layout and unchanged counts',async({page})=>{
  await login(page);expect(calls.filter(c=>c.name==='setup_time_command')).toHaveLength(0);await begin(page);await card(page).getByRole('button',{name:'Extern rüsten',exact:true}).click();await expect(card(page)).toContainText('Extern');await card(page).getByRole('button',{name:'Eigene Rüstzeit beenden',exact:true}).click();await expect(card(page)).toContainText('Personalaufwand');await card(page).getByRole('button',{name:'Rüsten beenden',exact:true}).click();await expect(card(page).getByRole('button',{name:'Rüsten beenden',exact:true})).toHaveCount(0);
  expect((await api.db.query('select count(*)::int n from public.production_station_counts')).rows[0].n).toBe(1);expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1)).toBe(true);
 });
 test('general pause outside setup mask; return question no starts nothing; yes requires mode',async({page})=>{
  await login(page);await begin(page);await page.locator('#tabs').getByRole('button',{name:'Dashboard',exact:true}).click();await page.getByRole('button',{name:'Zur Pause abmelden',exact:true}).click();await page.getByRole('button',{name:'Aus Pause zurückmelden',exact:true}).click();await expect(page.getByRole('region',{name:'Rüstung fortsetzen'})).toBeVisible();await page.getByRole('button',{name:'Nein, später',exact:true}).click();
  expect((await api.db.query('select count(*)::int n from public.production_operation_activity_segments where ended_at is null')).rows[0].n).toBe(0);
  await page.getByRole('button',{name:'Zur Pause abmelden',exact:true}).click();await page.getByRole('button',{name:'Aus Pause zurückmelden',exact:true}).click();await page.getByRole('button',{name:'Ja, Modus auswählen',exact:true}).click();await page.getByRole('region',{name:'Rüstung fortsetzen'}).getByRole('button',{name:'Extern rüsten',exact:true}).click();await expect(page.locator('#setup-personal [role=status]')).toContainText('bestätigt');
  expect((await api.db.query('select count(*)::int n from public.production_operation_activity_segments where ended_at is null')).rows[0].n).toBe(1);
 });
 test('shared close from another participant removes resume offer; personal pause retains colleague',async({page})=>{
  await login(page);await begin(page);const rid=(await api.db.query('select id from public.production_setup_runs')).rows[0].id;
  await api.rpc(reviewer,'setup_time_command',{p_action:'start',p_request_id:randomUUID(),p_operation_id:snapshot.op_rows[0].id,p_run_id:rid,p_mode:'external'});
  await page.getByRole('button',{name:'Zur Pause abmelden',exact:true}).click();await expect(page.locator('#setup-personal [role=status]')).toContainText('bestätigt');expect((await api.db.query('select user_id from public.production_operation_activity_segments where ended_at is null')).rows[0].user_id).toBe(reviewer);
  await api.rpc(reviewer,'setup_time_command',{p_action:'finish',p_request_id:randomUUID(),p_operation_id:snapshot.op_rows[0].id,p_run_id:rid});await page.getByRole('button',{name:'Aus Pause zurückmelden',exact:true}).click();await expect(page.getByRole('region',{name:'Rüstung fortsetzen'})).toHaveCount(0);
 });
 test('correction stores reason and immutable old/new audit; failure preserves input for retry',async({page})=>{
  await login(page);await begin(page);await card(page).getByRole('button',{name:'Eigene Rüstzeit beenden',exact:true}).click();await card(page).getByText('Rüstzeit korrigieren',{exact:true}).click();const f=card(page).locator('form[data-setup-correction]');await f.getByLabel('Rüstmodus',{exact:true}).selectOption('setup_external');await f.getByLabel('Pflichtbegründung').fill('Synthetic audited correction');failNext=true;await f.getByRole('button',{name:'Korrektur speichern'}).click();await expect(page.locator('#setup-personal [role=alert]')).toContainText('nicht bestätigt');
  await card(page).getByText('Rüstzeit korrigieren',{exact:true}).click();await expect(f.getByLabel('Pflichtbegründung')).toHaveValue('Synthetic audited correction');await f.getByRole('button',{name:'Korrektur speichern'}).click();await expect(card(page)).toContainText('Korrekturen: 1');expect((await api.db.query('select activity from public.production_operation_activity_segments')).rows[0].activity).toBe('setup_internal');
 });
 test('actual logout fails closed on transport failure, retry ends personal timer; flags off preserve legacy',async({page})=>{
  await login(page);await begin(page);failNext=true;await page.locator('#tabs').getByRole('button',{name:'Abmelden',exact:true}).click();await expect(page.locator('#setup-personal [role=alert]')).toContainText('nicht bestätigt');await expect(page.locator('#tabs')).toBeVisible();await page.locator('#tabs').getByRole('button',{name:'Abmelden',exact:true}).click();await expect(page.locator('#tabs')).toBeHidden();
  expect((await api.db.query('select count(*)::int n from public.production_operation_activity_segments where ended_at is null')).rows[0].n).toBe(0);await expect(page.locator('#setup-personal')).toHaveCount(0);
  calls=[];await login(page,{enabled:false});await expect(page.locator('#setup-times')).toHaveCount(0);expect(calls.some(c=>c.name.startsWith('setup_time_'))).toBe(false);
 });
 test('workplace takeover ends predecessor only and asks successor to explicitly choose a mode',async({page})=>{
  const r=await api.rpc(reviewer,'setup_time_command',{p_action:'start',p_request_id:randomUUID(),p_operation_id:snapshot.op_rows[0].id,p_mode:'external'});
  await login(page);await card(page).getByRole('button',{name:'Rüsten beginnen',exact:true}).click();await expect(page.getByRole('region',{name:'Rüstung fortsetzen'})).toBeVisible();await page.getByRole('button',{name:'Nein, später',exact:true}).click();await card(page).getByRole('button',{name:/Arbeitsplatz übernehmen von/}).click();await expect(page.getByRole('region',{name:'Rüstung fortsetzen'})).toBeVisible();
  expect((await api.db.query('select count(*)::int n from public.production_operation_activity_segments where ended_at is null')).rows[0].n).toBe(0);
  await page.getByRole('button',{name:'Ja, Modus auswählen',exact:true}).click();await page.getByRole('region',{name:'Rüstung fortsetzen'}).getByRole('button',{name:'Extern rüsten',exact:true}).click();
  await expect(page.locator('#setup-personal [role=status]')).toContainText('bestätigt');
  const active=(await api.db.query('select user_id,setup_run_id from public.production_operation_activity_segments where ended_at is null')).rows;expect(active).toHaveLength(1);expect(active[0].user_id).toBe(worker);expect(active[0].setup_run_id).toBe(r.run_id);
 });
 test('own historical setup stays reachable and correctable after cross-department transfer',async({page})=>{
  const r=await command('start',null,{p_mode:'external'});await command('finish',r.run_id);const op=snapshot.op_rows[0].id,route=(await api.rpc(worker,'op_routing_view')).rows.find(x=>x.operation_id===op);
  await api.rpc(reviewer,'op_routing_command',{p_operation_id:op,p_action:'fix',p_request_id:randomUUID(),p_expected_version:route.row_version,p_machine_id:foreignMachine,p_reason:'Synthetic transfer before historical correction'});
  await login(page);await expect(card(page)).toBeVisible();await card(page).getByText('Rüstzeit korrigieren',{exact:true}).click();const f=card(page).locator('form[data-setup-correction]');await f.getByLabel('Rüstmodus',{exact:true}).selectOption('setup_internal');await f.getByLabel('Pflichtbegründung').fill('Synthetic historical correction');await f.getByRole('button',{name:'Korrektur speichern'}).click();await expect(card(page)).toContainText('Korrekturen: 1');
 });
 test('machine switch aborts open run with configured reason, next machine requires explicit setup start',async({page})=>{
  await login(page);await begin(page);const c=page.getByRole('article',{name:'Maschinenzuordnung BA 812200 OP 010'});await c.getByLabel('Zielmaschine',{exact:true}).selectOption(machineA2);await c.getByLabel('Begründung',{exact:true}).fill('Synthetic move during setup');await c.getByLabel('Rüst-Abbruchgrund').selectOption('synthetic-machine-change');await c.getByRole('button',{name:'Maschine wählen',exact:true}).click();await expect(page.locator('#op-routing [role=status]')).toContainText('Audit gespeichert');await page.getByRole('button',{name:'Rüstzeiten neu laden',exact:true}).click();await expect(page.locator('#setup-times')).toContainText('Alternative A');expect((await api.db.query('select count(*)::int n from public.production_operation_activity_segments where ended_at is null')).rows[0].n).toBe(0);
 });
});
