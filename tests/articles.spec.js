const { test, expect } = require('@playwright/test');
const { prepareMockPage, appUrl } = require('./helpers/mock-page');
const { articleDatabase,manager,worker,outsider,department,foreignDepartment,order } = require('./helpers/article-db');
const reference = require('./fixtures/article-reference.json');

test.describe('Article master',()=>{
 let api, uid, calls, failNext, delay;
 test.beforeEach(async({page})=>{
  api=await articleDatabase(); uid=manager; calls=[]; failNext=false; delay=0;
  await page.exposeFunction('__articleRpc',async(name,params)=>{
   calls.push({name,params});
   if(failNext){failNext=false;return{error:{code:'08006'}};}
   if(delay) await new Promise(resolve=>setTimeout(resolve,delay));
   try{return{data:await api.rpc(uid,name,params),error:null};}catch(e){return{data:null,error:{code:e.code}};}
  });
  await prepareMockPage(page);
  await page.addInitScript(()=>{
   const make=window.supabase.createClient;
   window.supabase.createClient=(...args)=>({...make(...args),rpc:(name,params)=>window.__articleRpc(name,params)});
  });
 });
 test.afterEach(async()=>{await api?.db.close();});
 async function login(page,{user=manager,role='admin',enabled=true}={}){
  uid=user;
  await page.goto(appUrl);
  await page.getByRole('button',{name:'Mido/Admin'}).click();
  await page.evaluate(({user,role,enabled})=>{
   window.HUMBEL_ARTICLE_MASTER_ENABLED=enabled; currentSupabaseUser={id:user}; currentUser.role=role; render();
  },{user,role,enabled});
 }
 async function open(page){
  await page.locator('#tabs').getByRole('button',{name:'Produktion',exact:true}).click();
  await page.getByRole('button',{name:'Artikelstamm',exact:true}).click();
  await expect(page.getByRole('heading',{name:'Artikelstamm',exact:true})).toBeVisible();
  await expect(page.getByRole('button',{name:'Suchen / Neu laden'})).toBeEnabled();
 }
 async function create(page,number='32505'){
  await page.getByRole('button',{name:'Neuer Artikel',exact:true}).click();
  await page.getByLabel('Artikelnummer',{exact:true}).fill(number);
  for(const [key,value] of Object.entries(reference.data)) await page.locator(`[name="${key}"]`).fill(value);
  await page.getByRole('button',{name:'Entwurf speichern',exact:true}).click();
  await expect(page.getByRole('status')).toContainText('Gespeichert');
 }
 test('disabled feature issues no article RPC and creates no broken menu',async({page})=>{
  await login(page,{enabled:false});
  await expect(page.locator('#versionFooter button')).toHaveText('Versionslog · v0.6.3');
  await page.locator('#tabs').getByRole('button',{name:'Produktion',exact:true}).click();
  await expect(page.getByRole('button',{name:'Artikelstamm',exact:true})).toHaveCount(0);
  expect(calls).toEqual([]);
 });
 test('verified reference creates reusable fields, drawing data and long notes without BA quantities',async({page})=>{
  await login(page);await open(page);await create(page);
  await page.getByRole('button',{name:'Entwurf bearbeiten'}).click();
  await page.getByLabel('Zeichnungsnummer',{exact:true}).fill('D-32505');
  await page.getByLabel('Zeichnungsrevision',{exact:true}).fill('A');
  const text='Technischer Hinweis <script>window.bad=true</script> & prüfen.\n'.repeat(50);
  await page.getByLabel('Technische Beschreibung / Bearbeitungshinweise',{exact:true}).fill(text);
  await page.getByRole('button',{name:'Entwurf speichern'}).click();
  await expect(page.getByRole('status')).toContainText('Gespeichert');
  await expect(page.getByRole('region',{name:'Artikeldetails'})).toContainText('D-32505');
  expect(await page.evaluate(()=>window.bad)).toBeUndefined();
  const [row]=await api.rpc(manager,'article_master_list',{p_query:'32505'});
  expect(row.data.technical_notes).toBe(text.trim());expect(row.data.drawing_revision).toBe('A');
  expect(row.data).not.toHaveProperty('ba_number');expect(row.data).not.toHaveProperty('ba_quantity');
  expect(JSON.stringify(row.data)).not.toMatch(/812150|"48"/);
  expect(await page.evaluate(()=>Object.values(localStorage).join(' '))).not.toContain('Spannflansch');
  expect(reference.provenance.complete_operation_plan).toBe(false);
 });
 test('release freezes contents, new revision copies and archive preserves linked BA and source',async({page})=>{
  await login(page);await open(page);await create(page);
  await page.getByRole('button',{name:'Freigeben',exact:true}).click();
  await expect(page.getByRole('region',{name:'Artikeldetails'})).toContainText('Freigegeben');
  await expect(page.getByRole('button',{name:'Entwurf bearbeiten'})).toHaveCount(0);
  const [first]=await api.rpc(manager,'article_master_list',{p_query:'32505'});
  await api.db.query('insert into public.article_master_ba_links values($1,$2)',[order,first.id]);
  await page.getByRole('button',{name:'Neue Revision',exact:true}).click();
  await expect(page.getByRole('region',{name:'Artikeldetails'})).toContainText('Revision 2');
  await page.getByRole('button',{name:'Entwurf bearbeiten'}).click();
  await expect(page.getByLabel('Material',{exact:true})).toHaveValue('90MnCrV8');
  await page.getByLabel('Material',{exact:true}).fill('Revision-2-Material');
  await page.getByRole('button',{name:'Entwurf speichern'}).click();
  await expect(page.getByRole('status')).toContainText('Gespeichert');
  await page.getByRole('button',{name:'Archivieren',exact:true}).click();
  await expect(page.getByRole('region',{name:'Artikeldetails'})).toContainText('Archiviert');
  const rows=await api.rpc(manager,'article_master_list',{p_query:'32505'});
  expect(rows.find(r=>r.id===first.id).data).toEqual(first.data);
  expect((await api.db.query('select ba_number,ba_quantity from public.production_orders')).rows).toEqual([{ba_number:'812150',ba_quantity:48}]);
 });
 test('normalized duplicate identity returns an error, preserves inputs and permits another scope only when authorized',async({page})=>{
  await login(page);await open(page);await create(page);
  await page.getByRole('button',{name:'Neuer Artikel',exact:true}).click();
  await page.getByLabel('Artikelnummer',{exact:true}).fill(' 32505 ');
  await page.getByLabel('Kunde',{exact:true}).fill('Eingabe bleibt');
  await page.getByRole('button',{name:'Entwurf speichern'}).click();
  await expect(page.getByRole('alert')).toContainText('bereits vorhanden');
  await expect(page.getByLabel('Kunde',{exact:true})).toHaveValue('Eingabe bleibt');
  expect((await api.db.query('select count(*)::int n from public.article_master_items')).rows[0].n).toBe(1);
 });
 test('assigned first OP employee can create own draft but cannot release or enter legacy production',async({page})=>{
  await login(page,{user:worker,role:'employee'});await open(page);await create(page);
  await expect(page.getByRole('button',{name:'Freigeben',exact:true})).toHaveCount(0);
  await expect(page.getByRole('button',{name:'Neue Revision',exact:true})).toHaveCount(0);
  await expect(page.getByRole('button',{name:'Stückzahl',exact:true})).toHaveCount(0);
  expect(await page.evaluate(()=>canAccessProduction())).toBe(false);
  expect((await api.rpc(worker,'article_master_list',{p_query:'32505'}))[0].created_by).toBe(worker);
  await page.getByRole('button',{name:'Neuer Artikel',exact:true}).click();
  await page.getByLabel('Artikelnummer',{exact:true}).fill('32506');
  await page.getByRole('button',{name:'Entwurf speichern'}).click();
  await expect(page.getByRole('alert')).toContainText('zugewiesenen Artikel');
 });
 for(const role of ['admin','department_admin','employee','tool_scanner']){
  test(`${role} without separate article grant has no article access`,async({page})=>{
   await login(page,{user:outsider,role});
   await expect.poll(()=>calls.filter(c=>c.name==='article_master_capabilities').length).toBe(role==='tool_scanner'?0:1);
   if(['admin','department_admin'].includes(role)) await page.locator('#tabs').getByRole('button',{name:'Produktion',exact:true}).click();
   await expect(page.getByRole('button',{name:'Artikelstamm',exact:true})).toHaveCount(0);
   expect(await page.evaluate(()=>articleMaster.hasAccess())).toBe(false);
  });
 }
 test('provisioned department responsible sees only own scope',async({page})=>{
  await login(page,{role:'department_admin'});await open(page);await create(page);
  await page.getByRole('button',{name:'Neuer Artikel',exact:true}).click();
  await expect(page.getByLabel('Gültigkeitsbereich').locator('option')).toHaveCount(1);
  const result=await api.rpc(manager,'article_master_command',{p_action:'create',p_department_id:department,p_article_number:'32506',p_data:{designation:'Zweiter Artikel'}});
  expect(result.article_number).toBe('32506');
  await expect(api.rpc(manager,'article_master_command',{p_action:'create',p_department_id:foreignDepartment,p_article_number:'foreign',p_data:{}})).rejects.toThrow();
 });
 test('stale edit and failed save preserve input without success or silent overwrite',async({page})=>{
  await login(page);await open(page);await create(page);
  const [row]=await api.rpc(manager,'article_master_list',{p_query:'32505'});
  await page.getByRole('button',{name:'Entwurf bearbeiten'}).click();
  await page.getByLabel('Material',{exact:true}).fill('Lokale Eingabe');
  await api.rpc(manager,'article_master_command',{p_action:'update',p_revision_id:row.id,p_expected_version:row.row_version,p_data:{...reference.data,material:'Anderer Bearbeiter'}});
  await page.getByRole('button',{name:'Entwurf speichern'}).click();
  await expect(page.getByRole('alert')).toContainText('inzwischen geändert');
  await expect(page.getByLabel('Material',{exact:true})).toHaveValue('Lokale Eingabe');
  await expect(page.getByRole('status')).toHaveCount(0);
  expect((await api.rpc(manager,'article_master_list',{p_query:'32505'}))[0].data.material).toBe('Anderer Bearbeiter');
 });
 test('network failure keeps draft; in-flight save disables repeated submission',async({page})=>{
  await login(page);await open(page);
  await page.getByRole('button',{name:'Neuer Artikel',exact:true}).click();
  await page.getByLabel('Artikelnummer',{exact:true}).fill('32505');
  await page.getByLabel('Material',{exact:true}).fill('90MnCrV8');
  failNext=true;
  await page.getByRole('button',{name:'Entwurf speichern'}).click();
  await expect(page.getByRole('alert')).toContainText('nicht verfügbar');
  await expect(page.getByLabel('Material',{exact:true})).toHaveValue('90MnCrV8');
  delay=500;
  await page.getByRole('button',{name:'Entwurf speichern'}).click();
  await expect(page.getByRole('button',{name:'Entwurf speichern'})).toBeDisabled();
  await expect(page.getByRole('status')).toContainText('Gespeichert');
  expect((await api.db.query('select count(*)::int n from public.article_master_revisions')).rows[0].n).toBe(1);
 });
 test('search, reload and revoked assignment return fresh data without local article storage',async({page})=>{
  await login(page,{user:worker,role:'employee'});await open(page);await create(page);
  await page.getByLabel('Artikelnummer suchen',{exact:true}).fill(' 325 ');
  await page.getByRole('button',{name:'Suchen / Neu laden'}).click();
  await expect(page.getByRole('list',{name:'Artikelrevisionen'})).toContainText('32505');
  await login(page,{user:worker,role:'employee'});await open(page);
  await expect(page.getByRole('list',{name:'Artikelrevisionen'})).toContainText('Spannflansch');
  expect(await page.evaluate(()=>Object.values(localStorage).join(' '))).not.toContain('Spannflansch');
  await api.db.query('update public.article_master_assignments set revoked=true where user_id=$1',[worker]);
  await page.getByRole('button',{name:'Suchen / Neu laden'}).click();
  await expect(page.getByRole('list',{name:'Artikelrevisionen'})).not.toContainText('32505');
  await login(page,{user:worker,role:'employee'});
  await expect.poll(()=>page.evaluate(()=>articleMaster.hasAccess())).toBe(false);
 });
 test('responsive editor has long fields and no horizontal page overflow',async({page})=>{
  await login(page);await open(page);
  await page.getByRole('button',{name:'Neuer Artikel',exact:true}).click();
  await expect(page.getByLabel('Spezifikation',{exact:true})).toHaveAttribute('rows','5');
  expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth)).toBe(true);
  await page.getByRole('button',{name:'Abbrechen',exact:true}).click();
  await expect(page.locator('#article-editor')).toHaveCount(0);
 });
});
