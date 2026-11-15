const {test,expect}=require('@playwright/test');
const {pathToFileURL}=require('node:url');
const path=require('node:path');
const {accessFixture,ids}=require('./security/access-fixture');
const url=pathToFileURL(path.resolve(__dirname,'fixtures/secure-access.html')).href;
// The existing matrix's Operations plan group includes every browser/device project.
test.describe('Operations plan secure access integration',()=>{
 let f;
 test.beforeEach(async({page})=>{
  f=await accessFixture();
  await page.route('**/*',route=>new URL(route.request().url()).protocol==='file:'?route.continue():route.abort());
  await page.exposeFunction('__sqlAccess',async({kind,name,args})=>kind==='auth'?f.client.auth.getUser():kind==='materials'?await f.client.from('tool_materials').select('id,name,sort_order,is_active').eq('is_active',true).order('name'):f.client.rpc(name,args));
  await page.addInitScript(()=>{
   window.__ISOLATED_SECURE_ACCESS__=true;window.__sessionListeners=new Set();
   window.testSecureClient={auth:{getUser:()=>window.__sqlAccess({kind:'auth'}),onAuthStateChange:fn=>{window.__sessionListeners.add(fn);return{data:{subscription:{unsubscribe:()=>window.__sessionListeners.delete(fn)}}};}},rpc:(name,args)=>window.__sqlAccess({kind:'rpc',name,args}),from:()=>{
    const builder={select:()=>builder,eq:()=>builder,order:()=>builder,then:(resolve,reject)=>window.__sqlAccess({kind:'materials'}).then(resolve,reject)};return builder;
   }};
  });
  await page.goto(url);
 });
 test.afterEach(async()=>f?.db.close());
 test('own contacts and minimal published schedule use actual isolated SQL without persistent private data',async({page})=>{
  await f.client.rpc('security_own_contacts',{p_values:{synthetic_contact:'Synthetic private contact'}});
  await page.getByRole('button',{name:'Eigene Kontakte laden'}).click();await expect(page.locator('#output')).toContainText('Synthetic private contact');
  await page.getByRole('button',{name:'Freigegebene Schichten laden'}).click();await expect(page.locator('#output')).toHaveText('[]');
  const stored=await page.evaluate(()=>[...Object.values(localStorage),...Object.values(sessionStorage)].join(' '));expect(stored).not.toContain('Synthetic private contact');
  expect(f.calls).toContain('security_actor_context');expect(f.calls).not.toContain('table:employees');expect(f.calls).not.toContain('table:planner_assignments');
 });
 test('session switch and offline invalidate cached contact data and rendered output',async({page,context})=>{
  await f.client.rpc('security_own_contacts',{p_values:{synthetic_contact:'Synthetic old actor contact'}});
  await page.getByRole('button',{name:'Eigene Kontakte laden'}).click();await expect(page.locator('#output')).toContainText('Synthetic old actor contact');
  f.switchUser(ids.other);await page.evaluate(()=>{for(const fn of window.__sessionListeners)fn('SIGNED_IN');});await expect(page.locator('#output')).toHaveText('');
  await page.getByRole('button',{name:'Eigene Kontakte laden'}).click();await expect(page.locator('#output')).toHaveText('{}');
  await context.setOffline(true);await page.evaluate(()=>window.dispatchEvent(new Event('offline')));await expect(page.locator('#output')).toHaveText('');
  expect(await page.evaluate(()=>access.getState().contacts)).toBe(null);
 });
 test('unverified identity is denied and scanner retains only permitted material references',async({page})=>{
  f.switchUser(ids.orphan);await page.getByRole('button',{name:'Identität prüfen'}).click();await expect(page.getByRole('status')).toHaveText('Zugriff abgewiesen.');
  f.switchUser(ids.scanner);await page.evaluate(()=>{for(const fn of window.__sessionListeners)fn('SIGNED_IN');});
  await f.db.query("insert into public.tool_materials(name) values('Synthetic scanner material')");
  await page.getByRole('button',{name:'Materialreferenzen laden'}).click();await expect(page.locator('#output')).toContainText('Synthetic scanner material');
  await page.getByRole('button',{name:'Freigegebene Schichten laden'}).click();await expect(page.getByRole('status')).toHaveText('Zugriff abgewiesen.');
 });
});
