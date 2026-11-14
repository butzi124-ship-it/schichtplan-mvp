const {test,expect}=require('@playwright/test');
const {prepareMockPage,appUrl}=require('./helpers/mock-page');
const {setup,ids}=require('./helpers/learning-db');
test.describe('Gamification real isolated SQL integration',()=>{
 let api,calls;
 test.beforeEach(async({page})=>{
  api=await setup();await api.ready();await api.safe('worker');await api.question();calls=[];
  // One synthetic DB connection: serialize auth contexts, no claim of native races.
  let chain=Promise.resolve();
  await page.exposeFunction('__learningRpc',(name,body={})=>{
   const work=async()=>{calls.push(name);try{let result;
    if(name==='learning_command')result=await api.call('worker',body.p_action,body.p_body,body.p_request);
    else if(name==='learning_view')result=await api.view('worker');
    else throw Object.assign(new Error('No fallback'),{code:'42501'});
    return{data:result,error:null};
   }catch(e){return{data:null,error:{code:e.code||'XX000'}};}};
   const p=chain.then(work);chain=p.catch(()=>{});return p;
  });
  await prepareMockPage(page);await page.addInitScript(()=>{const make=window.supabase.createClient;window.supabase.createClient=(...args)=>({...make(...args),rpc:(name,params)=>window.__learningRpc(name,params)});});
 });
 test.afterEach(async()=>api?.db.close());
 async function login(page,enabled=true){await page.goto(appUrl);await page.getByRole('button',{name:'Mido/Admin'}).click();await page.evaluate(({uid,enabled})=>{window.HUMBEL_LEARNING_UI_ENABLED=enabled;currentSupabaseUser={id:uid};currentUser.role='employee';render();},{uid:ids.worker,enabled});}
 test('integrated navigation, actual scoring, private history and responsive surface',async({page})=>{
  await login(page);await page.locator('#tabs').getByRole('button',{name:'Gamification',exact:true}).click();
  await expect(page.locator('.learning')).toContainText('Synthetic theoretical question');
  await page.getByLabel('A',{exact:true}).check();await page.getByLabel('C',{exact:true}).check();await page.getByRole('button',{name:'Antwort verbindlich abgeben'}).click();
  await expect(page.locator('.learning')).toContainText('Vorgang serverseitig bestätigt');
  expect((await api.view('worker')).competence).toBe(5);
  await expect(page.locator('.learning')).toContainText('A and C are correct');expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1)).toBe(true);
 });
 test('separate safety navigation survives personal game deactivation',async({page})=>{
  await api.call('leader','person_enabled',{department:ids.d1,person:ids.worker,enabled:false});const q=await api.question({kind:'safety'});await api.call('leader','instruction_assign',{department:ids.d1,person:ids.worker,question:q,reason:'Synthetic entry'});
  await login(page);await page.locator('#tabs').getByRole('button',{name:'Sicherheitsunterweisungen',exact:true}).click();await expect(page.locator('.learning')).toContainText('Theorieprüfung offen');await expect(page.locator('.learning')).toContainText('ersetzt keine praktische Unterweisung');
 });
 test('rollout off leaves legacy navigation unchanged and makes no learning RPC',async({page})=>{
  await login(page,false);await expect(page.locator('#tabs').getByRole('button',{name:'Gamification',exact:true})).toHaveCount(0);expect(calls).toHaveLength(0);
 });
});
