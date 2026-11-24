'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const os=require('node:os');
const path=require('node:path');
const crypto=require('node:crypto');
const {chromium}=require('@playwright/test');
const {createSharedDemoServer}=require('../../scripts/pages/shared-demo-server');

test('real Chromium desktop/mobile sign-in, master CRUD, shared persistence, scoped worker views',{timeout:120000},async()=>{
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'fabos-shared-browser-'));
 const passwords={admin:crypto.randomBytes(32).toString('hex'),leader:crypto.randomBytes(32).toString('hex'),employee:crypto.randomBytes(32).toString('hex')};
 let server,browser;
 const screenshots=path.join(__dirname,'../../artifacts/fabos-shared');fs.mkdirSync(screenshots,{recursive:true});
 try{
  server=await createSharedDemoServer({passwords,filePath:path.join(dir,'store.json'),online:false});
  browser=await chromium.launch({executablePath:process.env.HUMBEL_CHROMIUM_PATH||undefined,args:['--no-sandbox']});
  for(const viewport of [{width:1280,height:850},{width:390,height:844}]){
   const context=await browser.newContext({viewport,locale:'de-DE',timezoneId:'Europe/Berlin'});
   const page=await context.newPage();const errors=[];page.on('pageerror',error=>errors.push(error.message));
   await page.goto(server.url+'/');
   await page.locator('#account').selectOption('admin');
   await page.locator('#password').fill(passwords.admin);
   await page.getByRole('button',{name:'Anmelden',exact:true}).click();
   await page.waitForURL('**/fertigung.html');
   await page.waitForFunction(()=>window.__FABOS_COCKPIT_READY===true,{timeout:30000});
   assert.match(await page.locator('#roleLabel').innerText(),/FABOS Master/);
   assert.equal(await page.locator('[data-machine]').count(),48);
   await page.getByRole('button',{name:'Mitarbeiterverwaltung'}).click();
   assert.equal(await page.locator('[data-person-dept]').count(),50);
   await page.getByRole('button',{name:'Betriebsaufträge'}).click();
   assert.equal(await page.locator('#orderCards article').count(),50);
   await page.locator('#newOrderPanel summary').click();
   await page.locator('#newOrderForm [name="article"]').fill('Chromium-Testwelle');
   await page.locator('#newOrderForm [name="article_no"]').fill('CHROMIUM-'+viewport.width);
   await page.locator('#newOrderForm button[type="submit"]').click();
   await page.waitForFunction(expected=>document.querySelectorAll('#orderCards article').length===expected,{timeout:20000},viewport.width===1280?51:52);
   await page.reload();
   await page.waitForFunction(()=>window.__FABOS_COCKPIT_READY===true);
   await page.getByRole('button',{name:'Betriebsaufträge'}).click();
   assert.equal(await page.locator('#orderCards article').count(),viewport.width===1280?51:52);
   await page.getByRole('button',{name:'Mitarbeiterverwaltung'}).click();
   await page.screenshot({path:path.join(screenshots,viewport.width+'-master.png')});
   assert.deepEqual(errors,[]);
   await context.close();
   const workerContext=await browser.newContext({viewport,locale:'de-DE'});
   const worker=await workerContext.newPage();
   await worker.goto(server.url+'/');
   await worker.locator('#account').selectOption('employee');
   await worker.locator('#password').fill(passwords.employee);
   await worker.getByRole('button',{name:'Anmelden',exact:true}).click();
   await worker.waitForURL('**/fertigung.html');
   await worker.waitForFunction(()=>window.__FABOS_COCKPIT_READY===true);
   assert.equal(await worker.locator('[data-machine]').count(),6);
   assert.equal(await worker.locator('#approvalNav').isVisible(),false);
   assert.equal(await worker.locator('#peopleNav').isVisible(),false);
   await worker.locator('[data-machine]').first().click();
   assert.match(await worker.locator('#machineDetail').innerText(),/Arbeitsgang|Arbeitsanweisung|Belegung/);
   await worker.screenshot({path:path.join(screenshots,viewport.width+'-worker.png')});
   await workerContext.close();
  }
 }finally{await browser?.close();await server?.close();fs.rmSync(dir,{force:true,recursive:true});}
});
