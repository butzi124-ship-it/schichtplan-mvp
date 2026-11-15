'use strict';
// Real Chromium UI acceptance: fixture counts alone are NOT a visible demo.
const {test}=require('node:test');
const assert=require('node:assert/strict');
const crypto=require('node:crypto');
const {chromium}=require('@playwright/test');
const {createPreview}=require('../../scripts/preview/server');

test('FABOS loads visible departments, staff, machines, orders and tools after login',{timeout:95000},async()=>{
 const before={demo:process.env.HUMBEL_PREVIEW_DEMO,online:process.env.HUMBEL_PREVIEW_ONLINE};
 process.env.HUMBEL_PREVIEW_DEMO='1';delete process.env.HUMBEL_PREVIEW_ONLINE;
 const passwords={admin:crypto.randomBytes(32).toString('hex'),leader:crypto.randomBytes(32).toString('hex'),employee:crypto.randomBytes(32).toString('hex')};
 const actors={
  admin:{role:'admin',uid:'00000000-0000-4000-8000-000000000001',name:'Demo Administration'},
  leader:{role:'department_admin',uid:'00000000-0000-4000-8000-000000000002',name:'Demo Abteilungsleitung'},
  employee:{role:'employee',uid:'00000000-0000-4000-8000-000000000003',name:'Demo Mitarbeiter'}
 };
 const backend={actors,async rpc(){throw Object.assign(Error('Read-only preview'),{code:'42501'});},async close(){}};
 let preview,browser;
 try{
  preview=await createPreview({passwords,backend});
  browser=await chromium.launch({executablePath:process.env.HUMBEL_CHROMIUM_PATH||undefined,args:['--no-sandbox']});
  const check=async(account)=>{
   const page=await browser.newPage({viewport:{width:1280,height:850},locale:'de-DE',timezoneId:'Europe/Berlin'});
   const errors=[];page.on('pageerror',error=>errors.push(error.message));
   try{
    await page.goto(preview.url,{waitUntil:'domcontentloaded'});
    await page.locator('select[name=account]').selectOption(account);
    await page.locator('input[name=password]').fill(passwords[account]);
    await page.getByRole('button',{name:'Anmelden',exact:true}).click();
    await page.locator('#isolated-preview-banner').waitFor({timeout:25000});
    const report=await page.evaluate(()=>({
      banner:document.querySelector('#isolated-preview-banner')?.innerText||'',
      departments:state.departments.length,
      employees:state.employeesList.length,
      machines:state.productionMachines.length,
      orders:state.productionOrders.length,
      tools:state.tools.length,
      actor:currentEmployeeRecord?.id||null,
      visibleDepartments:getVisibleProductionDepartments().length,
      visibleOrders:getVisibleProductionOrders().length,
      role:currentUser?.role||'',
    }));
    assert(!report.banner.includes('NICHT GELADEN'),JSON.stringify({report,errors}));
    assert.equal(report.departments,8,JSON.stringify({report,errors}));
    assert.equal(report.employees,41);
    assert.equal(report.machines,19);
    assert(report.orders>0);
    assert.equal(report.tools,20);
    assert(report.actor);
    assert.equal(report.role,actors[account].role);
    if(account==='admin'){
      assert.equal(report.visibleDepartments,8);
      await page.evaluate(()=>{setTab('produktion');setProductionSubTab('departments');});
      assert.match(await page.locator('#view').innerText(),/Fräserei/);
      assert.match(await page.locator('#view').innerText(),/Verzahnerei/);
      await page.evaluate(()=>{setTab('produktion');setProductionSubTab('machines');});
      assert.match(await page.locator('#view').innerText(),/Hermle C62/);
      await page.evaluate(()=>{setTab('produktion');setProductionSubTab('orders');});
      assert.match(await page.locator('#view').innerText(),/DEMO-BA-/);
      await page.evaluate(()=>setTab('personalverwaltung'));
      assert.match(await page.locator('#view').innerText(),/Nina Reuter/);
      await page.evaluate(()=>setTab('werkzeugverwaltung'));
      assert.match(await page.locator('#view').innerText(),/Schaftfräser/);
      await page.locator('#isolated-preview-banner a[href="/demo-plan"]').click();
      assert.match(await page.locator('body').innerText(),/Nächster Auftrag/);
      assert.match(await page.locator('body').innerText(),/DEMO-BA-810001/);
    }else{
      assert.equal(report.visibleDepartments,1,JSON.stringify({report,errors}));
      await page.evaluate(()=>{setTab('produktion');setProductionSubTab('departments');});
      assert.match(await page.locator('#view').innerText(),/Dreherei/);
    }
    console.log('DEMO_VISIBLE',account,JSON.stringify(report));
   }finally{await page.close();}
  };
  await check('admin');await check('leader');
 }finally{
  await browser?.close();await preview?.close();
  if(before.demo===undefined)delete process.env.HUMBEL_PREVIEW_DEMO;else process.env.HUMBEL_PREVIEW_DEMO=before.demo;
  if(before.online===undefined)delete process.env.HUMBEL_PREVIEW_ONLINE;else process.env.HUMBEL_PREVIEW_ONLINE=before.online;
 }
});
