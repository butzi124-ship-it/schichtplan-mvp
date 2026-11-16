'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const http=require('node:http');
const path=require('node:path');
const {chromium}=require('@playwright/test');
const {build}=require('../../scripts/pages/build');

test('GitHub Pages release uses full FABOS UI, coherent synthetic data and no demo banners', {timeout:100000},async()=>{
 const {DIST,data,rows}=build();
 assert.equal(data.operations.length,176);
 assert.equal(rows.departments.length,8);
 assert.equal(rows.employees.length,41);
 assert.equal(rows.production_machines.length,19);
 assert.equal(rows.tools.length,20);
 const mock=fs.readFileSync(path.join(DIST,'pages-bootstrap.js'),'utf8');
 assert(!mock.includes('sb_publishable_')&&!mock.includes('.supabase.co'));
 const mime={'.html':'text/html; charset=utf-8','.js':'application/javascript; charset=utf-8','.css':'text/css','.png':'image/png','.svg':'image/svg+xml','.json':'application/json'};
 const server=http.createServer((req,res)=>{
  try{
   const pathName=decodeURIComponent(new URL(req.url,'http://localhost').pathname);
   const target=path.resolve(DIST,'.'+(pathName==='/'?'/index.html':pathName));
   if(!target.startsWith(DIST+path.sep)||!fs.existsSync(target)||!fs.statSync(target).isFile()){res.writeHead(404);res.end('Not found');return;}
   res.writeHead(200,{'Content-Type':mime[path.extname(target)]||'application/octet-stream','Cache-Control':'no-store'});
   res.end(fs.readFileSync(target));
  }catch(e){res.writeHead(500);res.end('Error');}
 });
 await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
 let browser;
 try{
  browser=await chromium.launch({executablePath:process.env.HUMBEL_CHROMIUM_PATH||undefined,args:['--no-sandbox']});
  for(const viewport of [{width:1280,height:850},{width:390,height:844}]){
   const page=await browser.newPage({viewport,locale:'de-DE',timezoneId:'Europe/Berlin'});
   const errors=[];page.on('pageerror',e=>errors.push(e.message));
   try{
    await page.goto('http://127.0.0.1:'+server.address().port+'/app.html?role=admin',{waitUntil:'domcontentloaded'});
    await page.waitForFunction(()=>window.__FABOS_PAGES_READY===true||Boolean(window.__FABOS_PAGES_ERROR),{timeout:20000});
    const result=await page.evaluate(()=>({
      ready:window.__FABOS_PAGES_READY===true,error:window.__FABOS_PAGES_ERROR||null,
      role:currentUser&&currentUser.role,employee:currentEmployeeRecord&&currentEmployeeRecord.id,
      departments:state.departments.length,employees:state.employeesList.length,
      machines:state.productionMachines.length,orders:state.productionOrders.length,
      tools:state.tools.length,bodyText:document.body.innerText.slice(0,5000),
      loginVisible:document.getElementById('loginBox')&&getComputedStyle(document.getElementById('loginBox')).display!=='none',
      banner:!!document.querySelector('#isolated-preview-banner')
    }));
    assert(result.ready,JSON.stringify({result,errors}));
    assert.equal(result.role,'admin');
    assert(result.employee);
    assert.equal(result.departments,8);assert.equal(result.employees,41);
    assert.equal(result.machines,19);assert(result.orders>=8);assert.equal(result.tools,20);
    assert.equal(result.loginVisible,false);assert.equal(result.banner,false);
    assert(!/SYNTHETISCHE ENTWICKLUNGSANSICHT|nur Demo|nicht dauerhaft gespeichert/i.test(result.bodyText));
    assert(errors.length===0,JSON.stringify(errors));
    await page.evaluate(()=>{setTab('produktion');setProductionSubTab('departments');});
    const deptNames=await page.locator('#view input[id^="production-department-name-"]').evaluateAll(es=>es.map(e=>e.value));
    assert(deptNames.includes('Fräserei'),JSON.stringify(deptNames));
    await page.evaluate(()=>{setTab('produktion');setProductionSubTab('machines');});
    const machineNames=await page.locator('#view input[id^="production-machine-name-"]').evaluateAll(es=>es.map(e=>e.value));
    assert(machineNames.includes('Hermle C62'),JSON.stringify(machineNames));
    await page.evaluate(()=>{setTab('produktion');setProductionSubTab('orders');});
    assert.match(await page.locator('#view').innerText(),/BA-8100/);
    await page.evaluate(()=>setTab('personalverwaltung'));
    const firstNames=await page.locator('#view input[id^="pm-employee-first-"]').evaluateAll(es=>es.map(e=>e.value));
    assert(firstNames.includes('Nina'),JSON.stringify(firstNames));
    await page.evaluate(()=>setTab('werkzeugverwaltung'));
    assert.match(await page.locator('#view').innerText(),/Schaftfräser/);
    await page.getByRole('link',{name:'Maschinenbelegung & OP-Plan'}).click();
    await page.waitForFunction(()=>window.__FABOS_COCKPIT_READY===true);
    assert.match(await page.locator('body').innerText(),/Maschinenübersicht/);
    assert((await page.locator('[data-machine]').count())>0);
    await page.getByRole('button',{name:'Betriebsaufträge'}).click();
    assert.match(await page.locator('body').innerText(),/BA-810001/);
    await page.goto('http://127.0.0.1:'+server.address().port+'/fertigung.html?role=lead');
    await page.waitForFunction(()=>window.__FABOS_COCKPIT_READY===true);
    assert.match(await page.locator('#roleLabel').innerText(),/Jonas Lindner/);
    assert.equal(await page.locator('#approvalView').isVisible(),true);
    assert((await page.locator('[data-approve]').count())>0);
    await page.locator('[data-approve]').first().click();
    assert.match(await page.locator('#approvalCards').innerText(),/Freigegeben/);
    await page.getByRole('button',{name:'Maschinenübersicht'}).click();
    await page.locator('[data-machine]').first().click();
    assert.match(await page.locator('#machineDetail').innerText(),/Arbeitsgang|Arbeitsanweisung/);
    await page.goto('http://127.0.0.1:'+server.address().port+'/fertigung.html?role=employee');
    await page.waitForFunction(()=>window.__FABOS_COCKPIT_READY===true);
    assert.match(await page.locator('#roleLabel').innerText(),/Mila Fischer/);
    assert.equal(await page.locator('#approvalNav').isVisible(),false);
    assert((await page.locator('[data-machine]').count())>0);
    await page.locator('[data-machine]').first().click();
    assert.match(await page.locator('#machineDetail').innerText(),/BA-|Arbeitsgang|Belegung/);
    await page.goto('http://127.0.0.1:'+server.address().port+'/');
    assert.equal(await page.getByRole('link',{name:/Abteilungsleiter/}).count(),1);
    assert.equal(await page.getByRole('link',{name:/Mitarbeiter/}).count(),1);
    assert(!/keine echte PlanOS|Demo-Plan|Synthetisch|fiktive Personen/i.test(await page.locator('body').innerText()));
    console.log('PAGES_BROWSER_PASS',JSON.stringify({viewport,departments:result.departments,employees:result.employees,machines:result.machines,orders:result.orders,tools:result.tools}));
   }finally{await page.close();}
  }
 }finally{await browser?.close();await new Promise(resolve=>server.close(resolve));}
});
