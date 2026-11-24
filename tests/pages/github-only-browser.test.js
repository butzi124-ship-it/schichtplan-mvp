'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const http=require('node:http');
const path=require('node:path');
const {chromium}=require('@playwright/test');
const {build}=require('../../scripts/pages/build');

test('GitHub Pages fiction state survives browser tabs, saves and imports backups, and resets on desktop/mobile',{timeout:110000},async()=>{
 const {DIST}=build();
 const mime={'.html':'text/html; charset=utf-8','.js':'application/javascript; charset=utf-8','.css':'text/css; charset=utf-8'};
 const server=http.createServer((req,res)=>{
  try{
   const pathName=decodeURIComponent(new URL(req.url,'http://127.0.0.1').pathname);
   const file=path.resolve(DIST,'.'+(pathName==='/'?'/index.html':pathName));
   if(!file.startsWith(DIST+path.sep)||!fs.existsSync(file)||!fs.statSync(file).isFile()){res.writeHead(404);res.end('Not found');return;}
   res.writeHead(200,{'Content-Type':mime[path.extname(file)]||'application/octet-stream','Cache-Control':'no-store'});
   res.end(fs.readFileSync(file));
  }catch{res.writeHead(500);res.end('Error');}
 });
 await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
 const url='http://127.0.0.1:'+server.address().port;
 let browser;
 const artifact=path.join(__dirname,'../../artifacts/fabos-github-only');fs.mkdirSync(artifact,{recursive:true});
 try{
  browser=await chromium.launch({executablePath:process.env.HUMBEL_CHROMIUM_PATH||undefined,args:['--no-sandbox']});
  for(const viewport of [{width:1280,height:850},{width:390,height:844}]){
   const context=await browser.newContext({viewport,locale:'de-DE',timezoneId:'Europe/Berlin',acceptDownloads:true});
   const first=await context.newPage();const errors=[];first.on('pageerror',e=>errors.push(e.message));
   await first.goto(url+'/fertigung.html?role=master');
   await first.waitForFunction(()=>window.__FABOS_COCKPIT_READY===true);
   assert.equal(await first.locator('[data-machine]').count(),48);
   await first.getByRole('button',{name:'Betriebsaufträge'}).click();
   assert.equal(await first.locator('#orderCards article').count(),50);
   await first.locator('#newOrderPanel summary').click();
   await first.locator('#newOrderForm [name="article"]').fill('GitHub-Browser-Testwelle');
   await first.locator('#newOrderForm [name="article_no"]').fill('GITHUB-2026');
   await first.locator('#newOrderForm button[type="submit"]').click();
   assert.equal(await first.locator('#orderCards article').count(),51);
   assert.match(await first.locator('#storageStatus').innerText(),/gespeichert/);
   const second=await context.newPage();
   await second.goto(url+'/fertigung.html?role=master');
   await second.waitForFunction(()=>window.__FABOS_COCKPIT_READY===true);
   await second.getByRole('button',{name:'Betriebsaufträge'}).click();
   assert.equal(await second.locator('#orderCards article').count(),51);
   const downloadPromise=second.waitForEvent('download');
   await second.getByRole('button',{name:'Testdaten sichern'}).click();
   const download=await downloadPromise;
   const downloadPath=await download.path();
   const backup=fs.readFileSync(downloadPath,'utf8');
   assert.match(backup,/FABOS_GITHUB_PAGES_FICTITIOUS_BACKUP/);
   assert.equal(JSON.parse(backup).snapshot.data.orders.length,51);
   second.once('dialog',dialog=>dialog.accept());
   await second.getByRole('button',{name:'Testdaten zurücksetzen'}).click();
   await second.waitForFunction(()=>window.__FABOS_COCKPIT_READY===true);
   await second.getByRole('button',{name:'Betriebsaufträge'}).click();
   assert.equal(await second.locator('#orderCards article').count(),50);
   second.once('dialog',dialog=>dialog.accept());
   const importedPage=second.waitForEvent('load');
   await second.locator('#importDemo').setInputFiles({name:'fabos-backup.json',mimeType:'application/json',buffer:Buffer.from(backup)});
   await importedPage;
   await second.waitForFunction(()=>window.__FABOS_COCKPIT_READY===true);
   await second.getByRole('button',{name:'Betriebsaufträge'}).click();
   assert.equal(await second.locator('#orderCards article').count(),51);
   await second.goto(url+'/index.html');
   assert.match(await second.locator('main').innerText(),/keine echte Benutzeranmeldung|kein passwortgeschützter Login/i);
   await second.screenshot({path:path.join(artifact,viewport.width+'-github-pages.png')});
   assert.deepEqual(errors,[]);
   await context.close();
  }
 }finally{await browser?.close();await new Promise(resolve=>server.close(resolve));}
});
