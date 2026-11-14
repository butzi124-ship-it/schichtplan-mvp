#!/usr/bin/env node
// Real app screenshots; never navigates to a hosted/live app.
const {chromium}=require('@playwright/test'),fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const {createPreview}=require('../preview/server'),topics=require('../../js/help-topics');
async function capture({output='artifacts/documentation',only=null}={}){
 const out=path.resolve(output);fs.mkdirSync(path.join(out,'screenshots'),{recursive:true});
 const passwords={admin:crypto.randomBytes(32).toString('hex'),leader:crypto.randomBytes(32).toString('hex'),employee:crypto.randomBytes(32).toString('hex')};
 const priorPath=path.join(out,'screenshots.json');const prior=fs.existsSync(priorPath)?JSON.parse(fs.readFileSync(priorPath)):null;
 const preview=await createPreview({passwords});let browser;const shots=[],failures=[],network=[];
 const version=fs.readFileSync('schichtplan_mvp.js','utf8').match(/const APP_VERSION = "([^"]+)"/)[1];
 try{browser=await chromium.launch({executablePath:process.env.HUMBEL_CHROMIUM_PATH||undefined,args:['--no-sandbox']});
 for(const profile of [{id:'desktop',width:1440,height:1000},{id:'phone',width:390,height:844},{id:'tablet',width:1024,height:768}]){
  const context=await browser.newContext({viewport:{width:profile.width,height:profile.height},locale:'de-CH',timezoneId:'Europe/Berlin'}),page=await context.newPage();page.setDefaultTimeout(10000);
  await page.route('**/*',r=>r.request().url().startsWith(preview.url+'/')?r.continue():(network.push(r.request().url()),r.abort()));
  await page.goto(preview.url);await page.locator('select[name=account]').selectOption('admin');await page.locator('input[name=password]').fill(passwords.admin);await page.getByRole('button',{name:'Anmelden',exact:true}).click();
  await page.waitForFunction(()=>typeof currentUser!=='undefined'&&currentUser?.name==='Test Administration');
  const chosen=topics.filter(t=>t.tab&&(!only||only.includes(t.id))&&(profile.id==='desktop'||['home','counts','learning'].includes(t.id)));
  for(const t of chosen){try{
   const fingerprint=crypto.createHash('sha256').update(JSON.stringify(t)+JSON.stringify(require('./index').current())).digest('hex');const cached=prior?.screenshots.find(x=>x.topic===t.id&&x.profile===profile.id&&x.version===version&&x.fingerprint===fingerprint);if(cached&&fs.existsSync(path.join(out,cached.image))&&crypto.createHash('sha256').update(fs.readFileSync(path.join(out,cached.image))).digest('hex')===cached.sha256){shots.push(cached);console.log('UNCHANGED',t.id,profile.id);continue;}

   await page.evaluate(t=>{setTab(t.tab);if(t.subtab){if(t.tab==='dashboard')setDashboardSubTab(t.subtab);if(t.tab==='personalverwaltung')setPersonnelManagementSubTab(t.subtab);if(t.tab==='produktion')setProductionSubTab(t.subtab);if(t.tab==='werkzeugverwaltung')setToolManagementSubTab(t.subtab);if(t.tab==='lagerverwaltung')setStorageManagementSubTab(t.subtab);if(t.tab==='scanner')setScannerSubTab(t.subtab);if(t.tab==='auswertung')setAnalyticsSubTab(t.subtab);}},t);
   await page.waitForFunction(()=>document.querySelector('#view')?.innerText.trim().length>10);await page.evaluate(()=>document.fonts.ready);await page.locator('#context-help-launcher').waitFor();
   if(t.id==='learning')await page.getByRole('heading',{name:'Fachbezogene Auszeichnungen',exact:true}).waitFor();
   const fields=await page.locator('#view input,#view select,#view textarea').evaluateAll(es=>es.map(e=>({name:e.getAttribute('aria-label')||e.labels?.[0]?.innerText?.trim()||e.placeholder||e.id||e.name||'',type:e.type,required:e.required})));const visibleText=await page.locator('#view').innerText();
   const image=`screenshots/${t.id}-${profile.id}.png`;await page.screenshot({path:path.join(out,image),fullPage:true});shots.push({fingerprint,topic:t.id,profile:profile.id,role:'admin',version,image,sha256:crypto.createHash('sha256').update(fs.readFileSync(path.join(out,image))).digest('hex'),fields,caption:t.caption,textExcerpt:visibleText.slice(0,500),source:'running loopback app; synthetic data; browser legacy fixture and isolated SQL RPC adapters',nativeAuth:false});
   console.log('CAPTURE',t.id,profile.id);
  }catch(e){failures.push({topic:t.id,profile:profile.id,status:'BLOCKED_NOT_TESTED',reason:e.message.split('\n')[0]});console.error('CAPTURE BLOCKED',t.id,profile.id);}}
  await context.close();
 }
 }catch(e){failures.push({status:'BLOCKED_NOT_TESTED',reason:'Browser runtime failed: '+e.message.split('\n')[0]});}finally{await browser?.close();await preview.close();}
 const manifest={version,generatedAt:new Date().toISOString(),screenshots:shots,failures,externalRequestsBlocked:network.length};fs.writeFileSync(path.join(out,'screenshots.json'),JSON.stringify(manifest,null,2)+'\n');return manifest;
}
if(require.main===module)capture({only:process.argv.includes('--smoke')?['home','learning']:null}).then(r=>{console.log('CAPTURE RESULT',r.screenshots.length,r.failures.length);if(r.failures.length)process.exitCode=1;}).catch(e=>{console.error(e);process.exit(1);});module.exports={capture};
