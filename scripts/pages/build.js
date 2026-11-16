'use strict';
// Build a static, synthetic, browser-local FABOS showcase. No Supabase or Node service in output.
const fs=require('node:fs');
const path=require('node:path');
const {rowsByTable,installSupabaseMock}=require('../../tests/helpers/supabase-mock');
const {buildDemoPreview,renderDemoPlan}=require('../preview/fabos-demo-adapter');
const ROOT=path.resolve(__dirname,'../..');
const DIST=path.join(ROOT,'dist-pages');
function copyDir(dir){
 const dest=path.join(DIST,dir);fs.mkdirSync(dest,{recursive:true});
 for(const item of fs.readdirSync(path.join(ROOT,dir),{withFileTypes:true})){
  const relative=path.join(dir,item.name),from=path.join(ROOT,relative),to=path.join(DIST,relative);
  if(item.isDirectory())copyDir(relative);
  else if(item.isFile())fs.copyFileSync(from,to);
 }
}
function hideInternalNames(text){
 for(const [real,fiction] of [['Lavdrim','Finn'],['Roger','Emil'],['Dashmir','Nico'],['Thomas','Lukas'],['Musa','Leon'],['Ardian','Tim']]){
  text=text.replace(new RegExp('\\b'+real+'\\b','g'),fiction);
 }
 return text;
}
function visibleName(s){return String(s||'').replace(/\bDEMO[ -]?/gi,'').trim();}
function build(){
 fs.rmSync(DIST,{force:true,recursive:true});fs.mkdirSync(DIST,{recursive:true});
 for(const directory of ['js','css','img'])copyDir(directory);
 const tailwind=path.join(DIST,'tests','fixtures');fs.mkdirSync(tailwind,{recursive:true});
 fs.copyFileSync(path.join(ROOT,'tests/fixtures/app-tailwind.css'),path.join(tailwind,'app-tailwind.css'));
 let app=fs.readFileSync(path.join(ROOT,'schichtplan_mvp.js'),'utf8');
 app=hideInternalNames(app);
 fs.writeFileSync(path.join(DIST,'schichtplan_mvp.js'),app);
 const {data,rows}=buildDemoPreview(rowsByTable);
 const actors={
  admin:{role:'admin',uid:'00000000-0000-4000-8000-000000000001',name:'FABOS Administration'},
  lead:{role:'department_admin',uid:'00000000-0000-4000-8000-000000000002',name:data.employees[5].display_name},
  employee:{role:'employee',uid:'00000000-0000-4000-8000-000000000003',name:data.employees[6].display_name}
 };
 const actor=actors.admin;
 rows.employees[0]={...rows.employees[0],auth_user_id:actor.uid,role:actor.role,employee_type:'admin',display_name:actor.name,first_name:'FABOS',last_name:'Administration',department_id:data.departments[0].id};
 for(const [key,employeeIndex] of [['lead',5],['employee',6]]){
  const target=rows.employees.find(e=>e.id===data.employees[employeeIndex].id);
  if(!target)throw Error('Demo persona missing: '+key);
  target.auth_user_id=actors[key].uid;
 }
 for(const machine of rows.production_machines)machine.name=visibleName(machine.name);
 for(const order of rows.production_orders)order.ba_number=visibleName(order.ba_number);
 for(const tool of rows.tools){tool.t_number=visibleName(tool.t_number);tool.article_no=visibleName(tool.article_no);}
 for(const machine of data.machines)machine.name=visibleName(machine.name);
 for(const order of data.orders)order.ba_number=visibleName(order.ba_number);
 for(const op of data.operations)op.machine=visibleName(op.machine);
 const mock=`(${installSupabaseMock.toString()})(${JSON.stringify(rows)});window.__PREVIEW_ACTORS=${JSON.stringify(actors)};window.__PREVIEW_ACTOR=${JSON.stringify(actor)};window.__PREVIEW_DEMO=true;window.__PREVIEW_ONLINE=true;const _make=window.supabase.createClient;window.supabase.createClient=(...args)=>({..._make(...args),rpc:async()=>({data:null,error:{code:'42501',message:'No shared database in this browser-only preview'}})});`;
 fs.writeFileSync(path.join(DIST,'pages-bootstrap.js'),mock);
 let html=fs.readFileSync(path.join(ROOT,'schichtplan_mvp.html'),'utf8');
 const replaced = [
  ['<script src="https://cdn.tailwindcss.com"></script>','<link rel="stylesheet" href="./tests/fixtures/app-tailwind.css">'],
  ['<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>','<script src="./pages-bootstrap.js"></script>'],
 ];
 for(const [before,after] of replaced){if(!html.includes(before))throw Error('Expected HTML dependency not found');html=html.replace(before,after);}
 html=html.replace(/window.SUPABASE_URL = "[^"]+";/,'window.SUPABASE_URL = "https://invalid.example";');
 html=html.replace(/window.SUPABASE_PUBLISHABLE_KEY =\s*"[^"]+";/,'window.SUPABASE_PUBLISHABLE_KEY = "browser-only";');
 html=html.replace(/(window.HUMBEL_(?:ARTICLE_MASTER|OPERATIONS|BA_ENTRY|MACHINE_READINESS|OP_ROUTING|SETUP_TIMES|OP_QUANTITIES|BA_LIFECYCLE|LEARNING_UI|CONTEXT_HELP)_ENABLED = )false/g,'$1true');
 html=html.replace('<title>Humbel-App</title>','<title>FABOS</title>').replace(/(<h1 class="text-2xl font-bold">\s*)Humbel-App/,'$1FABOS');
 html=html.replace('<section id="loginBox"','<section style="display:none!important" id="loginBox"');
 html=hideInternalNames(html);
 html=html.replace('</body>','<script src="./pages-boot.js"></script></body>');
 fs.writeFileSync(path.join(DIST,'app.html'),html);
 fs.copyFileSync(path.join(ROOT,'scripts/pages/role-login.html'),path.join(DIST,'index.html'));
 fs.writeFileSync(path.join(DIST,'schichtplan_mvp.html'),'<!doctype html><html lang="de"><meta charset="utf-8"><meta http-equiv="refresh" content="0;url=./index.html"><title>FABOS</title><p>FABOS-Testanmeldung: <a href="./index.html">Weiter zur Rollenwahl</a></p></html>');
 fs.copyFileSync(path.join(ROOT,'scripts/pages/fertigung.html'),path.join(DIST,'fertigung.html'));
 fs.copyFileSync(path.join(ROOT,'scripts/pages/demo-ui-client.js'),path.join(DIST,'demo-ui-client.js'));
 fs.writeFileSync(path.join(DIST,'demo-dataset.js'),'window.__FABOS_DEMO_DATA='+JSON.stringify(data)+';');
 fs.copyFileSync(path.join(ROOT,'scripts/pages/browser-boot.js'),path.join(DIST,'pages-boot.js'));
 let plan=renderDemoPlan(data);
 plan=plan.replace(/<aside>[\s\S]*?<\/aside>/,'');
 plan=plan.replace(/Synthetische Maschinenbelegung/g,'Maschinenbelegung').replace(/\(nur Demo\)/g,'').replace(/\(Demodaten\)/g,'').replace(/fiktive Personen/g,'Mitarbeiter');
 plan=plan.replace(/href="\/"/g,'href="./index.html"').replace(/Keine Original-Firmendaten\./g,'');
 fs.writeFileSync(path.join(DIST,'maschinenbelegung.html'),plan);
 fs.writeFileSync(path.join(DIST,'.nojekyll'),'');
 // Public Pages must not ship real Supabase configuration, user passwords or live personal names.
 for(const name of ['index.html','schichtplan_mvp.html','app.html','fertigung.html','demo-dataset.js','demo-ui-client.js','schichtplan_mvp.js','pages-bootstrap.js','pages-boot.js']){
  const data=fs.readFileSync(path.join(DIST,name),'utf8');
  if(/supabase\\.co|sb_publishable_|HUMBEL_PREVIEW_PASSWORDS|Lavdrim|Dashmir|Ardian/.test(data))throw Error('Forbidden production reference in published '+name);
 }
 console.log('PAGES_READY',JSON.stringify({departments:rows.departments.length,employees:rows.employees.length,machines:rows.production_machines.length,orders:rows.production_orders.length,tools:rows.tools.length,operations:data.operations.length,files:['index.html','app.html','fertigung.html','maschinenbelegung.html','pages-bootstrap.js']}));
 return{DIST,data,rows};
}
if(require.main===module)build();
module.exports={build};
