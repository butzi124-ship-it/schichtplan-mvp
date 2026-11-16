#!/usr/bin/env node
// Local-only synthetic showcase. HTTPS/SSO and native DB acceptance remain gates.
const http=require('node:http'),fs=require('node:fs'),path=require('node:path'),crypto=require('node:crypto');
const {installSupabaseMock,rowsByTable}=require('../../tests/helpers/supabase-mock');
const {buildDemoPreview,renderDemoPlan}=require('./fabos-demo-adapter');
const ROOT=path.resolve(__dirname,'../..');
const csp="default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; font-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'";
function createOnlineDemoBackend(){
 const actors={
  admin:{role:'admin',uid:'00000000-0000-4000-8000-000000000001',name:'Demo Administration'},
  leader:{role:'department_admin',uid:'00000000-0000-4000-8000-000000000002',name:'Demo Abteilungsleitung'},
  employee:{role:'employee',uid:'00000000-0000-4000-8000-000000000003',name:'Demo Mitarbeiter'},
 };
 return {actors,async rpc(){throw Object.assign(Error('Online demonstration has no persistent RPC backend'),{code:'42501'});},async close(){}};
}
async function createPreview({passwords,backend,port=0}={}){
 const online=process.env.HUMBEL_PREVIEW_ONLINE==='1';
 const publicOrigin=online?process.env.HUMBEL_PREVIEW_ORIGIN:'';
 if(online){
  let parsed;
  try{parsed=new URL(publicOrigin);}catch{throw Error('Online preview requires HUMBEL_PREVIEW_ORIGIN HTTPS origin');}
  if(parsed.protocol!=='https:'||parsed.pathname!=='/'||parsed.search||parsed.hash||parsed.username||parsed.password||parsed.port||!parsed.hostname||parsed.hostname==='localhost'||parsed.hostname==='127.0.0.1')throw Error('Online preview requires clean public HTTPS origin');
  if(process.env.HUMBEL_PREVIEW_DEMO!=='1')throw Error('Online preview requires isolated demo dataset');
 }
 if(!passwords||Object.keys(passwords).some(k=>!['admin','leader','employee'].includes(k))||Object.values(passwords).some(x=>typeof x!=='string'||x.length<24)||!passwords.employee)throw Error('Explicit isolated test passwords of at least 24 characters required');
 backend=backend||(online?createOnlineDemoBackend():await require('./backend').createBackend());const sessions=new Map(),attempts=new Map();
 const demo=process.env.HUMBEL_PREVIEW_DEMO==='1'?buildDemoPreview(rowsByTable):null;
 const salt=crypto.randomBytes(32),hash=p=>crypto.scryptSync(p,salt,32),hashes=Object.fromEntries(Object.entries(passwords).map(([u,p])=>[u,hash(p)]));
 function send(res,status,body,type='text/html; charset=utf-8'){res.writeHead(status,{'Content-Type':type,'Content-Security-Policy':csp,'Cache-Control':'no-store','X-Content-Type-Options':'nosniff','Referrer-Policy':'no-referrer'});res.end(body);}
 const server=http.createServer(async(req,res)=>{try{
  if(online&&req.method==='GET'&&req.url==='/healthz')return send(res,200,'ok','text/plain; charset=utf-8');
  const origin=online?publicOrigin:`http://127.0.0.1:${server.address().port}`;if(req.headers.host!==new URL(origin).host)return send(res,403,'Host denied');
  if(online&&req.headers['x-forwarded-proto']!=='https')return send(res,403,'HTTPS required');
  const url=new URL(req.url,origin),documentation=url.pathname.startsWith('/documentation/'),cookies=req.headers.cookie||'',sid=cookies.match(/(?:^|; )humbel_preview=([a-f0-9]{64})(?:;|$)/)?.[1];const s=sessions.get(sid);
  if(s&&s.expires<Date.now()){sessions.delete(sid);return send(res,401,'Session expired');}
  if(req.method==='POST'&&req.headers.origin!==origin)return send(res,403,'Origin denied');
  let body={};if(req.method==='POST'){let raw='';for await(const chunk of req){raw+=chunk;if(Buffer.byteLength(raw)>32768)return send(res,413,'Body too large');}try{body=JSON.parse(raw);}catch{return send(res,400,'Invalid JSON');}}
  if(url.pathname==='/login'&&req.method==='POST'){
   const ip=req.socket.remoteAddress,old=attempts.get(ip);const a=old&&old.until>Date.now()?old:{count:0,until:Date.now()+600000};attempts.set(ip,a);if(++a.count>10)return send(res,429,'Too many attempts');
   if(typeof body.password!=='string'||body.password.length>200||!Object.hasOwn(hashes,body.account)||!crypto.timingSafeEqual(hash(body.password),hashes[body.account]))return send(res,401,'Invalid test credentials');
   const id=crypto.randomBytes(32).toString('hex');sessions.set(id,{account:body.account,expires:Date.now()+3600000});if(sid)sessions.delete(sid);
   res.setHeader('Set-Cookie',`humbel_preview=${id}; HttpOnly; SameSite=Strict; ${online?'Secure; ':''}Path=/; Max-Age=3600`);return send(res,200,'{}','application/json');
  }
  if(online&&documentation)return send(res,404,'Not found');
  if(documentation)url.pathname=url.pathname.slice('/documentation'.length);
  if(!s&&!documentation){if(req.method==='GET'&&url.pathname==='/'){
    let login=fs.readFileSync(path.join(__dirname,'login.html'),'utf8');
    if(online)login=login
     .replace('Humbel isolierter Testzugang','FABOS · Anmeldung')
     .replace('Humbel Testversion','FABOS')
     .replace('<p>Nur synthetische Daten. Keine Produktionsverbindung. Testdaten werden beim Neustart zurückgesetzt.</p>','')
     .replace('Testkonto','Benutzerrolle')
     .replace('Testpasswort','Passwort')
     .replace('Testzugang prüfen.','Zugangsdaten prüfen.');
    return send(res,200,login);
   }return send(res,401,'Test sign-in required');}
  if(demo&&url.pathname==='/demo-plan'&&req.method==='GET')return send(res,200,renderDemoPlan(demo.data,{clean:online}));
  if(url.pathname==='/logout'&&req.method==='POST'){sessions.delete(sid);res.setHeader('Set-Cookie',`humbel_preview=; HttpOnly; SameSite=Strict; ${online?'Secure; ':''}Path=/; Max-Age=0`);return send(res,200,'{}','application/json');}
  if(url.pathname==='/rpc'&&req.method==='POST'){
   const readKeys={article_master_capabilities:[],article_master_list:['p_query'],article_master_get:['p_revision_id'],operation_plan:['p_revision_id'],operation_order_read:['p_order_id'],machine_readiness_view:[],op_routing_view:[],setup_time_view:[],op_quantity_view:[],ba_lifecycle_view:[],learning_view:[],learning_admin_view:['p_department']};
   if(documentation&&(!Object.hasOwn(readKeys,body.name)||!body.params||Object.keys(body.params).some(k=>!readKeys[body.name].includes(k))))return send(res,200,JSON.stringify({data:null,error:{code:'42501'}}),'application/json');
   if(documentation&&!['article_master_capabilities','article_master_list','article_master_get','operation_plan','operation_order_read','machine_readiness_view','op_routing_view','setup_time_view','op_quantity_view','ba_lifecycle_view','learning_view','learning_admin_view'].includes(body.name))return send(res,200,JSON.stringify({data:null,error:{code:'42501'}}),'application/json');
   if(typeof body.name!=='string'||!body.params||typeof body.params!=='object'||Array.isArray(body.params))return send(res,400,'Invalid RPC');
   try{const data=await backend.rpc(documentation?'admin':s.account,body.name,body.params);return send(res,200,JSON.stringify({data,error:null}),'application/json');}
   catch(e){return send(res,200,JSON.stringify({data:null,error:{code:e.code||'42501'}}),'application/json');}
  }
  if(req.method!=='GET')return send(res,405,'Method denied');
  const actor=backend.actors[documentation?'admin':s.account];
  if(url.pathname==='/preview-bootstrap.js'){
   const rows=JSON.parse(JSON.stringify(demo?demo.rows:rowsByTable));
   if(online&&demo){
     for(const machine of rows.production_machines)machine.name=machine.name.replace(/\bDEMO\b[\s:–-]*/g,'').trim();
     for(const order of rows.production_orders)order.ba_number=order.ba_number.replace(/^DEMO-/,'');
   }
   rows.employees[0]={...rows.employees[0],auth_user_id:actor.uid,role:actor.role,display_name:actor.name,first_name:actor.name,last_name:'',department_id:demo?demo.data.departments[0].id:'department-one'};if(!demo)rows.employees[1].display_name=rows.employees[1].first_name='Test Mitarbeiter';
   return send(res,200,`(${installSupabaseMock.toString()})(${JSON.stringify(rows)});window.__PREVIEW_ACTOR=${JSON.stringify(actor)};window.__PREVIEW_READONLY=${documentation};window.__PREVIEW_DEMO=${!!demo};window.__PREVIEW_ONLINE=${online};`+fs.readFileSync(path.join(__dirname,'client.js'),'utf8'),'application/javascript');
  }
  if(url.pathname==='/'||url.pathname==='/schichtplan_mvp.html'){
   let html=fs.readFileSync(path.join(ROOT,'schichtplan_mvp.html'),'utf8');html=html.replace('<script src="https://cdn.tailwindcss.com"></script>','<link rel="stylesheet" href="/tests/fixtures/app-tailwind.css">');
   html=html.replace('<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>','<script src="/preview-bootstrap.js"></script>').replace(/window.SUPABASE_URL = "[^"]+";/,'window.SUPABASE_URL = location.origin;').replace(/"sb_publishable_[^"]+"/,'"isolated-test-placeholder"');
   html=html.replace(/(window.HUMBEL_(?:ARTICLE_MASTER|OPERATIONS|BA_ENTRY|MACHINE_READINESS|OP_ROUTING|SETUP_TIMES|OP_QUANTITIES|BA_LIFECYCLE|LEARNING_UI|CONTEXT_HELP)_ENABLED = )false/g,'$1true');
   if(documentation)html=html.replaceAll('src="/','src="/documentation/').replaceAll('href="/','href="/documentation/');
   return send(res,200,html.replace('</body>',`<script src="${documentation?'/documentation':''}/scripts/preview/boot.js"></script></body>`));
  }
  const name=decodeURIComponent(url.pathname.slice(1));if(!/^(?:js\/[\w./-]+\.(?:js|json)|css\/[\w.-]+\.css|img\/[\w.-]+\.(?:png|svg)|tests\/fixtures\/app-tailwind\.css|scripts\/preview\/boot\.js|schichtplan_mvp\.js)$/.test(name)||name.includes('..'))return send(res,404,'Not found');
  const file=path.join(ROOT,name);if(!fs.existsSync(file))return send(res,404,'Not found');const ext=path.extname(file);send(res,200,fs.readFileSync(file),({'.js':'application/javascript','.json':'application/json','.css':'text/css','.png':'image/png','.svg':'image/svg+xml'})[ext]);
 }catch{if(!res.headersSent)send(res,500,'Isolated preview error');else res.end();}});
 server.requestTimeout=15000;server.headersTimeout=10000;server.setTimeout(30000);await new Promise(r=>server.listen(port,online?'0.0.0.0':'127.0.0.1',r));
 return{server,url:online?publicOrigin:`http://127.0.0.1:${server.address().port}`,close:async()=>{server.closeAllConnections();await new Promise(r=>server.close(r));await backend.close();}};
}
if(require.main===module){let passwords;try{passwords=JSON.parse(process.env.HUMBEL_PREVIEW_PASSWORDS||(fs.existsSync('.env.preview-credentials.json')?fs.readFileSync('.env.preview-credentials.json','utf8'):'null'));}catch{}createPreview({passwords,port:Number(process.env.PORT||4173)}).then(p=>{console.log(`Synthetic preview listening; online mode: ${process.env.HUMBEL_PREVIEW_ONLINE==='1'}; configured URL: ${p.url}`);for(const sig of ['SIGTERM','SIGINT'])process.once(sig,()=>p.close().finally(()=>process.exit()));}).catch(e=>{console.error(e.message);process.exitCode=1;});}
module.exports={createPreview,csp};
