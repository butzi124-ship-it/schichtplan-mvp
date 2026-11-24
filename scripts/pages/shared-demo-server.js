#!/usr/bin/env node
'use strict';
// Isolated password-protected FABOS demo. Never connects to production Supabase.
const http=require('node:http');
const fs=require('node:fs');
const path=require('node:path');
const crypto=require('node:crypto');
const {createSharedDemoStore}=require('./shared-demo-store');
const {build}=require('./build');
const {renderDemoPlan}=require('../preview/fabos-demo-adapter');
const ROOT=path.resolve(__dirname,'../..');
const DIST=path.join(ROOT,'dist-pages');
const LOGIN=path.join(__dirname,'shared-demo-login.html');
const csp="default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; connect-src 'self'; font-src 'self'; object-src 'none'; base-uri 'self'; frame-ancestors 'none'; form-action 'self'";
const mime={'.html':'text/html; charset=utf-8','.js':'application/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json; charset=utf-8','.png':'image/png','.svg':'image/svg+xml','.ico':'image/x-icon'};
function readPass(raw){
 if(typeof raw==='string'){try{return JSON.parse(raw);}catch{throw Error('Invalid isolated demo password configuration');}}
 return raw;
}
async function createSharedDemoServer({passwords,filePath,port=0,online=process.env.HUMBEL_PREVIEW_ONLINE==='1',publicOrigin=process.env.HUMBEL_PREVIEW_ORIGIN}={}){
 const accounts=readPass(passwords);
 if(!accounts||!['admin','leader','employee'].every(k=>typeof accounts[k]==='string'&&accounts[k].length>=24)||
    Object.keys(accounts).some(k=>!['admin','leader','employee'].includes(k)))throw Error('Three dedicated synthetic test passwords of at least 24 characters are required');
 if(online){
  let u;try{u=new URL(publicOrigin);}catch{throw Error('Online demo requires configured HTTPS origin');}
  if(u.protocol!=='https:'||u.pathname!=='/'||u.search||u.hash||u.port||u.username||u.password||!u.hostname)throw Error('Online demo requires clean HTTPS origin');
  if(!filePath||!path.isAbsolute(filePath))throw Error('Online demo requires a persistent absolute state path');
  // Refuse deployment with an ephemeral container-only /data directory.
  if(!fs.existsSync('/data')||fs.statSync('/data').dev===fs.statSync('/').dev)throw Error('Persistent Railway volume must be mounted at /data before serving shared state');
  if(!path.resolve(filePath).startsWith('/data/'))throw Error('Online shared state must reside on persistent /data volume');
 }
 if(!fs.existsSync(path.join(DIST,'fertigung.html')))build();
 const store=createSharedDemoStore(filePath||path.join(ROOT,'artifacts','shared-demo-state.json'));
 const salt=crypto.randomBytes(32);
 const hash=value=>crypto.scryptSync(value,salt,32);
 const hashes=Object.fromEntries(Object.entries(accounts).map(([k,v])=>[k,hash(v)]));
 const sessions=new Map(),attempts=new Map();
 function send(res,status,body,type='text/html; charset=utf-8',extras={}){
  res.writeHead(status,{'Content-Type':type,'Content-Security-Policy':csp,'Cache-Control':'no-store, private',
   'X-Content-Type-Options':'nosniff','Referrer-Policy':'no-referrer','X-Frame-Options':'DENY','Vary':'Cookie',...extras});
  res.end(body);
 }
 function json(res,status,body){send(res,status,JSON.stringify(body),'application/json; charset=utf-8');}
 const server=http.createServer(async(req,res)=>{
  try{
   if(req.url==='/healthz'&&req.method==='GET')return send(res,200,'ok','text/plain; charset=utf-8');
   const origin=online?publicOrigin:'http://127.0.0.1:'+server.address().port;
   if(req.headers.host!==new URL(origin).host)return send(res,403,'Host denied');
   if(online&&req.headers['x-forwarded-proto']!=='https')return send(res,403,'HTTPS required');
   const url=new URL(req.url,origin);
   const token=(req.headers.cookie||'').match(/(?:^|; )fabos_demo_session=([a-f0-9]{64})(?:;|$)/)?.[1];
   const session=sessions.get(token);
   if(session&&session.expires<Date.now())sessions.delete(token);
   const signed=!!session&&session.expires>=Date.now();
   if(req.method==='POST'&&req.headers.origin!==origin)return send(res,403,'Origin denied');
   let body={};
   if(req.method==='POST'){
    let raw='';
    for await(const chunk of req){raw+=chunk;if(Buffer.byteLength(raw)>32768)return send(res,413,'Request too large');}
    try{body=JSON.parse(raw);}catch{return json(res,400,{error:'Ungültiges JSON'});}
   }
   if(url.pathname==='/login'&&req.method==='POST'){
    const ip=req.socket.remoteAddress||'unknown',old=attempts.get(ip);
    const state=old&&old.until>Date.now()?old:{count:0,until:Date.now()+600000};attempts.set(ip,state);
    if(++state.count>10)return json(res,429,{error:'Zu viele Anmeldeversuche'});
    if(typeof body.password!=='string'||body.password.length>200||!Object.hasOwn(hashes,body.account)||
       !crypto.timingSafeEqual(hash(body.password),hashes[body.account]))return json(res,401,{error:'Anmeldung fehlgeschlagen'});
    state.count=0;
    const sid=crypto.randomBytes(32).toString('hex');
    sessions.set(sid,{account:body.account,expires:Date.now()+3600000});
    if(token)sessions.delete(token);
    res.setHeader('Set-Cookie','fabos_demo_session='+sid+'; HttpOnly; SameSite=Strict; '+(online?'Secure; ':'')+'Path=/; Max-Age=3600');
    return json(res,200,{ok:true,account:body.account});
   }
   if(req.method==='GET'&&(url.pathname==='/'||url.pathname==='/index.html')){
    if(!signed)return send(res,200,fs.readFileSync(LOGIN));
    return send(res,302,'', 'text/plain; charset=utf-8',{'Location':'/fertigung.html'});
   }
   if(!signed)return json(res,401,{error:'Testanmeldung erforderlich'});
   if(req.method==='POST'&&url.pathname==='/logout'){
    sessions.delete(token);
    res.setHeader('Set-Cookie','fabos_demo_session=; HttpOnly; SameSite=Strict; '+(online?'Secure; ':'')+'Path=/; Max-Age=0');
    return json(res,200,{ok:true});
   }
   if(req.method==='GET'&&url.pathname==='/api/state')return json(res,200,store.snapshot(session.account));
   if(req.method==='POST'&&url.pathname==='/api/action'){
    try{
     if(typeof body.action!=='string'||!body.payload||typeof body.payload!=='object'||Array.isArray(body.payload))return json(res,400,{error:'Ungültige Aktion'});
     return json(res,200,store.change(session.account,body.action,body.payload));
    }catch(err){return json(res,err.status||500,{error:err.status?err.message:'Testdaten konnten nicht geändert werden'});}
   }
   if(req.method!=='GET')return send(res,405,'Method denied');
   if(url.pathname==='/schichtplan_mvp.html')return send(res,302,'','text/plain; charset=utf-8',{'Location':'/fertigung.html'});
   if(url.pathname==='/demo-dataset.js')
    return send(res,200,'window.__FABOS_DEMO_DATA='+JSON.stringify(store.snapshot(session.account).data)+';','application/javascript; charset=utf-8');
   if(url.pathname==='/maschinenbelegung.html'){
    return send(res,200,renderDemoPlan(store.snapshot(session.account).data));
   }
   if(url.pathname==='/fertigung.html'){
    let html=fs.readFileSync(path.join(DIST,'fertigung.html'),'utf8');
    const role=store.actor(session.account).role;
    html=html.replace('</head>','<script>window.__FABOS_SERVER_ROLE='+JSON.stringify(role)+';</script></head>');
    html=html.replace('<a class="fab-top-link" href="./index.html">Rolle wechseln</a>','<a class="fab-top-link" id="sharedDemoLogout" href="/">Abmelden</a>');
    return send(res,200,html);
   }
   if(session.account!=='admin'&&['/app.html','/pages-bootstrap.js','/pages-boot.js','/schichtplan_mvp.js'].includes(url.pathname))return send(res,403,'Master-only legacy preview');
   const name=decodeURIComponent(url.pathname.slice(1));
   if(name.includes('..')||!(/^(?:js|css|img)\/[\w./-]+\.(?:js|css|png|svg|ico)$/.test(name)||
      /^(?:demo-ui-client\.js|app\.html|pages-bootstrap\.js|pages-boot\.js|schichtplan_mvp\.js)$/.test(name)||
      /^tests\/fixtures\/app-tailwind\.css$/.test(name)))return send(res,404,'Not found');
   const full=path.resolve(DIST,name);
   if(!full.startsWith(DIST+path.sep)||!fs.existsSync(full)||!fs.statSync(full).isFile())return send(res,404,'Not found');
   return send(res,200,fs.readFileSync(full),mime[path.extname(full)]||'application/octet-stream');
  }catch(err){
   console.error('FABOS isolated demo request error',err.message);
   if(!res.headersSent)json(res,500,{error:'Testserver-Fehler'});else res.end();
  }
 });
 server.requestTimeout=15000;server.headersTimeout=10000;server.setTimeout(30000);
 await new Promise(resolve=>server.listen(port,online?'0.0.0.0':'127.0.0.1',resolve));
 return {server,store,url:online?publicOrigin:'http://127.0.0.1:'+server.address().port,close:()=>new Promise(resolve=>{server.closeAllConnections();server.close(resolve);})};
}
if(require.main===module){
 createSharedDemoServer({passwords:process.env.HUMBEL_PREVIEW_PASSWORDS,filePath:process.env.FABOS_DEMO_STATE_PATH,port:Number(process.env.PORT||4173)})
 .then(r=>{console.log('FABOS isolated shared demo ready',r.url);for(const sig of ['SIGTERM','SIGINT'])process.once(sig,()=>r.close().finally(()=>process.exit()));})
 .catch(e=>{console.error('FABOS isolated demo start refused:',e.message);process.exitCode=1;});
}
module.exports={createSharedDemoServer};
