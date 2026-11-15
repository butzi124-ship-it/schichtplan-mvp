'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const crypto=require('node:crypto');
const {createPreview}=require('../../scripts/preview/server');

test('external preview is opt-in and fails closed without HTTPS origin and synthetic data',{timeout:20000},async()=>{
 const old={online:process.env.HUMBEL_PREVIEW_ONLINE,origin:process.env.HUMBEL_PREVIEW_ORIGIN,demo:process.env.HUMBEL_PREVIEW_DEMO};
 const backend={actors:{employee:{uid:'synthetic',role:'employee',name:'Demo Employee'}},rpc:async()=>({}),close:async()=>{}};
 const passwords={employee:crypto.randomBytes(32).toString('hex')};
 try{
  process.env.HUMBEL_PREVIEW_ONLINE='1';delete process.env.HUMBEL_PREVIEW_ORIGIN;process.env.HUMBEL_PREVIEW_DEMO='1';
  await assert.rejects(()=>createPreview({backend,passwords}),/HTTPS origin/);
  process.env.HUMBEL_PREVIEW_ORIGIN='http://insecure.example.test';
  await assert.rejects(()=>createPreview({backend,passwords}),/public HTTPS origin/);
  process.env.HUMBEL_PREVIEW_ORIGIN='https://fabos.example.test';
  delete process.env.HUMBEL_PREVIEW_DEMO;
  await assert.rejects(()=>createPreview({backend,passwords}),/isolated demo/);
 }finally{for(const [key,val] of Object.entries(old)){const env={online:'HUMBEL_PREVIEW_ONLINE',origin:'HUMBEL_PREVIEW_ORIGIN',demo:'HUMBEL_PREVIEW_DEMO'}[key];if(val===undefined)delete process.env[env];else process.env[env]=val;}}
});

test('online preview restricts host and proxy protocol, authenticates protected plan and sets Secure cookie',{timeout:20000},async()=>{
 const old={online:process.env.HUMBEL_PREVIEW_ONLINE,origin:process.env.HUMBEL_PREVIEW_ORIGIN,demo:process.env.HUMBEL_PREVIEW_DEMO};
 const backend={actors:{employee:{uid:'synthetic',role:'employee',name:'Demo Employee'}},rpc:async()=>({}),close:async()=>{}};
 const password=crypto.randomBytes(32).toString('hex');
 let preview;
 try{
  process.env.HUMBEL_PREVIEW_ONLINE='1';process.env.HUMBEL_PREVIEW_ORIGIN='https://fabos.example.test';process.env.HUMBEL_PREVIEW_DEMO='1';
  preview=await createPreview({backend,passwords:{employee:password}});
  const url='http://127.0.0.1:'+preview.server.address().port;
  const valid={'Host':'fabos.example.test','X-Forwarded-Proto':'https'};
  let res=await fetch(url+'/healthz');assert.equal(res.status,200);
  res=await fetch(url+'/demo-plan');assert.equal(res.status,403);
  res=await fetch(url+'/demo-plan',{headers:valid});assert.equal(res.status,401);
  res=await fetch(url+'/demo-plan',{headers:{...valid,'X-Forwarded-Proto':'http'}});assert.equal(res.status,403);
  res=await fetch(url+'/documentation/');assert.equal(res.status,403);
  res=await fetch(url+'/documentation/',{headers:valid});assert.equal(res.status,404);
  res=await fetch(url+'/login',{method:'POST',headers:{...valid,'Origin':'https://fabos.example.test','Content-Type':'application/json'},body:JSON.stringify({account:'employee',password})});
  assert.equal(res.status,200);
  const cookie=res.headers.get('set-cookie');assert.match(cookie,/; Secure;/);assert.match(cookie,/HttpOnly/);assert.match(cookie,/SameSite=Strict/);
  res=await fetch(url+'/demo-plan',{headers:{...valid,Cookie:cookie.split(';')[0]}});
  assert.equal(res.status,200);assert.match(await res.text(),/Nächster Auftrag/);
  res=await fetch(url+'/logout',{method:'POST',headers:{...valid,Cookie:cookie.split(';')[0],Origin:'https://fabos.example.test','Content-Type':'application/json'},body:'{}'});
  assert.equal(res.status,200);assert.match(res.headers.get('set-cookie'),/; Secure;/);
 }finally{
  if(preview)await preview.close();
  for(const [key,val] of Object.entries(old)){const env={online:'HUMBEL_PREVIEW_ONLINE',origin:'HUMBEL_PREVIEW_ORIGIN',demo:'HUMBEL_PREVIEW_DEMO'}[key];if(val===undefined)delete process.env[env];else process.env[env]=val;}
 }
});
