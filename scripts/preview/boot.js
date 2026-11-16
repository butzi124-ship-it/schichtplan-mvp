// Presentation shell for the isolated preview; online appearance stays independent of stored data.
(async()=>{
 const a=window.__PREVIEW_ACTOR;
 loginAs('admin');
 currentSupabaseUser={id:a.uid};
 currentUser.role=a.role;
 currentUser.name=a.name;
 render();
 const logout=async()=>{
  await fetch('/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});
  localStorage.clear();
  location.reload();
 };
 if(window.__PREVIEW_ONLINE){
  const header=document.querySelector('.h-app-header')||document.querySelector('header');
  const nav=document.createElement('nav');
  nav.setAttribute('aria-label','Fertigungsnavigation');
  nav.style='display:flex;align-items:center;gap:12px;flex-wrap:wrap';
  if(window.__PREVIEW_DEMO){
   const link=document.createElement('a');
   link.href='/demo-plan';
   link.textContent='Maschinenbelegung';
   link.style='color:#0871dc;font-weight:600;text-decoration:underline';
   nav.append(link);
  }
  const out=document.createElement('button');
  out.type='button';
  out.textContent='Abmelden';
  out.style='background:#eff6ff;border:1px solid #cadbec;color:#102f52;border-radius:8px;padding:10px 14px;cursor:pointer';
  out.onclick=logout;
  nav.append(out);
  (header||document.body).append(nav);
 } else {
  const banner=document.createElement('aside');
  banner.id='isolated-preview-banner';
  banner.style='background:#fff3c4;color:#543600;padding:12px;font:14px system-ui;position:sticky;top:0;z-index:90';
  banner.textContent='SYNTHETISCHE ENTWICKLUNGSANSICHT · ausschließlich synthetische Daten · keine native Supabase-Abnahme · Neustart setzt Daten zurück';
  document.body.prepend(banner);
  const out=document.createElement('button');
  out.textContent='Testzugang beenden';
  out.onclick=logout;
  if(!window.__PREVIEW_READONLY)banner.append(' ',out);
  else banner.append(' · Dokumentationsansicht ohne schreibende Serveraktionen');
  if(window.__PREVIEW_DEMO){
   const link=document.createElement('a');
   link.href='/demo-plan';
   link.textContent='Maschinenbelegung und OP-Vorschau (Demodaten)';
   link.style='margin-left:12px;color:#0756a8;text-decoration:underline;font-weight:700';
   banner.append(' ',link);
  }
 }
 document.getElementById('loginBox').hidden=true;
})();
