// Synthetic preview adapter, served only after the server authenticated this session.
(async()=>{
 const a=window.__PREVIEW_ACTOR;
 let ready=false, issue='';
 try{
  // Do not use loginAs('admin'): it only sets the navigation role and does not
  // load employees, departments, machines, orders or tool inventory.
  currentSupabaseUser={id:a.uid};
  currentUser={name:a.name,role:a.role};
  await bootSupabase();
  currentEmployeeRecord=(state.employeesList||[]).find(x=>x.auth_user_id===a.uid&&x.is_active!==false)||null;
  if(!currentEmployeeRecord)throw Error('Kein zugeordneter synthetischer Mitarbeiter gefunden');
  if(window.__PREVIEW_DEMO){
   const data=window.__SUPABASE_MOCK_ROWS;
   if(!data||data.departments.length!==8||data.employees.length!==51||data.production_machines.length!==48||data.production_orders.length!==50)throw Error('Demo-Datensatz unvollständig');
   if(state.departments.length!==8||state.employeesList.length!==51||state.productionMachines.length!==48||state.productionOrders.length!==50)throw Error('Demo-Datensatz wurde nicht in FABOS geladen');
  }
  await loadToolPageData();
  if(window.__PREVIEW_DEMO&&(!state.tools||state.tools.length<10))throw Error('Demo-Werkzeugbestand wurde nicht in FABOS geladen');
  state.ui=state.ui||{};
  state.ui.supabaseReady=true;
  state.ui.toolsInitialLoaded=true;
  currentTab='produktion';
  state.productionSubTab='departments';
  render();
  ready=true;
 }catch(e){console.error('FABOS preview data initialization failed',e);issue=String(e?.message||e);}
 const banner=document.createElement('aside');banner.id='isolated-preview-banner';
 banner.style='background:#fff3c4;color:#543600;padding:12px;font:14px system-ui;position:sticky;top:0;z-index:90';
 banner.textContent='SYNTHETISCHE ENTWICKLUNGSANSICHT · keine Original-Firmendaten · nicht dauerhaft gespeichert';
 document.body.prepend(banner);
 if(window.__PREVIEW_ONLINE)banner.append(' · ONLINE-SIMULATION: keine gemeinsamen serverseitigen Rückmeldungen');
 if(!ready){
  banner.style.background='#fee2e2';banner.style.color='#991b1b';
  banner.append(' · DEMO-DATEN NICHT GELADEN: '+issue);
  const host=document.getElementById('view');
  if(host)host.innerHTML='<div role="alert" style="padding:20px;background:#fee2e2;color:#991b1b">Die Demo-Daten konnten nicht initialisiert werden. Bitte Administrator informieren.</div>';
 }
 const out=document.createElement('button');out.textContent='Testzugang beenden';
 out.onclick=async()=>{await fetch('/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});localStorage.clear();location.reload();};
 if(!window.__PREVIEW_READONLY)banner.append(' ',out);else banner.append(' · Nur lesende Dokumentationsansicht');
 if(window.__PREVIEW_DEMO){
  const link=document.createElement('a');link.href='/demo-plan';link.textContent='Maschinenbelegung und OP-Vorschau (Demodaten)';
  link.style='margin-left:12px;color:#0756a8;text-decoration:underline;font-weight:700';banner.append(' ',link);
 }
 if(ready&&window.__PREVIEW_DEMO){
  const data=window.__SUPABASE_MOCK_ROWS;
  const summary=document.createElement('span');
  summary.textContent=' · '+data.departments.length+' Abteilungen · 50 Fantasiemitarbeiter · '+data.production_machines.length+' Maschinen · '+data.production_orders.length+' BA einschließlich geplanter · '+state.tools.length+' Werkzeuge';
  banner.append(summary);
 }
 const login=document.getElementById('loginBox');if(login){login.hidden=true;login.classList.add('hidden');}
})();
