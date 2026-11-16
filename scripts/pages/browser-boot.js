// Bootstrap the isolated GitHub Pages edition with browser-local test data.
(async()=>{
 const requestedRole=new URLSearchParams(location.search).get('role')||'admin';
 const a=window.__PREVIEW_ACTORS?.[requestedRole]||window.__PREVIEW_ACTORS?.admin;
 if(!a)throw Error('Keine Testrolle vorhanden');
 const login=document.getElementById('loginBox');
 if(login){login.hidden=true;login.classList.add('hidden');}
 try{
  currentSupabaseUser={id:a.uid};
  currentUser={name:a.name,role:a.role};
  await bootSupabase();
  currentEmployeeRecord=(state.employeesList||[]).find(x=>x.auth_user_id===a.uid&&x.is_active!==false)||null;
  if(!currentEmployeeRecord)throw Error('Der Benutzer konnte nicht zugeordnet werden.');
  if(state.departments.length!==8||state.employeesList.length!==41||state.productionMachines.length!==19||!state.productionOrders.length)throw Error('Die Fertigungsdaten wurden nicht vollständig geladen.');
  await loadToolPageData();
  if(!state.tools||state.tools.length!==20)throw Error('Die Werkzeugdaten wurden nicht vollständig geladen.');
  state.ui=state.ui||{};
  state.ui.supabaseReady=true;
  state.ui.toolsInitialLoaded=true;
  currentTab='produktion';
  state.productionSubTab='machines';
  render();
  const header=document.querySelector('.h-app-header');
  if(header){
   const link=document.createElement('a');
   link.href='./fertigung.html?role='+encodeURIComponent(requestedRole);
   link.textContent='Maschinenbelegung & OP-Plan';
   link.className='px-3 py-2 rounded bg-blue-700 text-white font-semibold';
   header.appendChild(link);
   if(requestedRole==='lead'){
    const approval=document.createElement('a');
    approval.href='./fertigung.html?role=lead#freigaben';
    approval.textContent='Freigabebereich Fräserei';
    approval.className='px-3 py-2 rounded bg-emerald-700 text-white font-semibold';
    header.appendChild(approval);
   }
  }
  // A static release has no backend session to sign out of. Reload resets the browser fixture.
  if(typeof window.logout==='function')window.logout=()=>location.href='./index.html';
  window.__FABOS_PAGES_READY=true;
 }catch(e){
  console.error('FABOS initialization error',e);
  const view=document.getElementById('view');
  if(view)view.textContent='Die Anwendung konnte nicht geladen werden. Bitte Seite neu laden.';
  window.__FABOS_PAGES_ERROR=String(e&&e.message||e);
 }
})();
