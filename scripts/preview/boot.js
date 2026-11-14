// This adapter is served ONLY by the isolated loopback preview server.
(async()=>{const a=window.__PREVIEW_ACTOR;loginAs('admin');currentSupabaseUser={id:a.uid};currentUser.role=a.role;currentUser.name=a.name;render();
 const banner=document.createElement('aside');banner.id='isolated-preview-banner';banner.style='background:#fff3c4;color:#543600;padding:12px;font:14px system-ui;position:sticky;top:0;z-index:90';banner.textContent='SYNTHETISCHE ENTWICKLUNGSANSICHT · ausschließlich synthetische Daten · keine native Supabase-Abnahme · Neustart setzt Daten zurück';document.body.prepend(banner);
 const out=document.createElement('button');out.textContent='Testzugang beenden';out.onclick=async()=>{await fetch('/logout',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});localStorage.clear();location.reload();};if(!window.__PREVIEW_READONLY)banner.append(' ',out);else banner.append(' · Dokumentationsansicht ohne schreibende Serveraktionen');
 document.getElementById('loginBox').hidden=true;
})();
