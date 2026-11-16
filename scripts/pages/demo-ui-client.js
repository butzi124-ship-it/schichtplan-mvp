'use strict';
// Public synthetic production cockpit. Role choice is presentation-only, NOT authentication.
(()=>{
 const data=window.__FABOS_DEMO_DATA;
 if(!data||!Array.isArray(data.operations))throw Error('Demo-Datensatz fehlt');
 const params=new URLSearchParams(location.search);
 const role=['admin','lead','employee'].includes(params.get('role'))?params.get('role'):'employee';
 const fr=data.departments.find(d=>d.code==='FR');
 const lead=data.employees.find(e=>e.id===fr.leader_employee_id);
 const worker=data.employees.find(e=>e.department_id===fr.id&&e.role==='employee');
 const actor=role==='admin'?{name:'FABOS Administration',role:'Administrator'}:role==='lead'?{name:lead.display_name,role:'Abteilungsleiter Fräserei'}:{name:worker.display_name,role:'Mitarbeiter Fräserei'};
 const esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;','\'':'&#39;'}[c]));
 const byOrder=new Map(data.orders.map(o=>[o.id,o]));
 const byEmp=new Map(data.employees.map(e=>[e.id,e]));
 const byDept=new Map(data.departments.map(d=>[d.id,d]));
 const fmt=s=>new Date(s).toLocaleString('de-DE',{timeZone:'Europe/Berlin',day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'});
 const status={running:'Läuft',queued:'Geplant',completed:'Fertig',idle:'Frei'};
 const state={dept:role==='admin'?'all':fr.id,search:'',machine:null,tab:'machines',approved:new Set()};
 const $=id=>document.getElementById(id);
 $('roleLabel').textContent=actor.name+' · '+actor.role;
 $('appLink').href='./app.html?role='+role;
 $('asof').textContent='Planungsstand: '+fmt(data.as_of);
 if(role==='employee'){$('approvalNav').hidden=true;$('appLink').hidden=true;}
 const departmentFilter=$('departmentFilter');
 departmentFilter.innerHTML=(role==='admin'?'<option value="all">Alle Abteilungen</option>':'')+data.departments.filter(d=>role==='admin'||d.id===fr.id).map(d=>'<option value="'+esc(d.id)+'">'+esc(d.name)+'</option>').join('');
 departmentFilter.value=state.dept;
 const readable=op=>op?status[op.state]||op.state:'Frei';
 const current=m=>data.operations.find(o=>o.machine_id===m.id&&o.state==='running');
 const queued=m=>data.operations.filter(o=>o.machine_id===m.id&&o.state==='queued').sort((a,b)=>a.planned_start.localeCompare(b.planned_start));
 const lastCompleted=m=>data.operations.filter(o=>o.machine_id===m.id&&o.state==='completed').sort((a,b)=>b.planned_end.localeCompare(a.planned_end))[0];
 const primary=m=>current(m)||queued(m)[0]||null;
 const visibleMachines=()=>data.machines.filter(m=>(state.dept==='all'||m.department_id===state.dept)&&(!state.search||[m.name,m.machine_code,byDept.get(m.department_id)?.name].join(' ').toLocaleLowerCase('de-DE').includes(state.search)));
 const visibleOps=()=>data.operations.filter(op=>state.dept==='all'||op.department_id===state.dept);
 function renderMetrics(){
  const machines=data.machines.filter(m=>state.dept==='all'||m.department_id===state.dept);
  const ops=visibleOps(),orders=new Set(ops.map(op=>op.order_id));
  const nums=[[machines.length,'Maschinen'],[machines.filter(m=>!!current(m)).length,'Aktuell belegt'],[orders.size,'Betriebsaufträge'],[ops.filter(o=>o.state==='queued').length,'Geplante Arbeitsgänge']];
  $('metrics').innerHTML=nums.map(([n,label])=>'<div class="fab-metric"><strong>'+n+'</strong><span>'+label+'</span></div>').join('');
  const approvals=approvalOps();
  $('approvalCount').textContent=approvals.filter(op=>!state.approved.has(op.id)).length;
 }
 function renderMachines(){
  const machines=visibleMachines();
  if(!machines.some(m=>m.id===state.machine))state.machine=(machines.find(m=>primary(m)&&primary(m).assigned_employee_id===worker.id)||machines.find(m=>current(m))||machines[0])?.id||null;
  $('machineCount').textContent='('+machines.length+')';
  $('machineCards').innerHTML=machines.map(m=>{
   const op=primary(m),order=op?byOrder.get(op.order_id):null,active=current(m),next=queued(m).find(q=>q.id!==op?.id);
   const label=op?(active?'Läuft':'Als Nächstes'):'Frei';
   return '<button type="button" class="fab-machine '+(state.machine===m.id?'selected':'')+'" data-machine="'+esc(m.id)+'" aria-pressed="'+(state.machine===m.id)+'">'+
    '<span class="fab-machine-head"><strong>'+esc(m.name)+'</strong><span class="fab-status '+(op?.state||'idle')+'">'+label+'</span></span>'+
    '<span class="fab-machine-meta">'+esc(byDept.get(m.department_id)?.name||'')+' · '+esc(m.machine_code)+'</span>'+
    '<span class="fab-article">'+esc(order?.article_name||'Keine offene Belegung')+'</span>'+
    '<span class="fab-bano">'+esc(order?.ba_number||'–')+(op?' · OP '+esc(op.op_number):'')+'</span>'+
    '<span class="fab-subline">'+(op?'Soll: '+esc(byOrder.get(op.order_id)?.ba_quantity)+' Stk. · '+esc(byEmp.get(op.assigned_employee_id)?.display_name||'–'):'Aktuell kein laufender oder geplanter Auftrag')+'</span>'+
    (next?'<span class="fab-subline">Danach: '+esc(byOrder.get(next.order_id)?.ba_number)+' · OP '+esc(next.op_number)+'</span>':'')+'</button>';
  }).join('')||'<div class="fab-empty">Keine Maschine zu diesem Filter gefunden.</div>';
  renderMachineDetail();
 }
 function opDetails(op,heading){
  if(!op)return '<div class="fab-empty">Kein '+esc(heading.toLocaleLowerCase('de-DE'))+' vorhanden.</div>';
  const order=byOrder.get(op.order_id),person=byEmp.get(op.assigned_employee_id),ratio=Math.max(0,Math.min(100,Math.round(100*op.good_qty/Math.max(1,order.ba_quantity))));
  return '<h3>'+esc(heading)+' <span class="fab-status '+esc(op.state)+'">'+readable(op)+'</span></h3>'+
   '<div class="fab-field-grid"><div class="fab-field"><small>Auftrag / Arbeitsgang</small><strong>'+esc(order.ba_number)+' / OP '+esc(op.op_number)+'</strong></div>'+
   '<div class="fab-field"><small>Artikel / Material</small><strong>'+esc(order.article_name)+' · '+esc(order.material)+'</strong></div>'+
   '<div class="fab-field"><small>Auftragsmenge</small><strong>'+order.ba_quantity+' Stück</strong></div>'+
   '<div class="fab-field"><small>Zugeordneter Mitarbeiter</small><strong>'+esc(person?.display_name||'Nicht zugeteilt')+'</strong></div>'+
   '<div class="fab-field"><small>Geplanter Start</small><strong>'+esc(fmt(op.planned_start))+'</strong></div>'+
   '<div class="fab-field"><small>Geplantes Ende</small><strong>'+esc(fmt(op.planned_end))+'</strong></div>'+
   '<div class="fab-field"><small>Rüstzeit</small><strong>'+op.setup_minutes+' Minuten</strong></div>'+
   '<div class="fab-field"><small>Zykluszeit / Teil</small><strong>'+op.cycle_seconds+' Sekunden</strong></div></div>'+
   '<h3>Arbeitsanweisung</h3><p class="fab-instruction">'+esc(op.description)+'</p>'+
   '<div class="fab-field"><small>Simulierter Fertigungsfortschritt</small><strong>'+op.good_qty+' / '+order.ba_quantity+' Gutteile ('+ratio+' %)</strong></div><div class="fab-progress" role="progressbar" aria-valuenow="'+ratio+'" aria-valuemin="0" aria-valuemax="100"><span style="width:'+ratio+'%"></span></div>';
 }
 function renderMachineDetail(){
  const m=data.machines.find(x=>x.id===state.machine);
  if(!m){$('machineDetail').innerHTML='<div class="fab-empty">Maschine auswählen, um Arbeitsgänge zu sehen.</div>';return;}
  const running=current(m),q=queued(m),main=running||q[0],next=running?q[0]:q[1],past=lastCompleted(m);
  $('machineDetail').innerHTML='<div class="fab-eyebrow">MASCHINENDETAIL</div><h2>'+esc(m.name)+'</h2><p class="fab-muted">'+esc(byDept.get(m.department_id)?.name||'')+' · '+esc(m.machine_code)+'</p>'+
    (main?opDetails(main,running?'Aktueller Arbeitsgang':'Nächster eingeplanter Arbeitsgang'):'<div class="fab-empty">An dieser Maschine ist zur Demo-Bezugszeit kein Arbeitsgang offen.</div>')+
    '<hr class="fab-divider"><h3>Maschinenfolge</h3>'+(next?'<div class="fab-order-detail"><strong>Danach:</strong> '+esc(byOrder.get(next.order_id).ba_number)+' · OP '+esc(next.op_number)+' · '+esc(byOrder.get(next.order_id).article_name)+'<br>Start '+fmt(next.planned_start)+'</div>':'<p class="fab-muted">Kein weiterer Auftrag eingeplant.</p>')+
    (past?'<div class="fab-order-detail fab-muted">Zuletzt abgeschlossen: '+esc(byOrder.get(past.order_id).ba_number)+' · OP '+esc(past.op_number)+'</div>':'');
 }
 function visibleOrders(){
  const allowed=new Set(visibleOps().map(o=>o.order_id));
  return data.orders.filter(o=>allowed.has(o.id));
 }
 function renderOrders(){
  $('orderCards').innerHTML=visibleOrders().map(o=>{
   const ops=data.operations.filter(op=>op.order_id===o.id&&(state.dept==='all'||op.department_id===state.dept));
   const rows=ops.map(op=>'<div class="fab-order-detail">OP '+esc(op.op_number)+' · <strong>'+esc(op.machine)+'</strong> · '+esc(op.department)+' · '+fmt(op.planned_start)+' · <span class="fab-status '+op.state+'">'+readable(op)+'</span></div>').join('');
   return '<article class="fab-order"><div class="fab-order-head"><div><h3>'+esc(o.ba_number)+' · '+esc(o.article_name)+'</h3><p>'+esc(o.article_number)+' · '+esc(o.material)+' · '+o.ba_quantity+' Stk.</p></div><span class="fab-status '+o.status+'">'+readable(o)+'</span></div>'+rows+'</article>';
  }).join('')||'<p class="fab-empty">Keine Betriebsaufträge zugeordnet.</p>';
 }
 function approvalOps(){
  if(role==='employee')return [];
  const dept=state.dept==='all'?fr.id:state.dept;
  return data.operations.filter(op=>op.department_id===dept&&(op.state==='running'||op.state==='queued')).sort((a,b)=>a.planned_start.localeCompare(b.planned_start)).slice(0,8);
 }
 function renderApprovals(){
  const ops=approvalOps();
  $('approvalCards').innerHTML=ops.map(op=>{
   const o=byOrder.get(op.order_id),done=state.approved.has(op.id);
   return '<article class="fab-order"><div class="fab-order-head"><div><h3>'+esc(o.ba_number)+' · OP '+esc(op.op_number)+'</h3><p>'+esc(op.machine)+' · '+esc(o.article_name)+' · '+esc(op.department)+'</p></div><span class="fab-status '+(done?'completed':'queued')+'">'+(done?'In dieser Ansicht freigegeben':'Freigabe offen')+'</span></div><p class="fab-order-detail">Prüfschritt: '+(op.state==='running'?'Erststückfreigabe':'Rüst- und Anlauffreigabe')+' · Geplanter Beginn '+fmt(op.planned_start)+'</p><button type="button" class="fab-approve" data-approve="'+esc(op.id)+'" '+(done?'disabled':'')+'>'+(done?'Freigegeben':'Freigeben (Simulation)')+'</button></article>';
  }).join('')||'<div class="fab-empty">In dieser Abteilung sind keine offenen Freigaben simuliert.</div>';
 }
 function switchTab(tab){
  if(tab==='approvals'&&role==='employee')return;
  state.tab=tab;
  document.querySelectorAll('.fab-nav[data-tab]').forEach(el=>el.classList.toggle('active',el.dataset.tab===tab));
  $('machineView').hidden=tab!=='machines';$('ordersView').hidden=tab!=='orders';$('approvalView').hidden=tab!=='approvals';
  $('pageTitle').textContent={machines:'Maschinenübersicht',orders:'Betriebsaufträge',approvals:'Freigabebereich'}[tab];
  $('pageIntro').textContent={machines:'Was läuft an meiner Maschine – und was kommt als Nächstes?',orders:'Alle 32 Testaufträge und ihre geplanten Maschinen-Arbeitsgänge.',approvals:'Anstehende Freigaben meiner Abteilung.'}[tab];
 }
 document.querySelectorAll('.fab-nav[data-tab]').forEach(el=>el.addEventListener('click',()=>switchTab(el.dataset.tab)));
 departmentFilter.addEventListener('change',e=>{state.dept=e.target.value;state.machine=null;render();});
 $('machineSearch').addEventListener('input',e=>{state.search=e.target.value.trim().toLocaleLowerCase('de-DE');renderMachines();});
 $('machineCards').addEventListener('click',e=>{const btn=e.target.closest('[data-machine]');if(!btn)return;state.machine=btn.dataset.machine;renderMachines();});
 $('approvalCards').addEventListener('click',e=>{const btn=e.target.closest('[data-approve]');if(!btn)return;state.approved.add(btn.dataset.approve);renderMetrics();renderApprovals();});
 function render(){renderMetrics();renderMachines();renderOrders();renderApprovals();}
 render();switchTab(role==='lead'?'approvals':'machines');
 window.__FABOS_COCKPIT_READY=true;
})();