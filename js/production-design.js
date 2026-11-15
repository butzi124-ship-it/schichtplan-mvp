/* Reference design package 1. Presentation of already authorized, loaded BA data only. */
(function (root) {
  'use strict';
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const icons = {factory:'M3 21V10l6 3V7l6 3V3h6v18H3Zm4-4h2m4 0h2m3 0h1',home:'m3 11 9-8 9 8M5 10v11h5v-7h4v7h5V10',calendar:'M4 5h16v16H4zM7 2v6m10-6v6M4 10h16',check:'M5 3h14v18H5zM8 12l3 3 5-6',tool:'m14 5 5 5M3 21l10-10M13 3a6 6 0 0 0 8 8',layers:'m12 3 10 6-10 6L2 9Zm-10 12 10 6 10-6M2 12l10 6 10-6',users:'M16 21v-3a4 4 0 0 0-8 0v3M12 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8M3 21v-3m18 3v-3',search:'M21 21l-6-6M10 17a7 7 0 1 0 0-14 7 7 0 0 0 0 14',gear:'M12 3v3m0 12v3M3 12h3m12 0h3M6 6l2 2m8 8 2 2M6 18l2-2M16 8l2-2M12 7a5 5 0 1 0 0 10 5 5 0 0 0 0-10',play:'m8 5 12 7-12 7Z',pause:'M8 4v16M16 4v16',alert:'m12 3 10 18H2ZM12 9v5m0 3v1'};
  function icon(name) { return `<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="${icons[name] || icons.layers}"/></svg>`; }
  const progress = o => o.target > 0 ? Math.min(100, Math.max(0, Math.round(o.good / o.target * 100))) : 0;
  function badge(o) { return `<span class="h-status h-status-${o.status === 'running' ? 'running' : 'paused'}">${esc(o.statusLabel)}</span>`; }
  function detail(o) {
    if (!o) return '<p class="h-empty">Kein aktiver Auftrag geladen. Bitte eine verfügbare Maschine auswählen.</p>';
    return `<header class="h-detail-heading"><span class="h-part-icon">${icon('layers')}</span><div><h3>BA ${esc(o.ba)}</h3><p>Artikel ${esc(o.article)}</p><small>${esc(o.machine)} · ${esc(o.department)}</small></div>${badge(o)}</header>
      <h4>Auftragsübersicht</h4><dl class="h-detail-data"><dt>Maschine</dt><dd>${esc(o.machine)}</dd><dt>Zielmenge</dt><dd>${esc(o.target)} Stück</dd><dt>Fertige Gutteile</dt><dd>${esc(o.good)} Stück</dd><dt>Fortschritt</dt><dd>${progress(o)} %<progress max="100" value="${progress(o)}">${progress(o)} %</progress></dd><dt>Ausschuss</dt><dd>${esc(o.scrap)}</dd><dt>In Abklärung</dt><dd>${esc(o.clarify)}</dd><dt>BA-Stückzahl</dt><dd>${esc(o.quantity)}</dd></dl>
      <button type="button" class="h-primary h-detail-open" onclick="selectProductionCounterOrder('${esc(o.id)}', true)">Stückzahlen / Auftrag öffnen ${icon('play')}</button>`;
  }
  function render({orders, machines, machineCards}) {
    const running = orders.filter(o => o.status === 'running').length;
    const paused = orders.filter(o => o.status === 'paused').length;
    const clarification = orders.filter(o => o.clarify > 0).length;
    const cards = [['Aktive BA',orders.length,'layers','blue'],['Laufend',running,'play','green'],['Pausiert',paused,'pause','amber'],['Maschinen',machines.length,'factory','gray'],['BA mit Abklärung',clarification,'alert','red']];
    return `<section class="h-production-overview" aria-label="Fertigungsübersicht"><header class="h-production-heading">${icon('factory')}<div><h3>Fertigung – Produktionsübersicht</h3><p>Aktive Aufträge, Maschinen und gebuchte Mengen</p><h4 class="h-overview-caption">Zähler-Dashboard</h4></div></header>
      <div class="h-kpi-grid">${cards.map(([name,n,ic,color])=>`<article class="h-kpi h-tone-${color}"><span class="h-kpi-icon">${icon(ic)}</span><div><h4>${name}</h4><strong>${n}</strong></div></article>`).join('')}</div>
      <form class="h-production-filters" onsubmit="return false" aria-label="Fertigung filtern"><label>Abteilung<select data-production-filter="department"><option value="">Alle Abteilungen</option>${[...new Set(orders.map(o=>o.department))].map(d=>`<option>${esc(d)}</option>`).join('')}</select></label><label>Maschine<select data-production-filter="machine"><option value="">Alle Maschinen</option>${[...new Set(orders.map(o=>o.machine))].map(d=>`<option>${esc(d)}</option>`).join('')}</select></label><label>Status<select data-production-filter="status"><option value="">Alle Status</option><option value="running">Laufend</option><option value="paused">Pausiert</option></select></label><label class="h-search-field">${icon('search')}<span class="h-sr-only">BA oder Artikel suchen</span><input type="search" data-production-filter="search" placeholder="BA oder Artikel suchen …"></label></form>
      <div class="h-production-split"><div class="h-production-table-card"><div class="h-table-scroll" data-table-label="Aktive Fertigungsaufträge"><table class="h-production-table"><thead><tr><th>BA</th><th>Artikel</th><th>Maschine</th><th>Status</th><th>Fortschritt</th><th>Ziel / Gut</th><th><span class="h-sr-only">Aktion</span></th></tr></thead><tbody>${orders.map((o,i)=>`<tr data-ref-order="${i}" data-department="${esc(o.department)}" data-machine="${esc(o.machine)}" data-status="${esc(o.status)}"><td><button type="button" class="h-order-preview" data-preview-index="${i}" aria-label="BA ${esc(o.ba)} Vorschau" aria-pressed="${i===0}">${icon('layers')} ${esc(o.ba)}</button></td><td>${esc(o.article)}</td><td>${esc(o.machine)}</td><td>${badge(o)}</td><td><span class="h-progress-label">${progress(o)} %</span><progress max="100" value="${progress(o)}">${progress(o)} %</progress></td><td>${esc(o.target)} / ${esc(o.good)}</td><td><button type="button" class="h-row-open" aria-label="BA ${esc(o.ba)} öffnen" onclick="selectProductionCounterOrder('${esc(o.id)}', true)">→</button></td></tr>`).join('')}</tbody></table></div><p class="h-empty" data-filter-empty ${orders.length?'hidden':''}>Keine aktiven Aufträge für diese Auswahl.</p></div><aside class="h-reference-detail" aria-label="Auftragsvorschau">${detail(orders[0])}</aside></div>
      <section class="h-machine-directory"><h3>Maschinenübersicht</h3><div class="h-machine-links">${machineCards || machines.map(m=>`<button type="button" onclick="selectProductionCounterMachine('${esc(m.id)}')">${icon('factory')}<span>${esc(m.name)}<small>${esc(m.department)}</small></span></button>`).join('') || '<p>Keine Maschinen sichtbar.</p>'}</div></section></section>`;
  }
  let currentOrders = [];
  root.HumbelProductionDesign = { render(model) { currentOrders = model.orders; return render(model); } };
  function filter() {
    const scope = document.querySelector('.h-production-overview'); if (!scope) return;
    const values = Object.fromEntries([...scope.querySelectorAll('[data-production-filter]')].map(e=>[e.dataset.productionFilter,e.value.trim().toLocaleLowerCase()]));
    let visible=0;
    scope.querySelectorAll('[data-ref-order]').forEach(row=>{
      const o=currentOrders[Number(row.dataset.refOrder)];
      row.hidden=!!((values.department&&o.department.toLocaleLowerCase()!==values.department)||(values.machine&&o.machine.toLocaleLowerCase()!==values.machine)||(values.status&&o.status!==values.status)||(values.search&&!`${o.ba} ${o.article}`.toLocaleLowerCase().includes(values.search)));
      if(!row.hidden)visible++;
    });
    scope.querySelector('[data-filter-empty]').hidden=visible>0;
    const selected=scope.querySelector('[data-preview-index][aria-pressed=true]');
    if(selected)scope.querySelector('.h-reference-detail').innerHTML=selected.closest('tr').hidden?'<p class="h-empty">Bitte einen sichtbaren Auftrag für die Vorschau auswählen.</p>':detail(currentOrders[Number(selected.dataset.previewIndex)]);
  }
  document.addEventListener('input', e=>{if(e.target.matches('[data-production-filter]')){if(e.target.dataset.productionFilter==='search')document.querySelectorAll('[data-production-filter=search]').forEach(x=>{if(x!==e.target)x.value=e.target.value;});filter();}});
  document.addEventListener('change', e=>{if(e.target.matches('[data-production-filter]'))filter();});
  document.addEventListener('click',e=>{
    const b=e.target.closest('[data-preview-index]');if(!b)return;
    const scope=b.closest('.h-production-overview');scope.querySelectorAll('[data-preview-index]').forEach(x=>x.setAttribute('aria-pressed',String(x===b)));
    scope.querySelector('.h-reference-detail').innerHTML=detail(currentOrders[Number(b.dataset.previewIndex)]);
  });
  function decorate() {
    const nav=document.getElementById('tabs');if(!nav)return;
    if(!nav.querySelector('.h-nav-brand')){const brand=document.createElement('div');brand.className='h-nav-brand';brand.innerHTML=icon('gear')+'<span>HUMBEL GEARS<small>Humbel-App</small></span>';nav.prepend(brand);}
    const map=[[/Dashboard/,'home'],[/Schicht/,'calendar'],[/To-Do|Aufgaben/,'check'],[/Produktion/,'factory'],[/Werkzeug/,'tool'],[/Mitarbeiter|Personal/,'users'],[/Verwaltung/,'gear']];
    nav.querySelectorAll('button').forEach(b=>{if(b.querySelector('svg'))return;const name=map.find(([r])=>r.test(b.textContent))?.[1]||'layers';const label=document.createElement('span');while(b.firstChild)label.appendChild(b.firstChild);b.appendChild(label);b.insertAdjacentHTML('afterbegin',icon(name));});
  }
  function workspace() {
    const view=document.getElementById('view'),nav=document.getElementById('tabs'),header=document.querySelector('.h-app-header');
    const production=!!view.querySelector('.h-production');document.body.classList.toggle('h-reference-production',production);
    let search=header.querySelector('.h-top-search');
    if(view.querySelector('.h-production-overview')&&!search){search=document.createElement('label');search.className='h-top-search h-search-field';search.innerHTML=icon('search')+'<span class="h-sr-only">Fertigung durchsuchen</span><input type="search" data-production-filter="search" placeholder="BA oder Artikel in der Fertigung suchen …">';header.prepend(search);}
    if(!view.querySelector('.h-production-overview'))search?.remove();
    const tabs=[...view.querySelectorAll('button[onclick^="setProductionSubTab"]')].filter(b=>!b.disabled);
    const key=production?tabs.map(b=>b.getAttribute('onclick')+b.className).join('|'):'';
    let sub=nav.querySelector('.h-production-subnav');
    if(sub?.dataset.key===key)return;
    sub?.remove();if(!production||!tabs.length)return;
    sub=document.createElement('div');sub.className='h-production-subnav';sub.dataset.key=key;
    const available=tabs.filter(b=>/counts|orders|readiness|protocol/.test(b.getAttribute('onclick')));
    for(const button of available){const action=button.getAttribute('onclick');const a=document.createElement('a');a.href='#view';a.textContent=action.includes("'counts'")?'Produktionsübersicht':action.includes("'orders'")?'BA / OP':button.textContent;a.className=button.classList.contains('humbel-subtab-active')?'h-subnav-active':'';a.addEventListener('click',e=>{e.preventDefault();[...view.querySelectorAll('button[onclick]')].find(b=>b.getAttribute('onclick')===action)?.click();});sub.appendChild(a);}
    const productionButton=[...nav.querySelectorAll('button')].find(b=>b.getAttribute('onclick')?.includes("'produktion'"));productionButton?.after(sub);
  }
  document.addEventListener('DOMContentLoaded',()=>{decorate();workspace();new MutationObserver(decorate).observe(document.getElementById('tabs'),{childList:true,subtree:true});new MutationObserver(workspace).observe(document.getElementById('view'),{childList:true,subtree:true});});
})(window);
