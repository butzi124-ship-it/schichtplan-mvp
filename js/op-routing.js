/* Phase 2: existing OP execution and counters remain authoritative. No localStorage. */
(function(root){
 'use strict';
 const html=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const date=v=>v?new Date(v).toLocaleString('de-DE',{timeZone:'Europe/Berlin'}):'Offen';
 function create({client,actor,role,redraw,changed=async()=>{}}){
  let pollTimer=null;
  let user=null,generation=0,loaded=false,loading=false,busy=false,rows=[],notes=[],reasons=[],error='',notice='',selected='';
  const forms=new Map(),requests=new Map();
  const enabled=()=>root.HUMBEL_OP_ROUTING_ENABLED===true&&root.HUMBEL_MACHINE_READINESS_ENABLED===true&&root.HUMBEL_OPERATIONS_ENABLED===true&&root.HUMBEL_ARTICLE_MASTER_ENABLED===true;
  function reset(){if(pollTimer)root.clearInterval(pollTimer);pollTimer=null;generation++;user=null;loaded=loading=busy=false;rows=[];notes=[];reasons=[];error=notice=selected='';forms.clear();requests.clear();}
  function valid(){return enabled()&&!!user&&user===actor()&&role()!=='tool_scanner';}
  function hasAccess(){return !!(valid()&&(rows.length||notes.length));}
  async function rpc(name,params){const r=await client.rpc(name,params);if(r.error){const messages={'42501':'Keine gültige OP-/Maschinenberechtigung oder verbindliche Maschinenbindung.','40001':'Stand inzwischen geändert. Neu laden; Eingaben bleiben erhalten.','22023':'Zuordnung nicht zulässig. Status, Abbruchgrund und empfangenden AL prüfen.'};throw Error(messages[r.error.code]||'Maschinenzuordnung nicht verfügbar. Eingaben bleiben erhalten.');}return r.data;}
  async function load(ticket){const data=await rpc('op_routing_view',{});if(ticket!==generation||!valid())return;rows=data.rows||[];notes=data.notifications||[];reasons=data.abort_reasons||[];}
  async function ensureSession(){
   if(!enabled()||!actor()||role()==='tool_scanner'){if(user)reset();return;}
   if(user!==actor()){reset();user=actor();}if(!pollTimer)pollTimer=root.setInterval(()=>{if(valid()&&!busy)void refresh();},30000);if(loaded||loading)return;
   loading=true;const t=generation;try{await load(t);}catch(e){if(t===generation)error=e.message;}finally{if(t===generation){loaded=true;loading=false;redraw();}}
  }
  async function refresh(){if(!valid()||busy)return;busy=true;error='';const t=generation;redraw();try{await load(t);}catch(e){if(t===generation)error=e.message;}finally{if(t===generation){busy=false;redraw();}}}
  function values(r){if(!forms.has(r.operation_id))forms.set(r.operation_id,{machine:r.machine_id||r.default_machine_id||'',reason:'',abort:''});return forms.get(r.operation_id);}
  function collect(form){const data=new FormData(form),r=rows.find(r=>r.operation_id===form.dataset.routingForm);if(!r)return;const f=values(r);for(const k of ['machine','reason','abort'])f[k]=String(data.get(k)||'');}
  function card(r){const f=values(r),number=String(r.op_number).padStart(3,'0');
   return `<article class="routing-card" aria-label="Maschinenzuordnung BA ${html(r.ba_number)} OP ${number}" data-routing-id="${html(r.operation_id)}">
    <h4>BA ${html(r.ba_number)} · OP ${number} · ${html(r.activity)}</h4><p>Ursprüngliche Abteilung: ${html(r.original_department)} · Ausführende Abteilung: ${html(r.executing_department)}</p>
    <p>Aktuelle Maschine: ${html(r.machine_name||'Noch nicht zugeordnet')}</p><p>BA-Hauptverantwortung bleibt unverändert. OP-Status: ${html(r.work_status)}. Mengen und bisherige Zeiten bleiben im vorhandenen Fertigungsprotokoll.</p>
    <p>${r.binding?'Verbindliche BA-OP-Maschine – für Mitarbeiter gesperrt.':'Maschine aus dem Artikel vorbelegt; zulässiger Wechsel innerhalb der ausführenden Abteilung möglich.'}${r.mandatory_machine_id?' Zwingende technische Maschinenbindung vorhanden.':''}</p>
    <form data-routing-form="${html(r.operation_id)}"><fieldset ${busy||!r.mutable?'disabled':''}><legend>Maschine und Verantwortung</legend>
    <label>Zielmaschine<select name="machine" ${!r.can_manage&&r.binding?'disabled':''}><option value="">Zulässige Maschine wählen</option>${r.machine_id&&!r.machines.some(m=>m.id===r.machine_id)?`<option value="${html(r.machine_id)}" selected>${html(r.machine_name||r.machine_id)} (gesperrt)</option>`:''}${r.machines.map(m=>`<option value="${html(m.id)}" ${f.machine===m.id?'selected':''}>${html(m.name)}</option>`).join('')}</select></label>
    ${!r.machines.length?'<p>Keine freigegebene Zielmaschine verfügbar. Berechtigung / Maschinenbindung prüfen.</p>':''}
    <label>Begründung<textarea name="reason" maxlength="4000">${html(f.reason)}</textarea></label>
    ${r.active_setup?`<p>Aktive Rüstung: Beim Wechsel werden laufende persönliche Segmente zum selben Zeitpunkt beendet. Neue Rüstung anschließend ausdrücklich starten.</p><label>Rüst-Abbruchgrund<select name="abort"><option value="">Konfigurierten Grund wählen</option>${reasons.map(a=>`<option value="${html(a.code)}" ${f.abort===a.code?'selected':''}>${html(a.label)}${a.requires_note?' – Erläuterung erforderlich':''}</option>`).join('')}</select></label>`:''}
    <button type="button" data-routing-action="select" ${r.binding||!r.machines.length||r.can_manage?'disabled':''}>Maschine wählen</button>
    ${r.can_manage?`<button type="button" data-routing-action="fix" ${!r.machines.length?'disabled':''}>Verbindlich festlegen (AL)</button>`:''}
    ${r.can_return?'<button type="button" data-routing-action="return">An ursprüngliche Abteilung zurückgeben</button>':''}
    </fieldset></form>
    <details><summary>Maschinenabschnitte (${r.sections.length})</summary><ol>${r.sections.map(s=>`<li>Maschine ${html(s.machine_id)} · ${html(date(s.opened_at))} → ${html(date(s.closed_at))}${s.close_reason?' · '+html(s.close_reason):''}</li>`).join('')}</ol><p>Die Abschnitte führen keinen zweiten Stückzahlzähler. Personenbezogene Zeitstarts und die genaue Spannungs-/Zählerzuordnung folgen in Phase 3/4.</p></details>
    <details><summary>Zuordnungshistorie (${r.events.length})</summary><ol>${r.events.map(e=>`<li>${html(date(e.happened_at))} · ${html(e.action)} · ${html(e.payload.reason)} · ${e.interrupted_segments.length} persönliche Segmente unterbrochen</li>`).join('')}</ol></details>
   </article>`;
  }
  function render(){if(!valid())return '';return `<section id="op-routing" aria-label="OP-Maschinenzuordnung"><style>
    #op-routing{max-width:100%;overflow-wrap:anywhere}#op-routing *{box-sizing:border-box;min-width:0}#op-routing h3,#op-routing h4{font-weight:700;margin:12px 0}
    #op-routing .routing-card{border:1px solid #94a3b8;border-radius:8px;padding:12px;margin:12px 0}#op-routing label{display:block;margin:10px 0}
    #op-routing input,#op-routing select,#op-routing textarea{display:block;width:100%;max-width:100%;padding:10px;border:1px solid #94a3b8;border-radius:5px}
    #op-routing textarea{min-height:75px}#op-routing button{min-height:44px;margin:4px;padding:8px 12px;border:1px solid #64748b;border-radius:5px}#op-routing button:disabled{opacity:.45}#op-routing fieldset{border:0;padding:0}
    </style><h3>OP-Maschinenzuordnung</h3><p>Zuordnung gilt für diese BA-OP. Artikel und BA-Hauptverantwortung bleiben erhalten. Ein Wechsel ist keine Arbeits- oder Materialfreigabe.</p>
    <button type="button" data-routing-reload ${busy?'disabled':''}>Zuordnungen neu laden</button><button type="button" data-routing-all>Alle berechtigten OPs</button>
    ${error?`<p role="alert">${html(error)}</p>`:''}${notice?`<p role="status">${html(notice)}</p>`:''}
    <section aria-label="AL-Benachrichtigungen"><h4>Benachrichtigungen</h4><ul>${notes.map(n=>`<li>BA ${html(n.ba_number)} · OP ${String(n.op_number).padStart(3,'0')} · ${n.kind==='returned'?'Zurückgegeben':'Operative Verantwortung übernommen'} · ${html(date(n.created_at))}</li>`).join('')||'<li>Keine Benachrichtigungen.</li>'}</ul></section>
    ${rows.filter(r=>!selected||r.operation_id===selected).map(card).join('')||'<p>Keine berechtigten OPs.</p>'}</section>`;}
  async function execute(action,r){
   if(!valid()||busy||!r.mutable)return;const f=values(r);if(!f.machine||!f.reason.trim()||r.active_setup&&f.machine!==r.machine_id&&!f.abort){error='Zielmaschine, Begründung und bei aktiver Rüstung ein konfigurierter Abbruchgrund erforderlich.';redraw();return;}
   const body={p_operation_id:r.operation_id,p_action:action,p_expected_version:r.row_version,p_machine_id:f.machine,p_reason:f.reason.trim(),p_abort_code:f.abort||null};
   const key=JSON.stringify(body);if(requests.get(r.operation_id)?.key!==key)requests.set(r.operation_id,{key,id:root.crypto.randomUUID()});body.p_request_id=requests.get(r.operation_id).id;
   busy=true;error=notice='';const t=generation;redraw();try{await rpc('op_routing_command',body);if(t!==generation||!valid())return;requests.delete(r.operation_id);forms.delete(r.operation_id);await load(t);await changed();notice='Maschinenzuordnung mit Audit gespeichert. Betroffene AL wurden innerhalb der App benachrichtigt; kein automatischer Rüst-/Fertigungsstart.';}
   catch(e){if(t===generation&&user===actor())error=e.message;}finally{if(t===generation){busy=false;redraw();}}
  }
  function bind(doc){
   for(const type of ['input','change'])doc.addEventListener(type,e=>{const form=e.target.closest('[data-routing-form]');if(form&&!busy)collect(form);});
   doc.addEventListener('submit',e=>{if(e.target.matches('[data-routing-form]'))e.preventDefault();});
   doc.addEventListener('click',e=>{
    const opener=e.target.closest('[data-routing-open]');if(opener&&enabled()){void (async()=>{await ensureSession();selected=opener.dataset.routingOpen;await refresh();if(hasAccess())root.setProductionSubTab('readiness');})();return;}
    if(e.target.closest('[data-routing-reload]'))void refresh();if(e.target.closest('[data-routing-all]')){selected='';redraw();}
    const b=e.target.closest('[data-routing-action]');if(!b||!valid()||busy)return;const form=b.closest('form');collect(form);const r=rows.find(r=>r.operation_id===form.dataset.routingForm);if(r)void execute(b.dataset.routingAction,r);
   });
  }
  return{enabled,reset,ensureSession,hasAccess,render,bind,refresh};
 }
 root.HumbelRouting={create};if(typeof module!=='undefined')module.exports=root.HumbelRouting;
})(typeof window==='undefined'?globalThis:window);
