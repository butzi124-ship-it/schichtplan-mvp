/* Issue #8: local UI only; both feature flags and reviewed SQL are required. */
(function(root){
 'use strict';
 const textFields=['activity','description','machine_group','technical_notes','inspection_notes','programs','fixtures','tools','drawings','equipment'];
 const labels={activity:'Tätigkeit',description:'Arbeitsgangbeschreibung',machine_group:'Maschinengruppe',technical_notes:'Technische Hinweise',inspection_notes:'Prüfanweisungen',programs:'CNC-Programme',fixtures:'Spannmittel',tools:'Werkzeuge',drawings:'Zeichnungen / Arbeitsanweisungen',equipment:'Betriebsmittel'};
 const escape=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 function opNumber(n){if(!Number.isInteger(n)||n<1||n>999)throw Error('OP-Nummer muss zwischen 001 und 999 liegen.');return String(n).padStart(3,'0');}
 function parseBarcode(value){
  if(typeof value!=='string')throw Error('Ungültiger BA-/OP-Barcode.');
  const s=value.replace(/[\r\n]+$/,'');
  const m=/^([0-9]{1,24})(?:\.([0-9]{3}))?$/.exec(s);
  if(!m||m[2]==='000')throw Error('Barcode erwartet BA oder BA.OP mit dreistelliger OP.');
  return{ba:m[1],op:m[2]?Number(m[2]):null,value:s};
 }
 function barcode(ba,n=null){const p=parseBarcode(ba);if(p.op!==null)throw Error('BA enthält bereits eine OP.');return p.ba+(n===null?'':'.'+opNumber(n));}
 function blank(number=10,predecessor=null){return{number,predecessor,...Object.fromEntries(textFields.map(k=>[k,''])),machine_id:null,tr:{value:null,unit:null},te:{value:null,unit:null}};}
 function normalize(ops){
  if(!Array.isArray(ops)||ops.length>99)throw Error('Höchstens 99 Operationen zulässig.');
  const seen=new Set();return ops.map(o=>{
   if(!o||typeof o!=='object'||Array.isArray(o)||Object.keys(o).length!==15||Object.keys(o).some(k=>!['number','predecessor','machine_id','tr','te',...textFields].includes(k)))throw Error('Unzulässige OP-Felder.');
   opNumber(o.number);if(seen.has(o.number))throw Error('OP-Nummern müssen eindeutig sein.');
   if(o.predecessor!==null&&!seen.has(o.predecessor))throw Error('Vorgänger muss vor dieser OP liegen.');
   const out={number:o.number,predecessor:o.predecessor,machine_id:o.machine_id};
   if(o.machine_id!==null&&!/^[0-9a-f]{8}(?:-[0-9a-f]{4}){3}-[0-9a-f]{12}$/i.test(o.machine_id))throw Error('Ungültige Maschinen-ID.');
   for(const k of textFields){if(typeof o[k]!=='string'||o[k].length>12000)throw Error('OP-Texte: höchstens 12000 Zeichen.');out[k]=o[k].trim();}
   if(!out.activity||!out.machine_group)throw Error('Tätigkeit und Maschinengruppe sind erforderlich.');
   for(const k of ['tr','te']){const t=o[k];if(!t||Object.keys(t).length!==2||!Object.hasOwn(t,'value')||!Object.hasOwn(t,'unit')||!(t.value===null&&t.unit===null)&&!(typeof t.value==='number'&&Number.isFinite(t.value)&&t.value>=0&&t.value<=1000000&&['s','min','h'].includes(t.unit)))throw Error('Tr/Te benötigen einen nichtnegativen Wert und s, min oder h; unbekannt bleibt leer.');out[k]={...t};}
   seen.add(o.number);return out;
  });
 }
 function drawBarcodes(document){
  if(typeof root.JsBarcode!=='function')return;
  for(const svg of document.querySelectorAll('svg[data-ba-barcode]')){
   if(svg.dataset.rendered===svg.dataset.baBarcode)continue;
   svg.dataset.rendered=svg.dataset.baBarcode;
   root.JsBarcode(svg,svg.dataset.baBarcode,{format:'CODE128',displayValue:true,width:2,height:55,margin:12,fontSize:16});
  }
 }
 function symbol(value){return `<figure class="operation-barcode"><svg data-ba-barcode="${escape(value)}" role="img" aria-label="Barcode ${escape(value)}"></svg><figcaption>${escape(value)}</figcaption></figure>`;}
 function create({client,actor,role,redraw,canCreate,canEnter=canCreate,machines,refreshOrders}){
  let session=null,generation=0,row=null,ops=null,busy=false,error='',notice='',released=[],chosen=null,snapshot=null;
  let search='',scan='',orderForm={ba:'',quantity:'',due:'',machine:''};
  const entry=root.HumbelBAEntry?.create({client,actor,role,redraw,hasAccess:canEnter,renderSnapshot});
  const enabled=()=>root.HUMBEL_OPERATIONS_ENABLED===true&&root.HUMBEL_ARTICLE_MASTER_ENABLED===true;
  function reset(){entry?.reset();generation++;session=null;row=null;ops=null;busy=false;error=notice=search=scan='';released=[];chosen=null;snapshot=null;orderForm={ba:'',quantity:'',due:'',machine:''};}
  function active(){if(session!==actor()){reset();session=actor();}return enabled()&&!!session&&role()!=='tool_scanner';}
  function allowed(){return active()&&canCreate();}
  async function rpc(name,params){const result=await client.rpc(name,params);if(result.error){const code=result.error.code;throw Error(code==='23505'?'Diese BA-Nummer oder Revision ist bereits vorhanden.':code==='40001'?'Stand inzwischen geändert. Neu laden; Eingaben bleiben erhalten.':code==='42501'?'Keine gültige Berechtigung.':code==='22023'||code==='22007'?'Ungültige Eingaben oder Operationsplan.':'Speichern / Laden fehlgeschlagen. Eingaben bleiben erhalten.');}return result.data;}
  async function run(fn){if(!active()||busy)return;busy=true;error=notice='';const ticket=generation;redraw();try{await fn(ticket);}catch(e){if(ticket===generation)error=e.message;}finally{if(ticket===generation){busy=false;redraw();drawBarcodes(root.document);}}}
  const current=t=>t===generation&&session===actor()&&enabled();
  function renderArticle(selected){
   if(!active()||!selected)return '';
   if(row?.id!==selected.id){row=selected;ops=null;error=notice='';}else row=selected;
   const editable=row.status==='draft';
   return `${styles()}<section id="operation-plan" aria-label="Operationsplan"><h3>Operationsplan · Artikelrevision ${row.revision} · ${escape(row.status)}</h3><p>OP ist ein Arbeitsgang, keine Spannung. Mehrere Spannungen und Mitarbeiterzähler bleiben möglich. Zeiten ohne Nachweis bleiben nicht angegeben.</p>
    <button type="button" data-ops="load" ${busy?'disabled':''}>Operationsplan laden</button>${messages()}
    ${ops===null?'':`<p>${ops.length} Operationen · Reihenfolge von oben nach unten</p>${ops.map((o,i)=>editor(o,i,editable)).join('')}
    ${editable?`<button type="button" data-ops="add" ${busy?'disabled':''}>OP hinzufügen</button><button type="button" data-ops="save" ${busy?'disabled':''}>Operationsplan speichern</button>`:''}`}</section>`;
  }
  function styles(){return `<style>#operation-plan,#operations-orders{max-width:100%;overflow-wrap:anywhere}#operation-plan *,#operations-orders *{box-sizing:border-box;min-width:0}#operation-plan fieldset{border:1px solid #94a3b8;padding:12px;margin:12px 0}#operation-plan label,#operations-orders label{display:block;margin:10px 0}#operation-plan input,#operation-plan textarea,#operation-plan select,#operations-orders input,#operations-orders select{display:block;width:100%;max-width:100%;padding:10px;border:1px solid #94a3b8;border-radius:6px}#operation-plan button,#operations-orders button{min-height:44px;margin:4px;padding:8px 12px;border:1px solid #94a3b8;border-radius:6px}#operation-plan button:disabled,#operations-orders button:disabled{opacity:.5}.operation-barcode{margin:12px 0;max-width:100%;overflow-x:auto}.operation-barcode svg{display:block;background:white}#operations-orders [data-selected-op=true]{outline:3px solid #0d4682;background:#eff6ff;padding:12px}#operations-orders dd{white-space:pre-wrap;margin:0 0 8px}</style>`;}
 function messages(){return `${error?`<p role="alert">${escape(error)}</p>`:''}${notice?`<p role="status">${escape(notice)}</p>`:''}`;}
  function editor(o,i,editable){
   const disabled=!editable||busy?'disabled':'';
   return `<fieldset data-operation="${i}" ${disabled}><legend>Position ${i+1} · OP ${escape(Number.isInteger(o.number)&&o.number>0&&o.number<1000?opNumber(o.number):o.number)} · ${escape(row.status)} / Revision ${row.revision}</legend>
    <label>OP-Nummer<input type="number" min="1" max="999" data-op-field="number" value="${o.number}"></label>
    <label>Vorgänger-OP<input type="number" min="1" max="999" data-op-field="predecessor" value="${o.predecessor??''}"></label>
    ${textFields.map(k=>`<label for="op-${i}-${k}">${labels[k]}</label><textarea id="op-${i}-${k}" data-op-field="${k}" maxlength="12000" rows="${['description','technical_notes','inspection_notes'].includes(k)?4:2}">${escape(o[k])}</textarea>`).join('')}
    <label for="op-${i}-machine">Konkrete Maschine (optional)</label><select id="op-${i}-machine" data-op-field="machine_id"><option value="">Keine konkrete Maschine</option>${machineOptions(row.department_id,o.machine_id)}</select>
    ${['tr','te'].map(k=>`<label>${k==='tr'?'Tr Soll-Rüstzeit':'Te Soll-Bearbeitungszeit je Stück'}<input type="number" min="0" max="1000000" step="any" data-op-field="${k}.value" value="${o[k].value??''}" placeholder="Nicht angegeben"></label><label for="op-${i}-${k}-unit">${k.toUpperCase()} Zeiteinheit</label><select id="op-${i}-${k}-unit" data-op-field="${k}.unit"><option value="">Nicht angegeben</option>${['s','min','h'].map(u=>`<option ${o[k].unit===u?'selected':''}>${u}</option>`).join('')}</select>`).join('')}
    ${editable?`<button type="button" data-ops="up" data-index="${i}" ${i===0?'disabled':''}>OP nach oben</button><button type="button" data-ops="down" data-index="${i}" ${i===ops.length-1?'disabled':''}>OP nach unten</button><button type="button" data-ops="remove" data-index="${i}">OP entfernen</button>`:''}</fieldset>`;
  }
  function machineOptions(dept,id){return machines().filter(m=>m.department_id===dept&&m.active!==false).map(m=>`<option value="${escape(m.id)}" ${m.id===id?'selected':''}>${escape(m.name||m.machine_code)}</option>`).join('');}
  function orderPanel(){
   if(entry?.enabled())return entry.render();
   if(!allowed())return '';
   return `${styles()}<section id="operations-orders" aria-label="BA aus Artikel"><h3>Neue BA aus freigegebenem Artikel</h3><p>Neue Nummer, Menge und Termin eingeben. Keine alten Zähler, Checklisten oder Spannungen werden kopiert. Der bisherige BA-Pfad bleibt unten verfügbar.</p>${messages()}
    <form id="ops-search"><label>Artikelnummer für BA<input name="article" value="${escape(search)}" maxlength="80"></label><button ${busy?'disabled':''}>Freigegebene Revision suchen</button></form>
    <label>Freigegebene Artikelrevision<select id="ops-revision" ${busy?'disabled':''}><option value="">Bitte auswählen</option>${released.map(r=>`<option value="${escape(r.id)}" ${chosen?.id===r.id?'selected':''}>${escape(r.article_number)} · Revision ${r.revision} · ${escape(r.data.designation)}</option>`).join('')}</select></label>
    ${chosen?`<form id="ops-create"><label>Neue BA-Nummer<input name="ba" inputmode="numeric" maxlength="24" value="${escape(orderForm.ba)}" required></label><label>Neue Stückzahl<input name="quantity" type="number" min="1" max="2147483647" step="1" value="${escape(orderForm.quantity)}" required></label><label>Liefer-/Endtermin<input name="due" type="date" value="${escape(orderForm.due)}" required></label><label>BA-Maschine<select name="machine" required><option value="">Bitte auswählen</option>${machineOptions(chosen.department_id,orderForm.machine)}</select></label><button ${busy?'disabled':''}>BA mit OP-Snapshot erzeugen</button></form>`:''}
    <form id="ops-scan"><label>BA-/OP-Barcode scannen<input name="barcode" value="${escape(scan)}" autocomplete="off" maxlength="30"></label><button ${busy?'disabled':''}>BA / OP öffnen</button></form>
    ${snapshot?renderSnapshot(snapshot):''}</section>`;
  }
  function renderSnapshot(s){return `<section aria-label="BA-Operationssnapshot"><h4>BA ${escape(s.ba_number)} · ${s.quantity} Stück · Termin ${s.due_date?escape(s.due_date):'Nicht geplant'} · Artikelrevision ${s.article_revision} · unveränderlicher Snapshot</h4><p>Freigabe: ${escape(s.released_at)} · Verantwortlich: ${escape(s.released_by)}</p>${symbol(barcode(s.ba_number))}<p>Artikel: ${escape(s.article_data.designation)} · ${escape(s.article_data.material)}</p>${s.op_rows.map(o=>`<section aria-label="BA OP ${opNumber(o.op_number)}" ${s.selected_op===o.op_number?'data-selected-op="true" aria-current="step"':''}><h5>OP ${opNumber(o.op_number)} · Position ${o.position} · Vorgänger ${o.predecessor===null?'–':opNumber(o.predecessor)}</h5>${symbol(barcode(s.ba_number,o.op_number))}<p>${escape(o.data.activity)} · Gruppe ${escape(o.data.machine_group)}</p><dl>${textFields.filter(k=>!['activity','machine_group'].includes(k)).map(k=>`<dt>${labels[k]}</dt><dd>${escape(o.data[k])||'–'}</dd>`).join('')}<dt>Maschine</dt><dd>${escape(o.data.machine_id)||'Nicht festgelegt'}</dd>${['tr','te'].map(k=>`<dt>${k==='tr'?'Tr':'Te je Stück'}</dt><dd>${o.data[k].value===null?'Nicht angegeben':escape(o.data[k].value)+' '+escape(o.data[k].unit)}</dd>`).join('')}</dl>${root.HUMBEL_OP_ROUTING_ENABLED===true?`<button type="button" data-routing-open="${escape(o.id)}">Maschinenzuordnung OP ${opNumber(o.op_number)}</button>`:''}<p>Spannungen separat im bestehenden Auftrag anlegen und mit dieser OP-Nummer kennzeichnen. Keine Zählerbuchung durch Scan.</p></section>`).join('')}</section>`;}
  function collectOrder(form){const f=new FormData(form);for(const k of ['ba','quantity','due','machine'])orderForm[k]=String(f.get(k)||'');}
  function bind(document){
   entry?.bind(document);
   const observer=new MutationObserver(()=>{if(enabled())drawBarcodes(document);});observer.observe(document.getElementById('view')||document.body,{childList:true,subtree:true});
   document.addEventListener('input',e=>{
    if(e.target.closest('#ops-create'))collectOrder(e.target.closest('form'));
    const box=e.target.closest('[data-operation]');if(!box||!ops||busy)return;const o=ops[Number(box.dataset.operation)],k=e.target.dataset.opField;if(!k)return;
    if(k.includes('.')){const [a,b]=k.split('.');o[a][b]=e.target.value===''?null:(b==='value'?Number(e.target.value):e.target.value);}
    else o[k]=['number','predecessor'].includes(k)?(e.target.value===''?null:Number(e.target.value)):k==='machine_id'?(e.target.value||null):e.target.value;
   });
   document.addEventListener('change',e=>{
    if(e.target.closest('#ops-create'))collectOrder(e.target.closest('form'));
    if(e.target.id==='ops-revision'){chosen=released.find(r=>r.id===e.target.value)||null;orderForm={ba:'',quantity:'',due:'',machine:''};snapshot=null;redraw();}
    if(e.target.dataset.opField)e.target.dispatchEvent(new Event('input',{bubbles:true}));
   });
   document.addEventListener('click',e=>{const b=e.target.closest('[data-ops]');if(!b||!active()||busy)return;
    const action=b.dataset.ops;
    if(action==='load')void run(async t=>{const r=await rpc('operation_plan',{p_revision_id:row.id});if(current(t)){ops=r.operations;row.row_version=r.row_version;}});
    if(action==='add'&&row?.status==='draft'&&ops){const n=ops.length?Math.max(...ops.map(o=>o.number))+10:10;if(n>999){error='Keine nächste dreistellige OP-Nummer verfügbar.';redraw();return;}ops.push(blank(n,ops.at(-1)?.number??null));redraw();}
    if(action==='save'&&row?.status==='draft')void run(async t=>{const data=normalize(ops);const r=await rpc('operation_plan',{p_revision_id:row.id,p_operations:data,p_expected_version:row.row_version});if(current(t)){ops=r.operations;row.row_version=r.row_version;notice='Operationsplan gespeichert.';}});
    if(['remove','up','down'].includes(action)&&ops&&row?.status==='draft'){const i=Number(b.dataset.index);if(action==='remove')ops.splice(i,1);else{const j=i+(action==='up'?-1:1);if(j<0||j>=ops.length)return;[ops[i],ops[j]]=[ops[j],ops[i]];}ops.forEach((o,j)=>o.predecessor=j?ops[j-1].number:null);notice='Reihenfolge geändert; Vorgänger als fortlaufende Kette neu gesetzt. Noch nicht gespeichert.';redraw();}
   });
   document.addEventListener('submit',e=>{
    if(!['ops-search','ops-create','ops-scan'].includes(e.target.id))return;e.preventDefault();if(!allowed()||busy)return;
    if(e.target.id==='ops-search'){search=String(new FormData(e.target).get('article')||'');void run(async t=>{const r=await rpc('article_master_list',{p_query:search});if(current(t)){released=r.filter(v=>v.status==='released');chosen=null;notice=released.length?'Freigegebene Revision auswählen.':'Keine freigegebene Revision gefunden.';}});}
    if(e.target.id==='ops-create'){collectOrder(e.target);void run(async t=>{if(!chosen)throw Error('Freigegebene Revision wählen.');const p=parseBarcode(orderForm.ba);if(p.op!==null)throw Error('BA darf keine OP enthalten.');const quantity=Number(orderForm.quantity);if(!Number.isSafeInteger(quantity)||quantity<1||quantity>2147483647)throw Error('Positive ganze Stückzahl erforderlich.');if(!/^\d{4}-\d{2}-\d{2}$/.test(orderForm.due))throw Error('Gültiger Termin erforderlich.');const r=await rpc('order_from_article',{p_revision_id:chosen.id,p_ba_number:p.ba,p_quantity:quantity,p_due_date:orderForm.due,p_machine_id:orderForm.machine});if(current(t)){snapshot=r;orderForm={ba:'',quantity:'',due:'',machine:''};notice='Neue BA mit unveränderlichem OP-Snapshot angelegt.';await refreshOrders();}});}
    if(e.target.id==='ops-scan'){scan=String(new FormData(e.target).get('barcode')||'');void run(async t=>{const p=parseBarcode(scan);const r=await rpc('operation_scan',{p_barcode:p.value});if(current(t)){snapshot=r;notice='BA / OP geöffnet. Keine Mengenbuchung.';}});}
   });
  }
  return{enabled,reset,renderArticle,orderPanel,bind,hasEntryAccess:()=>!!entry?.hasAccess(),articleCommandName:()=>enabled()?'operations_article_command':'article_master_command'};
 }
 const api={create,opNumber,parseBarcode,barcode,blank,normalize};root.HumbelOperations=api;if(typeof module!=='undefined')module.exports=api;
})(typeof window==='undefined'?globalThis:window);
