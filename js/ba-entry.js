/* Phase 1 entry adapter: reuses the operations barcode and immutable snapshot UI.
   Disabled until schema, trusted first-OP grants and security review are approved. */
(function (root) {
  'use strict';
  const html = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  function create({client, actor, role, redraw, hasAccess, renderSnapshot}) {
    let user = null, generation = 0, busy = false, error = '', notice = '', snapshot = null;
    let fields = {ba:'', article:'', quantity:''}, lookup = {barcode:'', ba:'', op:''};
    const enabled = () => root.HUMBEL_BA_ENTRY_ENABLED === true && root.HUMBEL_OPERATIONS_ENABLED === true && root.HUMBEL_ARTICLE_MASTER_ENABLED === true;
    function reset() {
      generation++; user = null; busy = false; error = notice = ''; snapshot = null;
      fields = {ba:'', article:'', quantity:''}; lookup = {barcode:'', ba:'', op:''};
    }
    function access() {
      if (user !== actor()) { reset(); user = actor(); }
      return !!(enabled() && user && role() !== 'tool_scanner' && hasAccess());
    }
    function collect(form) {
      const values = new FormData(form), target = form.id === 'ba-entry-create' ? fields : lookup;
      for (const key of Object.keys(target)) if (values.has(key)) target[key] = String(values.get(key) || '');
    }
    async function execute(action, params) {
      if (!access() || busy) return;
      const ticket = generation, identity = user;
      busy = true; error = notice = ''; redraw();
      try {
        const result = await client.rpc('ba_entry_command', {p_action:action, ...params});
        if (ticket !== generation || identity !== actor() || !enabled()) return;
        if (result.error) {
          const messages = {'42501':'Keine gültige Berechtigung für diesen BA oder die erste OP.', '22023':'Artikel oder BA nicht eindeutig bzw. Eingaben ungültig. Freigegebenen Artikel und erste OP prüfen.', '23505':'BA bereits vorhanden. Bitte über die BA-Suche öffnen.', '40001':'Stand inzwischen geändert. Bitte erneut versuchen.'};
          throw Error(messages[result.error.code] || 'Speichern / Laden fehlgeschlagen. Eingaben bleiben erhalten.');
        }
        snapshot = result.data;
        const operation=snapshot?.op_rows?.find(o=>o.op_number===snapshot.selected_op);
        if(operation&&root.HUMBEL_OP_ROUTING_ENABLED===true)root.dispatchEvent(new CustomEvent('humbel-operation-open',{detail:{operation_id:operation.id}}));
        if (action === 'create') {
          notice = snapshot.existing ? 'Vorhandenen BA geöffnet. Keine Daten geändert.' : 'BA mit freigegebenem Artikel- und OP-Snapshot angelegt.';
          if (!snapshot.existing) fields = {ba:'', article:'', quantity:''};
        } else notice = 'BA / OP geöffnet. Keine Mengenbuchung.';
      } catch (e) {
        if (ticket === generation && identity === actor()) error = e.message;
      } finally {
        if (ticket === generation && identity === actor()) { busy = false; redraw(); }
      }
    }
    function render() {
      if (!access()) return '';
      return `<section id="ba-entry" aria-label="BA-Erstanlage"><style>
        #ba-entry{max-width:100%;overflow-wrap:anywhere}#ba-entry *{box-sizing:border-box;min-width:0}
        #ba-entry label{display:block;margin:12px 0}#ba-entry input{display:block;width:100%;max-width:100%;padding:10px;border:1px solid #94a3b8;border-radius:6px}
        #ba-entry button{min-height:44px;margin:4px;padding:8px 12px;border:1px solid #94a3b8;border-radius:6px}
        #ba-entry fieldset{border:1px solid #94a3b8;padding:12px;margin:12px 0}#ba-entry dd{white-space:pre-wrap}
        #ba-entry [data-selected-op=true]{outline:3px solid #0d4682;padding:12px}.operation-barcode{max-width:100%;overflow-x:auto;margin:12px 0}.operation-barcode svg{display:block;background:white}
        </style><h3>BA an der ersten OP anlegen</h3>
        <p>Humbel-BA-Nummer, Artikelnummer und Stückzahl eingeben. Die aktuelle freigegebene Revision und ihre OPs werden übernommen. Termin und Maschinenwahl sind hier nicht erforderlich.</p>
        ${error ? `<p role="alert">${html(error)}</p>` : ''}${notice ? `<p role="status">${html(notice)}</p>` : ''}
        <form id="ba-entry-create"><fieldset ${busy?'disabled':''}><legend>Neuer BA</legend>
          <label>BA-Nummer<input name="ba" inputmode="numeric" maxlength="24" required value="${html(fields.ba)}"></label>
          <label>Artikelnummer<input name="article" maxlength="80" required value="${html(fields.article)}"></label>
          <label>Stückzahl<input name="quantity" type="number" min="1" max="2147483647" step="1" required value="${html(fields.quantity)}"></label>
          <button>BA anlegen / vorhandenen öffnen</button></fieldset></form>
        <form id="ba-entry-scan"><fieldset ${busy?'disabled':''}><legend>Barcode öffnen</legend>
          <label>BA-/OP-Barcode<input name="barcode" autocomplete="off" maxlength="30" value="${html(lookup.barcode)}"></label><button>Barcode öffnen</button></fieldset></form>
        <form id="ba-entry-search"><fieldset ${busy?'disabled':''}><legend>Manuelle BA-/OP-Suche</legend>
          <label>BA für Suche<input name="ba" inputmode="numeric" maxlength="24" value="${html(lookup.ba)}"></label>
          <label>OP-Nummer (optional)<input name="op" type="number" min="1" max="999" step="1" value="${html(lookup.op)}"></label><button>BA / OP suchen</button></fieldset></form>
        ${snapshot ? snapshot.legacy ? `<section aria-label="Bestehender BA"><h4>BA ${html(snapshot.ba_number)}</h4><p>Bestandsauftrag geöffnet. Fertigungsdaten und bestehende Zähler bleiben im bisherigen Auftragsbereich erhalten; kein OP-Snapshot wurde nachträglich erzeugt.</p></section>` : `<p>BA-Hauptverantwortung: Zuordnungsregel noch offen. Keine automatische AL-Zuordnung.</p>${renderSnapshot(snapshot)}` : ''}</section>`;
    }
    function bind(document) {
      for (const type of ['input','change']) document.addEventListener(type, event => {
        const form = event.target.closest('#ba-entry-create,#ba-entry-scan,#ba-entry-search');
        if (form && !busy) collect(form);
      });
      document.addEventListener('submit', event => {
        if (!['ba-entry-create','ba-entry-scan','ba-entry-search'].includes(event.target.id)) return;
        event.preventDefault(); if (!access() || busy) return; collect(event.target);
        try {
          const ops = root.HumbelOperations;
          if (event.target.id === 'ba-entry-create') {
            const parsed = ops.parseBarcode(fields.ba), quantity = Number(fields.quantity);
            if (parsed.op !== null || !fields.article.trim() || !Number.isSafeInteger(quantity) || quantity < 1 || quantity > 2147483647) throw Error('BA-Nummer, Artikelnummer und positive ganze Stückzahl erforderlich.');
            void execute('create', {p_ba_number:parsed.ba, p_article_number:fields.article.trim(), p_quantity:quantity});
          } else {
            const parsed = event.target.id === 'ba-entry-scan' ? ops.parseBarcode(lookup.barcode) : ops.parseBarcode(ops.barcode(lookup.ba, lookup.op === '' ? null : Number(lookup.op)));
            void execute('open', {p_ba_number:parsed.ba, p_op_number:parsed.op});
          }
        } catch (e) { error = e.message; notice = ''; redraw(); }
      });
    }
    return {enabled, hasAccess:access, reset, render, bind};
  }
  root.HumbelBAEntry = {create};
  if (typeof module !== 'undefined') module.exports = root.HumbelBAEntry;
})(typeof window === 'undefined' ? globalThis : window);
