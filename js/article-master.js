/* Issue #7: disabled until the reviewed DB contract and access provisioning are approved. */
(function (root) {
  'use strict';
  const fields = Object.freeze({
    designation: 'Artikelbezeichnung', customer: 'Kunde', drawing_number: 'Zeichnungsnummer',
    drawing_revision: 'Zeichnungsrevision', material: 'Material', dimensions: 'Abmessungen',
    specification: 'Spezifikation', technical_notes: 'Technische Beschreibung / Bearbeitungshinweise',
    programs: 'Programme', fixtures: 'Spannmittelreferenzen', tools: 'Werkzeugreferenzen',
    equipment: 'Betriebsmittelreferenzen', inspection_notes: 'Prüfhinweise', internal_notes: 'Interne Bemerkungen',
  });
  const statusLabel = { draft: 'Entwurf', released: 'Freigegeben', archived: 'Archiviert' };
  const html = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  function normalizeData(data) {
    if (!data || typeof data !== 'object' || Array.isArray(data) || Object.keys(data).some(k => !Object.hasOwn(fields, k))) throw Error('Nur wiederverwendbare Artikelfelder sind zulässig.');
    const result = {};
    for (const key of Object.keys(fields)) {
      if (data[key] !== undefined && typeof data[key] !== 'string') throw Error('Artikelfelder müssen Text enthalten.');
      result[key] = (data[key] || '').trim();
      if (result[key].length > 12000) throw Error('Ein Artikelfeld darf höchstens 12000 Zeichen enthalten.');
    }
    return result;
  }
  function create({ client, actor, role, redraw, enabled = () => root.HUMBEL_ARTICLE_MASTER_ENABLED === true }) {
    let user = null, generation = 0, loaded = false, loading = false, busy = false;
    let caps = { scopes: [], assignments: [] }, rows = [], selected = null, form = null, error = '', notice = '', query = '';
    function reset() {
      generation++; user = null; loaded = loading = busy = false;
      caps = { scopes: [], assignments: [] }; rows = []; selected = form = null; error = notice = query = '';
    }
    function valid() { return enabled() && user && user === actor() && role() !== 'tool_scanner'; }
    function access() { return !!(valid() && loaded && (caps.scopes.length || caps.assignments.length)); }
    function responsible(department) { return access() && caps.scopes.some(s => s.department_id === department); }
    function assigned(department, number) {
      return access() && caps.assignments.some(a => a.department_id === department && a.article_number.trim().toUpperCase() === number.trim().toUpperCase() && Date.parse(a.expires_at) > Date.now());
    }
    async function rpc(name, params) {
      const result = await client.rpc(name, params);
      if (result.error) {
        if (result.error.code === '23505') throw Error('Diese Artikelnummer oder eine offene Revision ist bereits vorhanden.');
        if (result.error.code === '40001') throw Error('Der Stand wurde inzwischen geändert. Bitte neu laden; Eingaben bleiben erhalten.');
        if (result.error.code === '42501') throw Error('Keine gültige Artikelberechtigung. Bitte Berechtigung prüfen und neu laden.');
        throw Error('Artikelstamm derzeit nicht verfügbar. Bitte erneut versuchen.');
      }
      return result.data;
    }
    async function ensureSession() {
      if (!enabled() || !actor() || role() === 'tool_scanner') { if (user) reset(); return; }
      if (user !== actor()) { reset(); user = actor(); }
      if (loaded || loading) return;
      loading = true; const ticket = generation;
      try {
        const result = await rpc('article_master_capabilities', {});
        if (ticket !== generation || user !== actor()) return;
        caps = { scopes: result.scopes || [], assignments: result.assignments || [] }; loaded = true;
      } catch (e) { if (ticket === generation) { error = e.message; loaded = true; } }
      finally { if (ticket === generation) { loading = false; redraw(); } }
    }
    async function list() {
      if (!access() || busy) return;
      busy = true; error = ''; const ticket = generation;
      try {
        const result = await rpc('article_master_list', { p_query: query });
        if (ticket !== generation || !valid()) return;
        rows = result || [];
        if (selected) selected = rows.find(r => r.id === selected.id) || null;
      } catch (e) { if (ticket === generation) { rows = []; selected = null; error = e.message; } }
      finally { if (ticket === generation) { busy = false; redraw(); } }
    }
    function open() { if (!access()) return; form = null; selected = null; return list(); }
    function edit(row) {
      if (!access()) return;
      if (row && (row.status !== 'draft' || !(responsible(row.department_id) || (row.created_by === user && assigned(row.department_id, row.article_number))))) return;
      form = row ? { ...row, data: { ...row.data } } : { department_id: caps.scopes[0]?.department_id || caps.assignments[0]?.department_id || '', article_number: caps.scopes.length ? '' : (caps.assignments[0]?.article_number || ''), data: normalizeData({}) };
      error = notice = ''; redraw();
    }
    function collect(element) {
      if (!form || !element) return;
      const values = new FormData(element);
      form.department_id = values.get('department_id') || '';
      form.article_number = values.get('article_number') || '';
      form.data = Object.fromEntries(Object.keys(fields).map(key => [key, values.get(key) || '']));
    }
    async function command(action, element) {
      if (!access() || busy) return;
      if (element) collect(element);
      const target = form || selected;
      if (!target) return;
      const isManager = responsible(target.department_id);
      if (['release', 'revise', 'archive'].includes(action) && !isManager) return;
      if (action === 'save' && !isManager && !assigned(target.department_id, target.article_number)) { error = 'Erfassung nur für den aktuell zugewiesenen Artikel erlaubt.'; redraw(); return; }
      let data;
      try {
        if (action === 'save') {
          if (!target.article_number.trim() || target.article_number.trim().length > 80) throw Error('Artikelnummer mit 1 bis 80 Zeichen eingeben.');
          data = normalizeData(target.data);
        }
      } catch (e) { error = e.message; redraw(); return; }
      busy = true; error = notice = ''; redraw(); const ticket = generation;
      try {
        const result = await rpc('article_master_command', {
          p_action: action === 'save' ? (target.id ? 'update' : 'create') : action,
          p_revision_id: target.id || null, p_department_id: target.department_id,
          p_article_number: target.article_number.trim(), p_data: data || null,
          p_expected_version: target.row_version || null,
        });
        if (ticket !== generation || !valid()) return;
        selected = result; form = null;
        rows = [result, ...rows.filter(r => r.id !== result.id)];
        notice = 'Gespeichert. Bestehende BA bleiben unverändert.';
      } catch (e) { if (ticket === generation) error = e.message; }
      finally { if (ticket === generation) { busy = false; redraw(); } }
    }
    function button(action, label) { return `<button type="button" data-article-action="${action}" ${busy ? 'disabled' : ''}>${label}</button>`; }
    function render() {
      if (!access()) return '<h2>Artikelstamm</h2><p>Keine gültige Artikelberechtigung.</p>';
      const scopes = new Map([...caps.scopes, ...caps.assignments].map(s => [s.department_id, s.department_name || s.department_id]));
      const row = selected;
      const canEdit = row?.status === 'draft' && (responsible(row.department_id) || (row.created_by === user && assigned(row.department_id, row.article_number)));
      const details = row ? `<section aria-label="Artikeldetails"><h3>${html(row.article_number)} · Revision ${row.revision} · ${statusLabel[row.status]}</h3>
        <p>Verantwortlich: ${html(row.created_by)} · Angelegt: ${html(row.created_at)} · Geändert: ${html(row.updated_at)}${row.released_at ? ` · Freigabe: ${html(row.released_by)} / ${html(row.released_at)}` : ''}${row.archived_at ? ` · Archivierung: ${html(row.archived_by)} / ${html(row.archived_at)}` : ''}</p>
        <dl>${Object.entries(fields).map(([key,label]) => `<dt>${label}</dt><dd>${html(row.data[key]) || '–'}</dd>`).join('')}</dl>
        ${canEdit ? button('edit', 'Entwurf bearbeiten') : ''}
        ${row.status === 'draft' && responsible(row.department_id) ? button('release', 'Freigeben') : ''}
        ${row.status === 'released' && responsible(row.department_id) ? button('revise', 'Neue Revision') : ''}
        ${row.status !== 'archived' && responsible(row.department_id) ? button('archive', 'Archivieren') : ''}</section>` : '';
      const editor = form ? `<form id="article-editor"><h3>${form.id ? 'Entwurf bearbeiten' : 'Neuer Artikel'}</h3>
        <label>Gültigkeitsbereich<select name="department_id" ${form.id ? 'disabled' : ''}>${[...scopes].map(([id,name]) => `<option value="${html(id)}" ${id===form.department_id?'selected':''}>${html(name)}</option>`).join('')}</select></label>
        ${form.id ? `<input type="hidden" name="department_id" value="${html(form.department_id)}">` : ''}
        <label>Artikelnummer<input name="article_number" required maxlength="80" value="${html(form.article_number)}" ${form.id ? 'readonly' : ''}></label>
        ${Object.entries(fields).map(([key,label]) => `<label for="article-${key}">${label}</label><textarea id="article-${key}" name="${key}" maxlength="12000" rows="${['technical_notes','specification','inspection_notes','internal_notes'].includes(key)?5:2}">${html(form.data[key])}</textarea>`).join('')}
        <p>Nur Artikeldaten. BA-Nummer, Mengen und Auftragstermine gehören in den jeweiligen Auftrag.</p>
        <button type="submit" ${busy?'disabled':''}>Entwurf speichern</button>${button('cancel','Abbrechen')}</form>` : details;
      return `<section id="article-master" aria-label="Artikelstamm"><style>
        #article-master{max-width:100%;overflow-wrap:anywhere}#article-master *{box-sizing:border-box;min-width:0}
        #article-master label{display:block;margin:12px 0}#article-master input,#article-master textarea,#article-master select{display:block;width:100%;max-width:100%;padding:10px;border:1px solid #94a3b8;border-radius:6px}
        #article-master textarea{resize:vertical}#article-master button{min-height:44px;padding:8px 14px;margin:4px;border:1px solid #94a3b8;border-radius:6px}#article-master button:disabled{opacity:.5}
        #article-master dl dd{white-space:pre-wrap;margin:0 0 12px}#article-master dl dt{font-weight:600}#article-master ul{list-style:none;padding:0}
        </style><h2>Artikelstamm</h2><p>Versionierte Stammdaten · Entwurf → Freigegeben → Archiviert</p>
        ${error?`<p role="alert">${html(error)}</p>`:''}${notice?`<p role="status">${html(notice)}</p>`:''}
        <form id="article-search"><label>Artikelnummer suchen<input name="query" maxlength="80" value="${html(query)}"></label><button ${busy?'disabled':''}>Suchen / Neu laden</button></form>
        ${button('new','Neuer Artikel')}<p>Bis zu 200 Revisionen. Suche nach Artikelnummer eingrenzen. Änderungen an freigegebenen Inhalten erfolgen ausschließlich als neue Revision.</p>
        <ul aria-label="Artikelrevisionen">${rows.map(r => `<li><button data-article-id="${html(r.id)}">${html(r.article_number)} · ${html(r.data.designation)} · Revision ${r.revision} · ${statusLabel[r.status]}</button></li>`).join('') || '<li>Keine Artikel geladen.</li>'}</ul>
        ${editor}</section>`;
    }
    function bind(document) {
      document.addEventListener('input', event => { if (event.target.closest('#article-editor')) collect(event.target.closest('form')); });
      document.addEventListener('change', event => { if (event.target.closest('#article-editor')) collect(event.target.closest('form')); });
      document.addEventListener('submit', event => {
        if (event.target.id === 'article-search') { event.preventDefault(); query = String(new FormData(event.target).get('query') || '').trim(); void list(); }
        if (event.target.id === 'article-editor') { event.preventDefault(); void command('save', event.target); }
      });
      document.addEventListener('click', event => {
        const element = event.target.closest('[data-article-action],[data-article-id]');
        if (!element || !access() || busy) return;
        if (element.dataset.articleId) { selected = rows.find(r => r.id === element.dataset.articleId) || null; form = null; error = notice = ''; redraw(); return; }
        const action = element.dataset.articleAction;
        if (action === 'new') return edit(null);
        if (action === 'edit') return edit(selected);
        if (action === 'cancel') { form = null; error = ''; redraw(); return; }
        if (['release','archive','revise'].includes(action)) void command(action);
      });
    }
    return { ensureSession, hasAccess: access, open, render, reset, bind };
  }
  root.HumbelArticleMaster = { create, fields, normalizeData };
  if (typeof module !== 'undefined') module.exports = root.HumbelArticleMaster;
})(typeof window === 'undefined' ? globalThis : window);
