function renderTools() {
  const labels = getToolLabels();
  const manufacturers = getToolManufacturers();
  const holders = getToolHolders();
  const filters = state.toolFilters || {
    search: "",
    label: "",
    tNumber: "",
    diameter: "",
    holder: "",
  };
  const search = (filters.search || "").toLowerCase();
  const filterLabel = filters.label || "";
  const filterT = filters.tNumber || "";
  const filterD = filters.diameter || "";
  const filterHolder = filters.holder || "";

  const labelOptions = labels
    .map((l) => `<option value="${l}">${l}</option>`)
    .join("");
  const filterLabelOptions = labels
    .map(
      (l) =>
        `<option value="${l}" ${l === filterLabel ? "selected" : ""}>${l}</option>`,
    )
    .join("");
  const manufacturerOptions = manufacturers
    .map((m) => `<option value="${m}">${m}</option>`)
    .join("");
  const holderOptions = holders
    .map((h) => `<option value="${h}">${h}</option>`)
    .join("");

  const tools = state.tools.filter((t) => {
    const bySearch =
      !search ||
      `${t.tNumber} ${t.label} ${t.diameter} ${t.articleNo} ${t.holder || ""}`
        .toLowerCase()
        .includes(search);
    const byLabel = !filterLabel || t.label === filterLabel;
    const byT = !filterT || String(t.tNumber).includes(filterT);
    const byD = !filterD || String(t.diameter).includes(filterD);
    const byHolder = !filterHolder || t.holder === filterHolder;
    return bySearch && byLabel && byT && byD && byHolder;
  });

  const toolRows = tools
    .map(
      (t) => `<tr class='border-b'>
      <td class='p-2'>T ${t.tNumber}</td>
      <td class='p-2'>${t.label}</td>
      <td class='p-2'>${formatToolSize(t)}</td>
      <td class='p-2'>${t.shelf}</td>
      <td class='p-2'>${t.articleNo}</td>
      <td class='p-2'>${t.holder || "-"}</td>
      <td class='p-2'>${t.stock}</td>
      <td class='p-2'>${t.minStock}</td>
      <td class='p-2'>${t.manufacturer}</td>
      <td class='p-2'>${t.ordered ? "Bestellt" : "-"}</td>
      <td class='p-2'>
        <button class='px-2 py-1 rounded bg-emerald-700 text-white mr-1' onclick="bookToolChange('${t.id}')">Wechsel</button>
        ${
          currentUser.role === "admin"
            ? `<button class='px-2 py-1 rounded bg-amber-600 text-white mr-1' onclick="editTool('${t.id}')">Bearbeiten</button><button class='px-2 py-1 rounded bg-rose-700 text-white mr-1' onclick="deleteTool('${t.id}')">Löschen</button>`
            : ""
        }
      </td>
    </tr>`,
    )
    .join("");

  const todoTools = state.tools.filter((t) => shouldOrderTool(t) && !t.ordered);
  const orderedTools = state.tools.filter((t) => t.ordered);
  const orderCandidates = todoTools.filter((t) => suggestedOrderQty(t) > 0);

  const todoRows = todoTools
    .map(
      (t) => `<tr class='border-b'>
    <td class='p-2'>T ${t.tNumber}</td>
    <td class='p-2'>${t.label}</td>
    <td class='p-2'>${formatToolSize(t)}</td>
    <td class='p-2'>${t.stock}/${t.minStock}</td>
    <td class='p-2 whitespace-nowrap'><button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="markToolOrdered('${t.id}', true)">Bestellt markieren</button></td>
  </tr>`,
    )
    .join("");

  const orderedCards = orderedTools
    .map(
      (t) => `<div class='border rounded-lg p-3 bg-emerald-50 space-y-3'>
    <div class='grid md:grid-cols-2 gap-x-6 gap-y-2 text-sm'>
      <div><span class='font-semibold'>Bezeichnung:</span> ${t.label}</div>
      <div><span class='font-semibold'>Durchmesser:</span> ${formatToolSize(t)}</div>
      <div><span class='font-semibold'>Steigung:</span> ${t.threadPrefix === "MF" && t.threadPitch ? `P ${t.threadPitch}` : "-"}</div>
      <div><span class='font-semibold'>Menge:</span> ${t.orderedQty || effectiveOrderQty(t)}</div>
      <div><span class='font-semibold'>Artikelnummer:</span> ${t.articleNo}</div>
      <div><span class='font-semibold'>Lagerfach:</span> ${t.shelf}</div>
    </div>
    <div class='flex flex-col sm:flex-row gap-2 sm:justify-end'>
      <button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="markToolOrdered('${t.id}', false)">Bestellt</button>
      <button class='px-2 py-1 rounded bg-slate-700 text-white' onclick="restockTool('${t.id}')">Einlagern</button>
    </div>
  </div>`,
    )
    .join("");

  const journalRows = state.toolJournal
    .slice(0, 80)
    .map(
      (j) => `<tr class='border-b'>
    <td class='p-2'>${j.at}</td>
    <td class='p-2'>${j.user}</td>
    <td class='p-2'>T ${j.tNumber}</td>
    <td class='p-2'>${j.action}</td>
    <td class='p-2'><button class='px-2 py-1 rounded bg-rose-700 text-white' onclick="undoToolJournalEntry('${j.id}')">Rückgängig</button></td>
  </tr>`,
    )
    .join("");

  const orderGroups = orderCandidates.reduce((acc, tool) => {
    const maker = tool.manufacturer || "Ohne Hersteller";
    if (!acc[maker]) acc[maker] = [];
    acc[maker].push(tool);
    return acc;
  }, {});

  const orderListPopup = state.orderListPopupOpen
    ? `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-5xl max-h-[85vh] overflow-auto p-4">
      <div class="flex items-center justify-between mb-3">
        <h3 class="text-lg font-bold">Bestellliste nach Hersteller</h3>
        <button class="px-2 py-1 rounded bg-slate-200" onclick="closeOrderListPopup()">Schließen</button>
      </div>
      ${
        Object.keys(orderGroups).length
          ? Object.entries(orderGroups)
              .map(([maker, list]) => {
                const rows = list
                  .map(
                    (t) => `<tr class='border-b'>
          <td class='p-2'>${t.label}</td>
          <td class='p-2'>${formatToolSize(t)}</td>
          <td class='p-2'>${t.articleNo}</td>
          <td class='p-2'><input type='number' min='0' class='border rounded p-1 w-24' value='${effectiveOrderQty(t)}' onchange="setToolOrderOverride('${t.id}', this.value)" /></td>
        </tr>`,
                  )
                  .join("");
                return `<div class='border rounded p-3 mb-3'>
          <h4 class='font-semibold mb-2'>${maker}</h4>
          <table class='w-full text-sm'><thead class='bg-slate-100'><tr><th class='p-2 text-left'>Bezeichnung</th><th class='p-2 text-left'>Durchmesser</th><th class='p-2 text-left'>Artikelnummer</th><th class='p-2 text-left'>Menge</th></tr></thead><tbody>${rows}</tbody></table>
        </div>`;
              })
              .join("")
          : '<div class="text-sm text-slate-500">Keine bestellrelevanten Werkzeuge.</div>'
      }
    </div>
  </div>`
    : "";

  const suggestionTools = state.tools.filter((t) =>
    shouldShowOptimalQtySuggestion(t),
  );
  const suggestionRows = suggestionTools
    .map(
      (t) => `<tr class='border-b'>
      <td class='p-2'>${t.label}</td>
      <td class='p-2'>${formatToolSize(t)}</td>
      <td class='p-2'>${t.optimalStock || 0}</td>
      <td class='p-2'>${getOptimalQtySuggestion(t)}</td>
      <td class='p-2'>
        <div class='flex gap-2'>
          <button class='px-2 py-1 rounded bg-emerald-700 text-white' onclick="applyOptimalQtySuggestion('${t.id}')">Übernehmen</button>
          <button class='px-2 py-1 rounded bg-rose-700 text-white' onclick="rejectOptimalQtySuggestion('${t.id}')">Ablehnen</button>
        </div>
      </td>
    </tr>`,
    )
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <h2 class='text-xl font-bold mb-1'>Werkzeugverwaltung</h2>
    <p class='text-sm text-slate-600 mb-2'>Alle Bereiche sind getrennt dargestellt: Stammdaten, Bestand, To-Do, Bestellt und Journal.</p>

    ${
      currentUser.role === "admin"
        ? `<div class='border-2 border-slate-300 rounded-xl p-3 bg-slate-50'>
      <h3 class='text-lg font-bold mb-2'>1) Neues Werkzeug anlegen</h3>
      <div class='grid md:grid-cols-4 gap-2'>
        <input id='toolTNumber' class='border rounded p-1' placeholder='T-Nummer (z.B. 134)' />
        <select id='toolLabel' class='border rounded p-1'>${labelOptions}</select>
        <input id='toolDiameter' class='border rounded p-1' placeholder='Durchmesser' />
        <select id='toolThreadPrefix' class='border rounded p-1'>
          <option value=''>Kennung (nur Gewinde)</option>
          <option value='M'>M</option>
          <option value='MF'>MF</option>
          <option value='G'>G</option>
          <option value='UNF'>UNF</option>
          <option value='UNC'>UNC</option>
          <option value='Mx'>Mx</option>
        </select>
        <input id='toolThreadPitch' class='border rounded p-1' placeholder='Steigung P (nur MF)' />
        <input id='toolShelf' class='border rounded p-1' placeholder='A00' />
        <input id='toolArticle' class='border rounded p-1' placeholder='Artikel Nr.' />
        <select id='toolHolder' class='border rounded p-1'>${holderOptions}</select>
        <input id='toolStock' type='number' class='border rounded p-1' placeholder='Bestand' />
        <input id='toolMinStock' type='number' class='border rounded p-1' placeholder='Mindestbestand' />
        <input id='toolOptimalStock' type='number' class='border rounded p-1' placeholder='Optimale Stückzahl' />
        <label class='flex items-center gap-2 text-sm'><input id='toolInsertTool' type='checkbox' onchange='toggleInsertToolFields()' /> Wendeplattenwerkzeug</label>
        <input id='toolInsertEdges' type='number' class='border rounded p-1' placeholder='Anzahl Schneiden' disabled />
        <select id='toolManufacturer' class='border rounded p-1'>${manufacturerOptions}</select>
      </div>
      <div class='flex gap-2 mt-2 flex-wrap'>
        <button class='px-2 py-1 rounded bg-slate-900 text-white' onclick='createTool()'>Werkzeug speichern</button>
        <button class='px-2 py-1 rounded bg-slate-700 text-white' onclick='addToolLabel()'>Bezeichnung hinzufügen</button>
        <button class='px-2 py-1 rounded bg-slate-700 text-white' onclick='addToolManufacturer()'>Hersteller hinzufügen</button>
      </div>
    </div>`
        : ""
    }

    <div class='border-2 border-slate-300 rounded-xl p-3'>
      <h3 class='text-lg font-bold mb-2'>2) Filter & Suche</h3>
      <div class='grid md:grid-cols-7 gap-2 mb-2'>
        <input id='toolSearch' class='border rounded p-1' placeholder='Suche...' value='${filters.search || ""}' />
        <select id='toolFilterLabel' class='border rounded p-1'>
          <option value='' ${filterLabel ? "" : "selected"}>Alle Bezeichnungen</option>${filterLabelOptions}
        </select>
        <input id='toolFilterT' class='border rounded p-1' placeholder='Filter T-Nummer' value='${filterT}' />
        <input id='toolFilterD' class='border rounded p-1' placeholder='Filter Durchmesser' value='${filterD}' />
        <select id='toolFilterHolder' class='border rounded p-1'><option value='' ${filterHolder ? "" : "selected"}>Alle Aufnahmen</option><option value='HSK 100' ${filterHolder === "HSK 100" ? "selected" : ""}>HSK 100</option><option value='HSK 63' ${filterHolder === "HSK 63" ? "selected" : ""}>HSK 63</option></select>
        <button class='px-2 py-1 rounded bg-slate-900 text-white' onclick='applyToolFilters()'>Filter anwenden</button>
        <button class='px-2 py-1 rounded bg-slate-700 text-white' onclick='resetToolFilters()'>Filter zurücksetzen</button>
      </div>
      <p class='text-xs text-slate-500'>Gib alle Werte vollständig ein und klicke dann auf „Filter anwenden“.</p>
    </div>

    <div class='border-2 border-slate-300 rounded-xl p-3'>
      <h3 class='text-lg font-bold mb-2'>3) Werkzeugbestand</h3>
      <div class='overflow-auto max-h-[35vh] border rounded-lg'>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100 sticky top-0'>
            <tr><th class='p-2 text-left'>T</th><th class='p-2 text-left'>Bezeichnung</th><th class='p-2 text-left'>Ø</th><th class='p-2 text-left'>Fach</th><th class='p-2 text-left'>Artikel Nr.</th><th class='p-2 text-left'>Aufnahme</th><th class='p-2 text-left'>Bestand</th><th class='p-2 text-left'>Min</th><th class='p-2 text-left'>Hersteller</th><th class='p-2 text-left'>Status</th><th class='p-2'></th></tr>
          </thead>
          <tbody>${toolRows || '<tr><td class="p-2" colspan="11">Keine Werkzeuge.</td></tr>'}</tbody>
        </table>
      </div>
    </div>

    ${
      currentUser.role === "admin"
        ? `<div class='border-2 border-slate-300 rounded-xl p-3'>
      <h3 class='text-lg font-bold mb-2'>4) Admin-Bestandsaktionen</h3>
      <div class='grid md:grid-cols-2 gap-3'>
        <div class='border rounded p-3 bg-white overflow-auto'>
          <h4 class='font-semibold mb-2'>Werkzeug To-Do (bei Mindestbestand oder darunter)</h4>
          <table class='w-full text-sm'>
            <thead class='bg-slate-100'><tr><th class='p-2 text-left'>T</th><th class='p-2 text-left'>Bezeichnung</th><th class='p-2 text-left'>Größe</th><th class='p-2 text-left'>Bestand</th><th class='p-2'></th></tr></thead>
            <tbody>${todoRows || '<tr><td class="p-2" colspan="5">Keine offenen To-Dos.</td></tr>'}</tbody>
          </table>
        </div>
        <div class='border rounded p-3 bg-white'>
          <h4 class='font-semibold mb-2'>Bestellt</h4>
          <div class='space-y-3'>
            ${orderedCards || '<div class="text-sm text-slate-500">Keine bestellten Werkzeuge.</div>'}
          </div>
        </div>
      </div>

      <div class='mt-3 border rounded p-3 bg-slate-50'>
        <h4 class='font-semibold mb-2'>Bestellliste nach Hersteller</h4>
        <button class='px-2 py-1 rounded bg-slate-900 text-white' onclick='openOrderListPopup()'>Bestellliste anzeigen</button>
      </div>

      <div class='mt-3 border rounded p-3 bg-slate-50'>
        <h4 class='font-semibold mb-2'>Vorschläge für optimale Bestellmenge</h4>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100'><tr><th class='p-2 text-left'>Bezeichnung</th><th class='p-2 text-left'>Größe</th><th class='p-2 text-left'>Aktuell optimal</th><th class='p-2 text-left'>Vorschlag</th><th class='p-2'></th></tr></thead>
          <tbody>${suggestionRows || '<tr><td class="p-2" colspan="5">Noch keine aussagekräftigen Vorschläge vorhanden.</td></tr>'}</tbody>
        </table>
      </div>
    </div>`
        : ""
    }

    <div class='border-2 border-slate-300 rounded-xl p-3'>
      <h3 class='text-lg font-bold mb-2'>5) Schichtjournal – Werkzeugwechsel</h3>
      <div class='overflow-auto max-h-[25vh]'>
        <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Zeit</th><th class='p-2 text-left'>Wer</th><th class='p-2 text-left'>T-Nr</th><th class='p-2 text-left'>Was</th><th class='p-2'></th></tr></thead><tbody>${journalRows || '<tr><td class="p-2" colspan="5">Keine Einträge.</td></tr>'}</tbody></table>
      </div>
    </div>

    ${orderListPopup}
  </div>`;
}

function maybeTaskReminder() {
  if (!currentUser || currentUser.role !== "employee") return;
  const now = new Date();
  const today = isoDate(now);
  const hourKey = `${today}:${now.getHours()}:${currentUser.name}`;
  if (state[`_taskReminder_${hourKey}`]) return;
  const dueToday = state.tasks.filter(
    (t) =>
      t.assignee === currentUser.name &&
      t.status !== "done" &&
      t.deadline === today,
  );
  if (!dueToday.length) return;
  alert(`Erinnerung: ${dueToday.length} Aufgabe(n) heute fällig.`);
  state[`_taskReminder_${hourKey}`] = true;
  persist();
}

render();
window.loginAs = loginAs;
window.logout = logout;
window.setTab = setTab;
window.setStatsView = setStatsView;
window.markAbsent = markAbsent;
window.assignShift = assignShift;
window.cancelShift = cancelShift;
window.assignOptionalShift = assignOptionalShift;
window.approveSaturdayRequest = approveSaturdayRequest;
window.setAvailability = setAvailability;
window.openChecklist = openChecklist;
window.addVacation = addVacation;
window.addSickLeave = addSickLeave;
window.addSwap = addSwap;
window.resetActiveSwaps = resetActiveSwaps;
window.resetManualAssignments = resetManualAssignments;
window.resetPlanCurrentFuture = resetPlanCurrentFuture;
window.updateSlotAssignment = updateSlotAssignment;
window.addEmployee = addEmployee;
window.deactivateEmployee = deactivateEmployee;
window.activateEmployee = activateEmployee;
window.requestSaturdayEvening = requestSaturdayEvening;
window.setWeekendAvailability = setWeekendAvailability;
window.createTask = createTask;
window.completeTask = completeTask;
window.deleteTask = deleteTask;
window.reassignTaskPrompt = reassignTaskPrompt;
window.setConflictResolved = setConflictResolved;
window.stopDowntimeTimer = stopDowntimeTimer;
window.applyToolFilters = applyToolFilters;
window.resetToolFilters = resetToolFilters;
window.bookToolChange = bookToolChange;
window.undoToolJournalEntry = undoToolJournalEntry;
window.editTool = editTool;
window.deleteTool = deleteTool;
window.setToolOrderOverride = setToolOrderOverride;
window.setOrderStatsView = setOrderStatsView;
window.openOrderListPopup = openOrderListPopup;
window.closeOrderListPopup = closeOrderListPopup;
window.applyOptimalQtySuggestion = applyOptimalQtySuggestion;
window.rejectOptimalQtySuggestion = rejectOptimalQtySuggestion;
window.toggleInsertToolFields = toggleInsertToolFields;
