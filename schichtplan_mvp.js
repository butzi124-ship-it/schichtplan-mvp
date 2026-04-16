const { createClient } = window.supabase;

const supabaseClient = createClient(
  window.SUPABASE_URL,
  window.SUPABASE_PUBLISHABLE_KEY,
);

let supabaseReady = false;
let currentSupabaseUser = null;
let currentEmployeeRecord = null;

const USERS = [
  { name: "Lavdrim", slot: "A", type: "core" },
  { name: "Roger", slot: "B", type: "core" },
  { name: "Dashmir", slot: "C", type: "core" },
  { name: "Thomas", slot: "D", type: "springer" },
  { name: "Musa", slot: "E", type: "springer" },
  { name: "Ardian", slot: "F", type: "springer" },
];
const SLOT_CODES = ["A", "B", "C", "D", "E", "F"];
const DEFAULT_SLOT_ASSIGNMENTS = {
  A: "Lavdrim",
  B: "Roger",
  C: "Dashmir",
  D: "Thomas",
  E: "Musa",
  F: "Ardian",
};
const DEFAULT_TOOL_LABELS = [
  "Schaftfräser",
  "Trochodialfräser",
  "Radiusfräser",
  "Kugelfräser",
  "Bohrer",
  "NC Anbohrer",
  "Gewindebohrer",
  "Gewindefräser",
  "Gewindeformer",
  "Gewindewirbler",
  "Ausdrehkopf",
];
const DEFAULT_TOOL_MANUFACTURERS = ["SixSigma", "SFS", "THAA"];
const DEFAULT_TOOL_HOLDERS = ["HSK 100", "HSK 63"];

const STORAGE_KEY = "schichtplan_mvp_v_0_2";
const state = loadState();
let currentUser = null;
let currentTab = "schichtplan";
let statsViewPeriod = "week";

const WEEK_TEMPLATES = [
  makeWeek("A", "B", "C", "A", "B"),
  makeWeek("B", "C", "A", "B", "C"),
  makeWeek("C", "A", "B", "C", "A"),
  makeWeek("A", "B", "C", "A", "B"),
  makeWeek("B", "C", "A", "B", "C"),
  makeWeek("C", "A", "B", "C", "A"),
];
const ROTATION_ANCHOR_MONDAY = "2026-01-05";
const PLANNING_SUBTABS = [
  { id: "personal", label: "Personal" },
  { id: "abstinenz", label: "Abstinenz" },
  { id: "wochenende", label: "Wochenendeinsätze" },
  { id: "schichttausch", label: "Schichttausch" },
];

function setLoginStatus(message, isError = false) {
  const el = document.getElementById("loginStatus");
  if (!el) return;
  el.textContent = message;
  el.className = isError
    ? "text-sm rounded border border-rose-200 bg-rose-50 p-3 text-rose-700"
    : "text-sm rounded border border-slate-200 bg-slate-50 p-3 text-slate-600";
}

function fillLogin(type) {
  const emailInput = document.getElementById("loginEmail");
  const passwordInput = document.getElementById("loginPassword");
  if (!emailInput || !passwordInput) return;

  if (type === "admin" && window.TEST_ADMIN_EMAIL) {
    emailInput.value = window.TEST_ADMIN_EMAIL;
  }
  if (type === "employee" && window.TEST_EMPLOYEE_EMAIL) {
    emailInput.value = window.TEST_EMPLOYEE_EMAIL;
  }

  passwordInput.focus();
}

async function testSupabaseConnection() {
  try {
    const { error } = await supabaseClient
      .from("employees")
      .select("id")
      .limit(1);

    if (error) {
      console.error("Supabase-Test fehlgeschlagen:", error);
      setLoginStatus(`Supabase-Fehler: ${error.message}`, true);
      return false;
    }

    console.log("Supabase-Verbindung ok.");
    setLoginStatus("Supabase-Verbindung ok. Bitte anmelden.");
    return true;
  } catch (err) {
    console.error("Supabase konnte nicht initialisiert werden:", err);
    setLoginStatus("Supabase konnte nicht initialisiert werden.", true);
    return false;
  }
}

async function loadEmployeesFromSupabase() {
  const { data, error } = await supabaseClient
    .from("employees")
    .select("*")
    .order("display_name", { ascending: true });

  if (error) {
    console.error("Fehler beim Laden von employees:", error);
    return [];
  }

  return data || [];
}

function normalizeToolFromDb(row) {
  return {
    id: row.id,
    tNumber: row.t_number,
    label: row.label,
    diameter: row.diameter,
    threadPrefix: row.thread_prefix || "",
    threadPitch: row.thread_pitch || "",
    cornerRadius: row.corner_radius || "",
    shelf: row.shelf,
    articleNo: row.article_no,
    holder: row.holder,
    stock: Number(row.stock || 0),
    minStock: Number(row.min_stock || 0),
    optimalStock: Number(row.optimal_stock || 0),
    manufacturer: row.manufacturer || "",
    ordered: !!row.ordered,
    orderedQty: Number(row.ordered_qty || 0),
    insertTool: !!row.insert_tool,
    insertEdges: Number(row.insert_edges || 0),
  };
}

async function loadToolsFromSupabase() {
  const { data, error } = await supabaseClient
    .from("tools")
    .select("*")
    .order("t_number", { ascending: true });

  if (error) {
    console.error("Fehler beim Laden von tools:", error);
    return [];
  }

  return (data || []).map(normalizeToolFromDb);
}

async function loginWithSupabase() {
  const email = document.getElementById("loginEmail")?.value?.trim() || "";
  const password = document.getElementById("loginPassword")?.value || "";

  if (!email || !password) {
    setLoginStatus("Bitte E-Mail und Passwort eingeben.", true);
    return;
  }

  setLoginStatus("Anmeldung läuft ...");

  const { error } = await supabaseClient.auth.signInWithPassword({
    email,
    password,
  });

  if (error) {
    console.error("Supabase-Login fehlgeschlagen:", error);
    setLoginStatus(`Login fehlgeschlagen: ${error.message}`, true);
    return;
  }

  await syncSupabaseSessionToApp();
}

async function logoutSupabase() {
  await supabaseClient.auth.signOut();
  currentSupabaseUser = null;
  currentEmployeeRecord = null;
  currentUser = null;
  currentTab = "schichtplan";

  document.getElementById("loginBox")?.classList.remove("hidden");
  document.getElementById("tabs")?.classList.add("hidden");
  document.getElementById("sessionInfo").textContent = "";
  document.getElementById("view").innerHTML = "";
  setLoginStatus("Supabase-Session gelöscht.");
}

async function syncSupabaseSessionToApp() {
  const {
    data: { user },
    error,
  } = await supabaseClient.auth.getUser();

  if (error) {
    console.error("Fehler beim Abrufen des angemeldeten Users:", error);
    setLoginStatus(`Fehler beim Session-Lesen: ${error.message}`, true);
    return;
  }

  if (!user) {
    setLoginStatus("Keine aktive Supabase-Session vorhanden.");
    return;
  }

  currentSupabaseUser = user;
  currentEmployeeRecord = await getCurrentEmployeeRecord();

  if (!currentEmployeeRecord) {
    setLoginStatus(
      "Login ok, aber kein aktiver employees-Eintrag gefunden.",
      true,
    );
    return;
  }

  currentUser = {
    name: currentEmployeeRecord.display_name,
    role: currentEmployeeRecord.role,
  };

  document.getElementById("loginBox")?.classList.add("hidden");
  setLoginStatus(
    `Angemeldet als ${currentEmployeeRecord.display_name} (${currentEmployeeRecord.role}).`,
  );

  console.log("Supabase-User:", currentSupabaseUser);
  console.log("Employees-Datensatz:", currentEmployeeRecord);

  render();
}

async function bootSupabase() {
  supabaseReady = await testSupabaseConnection();

  if (!supabaseReady) {
    console.warn(
      "Supabase ist aktuell nicht bereit. App läuft vorerst lokal weiter.",
    );
    return;
  }

  const {
    data: { session },
  } = await supabaseClient.auth.getSession();

  if (session?.user) {
    await syncSupabaseSessionToApp();
    return;
  }

      const employees = await loadEmployeesFromSupabase();
  const tools = await loadToolsFromSupabase();

  state.tools = tools;
  persist();

  console.log("Employees aus Supabase:", employees);
  console.log("Tools aus Supabase:", tools);

  if (currentUser) render();
  
}

function allUsers() {
  return [...USERS, ...(state.extraUsers || [])];
}

function activeUsers() {
  return allUsers().filter((u) => !state.inactiveUsers?.[u.name]);
}

function makeWeek(early, late, night, satPrimary, satSecondary) {
  return {
    mondayToFriday: [
      { label: "Früh", start: "05:00", end: "13:00", options: [early] },
      { label: "Spät", start: "13:00", end: "21:00", options: [late] },
      { label: "Nacht", start: "21:00", end: "05:00", options: [night] },
    ],
    saturday: [
      {
        label: "Samstag S1",
        start: "05:00",
        end: "16:00",
        options: [satPrimary, "D"],
      },
      {
        label: "Samstag S2",
        start: "16:00",
        end: "05:00",
        options: [satSecondary, "NONE"],
      },
    ],
    sunday: [
      { label: "Sonntag Tag", start: "06:00", end: "18:00", options: [late] },
      {
        label: "Sonntag Abend",
        start: "18:00",
        end: "05:00",
        options: [early],
      },
    ],
  };
}

function loadState() {
  const base = {
    absences: {},
    assignments: {},
    availability: {},
    checklists: {},
    unmanned: {},
    vacations: [],
    sickLeaves: [],
    swaps: [],
    shiftCancellations: {},
    slotAssignments: { ...DEFAULT_SLOT_ASSIGNMENTS },
    saturdayEveningRequests: {},
    tasks: [],
    conflicts: {},
    shiftEndChecks: {},
    shiftStartChecks: {},
    machineDowntime: {},
    machinePromptSeen: {},
    extraUsers: [],
    inactiveUsers: {},
    tools: [],
    toolLabelsExtra: [],
    toolManufacturersExtra: [],
    toolJournal: [],
    toolFilters: {
      search: "",
      label: "",
      tNumber: "",
      diameter: "",
      holder: "",
    },
    toolOrderOverrides: {},
    orderArchive: [],
    orderHistory: [],
    orderStatsView: "week",
    orderSuggestionState: {},
    orderListPopupOpen: false,
    selectedOrderListManufacturer: "",
    planningSubTab: "personal",
    absenceReplacements: {},
  };
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return base;
    const parsed = JSON.parse(raw);
    return { ...base, ...parsed };
  } catch {
    return base;
  }
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loginAs(name) {
  if (name === "admin") {
    currentUser = { name: "Admin", role: "admin" };
  } else {
    currentUser = { name, role: "employee" };
  }
  document.getElementById("loginBox")?.classList.add("hidden");
  render();
}

function logout() {
  logoutSupabase();
}

function render() {
  if (!currentUser) return;
  cleanupOrderArchive();

  const tabs = ["schichtplan"];
  if (currentUser.role === "admin") {
    tabs.push(
      "planung",
      "werkzeuge",
      "bestellstatistik",
      "todo",
      "konflikte",
      "statistik",
    );
  }
  if (currentUser.role === "employee") tabs.push("meine", "werkzeuge", "todo");

  const tabsEl = document.getElementById("tabs");
  tabsEl.className = "flex gap-2 flex-wrap";
  tabsEl.innerHTML =
    tabs
      .map((t) => {
        const statusClass = tabNeedsAttention(t)
          ? "bg-rose-100 border-rose-400"
          : "bg-emerald-100 border-emerald-400";
        const activeClass = currentTab === t ? "ring-2 ring-slate-900" : "";
        return `<button class="px-3 py-2 rounded border ${statusClass} ${activeClass}" onclick="setTab('${t}')">${labelTab(t)}</button>`;
      })
      .join("") +
    `<button class="px-3 py-2 rounded bg-red-100" onclick="logout()">Abmelden</button>`;

  tabsEl.classList.remove("hidden");

  document.getElementById("sessionInfo").textContent =
    `Angemeldet: ${currentUser.name} (${currentUser.role})`;

  if (!tabs.includes(currentTab)) currentTab = tabs[0];
  const view = document.getElementById("view");
  if (currentTab === "schichtplan") view.innerHTML = renderSchedule();
  if (currentTab === "meine") view.innerHTML = renderMyShifts();
  if (currentTab === "planung") view.innerHTML = renderPlanning();
  if (currentTab === "werkzeuge") view.innerHTML = renderTools();
  if (currentTab === "bestellstatistik") view.innerHTML = renderOrderStats();
  if (currentTab === "todo") view.innerHTML = renderTodo();
  if (currentTab === "konflikte") view.innerHTML = renderConflicts();
  if (currentTab === "statistik") view.innerHTML = renderStats();
  maybeShowMachinePrompt();
  maybeTaskReminder();
  maybeShowShiftStartChecklist();
  maybeShowShiftEndChecklist();
}

function labelTab(tab) {
  return {
    schichtplan: "Schichtplan",
    meine: "Meine Schichten",
    planung: "Planung (Admin)",
    werkzeuge: "Werkzeuge",
    bestellstatistik: "Bestell-Statistik",
    todo: "To-Do",
    konflikte: "Konflikte",
    statistik: "Statistik",
  }[tab];
}

function setTab(tab) {
  currentTab = tab;
  render();
}

function setStatsView(period) {
  statsViewPeriod = period;
  render();
}

function setPlanningSubTab(subTab) {
  if (!PLANNING_SUBTABS.some((t) => t.id === subTab)) return;
  state.planningSubTab = subTab;
  persist();
  render();
}

function shouldOrderTool(tool) {
  return Number(tool.stock) <= Number(tool.minStock);
}

function tabNeedsAttention(tab) {
  if (tab === "planung") return generateThreeMonths().some((s) => s.open);
  if (tab === "todo") return state.tasks.some((t) => t.status !== "done");
  if (tab === "werkzeuge" && currentUser?.role === "admin")
    return state.tools.some((t) => shouldOrderTool(t));
  if (tab === "bestellstatistik" && currentUser?.role === "admin")
    return (state.orderHistory || []).length > 0;
  if (tab === "konflikte")
    return Object.values(state.conflicts).some((c) => c.resolved !== true);
  if (tab === "meine" && currentUser?.role === "employee")
    return generateThreeMonths().some(
      (s) => s.assigned === currentUser.name && s.open,
    );
  return false;
}

function generateThreeMonths() {
  const shifts = [];
  const today = new Date();
  const start = new Date(today);
  start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
  const anchor = new Date(`${ROTATION_ANCHOR_MONDAY}T00:00:00`);
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;

  for (let i = 0; i < 90; i++) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    const iso = d.toISOString().slice(0, 10);
    const weekIndex = ((Math.floor((d - anchor) / msPerWeek) % 6) + 6) % 6;
    const weekday = d.getDay();
    const template = WEEK_TEMPLATES[weekIndex];

    if (weekday >= 1 && weekday <= 5) {
      template.mondayToFriday.forEach((s, idx) =>
        shifts.push(buildShift(iso, `${iso}-mf-${idx}`, s)),
      );
    }
    if (weekday === 6) {
      template.saturday.forEach((s, idx) =>
        shifts.push(buildShift(iso, `${iso}-sa-${idx}`, s)),
      );
    }
    if (weekday === 0) {
      template.sunday.forEach((s, idx) =>
        shifts.push(buildShift(iso, `${iso}-su-${idx}`, s)),
      );
    }
  }

  return shifts;
}

function getReplacementForShift(shiftId) {
  return state.absenceReplacements?.[shiftId] || null;
}

function getReplacementEntriesForSource(sourceType, sourceId) {
  return Object.entries(state.absenceReplacements || {}).filter(([, entry]) => {
    return entry?.sourceType === sourceType && entry?.sourceId === sourceId;
  });
}

function getReplacementSummaryForSource(sourceType, sourceId) {
  const entries = getReplacementEntriesForSource(sourceType, sourceId);
  if (!entries.length) return "Kein Ersatz geplant";

  const normalized = entries.map(([, entry]) => {
    if (entry.mode === "cancel") {
      return `${formatDateDisplay(entry.weekFrom)} bis ${formatDateDisplay(entry.weekTo)}: AUSFALL`;
    }
    return `${formatDateDisplay(entry.weekFrom)} bis ${formatDateDisplay(entry.weekTo)}: ${entry.replacementUser}`;
  });

  return normalized.join("<br>");
}

function applyReplacementToShift(shiftId, replacementEntry) {
  if (!state.absenceReplacements) state.absenceReplacements = {};
  state.absenceReplacements[shiftId] = replacementEntry;

  if (replacementEntry.mode === "cancel") {
    state.shiftCancellations[shiftId] = true;
    delete state.assignments[shiftId];
    return;
  }

  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = replacementEntry.replacementUser;
}

function clearReplacementPlanForSource(sourceType, sourceId) {
  if (!state.absenceReplacements) return;
  Object.keys(state.absenceReplacements).forEach((shiftId) => {
    const entry = state.absenceReplacements[shiftId];
    if (entry?.sourceType === sourceType && entry?.sourceId === sourceId) {
      delete state.absenceReplacements[shiftId];
      delete state.shiftCancellations[shiftId];
      delete state.assignments[shiftId];
    }
  });
}

function clearAbsenceReplacementPlan(sourceType, sourceId) {
  clearReplacementPlanForSource(sourceType, sourceId);
  persist();
  render();
}

function buildShift(date, id, template) {
  const defaultAssigned = chooseDefault(template.options);
  const swappedDefault = defaultAssigned
    ? applySwap(date, defaultAssigned)
    : defaultAssigned;

  const isOptional = template.options.length > 1;
  const manualAssigned = state.assignments[id] || null;
  const replacement = getReplacementForShift(id);

  const absenceKey = `${date}:${swappedDefault}`;
  const calendarAbsenceType = getCalendarAbsenceType(swappedDefault, date);
  const manualAbsent = !!state.absences[absenceKey];
  const absent = manualAbsent || !!calendarAbsenceType;

  const canceled =
    !!state.shiftCancellations[id] || replacement?.mode === "cancel";

  let assigned = null;

  if (canceled) {
    assigned = null;
  } else if (replacement?.mode === "replace" && replacement?.replacementUser) {
    assigned = replacement.replacementUser;
  } else if (!isOptional && !absent) {
    assigned = swappedDefault;
  } else {
    assigned = manualAssigned || (isOptional ? swappedDefault : null);
  }

  const open =
    canceled ||
    (!assigned && (absent || !assigned || assigned === "NONE"));

  return {
    id,
    date,
    label: template.label,
    start: template.start,
    end: template.end,
    options: template.options,
    assigned,
    originalAssigned: swappedDefault,
    absenceType: calendarAbsenceType || (manualAbsent ? "abwesend" : null),
    replacement,
    open,
  };
}

function chooseDefault(options) {
  const slot = options[0];
  if (slot === "NONE") return null;
  return slotToName(slot);
}

function slotToName(slot) {
  return (
    state.slotAssignments?.[slot] ||
    allUsers().find((u) => u.slot === slot)?.name ||
    null
  );
}

function userByName(name) {
  return allUsers().find((u) => u.name === name);
}

function isCoreEmployee(name) {
  return userByName(name)?.type === "core";
}

function isSpringer(name) {
  return userByName(name)?.type === "springer";
}

function slotOfUser(name) {
  const mappedSlot = Object.entries(state.slotAssignments || {}).find(
    ([, assignedName]) => assignedName === name,
  )?.[0];
  return mappedSlot || allUsers().find((u) => u.name === name)?.slot || "-";
}

const PERSON_COLORS = {
  Lavdrim: "bg-green-100 text-green-900",
  Roger: "bg-violet-100 text-violet-900",
  Dashmir: "bg-rose-100 text-rose-900",
  Thomas: "bg-cyan-100 text-cyan-900",
  Musa: "bg-amber-100 text-amber-900",
  Ardian: "bg-indigo-100 text-indigo-900",
};

function personColorClasses(name) {
  const raw = PERSON_COLORS[name] || "bg-slate-100 text-slate-800";
  const parts = raw.split(" ");
  const bg = parts.find((p) => p.startsWith("bg-")) || "bg-slate-100";
  const text = parts.find((p) => p.startsWith("text-")) || "text-slate-800";
  return { bg, text, raw: `${bg} ${text}` };
}
function personBorderClass(name) {
  const slot = slotOfUser(name);
  if (slot === "A") return "border-green-500";
  if (slot === "B") return "border-violet-500";
  if (slot === "C") return "border-rose-500";
  if (slot === "D") return "border-cyan-500";
  if (slot === "E") return "border-amber-500";
  if (slot === "F") return "border-indigo-500";
  return "border-slate-400";
}

function formatSlot(options) {
  return options.map((slot) => (slot === "NONE" ? "0" : slot)).join("/");
}

function slotColor(slotText) {
  if (slotText.includes("A")) return "bg-green-100 text-green-900";
  if (slotText.includes("B")) return "bg-violet-100 text-violet-900";
  if (slotText.includes("C")) return "bg-rose-100 text-rose-900";
  return "bg-slate-100 text-slate-800";
}

function weekStartFromMonday(baseDate, addWeeks) {
  const d = new Date(baseDate);
  d.setDate(d.getDate() + addWeeks * 7);
  return d;
}

function isoDate(d) {
  return d.toISOString().slice(0, 10);
}

function todayIso() {
  return isoDate(new Date());
}

function formatDateDisplay(iso) {
  if (!iso || !iso.includes("-")) return iso || "";
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y.slice(2)}`;
}

function getWeekdayName(iso) {
  if (!iso) return "";
  const date = new Date(`${iso}T00:00:00`);
  return ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"][date.getDay()];
}

function formatDateWithWeekday(iso) {
  return `${getWeekdayName(iso)} ${formatDateDisplay(iso)}`;
}

function inRange(date, from, to) {
  return date >= from && date <= to;
}

function applySwap(date, name) {
  let resolved = name;
  state.swaps.forEach((swap) => {
    if (!swap || date < swap.startDate) return;
    if (swap.endDate && date > swap.endDate) return;
    if (resolved === swap.userA) resolved = swap.userB;
    else if (resolved === swap.userB) resolved = swap.userA;
  });
  return resolved;
}

function hasCalendarAbsence(name, date) {
  const inVacation = state.vacations.some(
    (v) => v.user === name && inRange(date, v.from, v.to),
  );
  const inSick = state.sickLeaves.some(
    (s) => s.user === name && inRange(date, s.from, s.to),
  );
  return inVacation || inSick;
}

function getCalendarAbsenceType(name, date) {
  if (!name) return null;
  if (
    state.sickLeaves.some((s) => s.user === name && inRange(date, s.from, s.to))
  )
    return "krank";
  if (
    state.vacations.some((v) => v.user === name && inRange(date, v.from, v.to))
  )
    return "urlaub";
  return null;
}

function resolveAssigned(shiftId, options) {
  const replacement = getReplacementForShift(shiftId);
  if (replacement?.mode === "cancel") return null;
  if (replacement?.replacementUser) return replacement.replacementUser;

  const date = shiftId.slice(0, 10);
  const defaultAssigned = chooseDefault(options);
  const swappedDefault = defaultAssigned
    ? applySwap(date, defaultAssigned)
    : defaultAssigned;
  const isOptional = options.length > 1;
  const absent =
    !!state.absences[`${date}:${swappedDefault}`] ||
    hasCalendarAbsence(swappedDefault, date);
  const manualAssigned = state.assignments[shiftId] || null;
  if (state.shiftCancellations[shiftId]) return null;
  return !isOptional && !absent
    ? swappedDefault
    : manualAssigned || (isOptional ? swappedDefault : null);
}

function resolveAbsenceType(shiftId, options) {
  const date = shiftId.slice(0, 10);
  const defaultAssigned = chooseDefault(options);
  const swappedDefault = defaultAssigned
    ? applySwap(date, defaultAssigned)
    : defaultAssigned;
  if (getCalendarAbsenceType(swappedDefault, date))
    return getCalendarAbsenceType(swappedDefault, date);
  if (state.absences[`${date}:${swappedDefault}`]) return "abwesend";
  return null;
}

function assignedDisplay(shiftId, options) {
  const assignedName = resolveAssigned(shiftId, options);
  if (!assignedName) return "-";
  return `${slotOfUser(assignedName)} • ${assignedName}`;
}

function getPrimaryAbsenceInfo(shiftId, options) {
  const date = shiftId.slice(0, 10);
  const defaultAssigned = chooseDefault(options);
  if (!defaultAssigned) return null;
  const swappedDefault = applySwap(date, defaultAssigned);
  const type = getCalendarAbsenceType(swappedDefault, date);
  if (type) return { name: swappedDefault, type };
  if (state.absences[`${date}:${swappedDefault}`])
    return { name: swappedDefault, type: "abwesend" };
  return null;
}

function assignedMeta(shiftId, options) {
  if (
    state.shiftCancellations?.[shiftId] ||
    getReplacementForShift(shiftId)?.mode === "cancel"
  ) {
    return {
      label: "AUSFALL",
      cls: "bg-rose-200 text-rose-900",
      borderCls: "border-rose-500",
      ringCls: "border-2",
    };
  }

  const shift = getShiftById(shiftId);
  const assignedName = shift?.assigned || resolveAssigned(shiftId, options);
  const absence = getPrimaryAbsenceInfo(shiftId, options);

  if (!assignedName) {
    const abs = absence?.type || resolveAbsenceType(shiftId, options);
    if (abs) {
      return {
        label: "-",
        cls: "bg-slate-100 text-slate-700",
        borderCls: absence?.name
          ? personBorderClass(absence.name)
          : "border-slate-400",
        ringCls: "border-2",
      };
    }
    return {
      label: "-",
      cls: "bg-slate-100 text-slate-700",
      borderCls: "border-slate-300",
      ringCls: "border",
    };
  }

  const assignedColors = personColorClasses(assignedName);
  if (absence?.name) {
    return {
      label: `${slotOfUser(assignedName)} • ${assignedName}`,
      cls: `${assignedColors.bg} ${assignedColors.text}`,
      borderCls: personBorderClass(absence.name),
      ringCls: "border-2",
    };
  }

  return {
    label: `${slotOfUser(assignedName)} • ${assignedName}`,
    cls: assignedColors.raw,
    borderCls: personBorderClass(assignedName),
    ringCls: "border",
  };
}

function renderOverviewPlan(weeksToShow = 12) {
  const today = new Date();
  const monday = new Date(today);
  monday.setDate(monday.getDate() - ((monday.getDay() + 6) % 7));
  const anchor = new Date(`${ROTATION_ANCHOR_MONDAY}T00:00:00`);
  const msPerWeek = 7 * 24 * 60 * 60 * 1000;

  const dayNames = [
    "Montag",
    "Dienstag",
    "Mittwoch",
    "Donnerstag",
    "Freitag",
    "Samstag",
    "Sonntag",
  ];
  let body = "";

  for (let w = 0; w < weeksToShow; w++) {
    const weekStart = weekStartFromMonday(monday, w);
    const templateWeekIndex =
      ((Math.floor((weekStart - anchor) / msPerWeek) % 6) + 6) % 6;
    const template = WEEK_TEMPLATES[templateWeekIndex];

    if (w === 0) {
      const startSunday = new Date(weekStart);
      startSunday.setDate(startSunday.getDate() - 1);
      const prevTemplateIndex = (templateWeekIndex + 5) % 6;
      const prevTemplate = WEEK_TEMPLATES[prevTemplateIndex];
      const startMeta = assignedMeta(
        `${isoDate(startSunday)}-su-1`,
        prevTemplate.sunday[1].options,
      );
      body += `<tr class="border-b bg-amber-50">
        <td class="p-2"></td>
        <td class="p-2 font-semibold">Start (Sonntag)</td>
        <td class="p-2 bg-slate-50" colspan="4"></td>
        <td class="p-2 ${startMeta.cls} ${startMeta.borderCls} ${startMeta.ringCls} font-semibold text-center">${startMeta.label}</td>
        <td class="p-2 font-semibold text-center">18:00-24:00</td>
      </tr>`;
    }

    dayNames.forEach((dayName, dayIndex) => {
      const date = new Date(weekStart);
      date.setDate(weekStart.getDate() + dayIndex);
      const weekLabel = dayIndex === 0 ? `Woche ${w + 1}` : "";

      if (dayIndex <= 4) {
        const s1 = template.mondayToFriday[0];
        const s2 = template.mondayToFriday[1];
        const s3 = template.mondayToFriday[2];
        const dateIso = isoDate(date);
        const m1 = assignedMeta(`${dateIso}-mf-0`, s1.options);
        const m2 = assignedMeta(`${dateIso}-mf-1`, s2.options);
        const m3 = assignedMeta(`${dateIso}-mf-2`, s3.options);
        body += `<tr class="border-b">
          <td class="p-2 font-semibold">${weekLabel}</td>
          <td class="p-2">${dayName}<div class="text-[11px] text-slate-500">${formatDateDisplay(dateIso)}</div></td>
          <td class="p-2 ${m1.cls} ${m1.borderCls} ${m1.ringCls} font-semibold text-center">${m1.label}</td>
          <td class="p-2 font-semibold text-center bg-white">05:00-11:00</td>
          <td class="p-2 ${m2.cls} ${m2.borderCls} ${m2.ringCls} font-semibold text-center">${m2.label}</td>
          <td class="p-2 font-semibold text-center bg-white">13:00-19:00</td>
          <td class="p-2 ${m3.cls} ${m3.borderCls} ${m3.ringCls} font-semibold text-center">${m3.label}</td>
          <td class="p-2 font-semibold text-center bg-white">21:00-03:00</td>
        </tr>`;
      } else if (dayIndex === 5) {
        const s1 = template.saturday[0];
        const s2 = template.saturday[1];
        const dateIso = isoDate(date);
        const m1 = assignedMeta(`${dateIso}-sa-0`, s1.options);
        const m2 = assignedMeta(`${dateIso}-sa-1`, s2.options);
        body += `<tr class="border-b bg-amber-50">
          <td class="p-2 font-semibold">${weekLabel}</td>
          <td class="p-2 font-semibold">${dayName}<div class="text-[11px] text-slate-500">${formatDateDisplay(dateIso)}</div></td>
          <td class="p-2 ${m1.cls} ${m1.borderCls} ${m1.ringCls} font-semibold text-center">${m1.label}</td>
          <td class="p-2 font-semibold text-center bg-amber-100">05:00-11:00 (Sa Morgen)</td>
          <td class="p-2 ${m2.cls} ${m2.borderCls} ${m2.ringCls} font-semibold text-center">${m2.label}</td>
          <td class="p-2 font-semibold text-center bg-amber-200">16:00-22:00 (Sa Abend)</td>
          <td class="p-2 bg-slate-50" colspan="2"></td>
        </tr>`;
      } else {
        const s1 = template.sunday[0];
        const s2 = template.sunday[1];
        const dateIso = isoDate(date);
        const m1 = assignedMeta(`${dateIso}-su-0`, s1.options);
        const m2 = assignedMeta(`${dateIso}-su-1`, s2.options);
        body += `<tr class="border-b bg-amber-50">
          <td class="p-2 font-semibold">${weekLabel}</td>
          <td class="p-2 font-semibold">${dayName}<div class="text-[11px] text-slate-500">${formatDateDisplay(dateIso)}</div></td>
          <td class="p-2 ${m1.cls} ${m1.borderCls} ${m1.ringCls} font-semibold text-center">${m1.label}</td>
          <td class="p-2 font-semibold text-center bg-blue-100">06:00-12:00 (So Morgen)</td>
          <td class="p-2 bg-slate-50" colspan="2"></td>
          <td class="p-2 ${m2.cls} ${m2.borderCls} ${m2.ringCls} font-semibold text-center">${m2.label}</td>
          <td class="p-2 font-semibold text-center bg-white">18:00-24:00</td>
        </tr>`;
      }
    });
  }

  return `<div class='overflow-auto max-h-[70vh] border rounded-lg'>
    <table class='min-w-[1200px] w-full text-sm'>
      <thead class='sticky top-0 bg-slate-200 z-10'>
        <tr>
          <th class='p-2 text-left'>Woche</th>
          <th class='p-2 text-left'>Tag</th>
          <th class='p-2 text-center'>S1</th>
          <th class='p-2 text-center'>Frühschicht</th>
          <th class='p-2 text-center'>S2</th>
          <th class='p-2 text-center'>Spätschicht</th>
          <th class='p-2 text-center'>S3</th>
          <th class='p-2 text-center'>Nachschicht</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
    </table>
  </div>`;
}

function renderSchedule() {
  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-3'>Gesamt-Schichtplan (90 Tage ab aktueller Woche)</h2>
    ${renderOverviewPlan(Math.ceil(90 / 7))}
  </div>`;
}

function assignShift(shiftId) {
  const select = document.getElementById(`sel-${shiftId}`);
  if (!select) return;
  if (!canAssignUserToShift(select.value, shiftId)) return;
  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = select.value;
  persist();
  render();
}

function cancelShift(shiftId) {
  state.shiftCancellations[shiftId] = true;
  delete state.assignments[shiftId];
  persist();
  render();
}

function renderMyShifts() {
  const shifts = generateThreeMonths()
    .filter((s) => s.assigned === currentUser.name)
    .slice(0, 90);
  const rows = shifts
    .map((s) => {
      const status =
        s.assigned === currentUser.name
          ? '<span class="text-emerald-700 font-semibold">Eingeplant</span>'
          : s.open
            ? '<span class="text-red-600 font-semibold">OFFEN</span>'
            : '<span class="text-emerald-700">Besetzt</span>';
      const saturdayEvening = s.id.includes("-sa-1");
      const reqKey = `${s.id}:${currentUser.name}`;
      const requested = !!state.saturdayEveningRequests[reqKey];
      return `<tr class="border-b">
      <td class="p-2">${formatDateWithWeekday(s.date)}</td>
      <td class="p-2">${s.label}</td>
      <td class="p-2">${s.start}–${s.end}</td>
      <td class="p-2">${status}</td>
      <td class="p-2">
        <button class='px-2 py-1 bg-amber-200 rounded mr-2' onclick="markAbsent('${s.id}','${s.date}','${currentUser.name}')">Abwesenheit</button>
        ${saturdayEvening ? `<button class='px-2 py-1 rounded ${requested ? "bg-emerald-200" : "bg-blue-200"}' onclick="requestSaturdayEvening('${s.id}')">${requested ? "Eingetragen" : "Sa-Abend eintragen"}</button>` : ""}
      </td>
    </tr>`;
    })
    .join("");

  const weekendRows = isSpringer(currentUser.name)
    ? generateThreeMonths()
        .filter((s) => s.id.includes("-sa-") || s.id.includes("-su-"))
        .slice(0, 90)
        .map((s) => {
          const key = `${s.id}:${currentUser.name}`;
          const val = state.availability[key] || "";
          return `<tr class='border-b'>
          <td class='p-2'>${formatDateWithWeekday(s.date)}</td>
          <td class='p-2'>${s.label}</td>
          <td class='p-2'>${s.start}–${s.end}</td>
          <td class='p-2'>
            <select class='border rounded p-1' onchange="setWeekendAvailability('${s.id}', this.value)">
              <option value='' ${val === "" ? "selected" : ""}>-</option>
              <option value='yes' ${val === "yes" ? "selected" : ""}>Kann</option>
              <option value='no' ${val === "no" ? "selected" : ""}>Kann nicht</option>
            </select>
          </td>
        </tr>`;
        })
        .join("")
    : "";

  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-3'>Meine individuellen Schichten (90 Tage)</h2>
    <div class='overflow-auto max-h-[65vh] border rounded-lg'>
      <table class='w-full text-sm'><thead class='sticky top-0 bg-slate-100'><tr>
        <th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Zeit</th><th class='p-2 text-left'>Status</th><th class='p-2 text-left'>Aktionen</th>
      </tr></thead><tbody>${rows || '<tr><td class="p-2" colspan="5">Keine Schichten gefunden.</td></tr>'}</tbody></table>
    </div>
    ${
      isSpringer(currentUser.name)
        ? `<h3 class='font-semibold mt-4 mb-2'>Wochenende-Verfügbarkeit (Samstag/Sonntag)</h3>
    <div class='overflow-auto max-h-[35vh] border rounded-lg'>
      <table class='w-full text-sm'><thead class='sticky top-0 bg-slate-100'><tr><th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Zeit</th><th class='p-2 text-left'>Kannst du?</th></tr></thead>
      <tbody>${weekendRows || '<tr><td class="p-2" colspan="4">Keine Wochenend-Schichten.</td></tr>'}</tbody></table>
    </div>`
        : ""
    }
  </div>`;
}

function requestSaturdayEvening(shiftId) {
  state.saturdayEveningRequests[`${shiftId}:${currentUser.name}`] = true;
  persist();
  render();
}

function setWeekendAvailability(shiftId, value) {
  state.availability[`${shiftId}:${currentUser.name}`] = value;
  persist();
  render();
}

function renderPlanning() {
  const subTab = PLANNING_SUBTABS.some((t) => t.id === state.planningSubTab)
    ? state.planningSubTab
    : "personal";

  const subTabButtons = PLANNING_SUBTABS.map((tab) => {
    const active = subTab === tab.id;
    return `<button class='px-3 py-2 rounded border ${active ? "bg-slate-900 text-white border-slate-900" : "bg-white border-slate-300"}' onclick="setPlanningSubTab('${tab.id}')">${tab.label}</button>`;
  }).join("");

  let content = "";
  if (subTab === "personal") content = renderPlanningPersonal();
  if (subTab === "abstinenz") content = renderPlanningAbstinenz();
  if (subTab === "wochenende") content = renderPlanningWochenende();
  if (subTab === "schichttausch") content = renderPlanningSchichttausch();

  return `<div class='bg-white rounded-xl shadow p-4'>
    <div class='flex items-center justify-between mb-4 gap-2 flex-wrap'>
      <div>
        <h2 class='text-lg font-semibold'>Planung (Admin)</h2>
        <p class='text-sm text-slate-500 mt-1'>Struktur: Unterregister für Personal, Abstinenz, Wochenendeinsätze und Schichttausch.</p>
      </div>
      <button class='px-2 py-1 rounded bg-red-700 text-white text-sm' onclick='resetPlanCurrentFuture()'>Gesamtplan zurücksetzen (aktuelle+zukünftige)</button>
    </div>
    <div class='flex gap-2 flex-wrap mb-5'>
      ${subTabButtons}
    </div>
    ${content}
  </div>`;
}

function renderPlanningPersonal() {
  const personnelRows = allUsers()
    .map(
      (u) => `<tr class='border-b'>
      <td class='p-2'>${u.name}</td>
      <td class='p-2'>${u.type === "core" ? "A/B/C" : "Springer"}</td>
      <td class='p-2'>${state.inactiveUsers?.[u.name] ? "Inaktiv" : "Aktiv"}</td>
      <td class='p-2'>
        ${
          state.inactiveUsers?.[u.name]
            ? `<button class='px-2 py-1 rounded bg-emerald-700 text-white' onclick="activateEmployee('${u.name}')">Reaktivieren</button>`
            : `<button class='px-2 py-1 rounded bg-rose-700 text-white' onclick="deactivateEmployee('${u.name}')">Löschen</button>`
        }
      </td>
    </tr>`,
    )
    .join("");

  const slotAssignmentRows = SLOT_CODES.map((slot) => {
    const options = activeUsers()
      .map(
        (u) =>
          `<option value='${u.name}' ${state.slotAssignments?.[slot] === u.name ? "selected" : ""}>${u.name}</option>`,
      )
      .join("");
    return `<tr class='border-b'>
      <td class='p-2 font-semibold'>${slot}</td>
      <td class='p-2'><select id='slot-${slot}' class='border rounded p-1 w-full'>${options}</select></td>
      <td class='p-2'><button class='px-2 py-1 rounded bg-slate-900 text-white' onclick="updateSlotAssignment('${slot}')">Speichern</button></td>
    </tr>`;
  }).join("");

  return `<div class='grid md:grid-cols-2 gap-4'>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <h3 class='font-semibold mb-2'>Personalverwaltung</h3>
      <div class='grid grid-cols-3 gap-2 mb-2'>
        <input id='newEmployeeName' class='border rounded p-1' placeholder='Neuer Name' />
        <select id='newEmployeeType' class='border rounded p-1'><option value='springer'>Springer</option><option value='core'>A/B/C</option></select>
        <button class='px-2 py-1 rounded bg-slate-900 text-white' onclick='addEmployee()'>Mitarbeiter hinzufügen</button>
      </div>
      <div class='overflow-auto max-h-[55vh]'>
        <table class='w-full text-sm'><thead class='bg-white sticky top-0'><tr><th class='p-2 text-left'>Name</th><th class='p-2 text-left'>Typ</th><th class='p-2 text-left'>Status</th><th class='p-2'></th></tr></thead>
        <tbody>${personnelRows}</tbody></table>
      </div>
    </div>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <h3 class='font-semibold mb-2'>Zuordnung</h3>
      <p class='text-sm text-slate-500 mb-2'>Admin kann festlegen, welcher Mitarbeiter aktuell A/B/C/D/E/F ist. Der Schichtplan passt sich danach direkt an.</p>
      <div class='overflow-auto max-h-[55vh]'>
        <table class='w-full text-sm'><thead class='bg-white sticky top-0'><tr><th class='p-2 text-left'>Slot</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2'></th></tr></thead>
        <tbody>${slotAssignmentRows}</tbody></table>
      </div>
    </div>
  </div>`;
}
function getWeekRanges(from, to) {
  const ranges = [];
  let current = new Date(`${from}T00:00:00`);
  const end = new Date(`${to}T00:00:00`);

  while (current <= end) {
    const start = new Date(current);
    const weekday = (start.getDay() + 6) % 7;
    const monday = new Date(start);
    monday.setDate(start.getDate() - weekday);

    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);

    const rangeFrom =
      start > new Date(`${from}T00:00:00`)
        ? start
        : new Date(`${from}T00:00:00`);
    const rangeTo = sunday < end ? sunday : end;

    ranges.push({
      from: isoDate(rangeFrom),
      to: isoDate(rangeTo),
    });

    current = new Date(sunday);
    current.setDate(current.getDate() + 1);
  }

  return ranges;
}

function getShiftsOfUserInRange(userName, from, to) {
  return generateThreeMonths().filter((s) => {
    return s.date >= from && s.date <= to && s.originalAssigned === userName;
  });
}

function chooseReplacementUser(absentUser, from, to, weekLabel = "") {
  const available = activeUsers().filter((u) => u.name !== absentUser);

  if (!available.length) {
    alert("Keine anderen aktiven Mitarbeiter verfügbar.");
    return Promise.resolve(null);
  }

  const host = getModalHost();

  return new Promise((resolve) => {
    host.innerHTML = `
      <div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
        <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
          <h3 class="text-lg font-bold mb-2">Ersatz auswählen</h3>
          <p class="text-sm text-slate-700 mb-3">
            ${weekLabel ? `${weekLabel}<br>` : ""}
            Wer ersetzt <b>${absentUser}</b> von <b>${formatDateDisplay(from)}</b> bis <b>${formatDateDisplay(to)}</b>?
          </p>

          <select id="replacementUserSelect" class="border rounded p-2 w-full mb-4">
            ${available
              .map((u) => `<option value="${u.name}">${u.name}</option>`)
              .join("")}
            <option value="AUSFALL">AUSFALL</option>
          </select>

          <div class="flex justify-end gap-2">
            <button id="replacementCancel" class="px-3 py-1 rounded bg-slate-200">Abbrechen</button>
            <button id="replacementSave" class="px-3 py-1 rounded bg-slate-900 text-white">Speichern</button>
          </div>
        </div>
      </div>
    `;

    host.querySelector("#replacementCancel")?.addEventListener("click", () => {
      host.innerHTML = "";
      resolve(null);
    });

    host.querySelector("#replacementSave")?.addEventListener("click", () => {
      const value = host.querySelector("#replacementUserSelect")?.value || "";

      host.innerHTML = "";

      if (value === "AUSFALL") {
        resolve({ mode: "cancel", replacementUser: null });
        return;
      }

      resolve({ mode: "replace", replacementUser: value });
    });
  });
}

async function planAbsenceReplacement(type, entryId) {
  const list = type === "vacation" ? state.vacations : state.sickLeaves;
  const entry = list.find((v) => v.id === entryId);
  if (!entry) return;

  const full = await askYesNoCentered(
    `Soll ein Mitarbeiter ${entry.user} für die komplette Zeit von ${formatDateDisplay(entry.from)} bis ${formatDateDisplay(entry.to)} ersetzen?`,
  );
  if (full === null) return;

  clearReplacementPlanForSource(type, entryId);

  if (full) {
    const choice = await chooseReplacementUser(entry.user, entry.from, entry.to);
    if (!choice) return;

    const shifts = getShiftsOfUserInRange(entry.user, entry.from, entry.to);
    shifts.forEach((shift) => {
      applyReplacementToShift(shift.id, {
        sourceType: type,
        sourceId: entryId,
        absentUser: entry.user,
        from: entry.from,
        to: entry.to,
        mode: choice.mode,
        replacementUser: choice.replacementUser,
        weekFrom: entry.from,
        weekTo: entry.to,
      });
    });

    persist();
    render();
    return;
  }

  const weeks = getWeekRanges(entry.from, entry.to);

  for (let i = 0; i < weeks.length; i++) {
    const week = weeks[i];
    const label = `Woche ${i + 1}: ${formatDateDisplay(week.from)} bis ${formatDateDisplay(week.to)}`;
    const choice = await chooseReplacementUser(
      entry.user,
      week.from,
      week.to,
      label,
    );
    if (!choice) continue;

    const shifts = getShiftsOfUserInRange(entry.user, week.from, week.to);
    shifts.forEach((shift) => {
      applyReplacementToShift(shift.id, {
        sourceType: type,
        sourceId: entryId,
        absentUser: entry.user,
        from: entry.from,
        to: entry.to,
        mode: choice.mode,
        replacementUser: choice.replacementUser,
        weekFrom: week.from,
        weekTo: week.to,
      });
    });
  }

  persist();
  render();
}

function renderPlanningAbstinenz() {
  const openShifts = generateThreeMonths().filter((s) => s.open);
  const absenceRows = Object.entries(state.absences)
    .slice(-100)
    .reverse()
    .map(([key]) => {
      const [date, user] = key.split(":");
      return `<tr class='border-b'><td class='p-2'>${formatDateWithWeekday(date)}</td><td class='p-2'>${user}</td><td class='p-2'>Kann nicht kommen</td></tr>`;
    })
    .join("");

  const monthValue = `${new Date().getFullYear()}-${String(
    new Date().getMonth() + 1,
  ).padStart(2, "0")}`;
  const personOptions = activeUsers()
    .map((u) => `<option value='${u.name}'>${u.name}</option>`)
    .join("");

  const vacationRows = state.vacations
    .slice(-20)
    .reverse()
    .map((v) => {
      const hasPlan =
        getReplacementEntriesForSource("vacation", v.id).length > 0;
      const summary = getReplacementSummaryForSource("vacation", v.id);

      return `<tr class='border-b'>
        <td class='p-2'>${v.user}</td>
        <td class='p-2'>${formatDateDisplay(v.from)}</td>
        <td class='p-2'>${formatDateDisplay(v.to)}</td>
        <td class='p-2 text-sm'>${summary}</td>
        <td class='p-2 whitespace-nowrap'>
          <button class='px-2 py-1 rounded bg-rose-700 text-white mr-2' onclick="deleteVacation('${v.id}')">Löschen</button>
          <button class='px-2 py-1 rounded ${
            hasPlan ? "bg-emerald-700" : "bg-blue-700"
          } text-white mr-2' onclick="planAbsenceReplacement('vacation','${v.id}')">${
            hasPlan ? "Bearbeiten" : "Ersetzen"
          }</button>
          ${
            hasPlan
              ? `<button class='px-2 py-1 rounded bg-slate-700 text-white' onclick="clearAbsenceReplacementPlan('vacation','${v.id}')">Ersatz löschen</button>`
              : ""
          }
        </td>
      </tr>`;
    })
    .join("");

  const sickRows = state.sickLeaves
    .slice(-20)
    .reverse()
    .map((v) => {
      const hasPlan = getReplacementEntriesForSource("sick", v.id).length > 0;
      const summary = getReplacementSummaryForSource("sick", v.id);

      return `<tr class='border-b'>
        <td class='p-2'>${v.user}</td>
        <td class='p-2'>${formatDateDisplay(v.from)}</td>
        <td class='p-2'>${formatDateDisplay(v.to)}</td>
        <td class='p-2 text-sm'>${summary}</td>
        <td class='p-2 whitespace-nowrap'>
          <button class='px-2 py-1 rounded bg-rose-700 text-white mr-2' onclick="deleteSickLeave('${v.id}')">Löschen</button>
          <button class='px-2 py-1 rounded ${
            hasPlan ? "bg-emerald-700" : "bg-blue-700"
          } text-white mr-2' onclick="planAbsenceReplacement('sick','${v.id}')">${
            hasPlan ? "Bearbeiten" : "Ersetzen"
          }</button>
          ${
            hasPlan
              ? `<button class='px-2 py-1 rounded bg-slate-700 text-white' onclick="clearAbsenceReplacementPlan('sick','${v.id}')">Ersatz löschen</button>`
              : ""
          }
        </td>
      </tr>`;
    })
    .join("");

  const rows = openShifts
    .map((s) => {
      const choices = activeUsers()
        .filter((u) => u.name !== s.assigned)
        .map((u) => `<option value="${u.name}">${u.name}</option>`)
        .join("");
      return `<tr class='border-b'>
      <td class='p-2'>${formatDateWithWeekday(s.date)}</td><td class='p-2'>${s.label}</td><td class='p-2'>${s.start}–${s.end}</td>
      <td class='p-2'>${s.originalAssigned || "-"}</td>
      <td class='p-2'>
        <select id='sel-${s.id}' class='border rounded p-1'>${choices}</select>
      </td>
      <td class='p-2 space-x-1'>
        <button class='px-2 py-1 rounded bg-slate-900 text-white' onclick="assignShift('${s.id}')">Übernehmen</button>
        <button class='px-2 py-1 rounded bg-rose-700 text-white' onclick="cancelShift('${s.id}')">Ausfall</button>
      </td>
    </tr>`;
    })
    .join("");

  return `<div class='space-y-4'>
    <div class='grid md:grid-cols-2 gap-4'>
      <div class='border rounded-lg p-3 bg-slate-50'>
        <h3 class='font-semibold mb-2'>Klärung Abstinz</h3>
        <div class='overflow-auto max-h-56'>
          <table class='w-full text-sm'><thead class='bg-white sticky top-0'><tr><th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Typ</th></tr></thead>
          <tbody>${absenceRows || '<tr><td class="p-2" colspan="3">Keine Meldungen.</td></tr>'}</tbody></table>
        </div>
      </div>
      <div class='border rounded-lg p-3 bg-slate-50 space-y-3'>
        <h3 class='font-semibold'>Urlaub / Krankmeldung eintragen</h3>
        <div class='grid grid-cols-2 gap-2 text-sm'>
          <label>Monat<input id='adminMonth' type='month' value='${monthValue}' class='border rounded p-1 w-full'/></label>
          <label>Mitarbeiter<select id='adminCalUser' class='border rounded p-1 w-full'>${personOptions}</select></label>
          <label>Von<input id='adminFrom' type='date' value='${todayIso()}' class='border rounded p-1 w-full'/></label>
          <label>Bis<input id='adminTo' type='date' value='${todayIso()}' class='border rounded p-1 w-full'/></label>
        </div>
        <div class='flex gap-2 flex-wrap'>
          <button class='px-2 py-1 rounded bg-emerald-700 text-white' onclick='addVacation()'>Urlaub speichern</button>
          <button class='px-2 py-1 rounded bg-amber-700 text-white' onclick='addSickLeave()'>Krank speichern</button>
        </div>
      </div>
    </div>

    <div class='grid grid-cols-1 gap-4'>
      <div class='border rounded-lg p-3'>
        <h4 class='font-semibold mb-2'>Geplante Urlaube</h4>
        <div class='overflow-auto max-h-56'>
          <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Von</th><th class='p-2 text-left'>Bis</th><th class='p-2 text-left'>Ersatz</th><th class='p-2'></th></tr></thead>
          <tbody>${vacationRows || '<tr><td class="p-2" colspan="5">Keine Einträge.</td></tr>'}</tbody></table>
        </div>
      </div>
      <div class='border rounded-lg p-3'>
        <h4 class='font-semibold mb-2'>Krankmeldungen</h4>
        <div class='overflow-auto max-h-56'>
          <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Von</th><th class='p-2 text-left'>Bis</th><th class='p-2 text-left'>Ersatz</th><th class='p-2'></th></tr></thead>
          <tbody>${sickRows || '<tr><td class="p-2" colspan="5">Keine Einträge.</td></tr>'}</tbody></table>
        </div>
      </div>
    </div>

    <div class='border rounded-lg p-3 bg-white'>
      <h3 class='text-md font-semibold mb-2'>Offene Schichten</h3>
      <div class='flex items-center justify-between mb-2 gap-2 flex-wrap'>
        <p class='text-sm text-slate-500'>Pro Tag kannst du entscheiden: ganze Schicht übernehmen lassen oder Ausfall markieren.</p>
        <button class='px-2 py-1 rounded bg-slate-800 text-white text-sm' onclick='resetManualAssignments()'>Zuordnungen zurücksetzen (aktuelle+zukünftige)</button>
      </div>
      <div class='overflow-auto max-h-[55vh]'>
        <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr>
        <th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Zeit</th><th class='p-2 text-left'>Vorher</th><th class='p-2 text-left'>Übernahme durch</th><th class='p-2'>Aktion</th></tr></thead>
        <tbody>${rows || '<tr><td class="p-2" colspan="6">Keine offenen Schichten.</td></tr>'}</tbody></table>
      </div>
    </div>
  </div>`;
}

function renderPlanningWochenende() {
  const saturdayRequestRows = Object.entries(state.saturdayEveningRequests)
    .filter(([, requested]) => requested)
    .map(([key]) => {
      const [shiftId, user] = key.split(":");
      const shift = generateThreeMonths().find((s) => s.id === shiftId);
      if (!shift) return "";
      return `<tr class='border-b'>
        <td class='p-2'>${formatDateWithWeekday(shift.date)}</td>
        <td class='p-2'>${user}</td>
        <td class='p-2'>${shift.label} (${shift.start}–${shift.end})</td>
        <td class='p-2'><button class='px-2 py-1 rounded bg-emerald-700 text-white' onclick="approveSaturdayRequest('${shiftId}','${user}')">Genehmigen</button></td>
      </tr>`;
    })
    .join("");

  let lastWeekend = "";
  const optionalRows = generateThreeMonths()
    .filter((s) => s.options.length > 1)
    .slice(0, 120)
    .map((s) => {
      const canUsers = activeUsers()
        .filter((u) => state.availability[`${s.id}:${u.name}`] === "yes")
        .map((u) => u.name);
      const cantUsers = activeUsers()
        .filter((u) => state.availability[`${s.id}:${u.name}`] === "no")
        .map((u) => u.name);
      const options = (
        canUsers.length ? canUsers : activeUsers().map((u) => u.name)
      )
        .map((name) => `<option value="${name}">${name}</option>`)
        .join("");
      const weekendId = s.date.slice(0, 8);
      const divider =
        weekendId !== lastWeekend
          ? `<tr class='bg-slate-200'><td class='p-2 font-semibold' colspan='7'>Wochenende ab ${formatDateDisplay(s.date)}</td></tr>`
          : "";
      lastWeekend = weekendId;
      return `${divider}<tr class='border-b'>
      <td class='p-2'>${formatDateWithWeekday(s.date)}</td>
      <td class='p-2'>${s.label}</td>
      <td class='p-2'>${formatSlot(s.options)}</td>
      <td class='p-2'>${canUsers.length ? canUsers.map((n) => `<span class='px-1 rounded bg-emerald-100 text-emerald-700 mr-1'>${n}</span>`).join("") : "-"}</td>
      <td class='p-2'>${cantUsers.length ? cantUsers.map((n) => `<span class='px-1 rounded bg-rose-100 text-rose-700 mr-1'>${n}</span>`).join("") : "-"}</td>
      <td class='p-2'><select id='opt-${s.id}' class='border rounded p-1'>${options}</select></td>
      <td class='p-2'><button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="assignOptionalShift('${s.id}')">Einteilen</button></td>
    </tr>`;
    })
    .join("");

  return `<div class='space-y-4'>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <h3 class='font-semibold mb-2'>Samstag Abend</h3>
      <div class='overflow-auto max-h-56'>
        <table class='w-full text-sm'><thead class='bg-white sticky top-0'><tr><th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Schicht</th><th class='p-2'></th></tr></thead>
        <tbody>${saturdayRequestRows || '<tr><td class="p-2" colspan="4">Keine Anfragen.</td></tr>'}</tbody></table>
      </div>
    </div>

    <div class='border rounded-lg p-3 bg-white'>
      <h3 class='text-md font-semibold mb-2'>Wochenende</h3>
      <p class='text-sm text-slate-500 mb-2'>Zugesagt = grün, nicht zugesagt = rot. Beide Gruppen bleiben einsetzbar. Wochenenden sind getrennt dargestellt.</p>
      <div class='overflow-auto max-h-[60vh]'>
        <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr>
          <th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Option</th><th class='p-2 text-left'>Kann</th><th class='p-2 text-left'>Kann nicht</th><th class='p-2 text-left'>Einteilen</th><th class='p-2'></th>
        </tr></thead>
        <tbody>${optionalRows || '<tr><td class="p-2" colspan="7">Keine Wochenend-Schichten.</td></tr>'}</tbody></table>
      </div>
    </div>
  </div>`;
}
function renderPlanningSchichttausch() {
  const corePersonOptions = activeUsers()
    .filter((u) => u.type === "core")
    .map((u) => `<option value='${u.name}'>${u.name}</option>`)
    .join("");

  const swapRows = state.swaps
    .slice()
    .reverse()
    .map(
      (swap, index) => `<tr class='border-b'>
      <td class='p-2'>${swap.userA}</td>
      <td class='p-2'>${swap.userB}</td>
      <td class='p-2'>${formatDateDisplay(swap.startDate)}</td>
      <td class='p-2'>${swap.endDate ? formatDateDisplay(swap.endDate) : "-"}</td>
      <td class='p-2'>${!swap.endDate || swap.endDate >= todayIso() ? "Aktiv/Geplant" : "Beendet"}</td>
      <td class='p-2'><button class='px-2 py-1 rounded bg-rose-700 text-white' onclick="deleteSwap(${state.swaps.length - 1 - index})">Löschen</button></td>
    </tr>`,
    )
    .join("");

  return `<div class='space-y-4'>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <h3 class='text-md font-semibold mb-2'>Schichttausch</h3>
      <div class='grid md:grid-cols-5 gap-2 items-end'>
        <label class='text-sm'>Mitarbeiter 1 (A/B/C)<select id='swapA' class='border rounded p-1 w-full'>${corePersonOptions}</select></label>
        <label class='text-sm'>Mitarbeiter 2 (A/B/C)<select id='swapB' class='border rounded p-1 w-full'>${corePersonOptions}</select></label>
        <label class='text-sm'>Gültig ab<input id='swapDate' type='date' value='${todayIso()}' class='border rounded p-1 w-full'/></label>
        <label class='text-sm'>Bis (optional)<input id='swapEndDate' type='date' class='border rounded p-1 w-full'/></label>
        <button class='px-2 py-1 rounded bg-purple-700 text-white h-9' onclick='addSwap()'>Schichttausch speichern</button>
      </div>
      <div class='mt-3 flex gap-2 flex-wrap'>
        <button class='px-2 py-1 rounded bg-slate-800 text-white text-sm' onclick='resetActiveSwaps()'>Tausch zurücksetzen (aktuelle+zukünftige)</button>
      </div>
    </div>

    <div class='border rounded-lg p-3 bg-white'>
      <h4 class='font-semibold mb-2'>Aktuelle und geplante Tausche</h4>
      <div class='overflow-auto max-h-[55vh]'>
        <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Mitarbeiter 1</th><th class='p-2 text-left'>Mitarbeiter 2</th><th class='p-2 text-left'>Ab</th><th class='p-2 text-left'>Bis</th><th class='p-2 text-left'>Status</th><th class='p-2'></th></tr></thead>
        <tbody>${swapRows || '<tr><td class="p-2" colspan="6">Keine Tausche vorhanden.</td></tr>'}</tbody></table>
      </div>
    </div>
  </div>`;
}

function canAssignUserToShift(name, shiftId) {
  const shift = getShiftById(shiftId);
  if (!shift) return true;
  const absType = getCalendarAbsenceType(name, shift.date);
  if (absType) {
    alert(
      `${name} ist am ${formatDateDisplay(shift.date)} als ${absType} gemeldet und kann nicht eingesetzt werden.`,
    );
    return false;
  }
  const restCheck = checkRestTimeRule(name, shiftId);
  if (!restCheck.ok) {
    alert(restCheck.message);
    return false;
  }
  return true;
}

function checkRestTimeRule(name, shiftId) {
  const target = getShiftById(shiftId);
  if (!target) return { ok: true };
  const assigned = generateThreeMonths()
    .filter((s) => s.assigned === name && s.id !== shiftId)
    .map((s) => ({
      ...s,
      _start: shiftDateRange(s).start,
      _end: shiftDateRange(s).end,
    }));
  const tRange = shiftDateRange(target);
  assigned.push({
    ...target,
    assigned: name,
    _start: tRange.start,
    _end: tRange.end,
  });
  assigned.sort((a, b) => a._start - b._start);

  let shortRestCountWeek = 0;
  const targetWeek = weekKey(target.date);
  for (let i = 1; i < assigned.length; i++) {
    const prev = assigned[i - 1];
    const cur = assigned[i];
    const restHours =
      (cur._start.getTime() - prev._end.getTime()) / (1000 * 60 * 60);
    if (restHours < 8) {
      return {
        ok: false,
        message: `${name} hat zwischen Schichten weniger als 8 Stunden Ruhezeit.`,
      };
    }
    if (restHours < 11 && weekKey(cur.date) === targetWeek)
      shortRestCountWeek += 1;
  }
  if (shortRestCountWeek > 1) {
    return {
      ok: false,
      message: `${name} überschreitet die Ruhezeit-Regel: max. 1 Ausnahme mit 8h pro Woche.`,
    };
  }
  return { ok: true };
}

function weekKey(iso) {
  const d = new Date(`${iso}T00:00:00`);
  const weekday = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - weekday);
  return d.toISOString().slice(0, 10);
}

function readAdminCalendarForm() {
  const user = document.getElementById("adminCalUser")?.value;
  const from = document.getElementById("adminFrom")?.value;
  const to = document.getElementById("adminTo")?.value;
  if (!user || !from || !to || to < from) {
    alert("Bitte Mitarbeiter sowie gültiges Von/Bis-Datum angeben.");
    return null;
  }
  return { user, from, to };
}

function addVacation() {
  const data = readAdminCalendarForm();
  if (!data) return;
  state.vacations.push({ id: `vac-${Date.now()}`, ...data });
  persist();
  render();
}

function addSickLeave() {
  const data = readAdminCalendarForm();
  if (!data) return;
  state.sickLeaves.push({ id: `sick-${Date.now()}`, ...data });
  persist();
  render();
}

function deleteVacation(id) {
  state.vacations = state.vacations.filter((v) => v.id !== id);
  clearReplacementPlanForSource("vacation", id);
  persist();
  render();
}

function deleteSickLeave(id) {
  state.sickLeaves = state.sickLeaves.filter((s) => s.id !== id);
  clearReplacementPlanForSource("sick", id);
  persist();
  render();
}

function addSwap() {
  const userA = document.getElementById("swapA")?.value;
  const userB = document.getElementById("swapB")?.value;
  const startDate = document.getElementById("swapDate")?.value;
  const endDate = document.getElementById("swapEndDate")?.value || null;
  if (!isCoreEmployee(userA) || !isCoreEmployee(userB)) {
    alert("Beim Tausch sind nur A/B/C erlaubt.");
    return;
  }
  if (!userA || !userB || userA === userB || !startDate) {
    alert(
      "Bitte zwei verschiedene Mitarbeiter (A/B/C) und ein Startdatum wählen.",
    );
    return;
  }
  if (endDate && endDate < startDate) {
    alert("Enddatum muss nach dem Startdatum liegen.");
    return;
  }
  state.swaps.push({ userA, userB, startDate, endDate });
  persist();
  render();
}

function deleteSwap(index) {
  if (index < 0 || index >= state.swaps.length) return;
  state.swaps.splice(index, 1);
  persist();
  render();
}

function resetActiveSwaps(silent = false) {
  const today = todayIso();
  const y = new Date(`${today}T00:00:00`);
  y.setDate(y.getDate() - 1);
  const yesterday = isoDate(y);
  state.swaps = state.swaps.flatMap((swap) => {
    if (!swap) return [];
    if (swap.startDate >= today) return [];
    if (!swap.endDate || swap.endDate >= today)
      return [{ ...swap, endDate: yesterday }];
    return [swap];
  });
  persist();
  if (!silent) {
    render();
    alert("Aktive/zukünftige Tausche wurden zurückgesetzt.");
  }
}

function resetManualAssignments(silent = false) {
  const assignments = { ...state.assignments };
  Object.keys(assignments).forEach((shiftId) => {
    const shift = getShiftById(shiftId);
    if (!shift) return;
    if (isCurrentOrFutureShift(shift)) {
      delete state.assignments[shiftId];
      delete state.shiftCancellations[shiftId];
      delete state.absenceReplacements?.[shiftId];
    }
  });
  persist();
  if (!silent) {
    render();
    alert(
      "Manuelle Zuordnungen (aktuelle/zukünftige Schichten) wurden zurückgesetzt.",
    );
  }
}

function resetPlanCurrentFuture() {
  if (
    !confirm(
      "Gesamtplan wirklich für aktuelle und zukünftige Schichten auf Ursprung zurücksetzen?",
    )
  )
    return;
  state.assignments = {};
  state.shiftCancellations = {};
  state.absences = {};
  state.swaps = [];
  state.shiftCancellations = {};
  state.absenceReplacements = {};
  state.saturdayEveningRequests = {};
  state.vacations = [];
  state.sickLeaves = [];
  state.availability = {};
  state.unmanned = {};
  state.shiftEndChecks = {};
  state.shiftStartChecks = {};
  state.checklists = {};
  state.conflicts = {};
  state.machineDowntime = {};
  state.machinePromptSeen = {};
  state.extraUsers = [];
  state.inactiveUsers = {};
  state.slotAssignments = { ...DEFAULT_SLOT_ASSIGNMENTS };
  state.planningSubTab = "personal";
  persist();
  render();
  alert("Gesamtplan wurde auf den Ursprungsplan zurückgesetzt.");
}

function updateSlotAssignment(slot) {
  const select = document.getElementById(`slot-${slot}`);
  if (!select) return;
  state.slotAssignments[slot] = select.value;
  persist();
  render();
}

function addEmployee() {
  const name = document.getElementById("newEmployeeName")?.value?.trim();
  const type = document.getElementById("newEmployeeType")?.value || "springer";
  if (!name) return alert("Bitte Namen eingeben.");
  if (allUsers().some((u) => u.name.toLowerCase() === name.toLowerCase()))
    return alert("Mitarbeiter existiert bereits.");
  const nextSlot = type === "core" ? "A" : "D";
  state.extraUsers.push({ name, slot: nextSlot, type });
  delete state.inactiveUsers[name];
  persist();
  render();
}

function deactivateEmployee(name) {
  if (!confirm(`${name} wirklich als inaktiv markieren?`)) return;
  state.inactiveUsers[name] = true;
  Object.keys(state.slotAssignments).forEach((slot) => {
    if (state.slotAssignments[slot] === name)
      state.slotAssignments[slot] = DEFAULT_SLOT_ASSIGNMENTS[slot] || null;
  });
  persist();
  render();
}

function activateEmployee(name) {
  delete state.inactiveUsers[name];
  persist();
  render();
}

function getToolLabels() {
  return [...DEFAULT_TOOL_LABELS, ...(state.toolLabelsExtra || [])];
}

function getToolManufacturers() {
  return [
    ...DEFAULT_TOOL_MANUFACTURERS,
    ...(state.toolManufacturersExtra || []),
  ];
}

function getToolHolders() {
  return [...DEFAULT_TOOL_HOLDERS];
}

function addToolLabel() {
  if (currentUser.role !== "admin") return;
  const value = prompt("Neue Bezeichnung:");
  if (!value) return;
  if (!state.toolLabelsExtra.includes(value)) state.toolLabelsExtra.push(value);
  persist();
  render();
}

function addToolManufacturer() {
  if (currentUser.role !== "admin") return;
  const value = prompt("Neuer Hersteller:");
  if (!value) return;
  if (!state.toolManufacturersExtra.includes(value))
    state.toolManufacturersExtra.push(value);
  persist();
  render();
}

function isThreadToolLabel(label) {
  return [
    "Gewindebohrer",
    "Gewindefräser",
    "Gewindeformer",
    "Gewindewirbler",
  ].includes(label);
}

function formatToolSize(tool) {
  if (!isThreadToolLabel(tool.label)) return `⌀ ${tool.diameter}`;
  const prefix = tool.threadPrefix || "";
  const base = `${prefix}${prefix ? " " : ""}${tool.diameter}`;
  if (prefix === "MF" && tool.threadPitch)
    return `${base} P${tool.threadPitch}`;
  return base;
}
function isRadiusToolLabel(label) {
  return label === "Radiusfräser";
}

function updateToolTypeFields(prefix = "tool") {
  const labelEl = document.getElementById(`${prefix}Label`);
  const threadPrefixWrap = document.getElementById(`${prefix}ThreadPrefixWrap`);
  const threadPitchWrap = document.getElementById(`${prefix}ThreadPitchWrap`);
  const cornerRadiusWrap = document.getElementById(`${prefix}CornerRadiusWrap`);

  const label = labelEl?.value || "";
  const isThread = isThreadToolLabel(label);
  const isRadius = isRadiusToolLabel(label);

  if (threadPrefixWrap) threadPrefixWrap.style.display = isThread ? "" : "none";
  if (threadPitchWrap)
    threadPitchWrap.style.display =
      isThread && labelEl?.value && document.getElementById(`${prefix}ThreadPrefix`)?.value === "MF"
        ? ""
        : "none";
  if (cornerRadiusWrap) cornerRadiusWrap.style.display = isRadius ? "" : "none";
}

function updateThreadPitchVisibility(prefix = "tool") {
  const label = document.getElementById(`${prefix}Label`)?.value || "";
  const threadPrefix = document.getElementById(`${prefix}ThreadPrefix`)?.value || "";
  const threadPitchWrap = document.getElementById(`${prefix}ThreadPitchWrap`);

  const visible = isThreadToolLabel(label) && threadPrefix === "MF";
  if (threadPitchWrap) threadPitchWrap.style.display = visible ? "" : "none";
}

function collectToolFormData(root = document) {
  return {
    tNumber: root.getElementById("toolTNumber")?.value?.trim(),
    label: root.getElementById("toolLabel")?.value,
    diameter: root.getElementById("toolDiameter")?.value?.trim(),
    threadPrefix: root.getElementById("toolThreadPrefix")?.value || "",
    threadPitch: root.getElementById("toolThreadPitch")?.value?.trim() || "",
    cornerRadius: root.getElementById("toolCornerRadius")?.value?.trim() || "",
    shelf: root.getElementById("toolShelf")?.value?.trim().toUpperCase(),
    articleNo: root.getElementById("toolArticle")?.value?.trim(),
    holder: root.getElementById("toolHolder")?.value,
    stock: Number(root.getElementById("toolStock")?.value || 0),
    minStock: Number(root.getElementById("toolMinStock")?.value || 0),
    optimalStock: Number(root.getElementById("toolOptimalStock")?.value || 0),
    manufacturer: root.getElementById("toolManufacturer")?.value,
    insertTool: !!root.getElementById("toolInsertTool")?.checked,
    insertEdges: Number(root.getElementById("toolInsertEdges")?.value || 0),
  };
}

function validateToolData(data) {
  const isThreadTool = isThreadToolLabel(data.label);
  const isRadiusTool = isRadiusToolLabel(data.label);

  if (
    !data.tNumber ||
    !data.label ||
    !data.diameter ||
    !/^[A-Z]\d{2}$/.test(data.shelf) ||
    !data.articleNo ||
    !["HSK 100", "HSK 63"].includes(data.holder)
  ) {
    return {
      ok: false,
      message:
        "Bitte Felder korrekt ausfüllen (A-Z + 2-stellige Zahl für Fach, Aufnahme HSK 100 oder HSK 63).",
    };
  }

  if (isThreadTool && !data.threadPrefix) {
    return { ok: false, message: "Bitte Gewindekennung wählen." };
  }

  if (isThreadTool && data.threadPrefix === "MF" && !data.threadPitch) {
    return {
      ok: false,
      message: "Bitte bei MF die Steigung (P) angeben.",
    };
  }

  if (isRadiusTool && !data.cornerRadius) {
    return {
      ok: false,
      message: "Bitte Schneidenradius eingeben.",
    };
  }

  if (!isThreadTool && !Number.isFinite(Number(data.diameter))) {
    return {
      ok: false,
      message: "Bitte gültigen numerischen Durchmesser eingeben.",
    };
  }

  if (
    data.insertTool &&
    (!Number.isFinite(data.insertEdges) || data.insertEdges <= 0)
  ) {
    return {
      ok: false,
      message: "Bitte Anzahl der Schneiden > 0 eingeben.",
    };
  }

  return { ok: true };
}

function normalizeToolData(data) {
  const isThreadTool = isThreadToolLabel(data.label);
  const isRadiusTool = isRadiusToolLabel(data.label);

  return {
    tNumber: data.tNumber,
    label: data.label,
    diameter: isThreadTool ? data.diameter : String(data.diameter),
    threadPrefix: isThreadTool ? data.threadPrefix : "",
    threadPitch:
      isThreadTool && data.threadPrefix === "MF" ? data.threadPitch : "",
    cornerRadius: isRadiusTool ? data.cornerRadius : "",
    shelf: data.shelf,
    articleNo: data.articleNo,
    holder: data.holder,
    stock: data.stock,
    minStock: data.minStock,
    optimalStock: Math.max(0, data.optimalStock),
    manufacturer: data.manufacturer,
    ordered: false,
    orderedQty: 0,
    insertTool: data.insertTool,
    insertEdges: data.insertTool ? data.insertEdges : 0,
  };
}
async function createTool() {
  if (currentUser.role !== "admin")
    return alert("Nur Admin darf Werkzeuge anlegen.");

  const data = collectToolFormData(document);
  const validation = validateToolData(data);
  if (!validation.ok) return alert(validation.message);

  const normalized = normalizeToolData(data);

  const payload = {
    t_number: String(normalized.tNumber),
    label: normalized.label,
    diameter: String(normalized.diameter),
    thread_prefix: normalized.threadPrefix || null,
    thread_pitch: normalized.threadPitch || null,
    corner_radius: normalized.cornerRadius || null,
    shelf: normalized.shelf,
    article_no: normalized.articleNo,
    holder: normalized.holder,
    stock: Number(normalized.stock || 0),
    min_stock: Number(normalized.minStock || 0),
    optimal_stock: Number(normalized.optimalStock || 0),
    manufacturer: normalized.manufacturer || null,
    ordered: !!normalized.ordered,
    ordered_qty: Number(normalized.orderedQty || 0),
    insert_tool: !!normalized.insertTool,
    insert_edges: Number(normalized.insertEdges || 0),
    created_by_employee_id: currentEmployeeRecord?.id || null,
  };

  const { data: inserted, error } = await supabaseClient
    .from("tools")
    .insert(payload)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Anlegen des Werkzeugs in Supabase:", error);
    return alert(`Werkzeug konnte nicht gespeichert werden: ${error.message}`);
  }

  state.tools.push(normalizeToolFromDb(inserted));
  persist();
  render();
}

function getModalHost() {
  let host = document.getElementById("centerModalHost");
  if (!host) {
    host = document.createElement("div");
    host.id = "centerModalHost";
    document.body.appendChild(host);
  }
  return host;
}

function askYesNoCentered(message) {
  const host = getModalHost();
  return new Promise((resolve) => {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
        <h3 class="text-lg font-bold mb-2">Bestätigung</h3>
        <p class="text-sm text-slate-700 mb-4">${message}</p>
        <div class="flex justify-end gap-2">
          <button id="modalNo" class="px-3 py-1 rounded bg-slate-200">Nein</button>
          <button id="modalYes" class="px-3 py-1 rounded bg-emerald-700 text-white">Ja</button>
        </div>
      </div>
    </div>`;
    host.querySelector("#modalYes")?.addEventListener("click", () => {
      host.innerHTML = "";
      resolve(true);
    });
    host.querySelector("#modalNo")?.addEventListener("click", () => {
      host.innerHTML = "";
      resolve(false);
    });
  });
}

function askNumberCentered(message, initialValue = "1") {
  const host = getModalHost();
  return new Promise((resolve) => {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
        <h3 class="text-lg font-bold mb-2">Eingabe</h3>
        <p class="text-sm text-slate-700 mb-3">${message}</p>
        <input id="modalNumber" type="number" min="1" class="border rounded p-2 w-full mb-4" value="${initialValue}" />
        <div class="flex justify-end gap-2">
          <button id="modalNo" class="px-3 py-1 rounded bg-slate-200">Abbrechen</button>
          <button id="modalOk" class="px-3 py-1 rounded bg-slate-900 text-white">Bestätigen</button>
        </div>
      </div>
    </div>`;
    const input = host.querySelector("#modalNumber");
    input?.focus();
    host.querySelector("#modalOk")?.addEventListener("click", () => {
      const val = Number(input?.value || 0);
      host.innerHTML = "";
      resolve(Number.isFinite(val) ? val : null);
    });
    host.querySelector("#modalNo")?.addEventListener("click", () => {
      host.innerHTML = "";
      resolve(null);
    });
  });
}

function editToolCentered(tool) {
  const host = getModalHost();
  return new Promise((resolve) => {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-2xl p-4">
        <h3 class="text-lg font-bold mb-3">Werkzeug bearbeiten – T ${tool.tNumber}</h3>
        <div class="grid md:grid-cols-2 gap-2 mb-4">
          <input id="editLabel" class="border rounded p-2" value="${tool.label}" />
          <input id="editDiameter" class="border rounded p-2" value="${tool.diameter}" />
          <select id="editThreadPrefix" class="border rounded p-2">
            <option value="" ${!tool.threadPrefix ? "selected" : ""}>Kennung (nur Gewinde)</option>
            <option value="M" ${tool.threadPrefix === "M" ? "selected" : ""}>M</option>
            <option value="MF" ${tool.threadPrefix === "MF" ? "selected" : ""}>MF</option>
            <option value="G" ${tool.threadPrefix === "G" ? "selected" : ""}>G</option>
            <option value="UNF" ${tool.threadPrefix === "UNF" ? "selected" : ""}>UNF</option>
            <option value="UNC" ${tool.threadPrefix === "UNC" ? "selected" : ""}>UNC</option>
            <option value="Mx" ${tool.threadPrefix === "Mx" ? "selected" : ""}>Mx</option>
          </select>
          <input id="editThreadPitch" class="border rounded p-2" value="${tool.threadPitch || ""}" placeholder="Steigung P (nur MF)" />
          <input id="editShelf" class="border rounded p-2" value="${tool.shelf}" />
          <input id="editArticleNo" class="border rounded p-2" value="${tool.articleNo}" />
          <select id="editHolder" class="border rounded p-2">
            <option value="HSK 100" ${tool.holder === "HSK 100" ? "selected" : ""}>HSK 100</option>
            <option value="HSK 63" ${tool.holder === "HSK 63" ? "selected" : ""}>HSK 63</option>
          </select>
          <input id="editManufacturer" class="border rounded p-2" value="${tool.manufacturer}" />
          <input id="editStock" type="number" class="border rounded p-2" value="${tool.stock}" />
          <input id="editMinStock" type="number" class="border rounded p-2" value="${tool.minStock}" />
          <input id="editOptimalStock" type="number" class="border rounded p-2" value="${tool.optimalStock || 0}" />
          <label class="flex items-center gap-2 text-sm"><input id="editInsertTool" type="checkbox" ${tool.insertTool ? "checked" : ""}/> Wendeplattenwerkzeug</label>
          <input id="editInsertEdges" type="number" class="border rounded p-2" value="${tool.insertEdges || 0}" placeholder="Anzahl Schneiden" />
        </div>
        <div class="flex justify-end gap-2">
          <button id="modalNo" class="px-3 py-1 rounded bg-slate-200">Abbrechen</button>
          <button id="modalSave" class="px-3 py-1 rounded bg-slate-900 text-white">Speichern</button>
        </div>
      </div>
    </div>`;
    host.querySelector("#modalSave")?.addEventListener("click", () => {
      const payload = {
        label: host.querySelector("#editLabel")?.value?.trim() || "",
        diameter: host.querySelector("#editDiameter")?.value?.trim() || "",
        threadPrefix: host.querySelector("#editThreadPrefix")?.value || "",
        threadPitch:
          host.querySelector("#editThreadPitch")?.value?.trim() || "",
        shelf:
          host.querySelector("#editShelf")?.value?.trim().toUpperCase() || "",
        articleNo: host.querySelector("#editArticleNo")?.value?.trim() || "",
        holder: host.querySelector("#editHolder")?.value || "",
        manufacturer:
          host.querySelector("#editManufacturer")?.value?.trim() || "",
        stock: Number(host.querySelector("#editStock")?.value || 0),
        minStock: Number(host.querySelector("#editMinStock")?.value || 0),
        optimalStock: Number(
          host.querySelector("#editOptimalStock")?.value || 0,
        ),
        insertTool: !!host.querySelector("#editInsertTool")?.checked,
        insertEdges: Number(host.querySelector("#editInsertEdges")?.value || 0),
      };
      host.innerHTML = "";
      resolve(payload);
    });
    host.querySelector("#modalNo")?.addEventListener("click", () => {
      host.innerHTML = "";
      resolve(null);
    });
  });
}

function renderToolCreateForm(prefix = "tool") {
  const labels = getToolLabels();
  const manufacturers = getToolManufacturers();
  const holders = getToolHolders();

  const labelOptions = labels
    .map((l) => `<option value="${l}">${l}</option>`)
    .join("");
  const manufacturerOptions = manufacturers
    .map((m) => `<option value="${m}">${m}</option>`)
    .join("");
  const holderOptions = holders
    .map((h) => `<option value="${h}">${h}</option>`)
    .join("");

  return `<div class='grid md:grid-cols-2 gap-3'>
    <input id='${prefix}TNumber' class='border rounded p-2' placeholder='T-Nummer (z.B. 134)' />

    <select id='${prefix}Label' class='border rounded p-2' onchange='updateToolTypeFields("${prefix}")'>
      ${labelOptions}
    </select>

    <input id='${prefix}Diameter' class='border rounded p-2' placeholder='Durchmesser' />

    <div id='${prefix}ThreadPrefixWrap' style='display:none;'>
      <select id='${prefix}ThreadPrefix' class='border rounded p-2 w-full' onchange='updateThreadPitchVisibility("${prefix}")'>
        <option value=''>Kennung (nur Gewinde)</option>
        <option value='M'>M</option>
        <option value='MF'>MF</option>
        <option value='G'>G</option>
        <option value='UNF'>UNF</option>
        <option value='UNC'>UNC</option>
        <option value='Mx'>Mx</option>
      </select>
    </div>

    <div id='${prefix}ThreadPitchWrap' style='display:none;'>
      <input id='${prefix}ThreadPitch' class='border rounded p-2 w-full' placeholder='Steigung P (nur MF)' />
    </div>

    <div id='${prefix}CornerRadiusWrap' style='display:none;'>
      <input id='${prefix}CornerRadius' class='border rounded p-2 w-full' placeholder='Schneidenradius' />
    </div>

    <input id='${prefix}Shelf' class='border rounded p-2' placeholder='A00' />
    <input id='${prefix}Article' class='border rounded p-2' placeholder='Artikel Nr.' />
    <select id='${prefix}Holder' class='border rounded p-2'>${holderOptions}</select>
    <input id='${prefix}Stock' type='number' class='border rounded p-2' placeholder='Bestand' />
    <input id='${prefix}MinStock' type='number' class='border rounded p-2' placeholder='Mindestbestand' />
    <input id='${prefix}OptimalStock' type='number' class='border rounded p-2' placeholder='Optimale Stückzahl' />
    <select id='${prefix}Manufacturer' class='border rounded p-2'>${manufacturerOptions}</select>
    <label class='flex items-center gap-2 text-sm md:col-span-2'><input id='${prefix}InsertTool' type='checkbox' onchange='toggleInsertToolFieldsById("${prefix}InsertTool","${prefix}InsertEdges")' /> Wendeplattenwerkzeug</label>
    <input id='${prefix}InsertEdges' type='number' class='border rounded p-2 md:col-span-2' placeholder='Anzahl Schneiden' disabled />
  </div>`;
}

function openCreateToolModal() {
  if (currentUser?.role !== "admin") return;
  const host = getModalHost();
  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-3xl max-h-[90vh] overflow-auto p-4">
      <div class='flex items-center justify-between mb-4 gap-3'>
        <h3 class="text-lg font-bold">Neues Werkzeug anlegen</h3>
        <button id="toolCreateCloseTop" class="px-3 py-1 rounded bg-slate-200">Schließen</button>
      </div>
      <div class='space-y-4'>
        <div class='border rounded-lg p-3 bg-slate-50'>
          ${renderToolCreateForm("tool")}
        </div>
        <div class='flex justify-between gap-2 flex-wrap'>
          <div class='flex gap-2 flex-wrap'>
            <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='addToolLabel()'>Bezeichnung hinzufügen</button>
            <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='addToolManufacturer()'>Hersteller hinzufügen</button>
          </div>
          <div class='flex gap-2'>
            <button id="toolCreateCloseBottom" class="px-3 py-2 rounded bg-slate-200">Abbrechen</button>
            <button id="toolCreateSave" class="px-3 py-2 rounded bg-slate-900 text-white">Werkzeug speichern</button>
          </div>
        </div>
      </div>
    </div>
  </div>`;

  updateToolTypeFields("tool");

  const close = () => {
    host.innerHTML = "";
  };

  host.querySelector("#toolCreateCloseTop")?.addEventListener("click", close);
  host
    .querySelector("#toolCreateCloseBottom")
    ?.addEventListener("click", close);

  host.querySelector("#toolCreateSave")?.addEventListener("click", async () => {
    if (currentUser.role !== "admin")
      return alert("Nur Admin darf Werkzeuge anlegen.");

    const data = collectToolFormData(document);
    const validation = validateToolData(data);
    if (!validation.ok) return alert(validation.message);

    const normalized = normalizeToolData(data);

    const payload = {
      t_number: String(normalized.tNumber),
      label: normalized.label,
      diameter: String(normalized.diameter),
      thread_prefix: normalized.threadPrefix || null,
      thread_pitch: normalized.threadPitch || null,
      corner_radius: normalized.cornerRadius || null,
      shelf: normalized.shelf,
      article_no: normalized.articleNo,
      holder: normalized.holder,
      stock: Number(normalized.stock || 0),
      min_stock: Number(normalized.minStock || 0),
      optimal_stock: Number(normalized.optimalStock || 0),
      manufacturer: normalized.manufacturer || null,
      ordered: !!normalized.ordered,
      ordered_qty: Number(normalized.orderedQty || 0),
      insert_tool: !!normalized.insertTool,
      insert_edges: Number(normalized.insertEdges || 0),
      created_by_employee_id: currentEmployeeRecord?.id || null,
    };

    const { data: inserted, error } = await supabaseClient
      .from("tools")
      .insert(payload)
      .select()
      .single();

    if (error) {
      console.error("Fehler beim Anlegen des Werkzeugs in Supabase:", error);
      return alert(`Werkzeug konnte nicht gespeichert werden: ${error.message}`);
    }

    state.tools.push(normalizeToolFromDb(inserted));
    persist();
    close();
    render();
  });
}

function suggestedOrderQty(tool) {
  const stock = Number(tool.stock || 0);
  const optimal = Number(tool.optimalStock || 0);
  const minStock = Number(tool.minStock || 0);

  const baseFromOptimal = Math.max(0, optimal - stock);
  const statSuggestion = getOptimalQtySuggestion(tool);

  if (Number.isFinite(statSuggestion) && statSuggestion > 0) {
    return Math.max(baseFromOptimal, statSuggestion);
  }

  if (baseFromOptimal > 0) return baseFromOptimal;

  if (stock <= minStock) {
    return Math.max(1, minStock - stock + 1);
  }

  return 0;
}

function effectiveOrderQty(tool) {
  const override = Number(state.toolOrderOverrides?.[tool.id]);
  if (Number.isFinite(override) && override >= 0) return override;
  return suggestedOrderQty(tool);
}

function setToolOrderOverride(toolId, value) {
  if (currentUser.role !== "admin") return;
  const qty = Math.max(0, Number(value || 0));
  if (!state.toolOrderOverrides) state.toolOrderOverrides = {};
  state.toolOrderOverrides[toolId] = qty;
  const tool = state.tools.find((t) => t.id === toolId);
  if (tool?.ordered) tool.orderedQty = qty;
  persist();
}

function archiveOrderEvent(tool, qty, action) {
  if (!state.orderArchive) state.orderArchive = [];
  if (!state.orderHistory) state.orderHistory = [];
  const entry = {
    id: `order-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    at: new Date().toISOString(),
    action,
    toolId: tool.id,
    tNumber: tool.tNumber,
    label: tool.label,
    size: formatToolSize(tool),
    manufacturer: tool.manufacturer || "Ohne Hersteller",
    articleNo: tool.articleNo,
    shelf: tool.shelf,
    qty: Number(qty || 0),
    user: currentUser?.name || "System",
  };
  state.orderArchive.unshift(entry);
  state.orderHistory.unshift(entry);
}

function cleanupOrderArchive() {
  if (!state.orderArchive) state.orderArchive = [];
  const now = Date.now();
  const sixWeeksMs = 6 * 7 * 24 * 60 * 60 * 1000;
  state.orderArchive = state.orderArchive.filter(
    (e) => now - new Date(e.at).getTime() <= sixWeeksMs,
  );
}

function markToolOrdered(toolId, ordered) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool || currentUser.role !== "admin") return;
  tool.ordered = ordered;
  tool.orderedQty = ordered ? Math.max(1, effectiveOrderQty(tool)) : 0;
  if (ordered) archiveOrderEvent(tool, tool.orderedQty, "mark_ordered");
  persist();
  render();
}

async function restockTool(toolId) {
  if (currentUser.role !== "admin") return;
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  let add = 0;
  if (tool.ordered && Number(tool.orderedQty || 0) > 0) {
    const full = await askYesNoCentered("Bestellte Menge einlagern?");
    if (full) {
      add = Number(tool.orderedQty || 0);
    } else {
      add = await askNumberCentered(
        "Einzulagernde Menge eingeben:",
        String(tool.orderedQty || 1),
      );
      if (add === null) return;
    }
  } else {
    add = await askNumberCentered("Werkzeug einlagern (Anzahl):", "1");
    if (add === null) return;
  }

  if (!Number.isFinite(add) || add <= 0) return;

  tool.stock += add;

  if (tool.ordered && Number(tool.orderedQty || 0) > 0) {
    tool.orderedQty = Math.max(0, Number(tool.orderedQty || 0) - add);
    if (tool.orderedQty === 0) tool.ordered = false;
  }

  if (tool.stock >= tool.minStock && Number(tool.orderedQty || 0) === 0)
    tool.ordered = false;

  archiveOrderEvent(tool, add, "restock");
  persist();
  render();
}

async function bookToolChange(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const takeOut = await askYesNoCentered("Werkzeug Entnahme?");
  let qty = 0;

  if (takeOut) {
    const defaultQty =
      tool.insertTool && Number(tool.insertEdges || 0) > 0
        ? String(tool.insertEdges)
        : "1";
    qty = await askNumberCentered("Entnahmemenge eingeben:", defaultQty);
    if (qty === null) return;
    if (!Number.isFinite(qty) || qty <= 0)
      return alert("Bitte eine gültige Entnahmemenge eingeben.");
    tool.stock = Math.max(0, tool.stock - qty);
  }

  state.toolJournal.unshift({
    id: `journal-${Date.now()}`,
    user: currentUser.name,
    at: new Date().toISOString().slice(0, 16).replace("T", " "),
    toolId: tool.id,
    tNumber: tool.tNumber,
    qty,
    action: takeOut
      ? `Werkzeugwechsel + Entnahme ${qty}`
      : "Werkzeugwechsel ohne Entnahme",
  });

  persist();
  render();
}

async function editTool(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool || currentUser.role !== "admin") return;
  const data = await editToolCentered(tool);
  if (!data) return;

  const isThreadTool = isThreadToolLabel(data.label);
  if (
    !data.label ||
    !data.diameter ||
    !/^[A-Z]\d{2}$/.test(data.shelf) ||
    !data.articleNo ||
    !["HSK 100", "HSK 63"].includes(data.holder)
  ) {
    return alert(
      "Bitte Felder korrekt ausfüllen (A-Z + 2-stellige Zahl für Fach, Aufnahme HSK 100 oder HSK 63).",
    );
  }
  if (isThreadTool && !data.threadPrefix)
    return alert("Bitte Gewindekennung wählen.");
  if (isThreadTool && data.threadPrefix === "MF" && !data.threadPitch)
    return alert("Bitte bei MF die Steigung (P) angeben.");
  if (!isThreadTool && !Number.isFinite(Number(data.diameter)))
    return alert("Bitte gültigen numerischen Durchmesser eingeben.");
  if (
    data.insertTool &&
    (!Number.isFinite(Number(data.insertEdges)) ||
      Number(data.insertEdges) <= 0)
  )
    return alert("Bitte Anzahl der Schneiden > 0 eingeben.");

  data.diameter = isThreadTool ? data.diameter : Number(data.diameter);
  data.threadPitch =
    isThreadTool && data.threadPrefix === "MF" ? data.threadPitch : "";
  if (!isThreadTool) data.threadPrefix = "";

  Object.assign(tool, data);
  if (tool.stock >= tool.minStock && Number(tool.orderedQty || 0) === 0)
    tool.ordered = false;
  persist();
  render();
}
async function deleteTool(toolId) {
  if (currentUser.role !== "admin") return;
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;
  const yes = await askYesNoCentered(
    `Werkzeug T ${tool.tNumber} wirklich löschen?`,
  );
  if (!yes) return;
  state.tools = state.tools.filter((t) => t.id !== toolId);
  state.toolJournal = state.toolJournal.filter((j) => j.toolId !== toolId);
  persist();
  render();
}

function undoToolJournalEntry(entryId) {
  const idx = state.toolJournal.findIndex((j) => j.id === entryId);
  if (idx === -1) return;
  const entry = state.toolJournal[idx];
  const tool = state.tools.find((t) => t.id === entry.toolId);
  if (tool && Number(entry.qty) > 0) {
    tool.stock += Number(entry.qty);
    if (tool.stock >= tool.minStock && Number(tool.orderedQty || 0) === 0)
      tool.ordered = false;
  }
  state.toolJournal.splice(idx, 1);
  persist();
  render();
}

function resetToolFilters() {
  state.toolFilters = {
    search: "",
    label: "",
    tNumber: "",
    diameter: "",
    holder: "",
  };
  persist();
  render();
}

function applyToolFilters() {
  const search = document.getElementById("toolSearch")?.value || "";
  const label = document.getElementById("toolFilterLabel")?.value || "";
  const tNumber = document.getElementById("toolFilterT")?.value || "";
  const diameter = document.getElementById("toolFilterD")?.value || "";
  const holder = document.getElementById("toolFilterHolder")?.value || "";
  state.toolFilters = { search, label, tNumber, diameter, holder };
  persist();
  render();
}

function getOrderCandidateGroups() {
  return state.tools
    .filter((t) => shouldOrderTool(t) || t.ordered)
    .reduce((acc, tool) => {
      const maker = (tool.manufacturer || "").trim() || "Ohne Hersteller";
      if (!acc[maker]) acc[maker] = [];
      acc[maker].push(tool);
      return acc;
    }, {});
}

function ensureSelectedOrderListManufacturer(groups) {
  const makers = Object.keys(groups).sort((a, b) => a.localeCompare(b));
  if (!makers.length) {
    state.selectedOrderListManufacturer = "";
    return "";
  }
  if (!makers.includes(state.selectedOrderListManufacturer || "")) {
    state.selectedOrderListManufacturer = makers[0];
  }
  return state.selectedOrderListManufacturer;
}

function openOrderListPopup() {
  state.orderListPopupOpen = true;
  const groups = getOrderCandidateGroups();
  ensureSelectedOrderListManufacturer(groups);
  persist();
  render();
}

function closeOrderListPopup() {
  state.orderListPopupOpen = false;
  render();
}

function setOrderListManufacturer(value) {
  state.selectedOrderListManufacturer = value || "";
  persist();
  render();
}

function setOrderStatsView(view) {
  if (!["week", "month", "year"].includes(view)) return;
  state.orderStatsView = view;
  persist();
  render();
}

function getOrderStatsRange(view) {
  const now = new Date();
  const start = new Date(now);
  const end = new Date(now);
  if (view === "week") {
    start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
    end.setDate(start.getDate() + 6);
  } else if (view === "month") {
    start.setDate(1);
    end.setMonth(start.getMonth() + 1, 0);
  } else {
    start.setMonth(0, 1);
    end.setMonth(11, 31);
  }
  return { start, end };
}

function buildOrderFrequency(view) {
  const { start, end } = getOrderStatsRange(view);
  const rows = (state.orderHistory || []).filter((e) => {
    const d = new Date(e.at);
    return (
      d >= start &&
      d <= end &&
      (e.action === "mark_ordered" || e.action === "restock")
    );
  });

  const map = {};
  rows.forEach((e) => {
    const key = `${e.toolId}`;
    if (!map[key]) {
      map[key] = {
        toolId: e.toolId,
        label: e.label,
        size: e.size,
        manufacturer: e.manufacturer,
        articleNo: e.articleNo,
        count: 0,
        qtyTotal: 0,
      };
    }
    map[key].count += 1;
    map[key].qtyTotal += Number(e.qty || 0);
  });

  return Object.values(map).sort((a, b) => b.count - a.count);
}

function getOptimalQtySuggestion(tool) {
  const yearData = buildOrderFrequency("year").find(
    (r) => r.toolId === tool.id,
  );
  if (!yearData || yearData.count < 6) return null;
  const avg = Math.ceil(yearData.qtyTotal / yearData.count);
  return Math.max(1, avg);
}

function shouldShowOptimalQtySuggestion(tool) {
  const suggestion = getOptimalQtySuggestion(tool);
  if (!Number.isFinite(suggestion) || suggestion <= 0) return false;
  const stateEntry = state.orderSuggestionState?.[tool.id];
  if (!stateEntry) return true;
  if (stateEntry.accepted) return false;
  const last = new Date(stateEntry.lastDecisionAt || 0).getTime();
  const days30 = 30 * 24 * 60 * 60 * 1000;
  return Date.now() - last >= days30;
}

function applyOptimalQtySuggestion(toolId) {
  if (currentUser.role !== "admin") return;
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;
  const s = getOptimalQtySuggestion(tool);
  if (!Number.isFinite(s) || s <= 0) return;
  tool.optimalStock = s;
  if (!state.orderSuggestionState) state.orderSuggestionState = {};
  state.orderSuggestionState[toolId] = {
    lastDecisionAt: new Date().toISOString(),
    accepted: true,
  };
  persist();
  render();
}

function rejectOptimalQtySuggestion(toolId) {
  if (currentUser.role !== "admin") return;
  if (!state.orderSuggestionState) state.orderSuggestionState = {};
  state.orderSuggestionState[toolId] = {
    lastDecisionAt: new Date().toISOString(),
    accepted: false,
  };
  persist();
  render();
}

function renderOrderStats() {
  if (currentUser.role !== "admin")
    return `<div class='bg-white rounded-xl shadow p-4'><p>Kein Zugriff.</p></div>`;
  const view = state.orderStatsView || "week";
  const freq = buildOrderFrequency(view);

  const rows = freq
    .map(
      (r) => `<tr class='border-b'>
    <td class='p-2'>${r.label}</td>
    <td class='p-2'>${r.size}</td>
    <td class='p-2'>${r.manufacturer}</td>
    <td class='p-2'>${r.articleNo}</td>
    <td class='p-2'>${r.count}</td>
    <td class='p-2'>${r.qtyTotal}</td>
  </tr>`,
    )
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4 space-y-3'>
    <h2 class='text-lg font-semibold'>Bestell-Statistik</h2>
    <div class='flex gap-2'>
      <button class='px-2 py-1 rounded ${view === "week" ? "bg-slate-900 text-white" : "bg-slate-200"}' onclick="setOrderStatsView('week')">Woche</button>
      <button class='px-2 py-1 rounded ${view === "month" ? "bg-slate-900 text-white" : "bg-slate-200"}' onclick="setOrderStatsView('month')">Monat</button>
      <button class='px-2 py-1 rounded ${view === "year" ? "bg-slate-900 text-white" : "bg-slate-200"}' onclick="setOrderStatsView('year')">Jahr</button>
    </div>
    <div class='border rounded-lg overflow-auto max-h-[60vh]'>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100 sticky top-0'>
          <tr>
            <th class='p-2 text-left'>Bezeichnung</th>
            <th class='p-2 text-left'>Größe</th>
            <th class='p-2 text-left'>Hersteller</th>
            <th class='p-2 text-left'>Artikelnummer</th>
            <th class='p-2 text-left'>Bestellvorgänge</th>
            <th class='p-2 text-left'>Gesamtmenge</th>
          </tr>
        </thead>
        <tbody>${rows || '<tr><td class="p-2" colspan="6">Keine Daten im gewählten Zeitraum.</td></tr>'}</tbody>
      </table>
    </div>
  </div>`;
}

function renderTools() {
  const labels = getToolLabels();
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

  const filterLabelOptions = labels
    .map(
      (l) =>
        `<option value="${l}" ${l === filterLabel ? "selected" : ""}>${l}</option>`,
    )
    .join("");

  const tools = state.tools.filter((t) => {
    const bySearch =
      !search ||
      `${t.tNumber} ${t.label} ${t.diameter} ${t.articleNo} ${t.holder || ""} ${t.manufacturer || ""}`
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
      <td class='p-2'>${t.manufacturer || "-"}</td>
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
  const orderGroups = getOrderCandidateGroups();
  const availableManufacturers = Object.keys(orderGroups).sort((a, b) =>
    a.localeCompare(b),
  );
  const selectedManufacturer = ensureSelectedOrderListManufacturer(orderGroups);
  const selectedManufacturerTools = selectedManufacturer
    ? orderGroups[selectedManufacturer] || []
    : [];

  const todoRows = todoTools
    .map(
      (t) => `<tr class='border-b'>
    <td class='p-2'>T ${t.tNumber}</td>
    <td class='p-2'>${t.label}</td>
    <td class='p-2'>${formatToolSize(t)}</td>
    <td class='p-2'>${t.articleNo || "-"}</td>
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
    <div class='flex flex-col sm:flex-row gap-2'>
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

  const manufacturerOptionsForPopup = availableManufacturers
    .map(
      (maker) =>
        `<option value="${maker}" ${maker === selectedManufacturer ? "selected" : ""}>${maker}</option>`,
    )
    .join("");

  const selectedManufacturerRows = selectedManufacturerTools
    .map(
      (t) => `<tr class='border-b'>
        <td class='p-2'>T ${t.tNumber}</td>
        <td class='p-2'>${t.label}</td>
        <td class='p-2'>${formatToolSize(t)}</td>
        <td class='p-2'>${t.articleNo || "-"}</td>
        <td class='p-2'>${t.stock}/${t.minStock}</td>
        <td class='p-2'>
          <input
            type='number'
            min='0'
            class='border rounded p-1 w-24'
            value='${effectiveOrderQty(t)}'
            onchange="setToolOrderOverride('${t.id}', this.value)"
          />
        </td>
      </tr>`,
    )
    .join("");

  const orderListPopup = state.orderListPopupOpen
    ? `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-6xl max-h-[88vh] overflow-auto p-4">
      <div class="flex items-center justify-between mb-3 gap-3">
        <h3 class="text-lg font-bold">Bestellliste nach Hersteller</h3>
        <button class="px-2 py-1 rounded bg-slate-200" onclick="closeOrderListPopup()">Schließen</button>
      </div>
      ${
        availableManufacturers.length
          ? `<div class='border rounded p-3 mb-4 bg-slate-50'>
          <div class='grid md:grid-cols-[280px,1fr] gap-3 items-end'>
            <label class='text-sm font-medium'>Hersteller auswählen
              <select class='border rounded p-2 w-full mt-1' onchange="setOrderListManufacturer(this.value)">
                ${manufacturerOptionsForPopup}
              </select>
            </label>
            <div class='text-sm text-slate-500'>Es wird immer nur ein Hersteller gleichzeitig angezeigt.</div>
          </div>
        </div>
        <div class='border rounded p-3 bg-white'>
          <div class='flex items-center justify-between mb-2'>
            <h4 class='font-semibold text-base'>${selectedManufacturer}</h4>
            <span class='text-xs text-slate-500'>${selectedManufacturerTools.length} Werkzeug(e)</span>
          </div>
          <table class='w-full text-sm'>
            <thead class='bg-slate-100'>
              <tr>
                <th class='p-2 text-left'>T</th>
                <th class='p-2 text-left'>Bezeichnung</th>
                <th class='p-2 text-left'>Größe</th>
                <th class='p-2 text-left'>Artikelnummer</th>
                <th class='p-2 text-left'>Bestand</th>
                <th class='p-2 text-left'>Menge</th>
              </tr>
            </thead>
            <tbody>${selectedManufacturerRows || '<tr><td class="p-2" colspan="6">Keine bestellrelevanten Werkzeuge für diesen Hersteller.</td></tr>'}</tbody>
          </table>
        </div>`
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
      <div class='flex items-center justify-between gap-3 flex-wrap'>
        <div>
          <h3 class='text-lg font-bold mb-1'>1) Neues Werkzeug anlegen</h3>
          <p class='text-sm text-slate-500'>Die Dateneingabe öffnet sich in einem separaten Eingabefenster.</p>
        </div>
        <div class='flex gap-2 flex-wrap'>
          <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='openCreateToolModal()'>Neues Werkzeug erfassen</button>
          <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='addToolLabel()'>Bezeichnung hinzufügen</button>
          <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='addToolManufacturer()'>Hersteller hinzufügen</button>
        </div>
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
            <thead class='bg-slate-100'><tr><th class='p-2 text-left'>T</th><th class='p-2 text-left'>Bezeichnung</th><th class='p-2 text-left'>Größe</th><th class='p-2 text-left'>Artikelnummer</th><th class='p-2'></th></tr></thead>
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

function toggleInsertToolFields() {
  toggleInsertToolFieldsById("toolInsertTool", "toolInsertEdges");
}

function toggleInsertToolFieldsById(checkboxId, edgesId) {
  const checkbox = document.getElementById(checkboxId);
  const edges = document.getElementById(edgesId);
  if (!checkbox || !edges) return;
  edges.disabled = !checkbox.checked;
  if (!checkbox.checked) edges.value = "";
}

function renderTodo() {
  const nowIso = new Date().toISOString();
  const isAdmin = currentUser.role === "admin";
  const relevantTasks = state.tasks.filter(
    (t) => isAdmin || t.assignee === currentUser.name,
  );
  const openTasks = relevantTasks.filter((t) => t.status !== "done");
  const doneTasks = relevantTasks.filter((t) => t.status === "done");

  const openRows = openTasks
    .map((t) => {
      const overdue = t.deadline && nowIso > `${t.deadline}T23:59:59`;
      return `<tr class='border-b ${overdue ? "bg-rose-50" : ""}'>
      <td class='p-2'>${t.title}</td><td class='p-2'>${t.assignee}</td><td class='p-2'>${t.deadline ? formatDateDisplay(t.deadline) : "-"}</td>
      <td class='p-2'>${t.doneAt ? "Erledigt" : "Offen"}</td>
      <td class='p-2'>
        ${
          isAdmin
            ? `<button class='px-2 py-1 rounded bg-slate-900 text-white mr-1' onclick="deleteTask('${t.id}')">Löschen</button>
          <button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="reassignTaskPrompt('${t.id}')">Neu zuweisen</button>`
            : `<button class='px-2 py-1 rounded bg-emerald-600 text-white' onclick="completeTask('${t.id}')">Erledigt</button>`
        }
      </td>
    </tr>`;
    })
    .join("");

  const doneRows = doneTasks
    .map(
      (t) => `<tr class='border-b bg-emerald-50'>
      <td class='p-2'>✅ ${t.title}</td><td class='p-2'>${t.assignee}</td><td class='p-2'>${t.deadline ? formatDateDisplay(t.deadline) : "-"}</td><td class='p-2'>${t.doneAt || "-"}</td>
      <td class='p-2'>${isAdmin ? `<button class='px-2 py-1 rounded bg-slate-900 text-white' onclick="deleteTask('${t.id}')">Löschen</button>` : "-"}</td>
    </tr>`,
    )
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-3'>To-Do</h2>
    ${isAdmin ? renderTaskCreateBox() : ""}
    <h3 class='font-semibold mt-3 mb-2'>Offene Aufgaben</h3>
    <div class='overflow-auto max-h-[45vh] border rounded-lg'>
      <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Aufgabe</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Frist</th><th class='p-2 text-left'>Status</th><th class='p-2'></th></tr></thead><tbody>${openRows || '<tr><td class="p-2" colspan="5">Keine offenen Aufgaben.</td></tr>'}</tbody></table>
    </div>
    <h3 class='font-semibold mt-4 mb-2'>Erledigte Aufgaben</h3>
    <div class='overflow-auto max-h-[25vh] border rounded-lg'>
      <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Aufgabe</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Frist</th><th class='p-2 text-left'>Erledigt am</th><th class='p-2'></th></tr></thead><tbody>${doneRows || '<tr><td class="p-2" colspan="5">Noch keine erledigten Aufgaben.</td></tr>'}</tbody></table>
    </div>
  </div>`;
}

function renderTaskCreateBox() {
  const options = activeUsers()
    .map((u) => `<option value='${u.name}'>${u.name}</option>`)
    .join("");
  return `<div class='border rounded-lg p-3 bg-slate-50 mb-3'>
    <h3 class='font-semibold mb-2'>Neue Aufgabe erstellen</h3>
    <div class='grid md:grid-cols-4 gap-2'>
      <input id='taskTitle' class='border rounded p-2' placeholder='Aufgabe'/>
      <select id='taskAssignee' class='border rounded p-2'>${options}</select>
      <input id='taskDeadline' type='date' class='border rounded p-2'/>
      <button class='px-2 py-1 rounded bg-slate-900 text-white' onclick='createTask()'>Speichern</button>
    </div>
  </div>`;
}

function createTask() {
  const title = document.getElementById("taskTitle")?.value?.trim();
  const assignee = document.getElementById("taskAssignee")?.value;
  const deadline = document.getElementById("taskDeadline")?.value;
  if (!title || !assignee)
    return alert("Bitte Aufgabe und Mitarbeiter angeben.");
  state.tasks.push({
    id: `task-${Date.now()}`,
    title,
    assignee,
    deadline,
    status: "open",
    createdAt: new Date().toISOString(),
    doneAt: null,
  });
  persist();
  render();
}

function completeTask(taskId) {
  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) return;
  task.status = "done";
  task.doneAt = new Date().toISOString().slice(0, 16).replace("T", " ");
  persist();
  render();
}

function deleteTask(taskId) {
  state.tasks = state.tasks.filter((t) => t.id !== taskId);
  persist();
  render();
}

function reassignTaskPrompt(taskId) {
  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) return;
  const name = prompt("Neuer Mitarbeiter:", task.assignee);
  if (!name) return;
  task.assignee = name;
  task.status = "open";
  task.doneAt = null;
  persist();
  render();
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

function renderConflicts() {
  const rows = Object.entries(state.conflicts)
    .map(
      ([id, c]) => `<tr class='border-b'>
    <td class='p-2'>${formatDateWithWeekday(c.date)}</td><td class='p-2'>${c.user}</td><td class='p-2'>${c.text}</td>
    <td class='p-2'>${c.resolved ? "Ja" : "Nein"}</td>
    <td class='p-2'><button class='px-2 py-1 rounded bg-slate-900 text-white' onclick="setConflictResolved('${id}', ${c.resolved ? "false" : "true"})">${c.resolved ? "Auf Nein" : "Als gelöst markieren"}</button></td>
  </tr>`,
    )
    .join("");
  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-3'>Konflikte</h2>
    <div class='overflow-auto max-h-[70vh] border rounded-lg'>
      <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Konflikt</th><th class='p-2 text-left'>Gelöst</th><th class='p-2'></th></tr></thead>
      <tbody>${rows || '<tr><td class="p-2" colspan="5">Keine Konflikte.</td></tr>'}</tbody></table>
    </div>
  </div>`;
}

function setConflictResolved(id, resolved) {
  if (!state.conflicts[id]) return;
  state.conflicts[id].resolved = resolved;
  persist();
  render();
}

function getPeriodRange(period) {
  const now = new Date();
  const start = new Date(now);
  const end = new Date(now);
  if (period === "week") {
    start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
    end.setDate(start.getDate() + 6);
  } else if (period === "month") {
    start.setDate(1);
    end.setMonth(start.getMonth() + 1, 0);
  } else {
    start.setMonth(0, 1);
    end.setMonth(11, 31);
  }
  return { from: isoDate(start), to: isoDate(end), start, end };
}

function plannedMannedHoursForShift(shift) {
  if (shift.id.includes("-mf-")) return 6;
  if (shift.id.includes("-sa-0")) return 6;
  if (shift.id.includes("-sa-1")) return 6;
  if (shift.id.includes("-su-0")) return 6;
  if (shift.id.includes("-su-1")) return 6;
  return shiftHours(shift.start, shift.end);
}

function parseDurationHours(text) {
  if (!text || !text.includes(":")) return 0;
  const [h, m] = text.split(":").map(Number);
  if (Number.isNaN(h) || Number.isNaN(m)) return 0;
  return h + m / 60;
}

function shiftDateRange(shift) {
  const start = new Date(`${shift.date}T${shift.start}:00`);
  const end = new Date(`${shift.date}T${shift.end}:00`);
  if (end <= start) end.setDate(end.getDate() + 1);
  return { start, end };
}

function getShiftById(shiftId) {
  return generateThreeMonths().find((s) => s.id === shiftId) || null;
}

function isCurrentOrFutureShift(shift) {
  const now = new Date();
  const { start, end } = shiftDateRange(shift);
  return start >= now || (now >= start && now <= end);
}

function maxUnmannedHoursForShift(shiftId) {
  const all = generateThreeMonths()
    .filter((s) => s.assigned)
    .sort((a, b) => {
      const ra = shiftDateRange(a).start;
      const rb = shiftDateRange(b).start;
      return ra - rb;
    });
  const current = all.find((s) => s.id === shiftId);
  if (!current) return 0;
  const currentEnd = shiftDateRange(current).end;
  const next = all.find((s) => shiftDateRange(s).start > currentEnd);
  if (!next) return 0;
  const nextStart = shiftDateRange(next).start;
  const diffHours = Math.max(
    0,
    (nextStart.getTime() - currentEnd.getTime()) / (1000 * 60 * 60),
  );
  return diffHours;
}

function overlapHours(rangeStart, rangeEnd, shiftStart, shiftEnd) {
  const start = Math.max(rangeStart.getTime(), shiftStart.getTime());
  const end = Math.min(rangeEnd.getTime(), shiftEnd.getTime());
  if (end <= start) return 0;
  return (end - start) / (1000 * 60 * 60);
}

function computeStats(period = "week") {
  const { from, to, start, end } = getPeriodRange(period);
  const target = 154;
  const now = new Date();
  const shifts = generateThreeMonths().filter(
    (s) => s.date >= from && s.date <= to && s.assigned,
  );

  const plannedTotal = shifts.reduce(
    (acc, s) => acc + shiftHours(s.start, s.end),
    0,
  );
  const plannedManned = shifts.reduce(
    (acc, s) => acc + plannedMannedHoursForShift(s),
    0,
  );
  const plannedUnmanned = Math.max(0, plannedTotal - plannedManned);

  const elapsedPlanned = shifts.reduce((acc, s) => {
    const { start: sStart, end: sEnd } = shiftDateRange(s);
    return acc + overlapHours(start, now < end ? now : end, sStart, sEnd);
  }, 0);

  const recordedUnmanned = Object.entries(state.unmanned)
    .filter(([shiftId]) => {
      const date = shiftId.slice(0, 10);
      return date >= from && date <= to;
    })
    .reduce((acc, [shiftId, timeText]) => {
      const entered = parseDurationHours(timeText);
      const allowed = maxUnmannedHoursForShift(shiftId);
      return acc + Math.min(entered, allowed);
    }, 0);

  const downtime =
    Object.entries(state.machineDowntime)
      .filter(([k]) => {
        const date = k.split(":")[0];
        return date >= from && date <= to && date <= isoDate(now);
      })
      .reduce((acc, [, v]) => acc + (v.minutes || 0), 0) / 60;

  const istHours = Math.max(0, elapsedPlanned - downtime);
  const deviationPct = elapsedPlanned
    ? (Math.abs(elapsedPlanned - istHours) / elapsedPlanned) * 100
    : 0;
  const downtimePct = elapsedPlanned ? (downtime / elapsedPlanned) * 100 : 0;
  const targetPct = target ? (istHours / target) * 100 : 0;

  return {
    target,
    period,
    plannedTotal: round1(plannedTotal),
    elapsedPlanned: round1(elapsedPlanned),
    plannedManned: round1(plannedManned),
    plannedUnmanned: round1(plannedUnmanned),
    recordedUnmanned: round1(recordedUnmanned),
    downtime: round1(downtime),
    istHours: round1(istHours),
    deviationPct: round1(deviationPct),
    downtimePct: round1(downtimePct),
    targetPct: round1(targetPct),
  };
}

function round1(value) {
  return Math.round(value * 10) / 10;
}

function shiftHours(start, end) {
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let diff = eh * 60 + em - (sh * 60 + sm);
  if (diff <= 0) diff += 24 * 60;
  return diff / 60;
}

function renderStats() {
  const stats = computeStats(statsViewPeriod);
  const color =
    stats.deviationPct <= 8
      ? "text-emerald-600"
      : stats.deviationPct <= 12
        ? "text-orange-500"
        : "text-rose-600";
  const barPlan = Math.min(
    100,
    stats.target ? (stats.plannedTotal / stats.target) * 100 : 0,
  );
  const barIst = Math.min(
    100,
    stats.target ? (stats.istHours / stats.target) * 100 : 0,
  );
  const barStillstand = Math.min(
    100,
    stats.target ? (stats.downtime / stats.target) * 100 : 0,
  );

  return `<div class='bg-white rounded-xl shadow p-4'>
    <div class='flex items-center justify-between mb-3'>
      <h2 class='text-lg font-semibold'>Laufzeitstatistik (${stats.period === "week" ? "Woche" : stats.period === "month" ? "Monat" : "Jahr"})</h2>
      <div class='flex gap-2'>
        <button class='px-2 py-1 rounded ${statsViewPeriod === "week" ? "bg-slate-900 text-white" : "bg-slate-100"}' onclick="setStatsView('week')">Woche</button>
        <button class='px-2 py-1 rounded ${statsViewPeriod === "month" ? "bg-slate-900 text-white" : "bg-slate-100"}' onclick="setStatsView('month')">Monat</button>
        <button class='px-2 py-1 rounded ${statsViewPeriod === "year" ? "bg-slate-900 text-white" : "bg-slate-100"}' onclick="setStatsView('year')">Jahr</button>
      </div>
    </div>
    <div class='grid md:grid-cols-4 gap-3 text-center mb-4'>
      <div class='p-3 rounded bg-slate-100'><div class='text-sm'>Marker</div><div class='text-2xl font-bold'>${stats.target} h</div></div>
      <div class='p-3 rounded bg-blue-100'><div class='text-sm'>Geplante Stunden</div><div class='text-2xl font-bold'>${stats.plannedTotal} h</div><div class='text-xs text-slate-500'>${stats.targetPct}% vom Marker</div></div>
      <div class='p-3 rounded bg-emerald-100'><div class='text-sm'>Ist-Stunden (bis jetzt)</div><div class='text-2xl font-bold'>${stats.istHours} h</div><div class='text-xs text-slate-500'>Stillstand ${stats.downtimePct}%</div></div>
      <div class='p-3 rounded bg-white border'><div class='text-sm'>Abweichung Plan/Ist</div><div class='text-2xl font-bold ${color}'>${stats.deviationPct}%</div></div>
    </div>
    <div class='grid md:grid-cols-2 gap-4 mb-4'>
      <div class='p-3 rounded border bg-slate-50'>
        <h3 class='font-semibold mb-2'>Bemannt / Mannlos</h3>
        <div class='text-sm'>Bemannt (Plan): <b>${stats.plannedManned} h</b></div>
        <div class='text-sm'>Mannlos (Plan): <b>${stats.plannedUnmanned} h</b></div>
        <div class='text-sm'>Mannlos (eingetragen): <b>${stats.recordedUnmanned} h</b></div>
      </div>
      <div class='p-3 rounded border bg-slate-50'>
        <h3 class='font-semibold mb-2'>Grafik (h / Marker 154h)</h3>
        <div class='mb-2 text-xs'>Plan: ${stats.plannedTotal} h</div>
        <div class='w-full bg-slate-200 rounded h-4 mb-2'><div class='bg-blue-500 h-4 rounded' style='width:${barPlan}%'></div></div>
        <div class='mb-2 text-xs'>Ist: ${stats.istHours} h</div>
        <div class='w-full bg-slate-200 rounded h-4 mb-2'><div class='bg-emerald-500 h-4 rounded' style='width:${barIst}%'></div></div>
        <div class='mb-2 text-xs'>Stillstand: ${stats.downtime} h</div>
        <div class='w-full bg-slate-200 rounded h-4'><div class='bg-rose-500 h-4 rounded' style='width:${barStillstand}%'></div></div>
      </div>
    </div>
  </div>`;
}

function renderAvailability() {
  const shifts = generateThreeMonths().filter((s) => s.options.length > 1);
  const rows = shifts
    .slice(0, 80)
    .map((s) => {
      const key = `${s.id}:${currentUser.name}`;
      const val = state.availability[key] || "";
      return `<tr class='border-b'><td class='p-2'>${formatDateWithWeekday(s.date)}</td><td class='p-2'>${s.label}</td><td class='p-2'>${s.options.join(" / ")}</td>
      <td class='p-2'>
        <select class='border rounded p-1' onchange="setAvailability('${s.id}', this.value)">
          <option value="" ${val === "" ? "selected" : ""}>-</option>
          <option value="yes" ${val === "yes" ? "selected" : ""}>Kann</option>
          <option value="no" ${val === "no" ? "selected" : ""}>Kann nicht</option>
        </select>
      </td></tr>`;
    })
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-3'>Unklare Schichten (nur für Admin sichtbar in Planung)</h2>
    <p class='text-sm text-slate-500 mb-3'>Du kannst hier vorab eintragen, ob du bei Entweder/Oder-Schichten könntest.</p>
    <div class='overflow-auto max-h-[70vh]'><table class='w-full text-sm'><thead class='bg-slate-100'><tr><th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Option</th><th class='p-2 text-left'>Dein Status</th></tr></thead><tbody>${rows}</tbody></table></div>
  </div>`;
}

function setAvailability(shiftId, value) {
  state.availability[`${shiftId}:${currentUser.name}`] = value;
  persist();
}

function renderClosure() {
  const myShifts = generateThreeMonths()
    .filter((s) => s.assigned === currentUser.name)
    .slice(0, 20);
  const rows = myShifts
    .map((s) => {
      const done = state.checklists[s.id] ? "✅" : "⏳";
      return `<tr class='border-b'><td class='p-2'>${formatDateWithWeekday(s.date)}</td><td class='p-2'>${s.label}</td><td class='p-2'>${s.start}–${s.end}</td><td class='p-2'>${done}</td>
      <td class='p-2'><button class='px-2 py-1 rounded bg-indigo-700 text-white' onclick="openChecklist('${s.id}')">Checklist starten</button></td></tr>`;
    })
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4 space-y-3'>
    <h2 class='text-lg font-semibold'>Schichtabschluss</h2>
    <p class='text-sm text-slate-500'>Im Livebetrieb soll dieser Dialog 15 Minuten vor Schichtende erscheinen. Im MVP startest du ihn manuell.</p>
    <table class='w-full text-sm'><thead class='bg-slate-100'><tr><th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Zeit</th><th class='p-2'>Status</th><th></th></tr></thead><tbody>${rows}</tbody></table>
  </div>`;
}

function openChecklist(shiftId) {
  const hourOptions = Array.from(
    { length: 24 },
    (_, i) => `<option>${String(i).padStart(2, "0")}</option>`,
  ).join("");
  const minOptions = Array.from(
    { length: 60 },
    (_, i) => `<option>${String(i).padStart(2, "0")}</option>`,
  ).join("");

  const wrapper = document.createElement("div");
  wrapper.className =
    "fixed inset-0 bg-black/40 flex items-center justify-center p-4 z-50";
  wrapper.innerHTML = `<div class='bg-white rounded-xl p-4 max-w-lg w-full space-y-3'>
    <h3 class='text-lg font-semibold'>Pflicht-Checkliste</h3>
    ${checkbox("c1", "Arbeitsplatz ordentlich")}
    ${checkbox("c2", "Informationen für die nächste Schicht vorhanden")}
    ${checkbox("c3", "Material versorgt und Lagerorte eingetragen")}
    ${checkbox("c4", "Alle Aufträge gebucht")}
    <div class='border-t pt-3'>
      <p class='text-sm font-medium mb-2'>Zeit für Mannlosbetrieb erfassen</p>
      <div class='flex gap-2 items-center'>
        <select id='uh'>${hourOptions}</select> : <select id='um'>${minOptions}</select>
      </div>
    </div>
    <div class='flex justify-end gap-2'>
      <button class='px-3 py-2 rounded bg-slate-200' id='cancelBtn'>Abbrechen</button>
      <button class='px-3 py-2 rounded bg-slate-900 text-white' id='saveBtn'>Speichern</button>
    </div>
  </div>`;
  document.body.appendChild(wrapper);

  wrapper.querySelector("#cancelBtn").onclick = () => wrapper.remove();
  wrapper.querySelector("#saveBtn").onclick = () => {
    const checks = ["c1", "c2", "c3", "c4"].map(
      (id) => wrapper.querySelector(`#${id}`).checked,
    );
    if (checks.some((v) => !v)) {
      alert("Bitte alle 4 Punkte bestätigen.");
      return;
    }
    state.checklists[shiftId] = {
      by: currentUser.name,
      at: new Date().toISOString(),
      checks,
    };
    state.shiftEndChecks[shiftId] = {
      by: currentUser.name,
      at: new Date().toISOString(),
      checks,
    };
    const h = wrapper.querySelector("#uh").value;
    const m = wrapper.querySelector("#um").value;
    state.unmanned[shiftId] = `${h}:${m}`;
    persist();
    wrapper.remove();
    render();
  };
}

function getCurrentShiftForUser(name) {
  const now = new Date();
  const today = isoDate(now);
  const shifts = generateThreeMonths().filter(
    (s) => s.assigned === name && s.date === today,
  );
  return (
    shifts.find((s) => {
      const [sh, sm] = s.start.split(":").map(Number);
      const [eh, em] = s.end.split(":").map(Number);
      const start = new Date(now);
      start.setHours(sh, sm, 0, 0);
      const end = new Date(now);
      end.setHours(eh, em, 0, 0);
      if (end <= start) end.setDate(end.getDate() + 1);
      return now >= start && now <= end;
    }) || null
  );
}

function maybeShowMachinePrompt() {
  if (!currentUser || currentUser.role !== "employee") return;
  const today = isoDate(new Date());
  const key = `${today}:${currentUser.name}`;
  if (state.machinePromptSeen[key]) {
    renderDowntimeWidget();
    return;
  }
  state.machinePromptSeen[key] = true;
  const running = confirm("Läuft die Maschine zu Schichtbeginn?");
  if (!running) {
    state.machineDowntime[key] = {
      startAt: new Date().toISOString(),
      minutes: 0,
      running: true,
    };
  }
  persist();
  renderDowntimeWidget();
}

function renderDowntimeWidget() {
  const today = isoDate(new Date());
  const key = `${today}:${currentUser.name}`;
  const downtime = state.machineDowntime[key];
  let box = document.getElementById("downtimeWidget");
  if (!box) {
    box = document.createElement("div");
    box.id = "downtimeWidget";
    box.className = "fixed top-4 right-4 z-50";
    document.body.appendChild(box);
  }
  if (!downtime || !downtime.running) {
    box.innerHTML = "";
    return;
  }
  box.innerHTML = `<button onclick="stopDowntimeTimer()" class="px-4 py-3 rounded-full bg-emerald-600 text-white shadow-lg flex items-center gap-2">
    <span class="w-4 h-4 rounded-full bg-green-300 inline-block"></span>
    Stillstand läuft – Stoppen
  </button>`;
}

function stopDowntimeTimer() {
  const today = isoDate(new Date());
  const key = `${today}:${currentUser.name}`;
  const downtime = state.machineDowntime[key];
  if (!downtime?.running) return;
  const start = new Date(downtime.startAt);
  const minutes = Math.max(
    1,
    Math.round((Date.now() - start.getTime()) / 60000),
  );
  state.machineDowntime[key] = {
    ...downtime,
    running: false,
    minutes: (downtime.minutes || 0) + minutes,
  };
  persist();
  render();
}

function maybeShowShiftEndChecklist() {
  if (!currentUser || currentUser.role !== "employee") return;
  const shift = getCurrentShiftForUser(currentUser.name);
  if (!shift) return;
  const [eh, em] = shift.end.split(":").map(Number);
  const end = new Date();
  end.setHours(eh, em, 0, 0);
  const trigger = new Date(end.getTime() - 15 * 60000);
  const now = new Date();
  if (now >= trigger && !state.shiftEndChecks[shift.id])
    openChecklist(shift.id);
}

function maybeShowShiftStartChecklist() {
  if (!currentUser || currentUser.role !== "employee") return;
  const shift = getCurrentShiftForUser(currentUser.name);
  if (!shift) return;
  const key = `${shift.id}:${currentUser.name}`;
  if (state.shiftStartChecks[key]) return;
  const endData = state.shiftEndChecks[shift.id];
  const answers = [
    "Arbeitsplatz ordentlich",
    "Informationen für die nächste Schicht vorhanden",
    "Material versorgt und Lagerorte eingetragen",
    "Alle Aufträge gebucht",
  ].map((txt) => ({ txt, ok: confirm(`Schichtstart-Check: ${txt}?`) }));
  state.shiftStartChecks[key] = answers;
  answers.forEach((a, idx) => {
    if (endData && endData.checks && endData.checks[idx] !== a.ok) {
      state.conflicts[`${key}:${idx}`] = {
        date: shift.date,
        user: currentUser.name,
        text: `${a.txt} weicht von Schichtende-Angabe ab`,
        resolved: false,
      };
    }
  });
  persist();
}

function checkbox(id, label) {
  return `<label class='flex items-center gap-2'><input id='${id}' type='checkbox' class='w-4 h-4'/> ${label}</label>`;
}

function markAbsent(shiftId, date, userName) {
  state.absences[`${date}:${userName}`] = true;
  persist();
  render();
}

function assignOptionalShift(shiftId) {
  const select = document.getElementById(`opt-${shiftId}`);
  if (!select) return;
  const name = select.value;
  if (!canAssignUserToShift(name, shiftId)) return;
  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = name;
  persist();
  render();
}

function approveSaturdayRequest(shiftId, user) {
  if (!canAssignUserToShift(user, shiftId)) return;
  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = user;
  delete state.saturdayEveningRequests[`${shiftId}:${user}`];
  persist();
  render();
}

render();
bootSupabase();
window.loginWithSupabase = loginWithSupabase;
window.logoutSupabase = logoutSupabase;
window.fillLogin = fillLogin;
window.loginAs = loginAs;
window.logout = logout;
window.setTab = setTab;
window.setStatsView = setStatsView;
window.setPlanningSubTab = setPlanningSubTab;
window.markAbsent = markAbsent;
window.assignShift = assignShift;
window.cancelShift = cancelShift;
window.assignOptionalShift = assignOptionalShift;
window.approveSaturdayRequest = approveSaturdayRequest;
window.setAvailability = setAvailability;
window.openChecklist = openChecklist;
window.addVacation = addVacation;
window.addSickLeave = addSickLeave;
window.deleteVacation = deleteVacation;
window.deleteSickLeave = deleteSickLeave;
window.planAbsenceReplacement = planAbsenceReplacement;
window.clearAbsenceReplacementPlan = clearAbsenceReplacementPlan;
window.addSwap = addSwap;
window.deleteSwap = deleteSwap;
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
window.setOrderListManufacturer = setOrderListManufacturer;
window.applyOptimalQtySuggestion = applyOptimalQtySuggestion;
window.rejectOptimalQtySuggestion = rejectOptimalQtySuggestion;
window.toggleInsertToolFields = toggleInsertToolFields;
window.toggleInsertToolFieldsById = toggleInsertToolFieldsById;
window.markToolOrdered = markToolOrdered;
window.restockTool = restockTool;
window.createTool = createTool;
window.openCreateToolModal = openCreateToolModal;
window.addToolLabel = addToolLabel;
window.addToolManufacturer = addToolManufacturer;
