const { createClient } = window.supabase;

const supabase = createClient(
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
    const { error } = await supabase.from("employees").select("id").limit(1);

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
  const { data, error } = await supabase
    .from("employees")
    .select("*")
    .order("display_name", { ascending: true });

  if (error) {
    console.error("Fehler beim Laden von employees:", error);
    return [];
  }

  return data || [];
}

async function loadToolsFromSupabase() {
  const { data, error } = await supabase
    .from("tools")
    .select("*")
    .order("t_number", { ascending: true });

  if (error) {
    console.error("Fehler beim Laden von tools:", error);
    return [];
  }

  return data || [];
}

async function getCurrentEmployeeRecord() {
  const {
    data: { user },
    error: userError,
  } = await supabase.auth.getUser();

  if (userError) {
    console.error("Fehler bei auth.getUser():", userError);
    return null;
  }

  if (!user) return null;

  const { data, error } = await supabase
    .from("employees")
    .select("*")
    .eq("auth_user_id", user.id)
    .eq("is_active", true)
    .maybeSingle();

  if (error) {
    console.error("Fehler beim Laden des employees-Datensatzes:", error);
    return null;
  }

  return data || null;
}

async function loginWithSupabase() {
  const email = document.getElementById("loginEmail")?.value?.trim() || "";
  const password = document.getElementById("loginPassword")?.value || "";

  if (!email || !password) {
    setLoginStatus("Bitte E-Mail und Passwort eingeben.", true);
    return;
  }

  setLoginStatus("Anmeldung läuft ...");

  const { error } = await supabase.auth.signInWithPassword({
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
  await supabase.auth.signOut();
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
  } = await supabase.auth.getUser();

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
  } = await supabase.auth.getSession();

  if (session?.user) {
    await syncSupabaseSessionToApp();
    return;
  }

  const employees = await loadEmployeesFromSupabase();
  const tools = await loadToolsFromSupabase();

  console.log("Employees aus Supabase:", employees);
  console.log("Tools aus Supabase:", tools);
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
  document.getElementById("loginBox").classList.add("hidden");
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
