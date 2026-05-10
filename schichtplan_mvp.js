const { createClient } = window.supabase;

const supabaseClient = createClient(
  window.SUPABASE_URL,
  window.SUPABASE_PUBLISHABLE_KEY,
);

let supabaseReady = false;
let currentSupabaseUser = null;
let currentEmployeeRecord = null;
let toolMaterials = [];

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
const EMPLOYEE_COLOR_OPTIONS = [
  { key: "green", label: "Grün" },
  { key: "blue", label: "Blau" },
  { key: "yellow", label: "Gelb" },
  { key: "purple", label: "Lila" },
  { key: "orange", label: "Orange" },
  { key: "teal", label: "Türkis" },
  { key: "cyan", label: "Cyan" },
  { key: "pink", label: "Pink" },
  { key: "rose", label: "Rose" },
  { key: "lime", label: "Lime" },
  { key: "gray", label: "Grau" },
];
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

const APP_VERSION = "0.4.17";
const STORAGE_KEY = "schichtplan_mvp_v_0_2";
const state = loadState();
let currentUser = null;
let currentTab = "schichtplan";
let statsViewPeriod = "week";
let qrScannerStream = null;
let qrScannerTimer = null;
let qrScannerMode = null;

const HELP_TEXTS = {
  planningPersonal: {
    title: "Planung / Personal",
    body: [
      "Hier pflegst du Mitarbeiter, Rollen, Farben und die Slots A-F.",
      "Die Slot-Zuordnung steuert, wer in der automatischen Rotation eingesetzt wird.",
      "Nach dem Speichern werden die Mitarbeiterdaten in Supabase aktualisiert und der Plan nutzt die neue Zuordnung.",
    ],
  },
  abstinenz: {
    title: "Abstinenz",
    body: [
      "Hier trägst du Urlaub, Krankheit und manuelle Abwesenheiten ein.",
      "Abwesenheiten öffnen Schichten, wenn der ursprünglich geplante Mitarbeiter fehlt.",
      "Nach dem Speichern werden die Planungsdaten in Supabase gehalten und beim Neuladen wieder geladen.",
    ],
  },
  ersatzplanung: {
    title: "Ersatzplanung",
    body: [
      "Hier planst du Ersatz oder Ausfall für Schichten während Urlaub oder Krankheit.",
      "Es werden nur Schichten des abwesenden Mitarbeiters im betroffenen Zeitraum angezeigt.",
      "Ersatz schreibt eine manuelle Einteilung, Ausfall sperrt die Schicht. Beides wirkt direkt auf den Plan.",
    ],
  },
  wochenende: {
    title: "Wochenendeinsätze",
    body: [
      "Springer melden, ob sie für Wochenend-Schichten können oder nicht können.",
      "Der Admin teilt Springer nur ein, wenn der ursprüngliche A/B/C-Mitarbeiter abwesend ist.",
      "Die Einteilung wird gespeichert und bleibt nach einem Refresh erhalten.",
    ],
  },
  schichttausch: {
    title: "Schichttausch",
    body: [
      "Hier können A/B/C-Mitarbeiter für einen Zeitraum ihre Rotation tauschen.",
      "Der Tausch wirkt ab dem Startdatum bis zum Enddatum oder bis er zurückgesetzt wird.",
      "Gespeicherte Tausche verändern die automatisch berechnete Schichtzuordnung.",
    ],
  },
  werkzeuge: {
    title: "Werkzeuge",
    body: [
      "Hier verwaltest du Werkzeugbestand, Stammdaten, Bestellungen und Werkzeugwechsel.",
      "Bestand und Werkzeugdaten werden in Supabase gespeichert.",
      "Wechsel und Bestellhistorie sind aktuell noch lokal und dienen als MVP-Journal.",
    ],
  },
  statistik: {
    title: "Statistik",
    body: [
      "Die Statistik vergleicht geplante Stunden, Ist-Stunden und Stillstand.",
      "Du kannst zwischen Woche, Monat und Jahr wechseln.",
      "Die Werte helfen, Abweichungen im Plan und Maschinenlaufzeit sichtbar zu machen.",
    ],
  },
};

function makeWeek(early, late, night, satPrimary, satSecondary) {
  return {
    mondayToFriday: [
      { label: "Früh", start: "05:00", end: "11:00", options: [early] },
      { label: "Spät", start: "13:00", end: "19:00", options: [late] },
      { label: "Nacht", start: "21:00", end: "03:00", options: [night] },
    ],
    saturday: [
      {
        label: "Samstag Morgen",
        start: "05:00",
        end: "11:00",
        options: [satPrimary],
      },
      {
        label: "Samstag Abend",
        start: "16:00",
        end: "22:00",
        options: [satSecondary, "D", "E", "F"],
      },
    ],
    sunday: [
      {
        label: "Sonntag Morgen",
        start: "06:00",
        end: "12:00",
        options: [late],
      },
      {
        label: "Sonntag Nacht",
        start: "18:00",
        end: "24:00",
        options: [night],
      },
    ],
  };
}

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

function normalizeEmployeeFromDb(row) {
  return {
    id: row.id,
    authUserId: row.auth_user_id || null,
    name: row.display_name,
    display_name: row.display_name,
    role: row.role || "employee",
    type: row.employee_type || "springer",
    employee_type: row.employee_type || "springer",
    slot: row.slot_code || "",
    slot_code: row.slot_code || "",
    isActive: row.is_active !== false,
    is_active: row.is_active !== false,
    color_key: row.color_key || "",
    created_at: row.created_at || null,
    updated_at: row.updated_at || null,
  };
}

function applyEmployeesToState(rows) {
  const employees = (rows || []).map(normalizeEmployeeFromDb);
  const employeeMap = {};
  const nextSlotAssignments = { ...(state.slotAssignments || {}) };

  employees.forEach((employee) => {
    employeeMap[employee.id] = employee;
    if (
      employee.isActive &&
      employee.slot &&
      SLOT_CODES.includes(employee.slot)
    ) {
      nextSlotAssignments[employee.slot] = employee.name;
    }
  });

  SLOT_CODES.forEach((slot) => {
    const assignedName = nextSlotAssignments[slot];
    const assignedEmployee = employees.find((employee) => {
      return employee.name === assignedName && employee.isActive;
    });
    if (!assignedEmployee) {
      const defaultName = DEFAULT_SLOT_ASSIGNMENTS[slot] || "";
      const defaultEmployee = employees.find((employee) => {
        return employee.name === defaultName && employee.isActive;
      });
      nextSlotAssignments[slot] = defaultEmployee ? defaultName : "";
    }
  });

  state.employees = employeeMap;
  state.employeesList = employees;
  state.slotAssignments = nextSlotAssignments;
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
    materialId: row.material_id || null,
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
    insertRadius: row.insert_radius || "",
  };
}

function normalizeTaskFromDb(row) {
  return {
    id: row.id,
    title: row.title || "",
    description: row.description || "",
    assignee: row.assigned_to || "",
    assignedTo: row.assigned_to || "",
    dueDate: row.due_date || "",
    status: row.status || "open",
    createdAt: row.created_at || null,
    completedAt: row.completed_at || null,
  };
}

async function loadTasksFromSupabase() {
  const { data, error } = await supabaseClient
    .from("planner_tasks")
    .select("*")
    .order("created_at", { ascending: false });

  if (error) {
    console.error("Fehler beim Laden von planner_tasks:", error);
    return null;
  }

  console.log("planner_tasks raw data:", data);
  return (data || []).map(normalizeTaskFromDb);
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

async function getCurrentEmployeeRecord() {
  const {
    data: { user },
    error: userError,
  } = await supabaseClient.auth.getUser();

  if (userError) {
    console.error("Fehler bei auth.getUser():", userError);
    return null;
  }

  if (!user) return null;

  const { data, error } = await supabaseClient
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

async function loadToolMaterialsFromSupabase() {
  const { data, error } = await supabaseClient
    .from("tool_materials")
    .select("*")
    .eq("is_active", true)
    .order("sort_order", { ascending: true })
    .order("name", { ascending: true });

  if (error) {
    console.error("Fehler beim Laden von tool_materials:", error);
    return [];
  }

  return data || [];
}

function getToolMaterialNameById(materialId) {
  if (!materialId) return "-";
  const found = toolMaterials.find((m) => m.id === materialId);
  return found?.name || "-";
}

async function loadPlanningDataFromSupabase() {
  const [
    assignmentsRes,
    absencesRes,
    vacationsRes,
    sickLeavesRes,
    availabilityRes,
    swapsRes,
    cancellationsRes,
    saturdayRequestsRes,
    replacementsRes,
  ] = await Promise.all([
    supabaseClient
      .from("planner_assignments")
      .select("*")
      .order("shift_date", { ascending: true }),
    supabaseClient
      .from("planner_absences")
      .select("*")
      .order("absence_date", { ascending: true }),
    supabaseClient
      .from("planner_vacations")
      .select("*")
      .order("from_date", { ascending: true }),
    supabaseClient
      .from("planner_sick_leaves")
      .select("*")
      .order("from_date", { ascending: true }),
    supabaseClient
      .from("planner_availability")
      .select("*")
      .order("shift_date", { ascending: true }),
    supabaseClient
      .from("planner_swaps")
      .select("*")
      .order("start_date", { ascending: true }),
    supabaseClient
      .from("planner_shift_cancellations")
      .select("*")
      .order("shift_date", { ascending: true }),
    supabaseClient
      .from("planner_saturday_requests")
      .select("*")
      .order("shift_date", { ascending: true }),
    supabaseClient
      .from("planner_absence_replacements")
      .select("*")
      .order("shift_date", { ascending: true }),
  ]);

  const responses = [
    ["planner_assignments", assignmentsRes],
    ["planner_absences", absencesRes],
    ["planner_vacations", vacationsRes],
    ["planner_sick_leaves", sickLeavesRes],
    ["planner_availability", availabilityRes],
    ["planner_swaps", swapsRes],
    ["planner_shift_cancellations", cancellationsRes],
    ["planner_saturday_requests", saturdayRequestsRes],
    ["planner_absence_replacements", replacementsRes],
  ];

  const firstError = responses.find(([, res]) => res.error)?.[1]?.error;
  if (firstError) {
    console.error("Fehler beim Laden der Planungsdaten:", firstError);
    return null;
  }

  const tasks = await loadTasksFromSupabase();

  const assignments = {};
  (assignmentsRes.data || []).forEach((row) => {
    assignments[row.shift_id] = row.assigned_user;
  });

  const absences = {};
  (absencesRes.data || []).forEach((row) => {
    absences[row.absence_key] = true;
  });

  const vacations = (vacationsRes.data || []).map((row) => ({
    id: row.id,
    user: row.user_name,
    from: row.from_date,
    to: row.to_date,
  }));

  const sickLeaves = (sickLeavesRes.data || []).map((row) => ({
    id: row.id,
    user: row.user_name,
    from: row.from_date,
    to: row.to_date,
  }));

  const availability = {};
  (availabilityRes.data || []).forEach((row) => {
    availability[row.availability_key] = row.status;
  });

  const swaps = (swapsRes.data || []).map((row) => ({
    id: row.id,
    userA: row.user_a,
    userB: row.user_b,
    startDate: row.start_date,
    endDate: row.end_date || null,
  }));

  const shiftCancellations = {};
  (cancellationsRes.data || []).forEach((row) => {
    shiftCancellations[row.shift_id] = true;
  });

  const saturdayEveningRequests = {};
  (saturdayRequestsRes.data || []).forEach((row) => {
    saturdayEveningRequests[row.request_key] = true;
  });

  const { data: specialDaysData } = await supabaseClient
    .from("planner_special_days")
    .select("*");

  state.specialDays = {};
  (specialDaysData || []).forEach((d) => {
    state.specialDays[d.day_date] = d;
  });

  const absenceReplacements = {};
  (replacementsRes.data || []).forEach((row) => {
    absenceReplacements[row.shift_id] = {
      sourceType: row.source_type,
      sourceId: row.source_id,
      absentUser: row.absent_user,
      from: row.from_date,
      to: row.to_date,
      mode: row.mode,
      replacementUser: row.replacement_user,
      weekFrom: row.week_from,
      weekTo: row.week_to,
    };
  });

  return {
    assignments,
    absences,
    vacations,
    sickLeaves,
    availability,
    swaps,
    shiftCancellations,
    saturdayEveningRequests,
    absenceReplacements,
    tasks,
  };
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

  const employees = await loadEmployeesFromSupabase();
  const materials = await loadToolMaterialsFromSupabase();
  const tools = await loadToolsFromSupabase();
  const planning = await loadPlanningDataFromSupabase();

  applyEmployeesToState(employees);
  toolMaterials = materials;
  state.tools = tools;

  if (planning) {
    state.assignments = planning.assignments;
    state.absences = planning.absences;
    state.vacations = planning.vacations;
    state.sickLeaves = planning.sickLeaves;
    state.availability = planning.availability;
    state.swaps = planning.swaps;
    state.shiftCancellations = planning.shiftCancellations;
    state.saturdayEveningRequests = planning.saturdayEveningRequests;
    state.absenceReplacements = planning.absenceReplacements;
    if (Array.isArray(planning.tasks)) {
      state.tasks = planning.tasks;
      console.log("Tasks aus Supabase geladen:", state.tasks);
    } else {
      console.warn(
        "Tasks wurden nicht aktualisiert, alter state.tasks bleibt erhalten.",
      );
    }
  }

  persist();

  document.getElementById("loginBox")?.classList.add("hidden");
  setLoginStatus(
    `Angemeldet als ${currentEmployeeRecord.display_name} (${currentEmployeeRecord.role}).`,
  );

  console.log("Supabase-User:", currentSupabaseUser);
  console.log("Employees-Datensatz:", currentEmployeeRecord);
  console.log("Tool-Materials nach Login geladen:", materials);
  console.log("Tools nach Login geladen:", tools);
  console.log("Planungsdaten nach Login geladen:", planning);
  console.log("Tasks nach Session-Sync:", state.tasks);

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
  const materials = await loadToolMaterialsFromSupabase();
  const tools = await loadToolsFromSupabase();
  const planning = await loadPlanningDataFromSupabase();

  applyEmployeesToState(employees);
  toolMaterials = materials;
  state.tools = tools;

  if (planning) {
    state.assignments = planning.assignments;
    state.absences = planning.absences;
    state.vacations = planning.vacations;
    state.sickLeaves = planning.sickLeaves;
    state.availability = planning.availability;
    state.swaps = planning.swaps;
    state.shiftCancellations = planning.shiftCancellations;
    state.saturdayEveningRequests = planning.saturdayEveningRequests;
    state.absenceReplacements = planning.absenceReplacements;
  }

  persist();

  console.log("Employees aus Supabase:", employees);
  console.log("Tool-Materials aus Supabase:", materials);
  console.log("Tools aus Supabase:", tools);
  console.log("Planungsdaten aus Supabase:", planning);

  if (currentUser) render();
}

function allUsers() {
  if (Array.isArray(state.employeesList) && state.employeesList.length) {
    return state.employeesList.map((employee) => ({
      id: employee.id,
      name: employee.name || employee.display_name,
      display_name: employee.display_name || employee.name,
      slot: employee.slot || employee.slot_code || "",
      type: employee.type || employee.employee_type || "springer",
      role: employee.role || "employee",
      color_key: employee.color_key || "",
      isActive: employee.isActive !== false && employee.is_active !== false,
    }));
  }

  return [...USERS, ...(state.extraUsers || [])];
}

function activeUsers() {
  return allUsers().filter((u) => {
    const activeByDb = u.isActive !== false;
    const activeByLocalFallback = !state.inactiveUsers?.[u.name];
    return activeByDb && activeByLocalFallback;
  });
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
    employees: {},
    employeesList: [],
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
      imageStatus: "",
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
    replacementPlannerSelection: {},
    replacementPlannerChoice: {},
    ui: {
      pendingEmployeeEdits: {},
      pendingSlotAssignments: {},
    },
  };
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return base;
    const parsed = JSON.parse(raw);
    return {
      ...base,
      ...parsed,
      toolFilters: {
        ...base.toolFilters,
        ...(parsed.toolFilters || {}),
      },
      ui: {
        ...base.ui,
        ...(parsed.ui || {}),
      },
    };
  } catch {
    return base;
  }
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function helpButton(topic) {
  return `<button class="inline-flex items-center justify-center w-7 h-7 rounded-full bg-blue-600 text-white font-bold shadow-sm hover:bg-blue-700" onclick="openHelp('${topic}')" title="Hilfe">?</button>`;
}

function getHelpModalHost() {
  let host = document.getElementById("helpModalHost");
  if (!host) {
    host = document.createElement("div");
    host.id = "helpModalHost";
    document.body.appendChild(host);
  }
  return host;
}

function openHelp(topic) {
  const help = HELP_TEXTS[topic];
  if (!help) return;

  const body = help.body
    .map((paragraph) => `<p class="text-sm text-slate-700">${paragraph}</p>`)
    .join("");

  getHelpModalHost().innerHTML = `<div class="fixed inset-0 z-[60] bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-lg p-4">
      <div class="flex items-start justify-between gap-3 mb-3">
        <h3 class="text-lg font-bold">${help.title}</h3>
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeHelp()">Schließen</button>
      </div>
      <div class="space-y-2">${body}</div>
    </div>
  </div>`;
}

function closeHelp() {
  getHelpModalHost().innerHTML = "";
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

  const tabs =
    currentUser.role === "tool_scanner" ? ["toolscanner"] : ["schichtplan"];
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
  if (currentTab === "toolscanner") view.innerHTML = renderToolScanner();
  if (currentTab === "bestellstatistik") view.innerHTML = renderOrderStats();
  if (currentTab === "todo") view.innerHTML = renderTodo();
  if (currentTab === "konflikte") view.innerHTML = renderConflicts();
  if (currentTab === "statistik") view.innerHTML = renderStats();
  if (currentUser.role === "tool_scanner") return;
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
    toolscanner: "Werkzeug-Scanner",
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
    return (
      entry?.sourceType === sourceType &&
      String(entry?.sourceId) === String(sourceId)
    );
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

function getReplacementShiftIdsForSource(sourceType, sourceId) {
  return getReplacementEntriesForSource(sourceType, sourceId).map(
    ([shiftId]) => shiftId,
  );
}

function clearReplacementPlanForSource(sourceType, sourceId) {
  if (!state.absenceReplacements) return;
  Object.keys(state.absenceReplacements).forEach((shiftId) => {
    const entry = state.absenceReplacements[shiftId];
    if (
      entry?.sourceType === sourceType &&
      String(entry?.sourceId) === String(sourceId)
    ) {
      delete state.absenceReplacements[shiftId];
      delete state.shiftCancellations[shiftId];
      delete state.assignments[shiftId];
    }
  });
}

async function deleteReplacementPlanFromSupabase(
  sourceType,
  sourceId,
  shiftIds = null,
) {
  if (!supabaseReady) return null;

  const targetShiftIds =
    shiftIds || getReplacementShiftIdsForSource(sourceType, sourceId);

  const operations = [
    supabaseClient
      .from("planner_absence_replacements")
      .delete()
      .eq("source_type", sourceType)
      .eq("source_id", sourceId),
  ];

  if (targetShiftIds.length) {
    operations.push(
      supabaseClient
        .from("planner_assignments")
        .delete()
        .in("shift_id", targetShiftIds),
    );
    operations.push(
      supabaseClient
        .from("planner_shift_cancellations")
        .delete()
        .in("shift_id", targetShiftIds),
    );
  }

  const results = await Promise.all(operations);
  const firstError = results.find((res) => res.error)?.error;

  return firstError || null;
}

async function persistReplacementShiftEffects(replacementEntries) {
  if (!supabaseReady || !replacementEntries.length) return null;

  const assignmentRows = [];
  const cancellationRows = [];
  const replacedShiftIds = [];
  const canceledShiftIds = [];

  replacementEntries.forEach(([shiftId, entry]) => {
    const shift = getShiftById(shiftId);
    if (!shift) return;

    if (entry.mode === "cancel") {
      canceledShiftIds.push(shiftId);
      cancellationRows.push({
        shift_id: shiftId,
        shift_date: shift.date,
        created_by_employee_id: currentEmployeeRecord?.id || null,
      });
      return;
    }

    if (entry.mode === "replace" && entry.replacementUser) {
      replacedShiftIds.push(shiftId);
      assignmentRows.push({
        shift_id: shiftId,
        shift_date: shift.date,
        assigned_user: entry.replacementUser,
        created_by_employee_id: currentEmployeeRecord?.id || null,
      });
    }
  });

  const operations = [];

  if (replacedShiftIds.length) {
    operations.push(
      supabaseClient
        .from("planner_shift_cancellations")
        .delete()
        .in("shift_id", replacedShiftIds),
    );
  }

  if (canceledShiftIds.length) {
    operations.push(
      supabaseClient
        .from("planner_assignments")
        .delete()
        .in("shift_id", canceledShiftIds),
    );
  }

  if (assignmentRows.length) {
    operations.push(
      supabaseClient
        .from("planner_assignments")
        .upsert(assignmentRows, { onConflict: "shift_id" }),
    );
  }

  if (cancellationRows.length) {
    operations.push(
      supabaseClient
        .from("planner_shift_cancellations")
        .upsert(cancellationRows, { onConflict: "shift_id" }),
    );
  }

  const results = await Promise.all(operations);
  const firstError = results.find((res) => res.error)?.error;

  return firstError || null;
}

async function deleteReplacementEntriesFromSupabase(sourceType, sourceId) {
  if (!supabaseReady) return null;

  const { error } = await supabaseClient
    .from("planner_absence_replacements")
    .delete()
    .eq("source_type", sourceType)
    .eq("source_id", sourceId);

  return error || null;
}

async function saveReplacementPlanToSupabase(replacementEntries) {
  if (!supabaseReady || !replacementEntries.length) return null;

  const rows = replacementEntries.map(([shiftId, entry]) => {
    const shift = getShiftById(shiftId);
    return {
      shift_id: shiftId,
      shift_date: shift?.date || entry.weekFrom || entry.from,
      source_type: entry.sourceType,
      source_id: entry.sourceId,
      absent_user: entry.absentUser,
      from_date: entry.from,
      to_date: entry.to,
      mode: entry.mode,
      replacement_user: entry.replacementUser,
      week_from: entry.weekFrom,
      week_to: entry.weekTo,
    };
  });

  const { error } = await supabaseClient
    .from("planner_absence_replacements")
    .insert(rows);

  if (error) return error;

  return persistReplacementShiftEffects(replacementEntries);
}

async function clearAbsenceReplacementPlan(sourceType, sourceId) {
  const shiftIds = getReplacementShiftIdsForSource(sourceType, sourceId);
  const error = await deleteReplacementPlanFromSupabase(sourceType, sourceId);
  if (error) {
    console.error("Fehler beim Löschen der Ersatzplanung:", error);
    return alert(
      `Ersatzplanung konnte nicht gelöscht werden: ${error.message}`,
    );
  }

  clearReplacementPlanForSource(sourceType, sourceId);
  persist();
  render();
}

function getAbsenceEntry(type, entryId) {
  const list = type === "vacation" ? state.vacations : state.sickLeaves;
  return list.find((entry) => String(entry.id) === String(entryId)) || null;
}

function getPendingReplacementShifts(type, entryId) {
  const entry = getAbsenceEntry(type, entryId);
  if (!entry) return [];

  return getShiftsOfUserInRange(entry.user, entry.from, entry.to).filter(
    (shift) => {
      const replacement = state.absenceReplacements?.[shift.id];
      return !(
        replacement?.sourceType === type &&
        String(replacement?.sourceId) === String(entryId)
      );
    },
  );
}

function groupReplacementShiftsByWeek(shifts) {
  const grouped = {};

  shifts.forEach((shift) => {
    const key = weekKey(shift.date);
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(shift);
  });

  return Object.entries(grouped).map(([weekStart, weekShifts]) => ({
    weekStart,
    shifts: weekShifts,
  }));
}

function getReplacementPlannerSelectionKey(type, entryId) {
  return `${type}:${entryId}`;
}

function getReplacementPlannerChoice(type, entryId) {
  const key = getReplacementPlannerSelectionKey(type, entryId);
  return state.replacementPlannerChoice?.[key] || "AUSFALL";
}

function setReplacementPlannerChoice(type, entryId, value) {
  const key = getReplacementPlannerSelectionKey(type, entryId);
  if (!state.replacementPlannerChoice) state.replacementPlannerChoice = {};
  state.replacementPlannerChoice[key] = value || "AUSFALL";
  persist();
}

function ensureReplacementPlannerSelection(type, entryId) {
  const key = getReplacementPlannerSelectionKey(type, entryId);
  if (!state.replacementPlannerSelection)
    state.replacementPlannerSelection = {};
  if (!state.replacementPlannerSelection[key]) {
    state.replacementPlannerSelection[key] = {};
  }
  return state.replacementPlannerSelection[key];
}

function renderReplacementCalendar(type, entryId) {
  const pendingShifts = getPendingReplacementShifts(type, entryId);
  const selection = ensureReplacementPlannerSelection(type, entryId);
  const groupedWeeks = groupReplacementShiftsByWeek(pendingShifts);

  if (!pendingShifts.length) {
    return `<div class="text-sm rounded border border-emerald-200 bg-emerald-50 p-3 text-emerald-800">
      Alle Schichten in diesem Zeitraum sind ersetzt oder als Ausfall markiert.
    </div>`;
  }

  return groupedWeeks
    .map(({ weekStart, shifts }, index) => {
      const weekEnd = new Date(`${weekStart}T00:00:00`);
      weekEnd.setDate(weekEnd.getDate() + 6);

      const shiftButtons = shifts
        .map((shift) => {
          const selected = !!selection[shift.id];
          return `<button class="text-left border rounded p-2 ${selected ? "bg-slate-900 text-white border-slate-900" : "bg-white border-slate-300"}" onclick="toggleReplacementDay('${type}', '${entryId}', '${shift.id}')">
            <div class="font-semibold">${formatDateWithWeekday(shift.date)}</div>
            <div class="text-xs">${shift.label} · ${shift.start}-${shift.end}</div>
          </button>`;
        })
        .join("");

      return `<div class="border rounded-lg p-3 bg-slate-50">
        <div class="flex items-center justify-between gap-2 mb-2">
          <div class="font-semibold">KW ${index + 1}: ${formatDateDisplay(weekStart)} bis ${formatDateDisplay(isoDate(weekEnd))}</div>
          <button class="px-2 py-1 rounded bg-slate-700 text-white text-sm" onclick="applyReplacementForWeek('${type}', '${entryId}', '${weekStart}')">Diese Woche</button>
        </div>
        <div class="grid md:grid-cols-2 gap-2">${shiftButtons}</div>
      </div>`;
    })
    .join("");
}

function toggleReplacementDay(type, entryId, shiftId) {
  const selection = ensureReplacementPlannerSelection(type, entryId);
  selection[shiftId] = !selection[shiftId];
  openAbsenceReplacementPlanner(type, entryId);
}

async function applyReplacementForSelectedDays(
  type,
  entryId,
  selectedShiftIds,
) {
  if (!supabaseReady) return;

  const entry = getAbsenceEntry(type, entryId);
  if (!entry) return;

  const modeValue =
    document.getElementById("replacementUserSelect")?.value ||
    getReplacementPlannerChoice(type, entryId);
  setReplacementPlannerChoice(type, entryId, modeValue);
  const mode = modeValue === "AUSFALL" ? "cancel" : "replace";
  const replacementUser = mode === "replace" ? modeValue : null;

  if (!selectedShiftIds.length) {
    alert("Bitte mindestens eine Schicht auswählen.");
    return;
  }

  const newEntries = selectedShiftIds
    .map((shiftId) => {
      const shift = getShiftById(shiftId);
      if (!shift) return null;

      return [
        shiftId,
        {
          sourceType: type,
          sourceId: entryId,
          absentUser: entry.user,
          from: entry.from,
          to: entry.to,
          mode,
          replacementUser,
          weekFrom: weekKey(shift.date),
          weekTo: shift.date,
        },
      ];
    })
    .filter(Boolean);

  const deleteError = await deleteReplacementEntriesFromSupabase(type, entryId);
  if (deleteError) {
    console.error("Fehler beim Vorbereiten der Ersatzplanung:", deleteError);
    return alert(
      `Ersatzplanung konnte nicht vorbereitet werden: ${deleteError.message}`,
    );
  }

  newEntries.forEach(([shiftId, replacementEntry]) => {
    applyReplacementToShift(shiftId, replacementEntry);
  });

  const allEntries = getReplacementEntriesForSource(type, entryId);
  const saveError = await saveReplacementPlanToSupabase(allEntries);
  if (saveError) {
    console.error("Fehler beim Speichern der Ersatzplanung:", saveError);
    return alert(
      `Ersatzplanung konnte nicht gespeichert werden: ${saveError.message}`,
    );
  }

  const selectionKey = getReplacementPlannerSelectionKey(type, entryId);
  if (state.replacementPlannerSelection) {
    state.replacementPlannerSelection[selectionKey] = {};
  }

  persist();

  if (getPendingReplacementShifts(type, entryId).length) {
    openAbsenceReplacementPlanner(type, entryId);
    return;
  }

  if (state.replacementPlannerSelection) {
    delete state.replacementPlannerSelection[selectionKey];
  }
  if (state.replacementPlannerChoice) {
    delete state.replacementPlannerChoice[selectionKey];
  }

  getModalHost().innerHTML = "";
  render();
}

function applyReplacementForWeek(type, entryId, weekStart) {
  const pendingShiftIds = getPendingReplacementShifts(type, entryId)
    .filter((shift) => weekKey(shift.date) === weekStart)
    .map((shift) => shift.id);

  applyReplacementForSelectedDays(type, entryId, pendingShiftIds);
}

function applyReplacementForAll(type, entryId) {
  const pendingShiftIds = getPendingReplacementShifts(type, entryId).map(
    (shift) => shift.id,
  );

  applyReplacementForSelectedDays(type, entryId, pendingShiftIds);
}

function openAbsenceReplacementPlanner(type, entryId) {
  const entry = getAbsenceEntry(type, entryId);
  if (!entry) return;

  const host = getModalHost();
  const pendingShifts = getPendingReplacementShifts(type, entryId);
  const selection = ensureReplacementPlannerSelection(type, entryId);
  const selectedReplacementValue = getReplacementPlannerChoice(type, entryId);
  const selectedShiftIds = Object.entries(selection)
    .filter(([, selected]) => selected)
    .map(([shiftId]) => shiftId)
    .filter((shiftId) => pendingShifts.some((shift) => shift.id === shiftId));

  const replacementOptions = [
    `<option value="AUSFALL" ${selectedReplacementValue === "AUSFALL" ? "selected" : ""}>Ausfall</option>`,
    ...activeUsers()
      .filter((user) => user.name !== entry.user)
      .map(
        (user) =>
          `<option value="${user.name}" ${selectedReplacementValue === user.name ? "selected" : ""}>${user.name}</option>`,
      ),
  ].join("");

  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-5xl max-h-[90vh] overflow-auto p-4">
      <div class="flex items-start justify-between gap-3 mb-4">
        <div>
          <h3 class="text-lg font-bold">Ersatzplanung</h3>
          <p class="text-sm text-slate-600">${entry.user}: ${formatDateDisplay(entry.from)} bis ${formatDateDisplay(entry.to)}</p>
        </div>
        <div class="flex items-center gap-2">
          ${helpButton("ersatzplanung")}
          <button class="px-3 py-1 rounded bg-slate-200" onclick="closeReplacementPlanner('${type}', '${entryId}')">Abbrechen</button>
        </div>
      </div>

      <div class="grid md:grid-cols-[minmax(180px,260px)_1fr] gap-4">
        <div class="space-y-3">
          <label class="block text-sm font-medium">
            Ersatz-Mitarbeiter
            <select id="replacementUserSelect" class="border rounded p-2 w-full mt-1" onchange="setReplacementPlannerChoice('${type}', '${entryId}', this.value)">${replacementOptions}</select>
          </label>
          <button class="px-3 py-2 rounded bg-slate-900 text-white w-full" onclick="applyReplacementForAll('${type}', '${entryId}')">Alle</button>
          <button class="px-3 py-2 rounded bg-emerald-700 text-white w-full" onclick="applyReplacementForSelectedDays('${type}', '${entryId}', ${JSON.stringify(selectedShiftIds).replaceAll('"', "&quot;")})">OK</button>
          <div class="text-xs text-slate-500">${pendingShifts.length} offene Schicht(en), ${selectedShiftIds.length} ausgewählt.</div>
        </div>
        <div class="space-y-3">${renderReplacementCalendar(type, entryId)}</div>
      </div>
    </div>
  </div>`;
}

function closeReplacementPlanner(type = null, entryId = null) {
  if (type && entryId && state.replacementPlannerSelection) {
    const key = getReplacementPlannerSelectionKey(type, entryId);
    delete state.replacementPlannerSelection[key];
    delete state.replacementPlannerChoice?.[key];
    persist();
  }

  getModalHost().innerHTML = "";
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
  const specialDay = getSpecialDay(date);
  const blocked = isBlockedDay(date);

  const absenceKey = swappedDefault ? `${date}:${swappedDefault}` : null;
  const calendarAbsenceType = swappedDefault
    ? getCalendarAbsenceType(swappedDefault, date)
    : null;
  const manualAbsent = absenceKey ? !!state.absences[absenceKey] : false;
  const absent = manualAbsent || !!calendarAbsenceType;

  const canceled =
    !!state.shiftCancellations[id] || replacement?.mode === "cancel";

  let assigned = null;

  if (manualAssigned) {
    assigned = manualAssigned;
  } else if (blocked || canceled) {
    assigned = null;
  } else if (replacement?.mode === "replace" && replacement?.replacementUser) {
    assigned = replacement.replacementUser;
  } else if (!isOptional && !absent) {
    assigned = swappedDefault;
  } else {
    assigned = null;
  }

  const open = !blocked && (canceled || !assigned);

  return {
    id,
    date,
    label: template.label,
    start: template.start,
    end: template.end,
    options: template.options,
    assigned,
    originalAssigned: blocked ? null : swappedDefault,
    absenceType: calendarAbsenceType || (manualAbsent ? "abwesend" : null),
    replacement,
    open,
    blocked,
    specialDay,
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
  const fixedColors = {
    Lavdrim: "bg-green-200 text-green-900",
    Roger: "bg-blue-200 text-blue-900",
    Dashmir: "bg-yellow-200 text-yellow-900",
    Thomas: "bg-purple-200 text-purple-900",
    Musa: "bg-orange-200 text-orange-900",
    Ardian: "bg-teal-200 text-teal-900",
  };

  const colorByKey = {
    green: "bg-green-200 text-green-900",
    blue: "bg-blue-200 text-blue-900",
    yellow: "bg-yellow-200 text-yellow-900",
    purple: "bg-purple-200 text-purple-900",
    orange: "bg-orange-200 text-orange-900",
    teal: "bg-teal-200 text-teal-900",
    cyan: "bg-cyan-200 text-cyan-900",
    pink: "bg-pink-200 text-pink-900",
    rose: "bg-rose-200 text-rose-900",
    lime: "bg-lime-200 text-lime-900",
    gray: "bg-gray-200 text-gray-900",
  };

  const emp = Object.values(state.employees || {}).find(
    (e) => e.display_name === name || e.name === name,
  );

  const fromDb = emp?.color_key ? colorByKey[emp.color_key] : null;
  const raw = fromDb || fixedColors[name] || "bg-gray-200 text-gray-900";

  const parts = raw.split(" ");

  return {
    bg: parts[0],
    text: parts[1],
    raw,
  };
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

function getSpecialDay(date) {
  return state.specialDays?.[date] || null;
}

function isBlockedDay(date) {
  const d = getSpecialDay(date);
  if (!d) return false;

  return d.type === "holiday" || d.type === "bridge" || d.type === "company";
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
  const date = shiftId.slice(0, 10);
  const specialDay = getSpecialDay(date);

  if (specialDay) {
    const label =
      specialDay.type === "holiday"
        ? `FEIERTAG${specialDay.label ? ` – ${specialDay.label}` : ""}`
        : specialDay.type === "bridge"
          ? `BRÜCKENTAG${specialDay.label ? ` – ${specialDay.label}` : ""}`
          : `BETRIEBSFERIEN${specialDay.label ? ` – ${specialDay.label}` : ""}`;

    return {
      label,
      cls: "bg-red-100 text-red-900",
      borderCls: "border-red-500",
      ringCls: "border-2",
    };
  }

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
  const isAdmin = currentUser?.role === "admin";

  const dayNames = [
    "Montag",
    "Dienstag",
    "Mittwoch",
    "Donnerstag",
    "Freitag",
    "Samstag",
    "Sonntag",
  ];

  function renderShiftCell(shiftId, options) {
    const meta = assignedMeta(shiftId, options);
    const hasManualAssignment = !!state.assignments?.[shiftId];
    const hasManualCancellation = !!state.shiftCancellations?.[shiftId];
    const hasReplacement = !!state.absenceReplacements?.[shiftId];
    const isManual =
      hasManualAssignment || hasManualCancellation || hasReplacement;

    return `<div class="flex flex-col items-center gap-1">
      <span>${meta.label}</span>
      ${
        isAdmin && isManual
          ? `<button class="px-1.5 py-0.5 rounded bg-slate-800 text-white text-[10px] leading-none" onclick="clearManualShift('${shiftId}')">Zurück</button>`
          : ""
      }
    </div>`;
  }

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
      const startShiftId = `${isoDate(startSunday)}-su-1`;

      body += `<tr class="border-b bg-amber-50">
        <td class="p-2"></td>
        <td class="p-2 font-semibold">Start (Sonntag)</td>
        <td class="p-2 bg-slate-50" colspan="4"></td>
        <td class="p-2 font-semibold text-center ${assignedMeta(startShiftId, prevTemplate.sunday[1].options).cls} ${assignedMeta(startShiftId, prevTemplate.sunday[1].options).borderCls} ${assignedMeta(startShiftId, prevTemplate.sunday[1].options).ringCls}">
          ${renderShiftCell(startShiftId, prevTemplate.sunday[1].options)}
        </td>
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

        const id1 = `${dateIso}-mf-0`;
        const id2 = `${dateIso}-mf-1`;
        const id3 = `${dateIso}-mf-2`;

        const m1 = assignedMeta(id1, s1.options);
        const m2 = assignedMeta(id2, s2.options);
        const m3 = assignedMeta(id3, s3.options);

        body += `<tr class="border-b">
          <td class="p-2 font-semibold">${weekLabel}</td>
          <td class="p-2">${dayName}<div class="text-[11px] text-slate-500">${formatDateDisplay(dateIso)}</div></td>
          <td class="p-2 ${m1.cls} ${m1.borderCls} ${m1.ringCls} font-semibold text-center">${renderShiftCell(id1, s1.options)}</td>
          <td class="p-2 font-semibold text-center bg-white">05:00-11:00</td>
          <td class="p-2 ${m2.cls} ${m2.borderCls} ${m2.ringCls} font-semibold text-center">${renderShiftCell(id2, s2.options)}</td>
          <td class="p-2 font-semibold text-center bg-white">13:00-19:00</td>
          <td class="p-2 ${m3.cls} ${m3.borderCls} ${m3.ringCls} font-semibold text-center">${renderShiftCell(id3, s3.options)}</td>
          <td class="p-2 font-semibold text-center bg-white">21:00-03:00</td>
        </tr>`;
      } else if (dayIndex === 5) {
        const s1 = template.saturday[0];
        const s2 = template.saturday[1];
        const dateIso = isoDate(date);

        const id1 = `${dateIso}-sa-0`;
        const id2 = `${dateIso}-sa-1`;

        const m1 = assignedMeta(id1, s1.options);
        const m2 = assignedMeta(id2, s2.options);

        body += `<tr class="border-b bg-amber-50">
          <td class="p-2 font-semibold">${weekLabel}</td>
          <td class="p-2 font-semibold">${dayName}<div class="text-[11px] text-slate-500">${formatDateDisplay(dateIso)}</div></td>
          <td class="p-2 ${m1.cls} ${m1.borderCls} ${m1.ringCls} font-semibold text-center">${renderShiftCell(id1, s1.options)}</td>
          <td class="p-2 font-semibold text-center bg-amber-100">05:00-11:00 (Sa Morgen)</td>
          <td class="p-2 ${m2.cls} ${m2.borderCls} ${m2.ringCls} font-semibold text-center">${renderShiftCell(id2, s2.options)}</td>
          <td class="p-2 font-semibold text-center bg-amber-200">16:00-22:00 (Sa Abend)</td>
          <td class="p-2 bg-slate-50" colspan="2"></td>
        </tr>`;
      } else {
        const s1 = template.sunday[0];
        const s2 = template.sunday[1];
        const dateIso = isoDate(date);

        const id1 = `${dateIso}-su-0`;
        const id2 = `${dateIso}-su-1`;

        const m1 = assignedMeta(id1, s1.options);
        const m2 = assignedMeta(id2, s2.options);

        body += `<tr class="border-b bg-amber-50">
          <td class="p-2 font-semibold">${weekLabel}</td>
          <td class="p-2 font-semibold">${dayName}<div class="text-[11px] text-slate-500">${formatDateDisplay(dateIso)}</div></td>
          <td class="p-2 ${m1.cls} ${m1.borderCls} ${m1.ringCls} font-semibold text-center">${renderShiftCell(id1, s1.options)}</td>
          <td class="p-2 font-semibold text-center bg-blue-100">06:00-12:00 (So Morgen)</td>
          <td class="p-2 bg-slate-50" colspan="2"></td>
          <td class="p-2 ${m2.cls} ${m2.borderCls} ${m2.ringCls} font-semibold text-center">${renderShiftCell(id2, s2.options)}</td>
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
          <th class='p-2 text-center'>Nachtschicht</th>
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

async function assignShift(shiftId) {
  if (!supabaseReady) return;

  const select = document.getElementById(`sel-${shiftId}`);
  if (!select) return;

  const name = select.value;
  if (!canAssignUserToShift(name, shiftId)) return;

  const shift = getShiftById(shiftId);
  if (!shift) return;

  const { error } = await supabaseClient.from("planner_assignments").upsert(
    {
      shift_id: shiftId,
      shift_date: shift.date,
      assigned_user: name,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "shift_id" },
  );

  if (error) {
    console.error("Fehler bei Zuweisung:", error);
    return alert(`Zuweisung fehlgeschlagen: ${error.message}`);
  }

  // falls vorher storniert → entfernen
  await supabaseClient
    .from("planner_shift_cancellations")
    .delete()
    .eq("shift_id", shiftId);

  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = name;

  persist();
  render();
}

async function cancelShift(shiftId) {
  if (!supabaseReady) return;

  const shift = getShiftById(shiftId);
  if (!shift) return;

  const { error } = await supabaseClient
    .from("planner_shift_cancellations")
    .upsert(
      {
        shift_id: shiftId,
        shift_date: shift.date,
        created_by_employee_id: currentEmployeeRecord?.id || null,
      },
      { onConflict: "shift_id" },
    );

  if (error) {
    console.error("Fehler beim Stornieren:", error);
    return alert(`Stornierung fehlgeschlagen: ${error.message}`);
  }

  const { error: deleteError } = await supabaseClient
    .from("planner_assignments")
    .delete()
    .eq("shift_id", shiftId);

  if (deleteError) {
    console.error("Fehler beim Entfernen der Zuweisung:", deleteError);
    return alert(
      `Ausfall gespeichert, aber Zuweisung konnte nicht entfernt werden: ${deleteError.message}`,
    );
  }

  delete state.assignments[shiftId];
  state.shiftCancellations[shiftId] = true;

  persist();
  render();
}

async function clearManualShift(shiftId) {
  if (currentUser?.role !== "admin") return;
  if (!supabaseReady) return;

  const shift = getShiftById(shiftId);
  if (!shift) {
    return alert("Schicht konnte nicht gefunden werden.");
  }

  const ok = confirm(
    `Manuelle Planung für ${formatDateWithWeekday(shift.date)} – ${shift.label} zurücksetzen?`,
  );
  if (!ok) return;

  const [assignmentDelete, cancellationDelete, replacementDelete] =
    await Promise.all([
      supabaseClient
        .from("planner_assignments")
        .delete()
        .eq("shift_id", shiftId),
      supabaseClient
        .from("planner_shift_cancellations")
        .delete()
        .eq("shift_id", shiftId),
      supabaseClient
        .from("planner_absence_replacements")
        .delete()
        .eq("shift_id", shiftId),
    ]);

  const firstError =
    assignmentDelete.error ||
    cancellationDelete.error ||
    replacementDelete.error;

  if (firstError) {
    console.error("Fehler beim Zurücksetzen der Schicht:", firstError);
    return alert(
      `Schicht konnte nicht zurückgesetzt werden: ${firstError.message}`,
    );
  }

  delete state.assignments[shiftId];
  delete state.shiftCancellations[shiftId];
  if (state.absenceReplacements) {
    delete state.absenceReplacements[shiftId];
  }

  persist();
  render();
}

function renderMyShifts() {
  const isSpringerUser = isSpringer(currentUser.name);

  const shifts = generateThreeMonths()
    .filter((s) => s.assigned === currentUser.name)
    .slice(0, 180);

  const rows = shifts
    .map((s) => {
      const status =
        s.assigned === currentUser.name
          ? '<span class="text-emerald-700 font-semibold">Eingeplant</span>'
          : s.open
            ? '<span class="text-red-600 font-semibold">OFFEN</span>'
            : '<span class="text-emerald-700">Besetzt</span>';

      return `<tr class="border-b">
        <td class="p-2">${formatDateWithWeekday(s.date)}</td>
        <td class="p-2">${s.label}</td>
        <td class="p-2">${s.start}–${s.end}</td>
        <td class="p-2">${status}</td>
        <td class="p-2">
          ${
            !isSpringerUser
              ? `<button class='px-2 py-1 bg-amber-200 rounded mr-2' onclick="markAbsent('${s.id}','${s.date}','${currentUser.name}')">Abwesenheit</button>`
              : "-"
          }
        </td>
      </tr>`;
    })
    .join("");

  const weekendRows = isSpringerUser
    ? generateThreeMonths()
        .filter((s) => s.id.includes("-sa-1") || s.id.includes("-su-0"))
        .slice(0, 180)
        .map((s) => {
          const key = `${s.id}:${currentUser.name}`;
          const val = state.availability[key] || "";

          return `<tr class='border-b'>
            <td class='p-2'>${formatDateWithWeekday(s.date)}</td>
            <td class='p-2'>${s.label}</td>
            <td class='p-2'>${s.start}–${s.end}</td>
            <td class='p-2'>
              <select class='border rounded p-1' onchange="setWeekendAvailability('${s.id}', '${currentUser.name}', '${s.date}', this.value)">
                <option value='' ${val === "" ? "selected" : ""}>-</option>
                <option value='yes' ${val === "yes" ? "selected" : ""}>Kann</option>
                <option value='no' ${val === "no" ? "selected" : ""}>Kann nicht</option>
              </select>
            </td>
          </tr>`;
        })
        .join("")
    : "";

  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-lg font-semibold mb-3'>Meine Schichten (90 Tage)</h2>
      <div class='overflow-auto border rounded-lg'>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100'>
            <tr>
              <th class='p-2 text-left'>Datum</th>
              <th class='p-2 text-left'>Schicht</th>
              <th class='p-2 text-left'>Zeit</th>
              <th class='p-2 text-left'>Status</th>
              <th class='p-2 text-left'>Aktionen</th>
            </tr>
          </thead>
          <tbody>${rows || `<tr><td class='p-2' colspan='5'>Keine Schichten gefunden.</td></tr>`}</tbody>
        </table>
      </div>
    </div>

    ${
      isSpringerUser
        ? `<div>
            <h3 class='text-lg font-semibold mb-3'>Wochenend-Verfügbarkeit</h3>
            <div class='overflow-auto border rounded-lg'>
              <table class='w-full text-sm'>
                <thead class='bg-slate-100'>
                  <tr>
                    <th class='p-2 text-left'>Datum</th>
                    <th class='p-2 text-left'>Schicht</th>
                    <th class='p-2 text-left'>Zeit</th>
                    <th class='p-2 text-left'>Verfügbarkeit</th>
                  </tr>
                </thead>
                <tbody>${weekendRows || `<tr><td class='p-2' colspan='4'>Keine relevanten Wochenendschichten gefunden.</td></tr>`}</tbody>
              </table>
            </div>
          </div>`
        : ""
    }
  </div>`;
}

async function requestSaturdayEvening(shiftId) {
  if (!supabaseReady) return;

  const shift = getShiftById(shiftId);
  if (!shift) return alert("Schicht konnte nicht gefunden werden.");

  const requestKey = `${shiftId}:${currentUser.name}`;

  const { error } = await supabaseClient
    .from("planner_saturday_requests")
    .upsert(
      {
        request_key: requestKey,
        shift_id: shiftId,
        shift_date: shift.date,
        user_name: currentUser.name,
      },
      { onConflict: "request_key" },
    );

  if (error) {
    console.error("Fehler beim Speichern der Samstags-Anfrage:", error);
    return alert(
      `Samstags-Anfrage konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  state.saturdayEveningRequests[requestKey] = true;
  persist();
  render();
}

async function setWeekendAvailability(shiftId, userName, date, status) {
  if (!supabaseReady) return;

  const availabilityKey = `${shiftId}:${userName}`;

  if (!status) {
    const { error } = await supabaseClient
      .from("planner_availability")
      .delete()
      .eq("availability_key", availabilityKey);

    if (error) {
      console.error("Fehler beim Löschen der Verfügbarkeit:", error);
      return alert(
        `Verfügbarkeit konnte nicht gelöscht werden: ${error.message}`,
      );
    }

    delete state.availability[availabilityKey];
    persist();
    render();
    return;
  }

  const { error } = await supabaseClient.from("planner_availability").upsert(
    {
      availability_key: availabilityKey,
      shift_id: shiftId,
      shift_date: date,
      user_name: userName,
      status,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "availability_key" },
  );

  if (error) {
    console.error("Fehler bei Verfügbarkeit:", error);
    return alert(
      `Verfügbarkeit konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  state.availability[availabilityKey] = status;
  persist();
  render();
}

async function loadAvailabilityFromSupabase() {
  const { data, error } = await supabaseClient
    .from("planner_availability")
    .select("*");

  if (error) {
    console.error("Fehler beim Laden der Verfügbarkeit:", error);
    return {};
  }

  const map = {};
  (data || []).forEach((row) => {
    map[row.availability_key] = row.status;
  });

  return map;
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

function ensurePersonnelPendingState() {
  if (!state.ui) state.ui = {};
  if (!state.ui.pendingEmployeeEdits) state.ui.pendingEmployeeEdits = {};
  if (!state.ui.pendingSlotAssignments) state.ui.pendingSlotAssignments = {};
  return state.ui;
}

function queueEmployeeEdit(id, field, value) {
  if (!id) return;
  const ui = ensurePersonnelPendingState();
  ui.pendingEmployeeEdits[id] = {
    ...(ui.pendingEmployeeEdits[id] || {}),
    [field]: value,
  };
  persist();
  render();
}

function queueEmployeeActive(id, isActive) {
  if (!id) return;
  const employee = state.employees?.[id];
  const name = employee?.display_name || employee?.name || "Mitarbeiter";
  if (!isActive && !confirm(`${name} als inaktiv vormerken?`)) return;
  queueEmployeeEdit(id, "is_active", isActive);
}

function queueSlotAssignment(slot, value) {
  const ui = ensurePersonnelPendingState();
  ui.pendingSlotAssignments[slot] = value;
  persist();
  render();
}

function hasPersonnelPendingChanges() {
  const ui = ensurePersonnelPendingState();
  return (
    Object.keys(ui.pendingEmployeeEdits).length > 0 ||
    Object.keys(ui.pendingSlotAssignments).length > 0
  );
}

function renderPlanningPersonal() {
  const ui = ensurePersonnelPendingState();
  const pendingEmployees = ui.pendingEmployeeEdits;
  const pendingSlots = ui.pendingSlotAssignments;
  const hasPending = hasPersonnelPendingChanges();

  const personnelCards = allUsers()
    .map((u) => {
      const pending = pendingEmployees[u.id] || {};
      const name = pending.display_name ?? u.name;
      const type = pending.employee_type ?? u.type;
      const isActive =
        pending.is_active ??
        (u.isActive !== false && !state.inactiveUsers?.[u.name]);
      const colorKey = pending.color_key ?? u.color_key ?? "gray";
      const role = pending.role ?? u.role ?? "employee";
      const changed = Object.keys(pending).length > 0;
      const cardClass = changed
        ? "border-amber-300 bg-amber-50"
        : "border-slate-200 bg-white";
      const statusClass = isActive
        ? "bg-emerald-100 text-emerald-800"
        : "bg-slate-200 text-slate-700";

      return `<div class='border ${cardClass} rounded-lg p-3 shadow-sm'>
        <div class='flex items-start justify-between gap-3 mb-3'>
          <div class='min-w-0'>
            <div class='font-semibold text-slate-900 truncate'>${escapeHtml(name)}</div>
            ${changed ? `<div class='text-xs text-amber-700 mt-1'>Ungespeichert</div>` : ""}
          </div>
          <span class='shrink-0 px-2 py-1 rounded-full text-xs font-semibold ${statusClass}'>${isActive ? "Aktiv" : "Inaktiv"}</span>
        </div>
        <div class='grid sm:grid-cols-2 gap-3'>
          <label class='text-xs text-slate-500'>Name
            <input id='employee-name-${u.id}' class='mt-1 border rounded p-2 w-full text-sm bg-white' value='${escapeHtml(name)}' onchange="queueEmployeeEdit('${u.id}', 'display_name', this.value.trim())" />
          </label>
          <label class='text-xs text-slate-500'>Typ
            <select id='employee-type-${u.id}' class='mt-1 border rounded p-2 w-full text-sm bg-white' onchange="queueEmployeeEdit('${u.id}', 'employee_type', this.value)">
              <option value='core' ${type === "core" ? "selected" : ""}>A/B/C</option>
              <option value='springer' ${type === "springer" ? "selected" : ""}>Springer</option>
            </select>
          </label>
          <label class='text-xs text-slate-500'>Farbe
            <select id='employee-color-${u.id}' class='mt-1 border rounded p-2 w-full text-sm bg-white' onchange="queueEmployeeEdit('${u.id}', 'color_key', this.value)">
              ${EMPLOYEE_COLOR_OPTIONS.map((color) => `<option value='${color.key}' ${colorKey === color.key ? "selected" : ""}>${color.label}</option>`).join("")}
            </select>
          </label>
          <label class='text-xs text-slate-500'>Rolle
            <select id='employee-role-${u.id}' class='mt-1 border rounded p-2 w-full text-sm bg-white' onchange="queueEmployeeEdit('${u.id}', 'role', this.value)">
              <option value='employee' ${role === "employee" ? "selected" : ""}>Mitarbeiter</option>
              <option value='admin' ${role === "admin" ? "selected" : ""}>Admin</option>
            </select>
          </label>
        </div>
        <div class='flex justify-end mt-3'>
          ${
            isActive
              ? `<button class='px-3 py-2 rounded bg-rose-700 text-white text-sm' onclick="queueEmployeeActive('${u.id}', false)">Inaktiv vormerken</button>`
              : `<button class='px-3 py-2 rounded bg-emerald-700 text-white text-sm' onclick="queueEmployeeActive('${u.id}', true)">Reaktivieren vormerken</button>`
          }
        </div>
      </div>`;
    })
    .join("");

  const slotAssignmentRows = SLOT_CODES.map((slot) => {
    const currentValue = state.slotAssignments?.[slot] || "";
    const selectedValue = pendingSlots[slot] ?? currentValue;
    const changed = Object.prototype.hasOwnProperty.call(pendingSlots, slot);
    const rowClass = changed ? "border-b bg-amber-50" : "border-b";
    const options = activeUsers()
      .map(
        (u) =>
          `<option value='${escapeHtml(u.name)}' ${selectedValue === u.name ? "selected" : ""}>${escapeHtml(u.name)}</option>`,
      )
      .join("");
    return `<tr class='${rowClass}'>
      <td class='p-2 font-semibold'>${slot}</td>
      <td class='p-2'><select id='slot-${slot}' class='border rounded p-1 w-full' onchange="queueSlotAssignment('${slot}', this.value)">${options}</select></td>
      <td class='p-2 text-xs text-amber-700'>${changed ? "ungespeichert" : ""}</td>
    </tr>`;
  }).join("");

  return `<div class='space-y-4'>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <div class='flex items-center justify-between gap-2 mb-3'>
        <h3 class='font-semibold'>Personalverwaltung</h3>
        ${helpButton("planningPersonal")}
      </div>
      <div class='border rounded-lg bg-white p-3 mb-4'>
        <div class='font-semibold text-sm mb-3'>Mitarbeiter hinzufügen</div>
        <div class='grid sm:grid-cols-2 lg:grid-cols-3 gap-3'>
          <input id='newEmployeeName' class='border rounded p-2' placeholder='Neuer Name' />
          <select id='newEmployeeType' class='border rounded p-2'><option value='springer'>Springer</option><option value='core'>A/B/C</option></select>
          <select id='newEmployeeSlot' class='border rounded p-2'>
            <option value=''>Kein Slot</option>
            ${SLOT_CODES.map((slot) => `<option value='${slot}'>${slot}</option>`).join("")}
          </select>
          <select id='newEmployeeColor' class='border rounded p-2'>
            ${EMPLOYEE_COLOR_OPTIONS.map((color) => `<option value='${color.key}'>${color.label}</option>`).join("")}
          </select>
          <select id='newEmployeeRole' class='border rounded p-2'><option value='employee'>Mitarbeiter</option><option value='admin'>Admin</option></select>
          <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addEmployee()'>Mitarbeiter hinzufügen</button>
        </div>
      </div>
      <div class='grid lg:grid-cols-2 gap-3'>
        ${personnelCards}
      </div>
    </div>
    <div class='border rounded-lg p-3 bg-slate-50'>
      <h3 class='font-semibold mb-2'>Zuordnung</h3>
      <p class='text-sm text-slate-500 mb-2'>Admin kann festlegen, welcher Mitarbeiter aktuell A/B/C/D/E/F ist. Der Schichtplan passt sich nach dem zentralen Speichern an.</p>
      <div class='overflow-auto max-h-[55vh]'>
        <table class='w-full text-sm'><thead class='bg-white sticky top-0'><tr><th class='p-2 text-left'>Slot</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Status</th></tr></thead>
        <tbody>${slotAssignmentRows}</tbody></table>
      </div>
    </div>
    <div class='flex items-center justify-end gap-3 border rounded-lg p-3 bg-white'>
      ${hasPending ? `<span class='text-sm text-amber-700'>Es gibt ungespeicherte Änderungen</span>` : `<span class='text-sm text-slate-500'>Keine ungespeicherten Änderungen</span>`}
      <button class='px-4 py-2 rounded ${hasPending ? "bg-slate-900 text-white" : "bg-slate-200 text-slate-500 cursor-not-allowed"}' ${hasPending ? "" : "disabled"} onclick='saveAllPersonnelChanges()'>Alle Änderungen speichern</button>
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
  if (!supabaseReady) return;
  openAbsenceReplacementPlanner(type, entryId);
}

async function deleteManualAbsence(date, userName) {
  if (!supabaseReady) return;

  const confirmDelete = confirm(
    `Abwesenheit von ${userName} am ${formatDateDisplay(date)} wirklich löschen?`,
  );
  if (!confirmDelete) return;

  const { error } = await supabaseClient
    .from("planner_absences")
    .delete()
    .eq("absence_date", date)
    .eq("user_name", userName);

  if (error) {
    console.error("Fehler beim Löschen der Abwesenheit:", error);
    return alert(`Fehler: ${error.message}`);
  }

  delete state.absences[`${date}:${userName}`];

  persist();
  render();
}

function renderPlanningAbstinenz() {
  const openShifts = generateThreeMonths().filter((s) => s.open);

  const monthValue = `${new Date().getFullYear()}-${String(
    new Date().getMonth() + 1,
  ).padStart(2, "0")}`;

  const personOptions = activeUsers()
    .map((u) => `<option value='${u.name}'>${u.name}</option>`)
    .join("");

  const absenceRows = Object.entries(state.absences || {})
    .slice(-100)
    .reverse()
    .map(([key, value]) => {
      const [date, user] = key.split(":");
      return `<tr class='border-b'>
        <td class='p-2'>${formatDateWithWeekday(date)}</td>
        <td class='p-2'>${user}</td>
        <td class='p-2'>Kann nicht kommen</td>
        <td class='p-2'>
          <button class='px-2 py-1 rounded bg-red-600 text-white' onclick="deleteManualAbsence('${date}','${user}')">Löschen</button>
        </td>
      </tr>`;
    })
    .join("");

  const vacationRows = state.vacations
    .slice(-20)
    .reverse()
    .map(
      (v) => `<tr class='border-b'>
        <td class='p-2'>${v.user}</td>
        <td class='p-2'>${formatDateDisplay(v.from)}</td>
        <td class='p-2'>${formatDateDisplay(v.to)}</td>
        <td class='p-2'>
          <button class='px-2 py-1 rounded bg-red-600 text-white' onclick="deleteVacation('${v.id}')">Löschen</button>
          <button class='px-2 py-1 rounded bg-slate-900 text-white ml-1' onclick="planAbsenceReplacement('vacation','${v.id}')">Ersatz</button>
        </td>
      </tr>`,
    )
    .join("");

  const sickRows = state.sickLeaves
    .slice(-20)
    .reverse()
    .map(
      (v) => `<tr class='border-b'>
        <td class='p-2'>${v.user}</td>
        <td class='p-2'>${formatDateDisplay(v.from)}</td>
        <td class='p-2'>${formatDateDisplay(v.to)}</td>
        <td class='p-2'>
          <button class='px-2 py-1 rounded bg-red-600 text-white' onclick="deleteSickLeave('${v.id}')">Löschen</button>
          <button class='px-2 py-1 rounded bg-slate-900 text-white ml-1' onclick="planAbsenceReplacement('sick','${v.id}')">Ersatz</button>
        </td>
      </tr>`,
    )
    .join("");

  const openRows = openShifts
    .map((s) => {
      const options = activeUsers()
        .map((u) => `<option value="${u.name}">${u.name}</option>`)
        .join("");

      return `<tr class='border-b'>
        <td class='p-2'>${formatDateWithWeekday(s.date)}</td>
        <td class='p-2'>${s.label}</td>
        <td class='p-2'>${s.start}–${s.end}</td>
        <td class='p-2'>${s.originalAssigned || "-"}</td>
        <td class='p-2'>
          <select id='sel-${s.id}' class='border rounded p-1'>${options}</select>
        </td>
        <td class='p-2'>
          <button class='px-2 py-1 rounded bg-blue-700 text-white mr-2' onclick="assignShift('${s.id}')">Übernehmen</button>
          <button class='px-2 py-1 rounded bg-red-700 text-white' onclick="cancelShift('${s.id}')">Ausfall</button>
        </td>
      </tr>`;
    })
    .join("");

  return `
  <div class='space-y-4'>
    <div class='flex justify-end'>${helpButton("abstinenz")}</div>

    <div class='grid md:grid-cols-2 gap-4'>
      <div class='border rounded-lg p-3 bg-white'>
        <h3 class='font-semibold mb-2'>Klärung Abstinenz</h3>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100'>
            <tr>
              <th class='p-2'>Datum</th>
              <th class='p-2'>Mitarbeiter</th>
              <th class='p-2'>Typ</th>
              <th class='p-2'>Aktion</th>
            </tr>
          </thead>
          <tbody>
            ${absenceRows || "<tr><td class='p-2' colspan='4'>Keine Einträge</td></tr>"}
          </tbody>
        </table>
      </div>

      <div class='border rounded-lg p-3 bg-white'>
        <h3 class='font-semibold mb-3'>Urlaub / Krankmeldung eintragen</h3>

        <div class='grid grid-cols-2 gap-3 text-sm'>
          <label>
            Monat
            <input id='adminMonth' type='month' value='${monthValue}' class='border rounded p-2 w-full mt-1'/>
          </label>

          <label>
            Mitarbeiter
            <select id='adminCalUser' class='border rounded p-2 w-full mt-1'>
              ${personOptions}
            </select>
          </label>

          <label>
            Von
            <input id='adminFrom' type='date' value='${todayIso()}' class='border rounded p-2 w-full mt-1'/>
          </label>

          <label>
            Bis
            <input id='adminTo' type='date' value='${todayIso()}' class='border rounded p-2 w-full mt-1'/>
          </label>
        </div>

        <div class='flex gap-2 mt-4'>
          <button class='px-3 py-2 rounded bg-emerald-700 text-white' onclick='addVacation()'>Urlaub speichern</button>
          <button class='px-3 py-2 rounded bg-amber-700 text-white' onclick='addSickLeave()'>Krank speichern</button>
        </div>
      </div>
    </div>

    <div class='border rounded-lg p-3 bg-white'>
      <h3 class='font-semibold mb-2'>Geplante Urlaube</h3>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100'>
          <tr>
            <th class='p-2'>Mitarbeiter</th>
            <th class='p-2'>Von</th>
            <th class='p-2'>Bis</th>
            <th class='p-2'>Aktion</th>
          </tr>
        </thead>
        <tbody>
          ${vacationRows || "<tr><td class='p-2' colspan='4'>Keine Einträge</td></tr>"}
        </tbody>
      </table>
    </div>

    <div class='border rounded-lg p-3 bg-white'>
      <h3 class='font-semibold mb-2'>Krankmeldungen</h3>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100'>
          <tr>
            <th class='p-2'>Mitarbeiter</th>
            <th class='p-2'>Von</th>
            <th class='p-2'>Bis</th>
            <th class='p-2'>Aktion</th>
          </tr>
        </thead>
        <tbody>
          ${sickRows || "<tr><td class='p-2' colspan='4'>Keine Einträge</td></tr>"}
        </tbody>
      </table>
    </div>

    <div class='border rounded-lg p-3 bg-white'>
      <h3 class='font-semibold mb-2'>Offene Schichten</h3>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100'>
          <tr>
            <th class='p-2'>Datum</th>
            <th class='p-2'>Schicht</th>
            <th class='p-2'>Zeit</th>
            <th class='p-2'>Vorher</th>
            <th class='p-2'>Übernehmen</th>
            <th class='p-2'>Aktion</th>
          </tr>
        </thead>
        <tbody>
          ${openRows || "<tr><td class='p-2' colspan='6'>Keine offenen Schichten</td></tr>"}
        </tbody>
      </table>
    </div>

  </div>`;
}

function isPrimaryCoreAbsentForShift(shift) {
  const primary = shift.originalAssigned;
  if (!primary || !isCoreEmployee(primary)) return false;

  const manualAbsent = !!state.absences?.[`${shift.date}:${primary}`];
  const calendarAbsent = !!getCalendarAbsenceType(primary, shift.date);

  return manualAbsent || calendarAbsent;
}

function getSuggestedSpringerForShift(shift) {
  if (!shift) return null;

  const availableSpringer = activeUsers()
    .filter((u) => u.type === "springer")
    .filter((u) => state.availability[`${shift.id}:${u.name}`] === "yes")
    .filter((u) => canAssignUserToShift(u.name, shift.id, true));

  if (!availableSpringer.length) return null;

  const shifts = generateThreeMonths();

  const ranked = availableSpringer
    .map((u) => {
      const weekendCount = shifts.filter(
        (s) =>
          s.assigned === u.name &&
          (s.id.includes("-sa-") || s.id.includes("-su-")),
      ).length;

      const totalCount = shifts.filter((s) => s.assigned === u.name).length;

      return {
        name: u.name,
        weekendCount,
        totalCount,
      };
    })
    .sort((a, b) => {
      if (a.weekendCount !== b.weekendCount) {
        return a.weekendCount - b.weekendCount;
      }
      if (a.totalCount !== b.totalCount) {
        return a.totalCount - b.totalCount;
      }
      return a.name.localeCompare(b.name);
    });

  return ranked[0]?.name || null;
}

function renderPlanningWochenende() {
  const weekendShifts = generateThreeMonths().filter(
    (s) => s.id.includes("-sa-1") || s.id.includes("-su-0"),
  );

  const optionalRows = weekendShifts
    .map((s) => {
      const primary = s.originalAssigned || "-";
      const primaryAbsent = isPrimaryCoreAbsentForShift(s);

      const yesUsers = activeUsers()
        .filter((u) => u.type === "springer")
        .filter((u) => state.availability[`${s.id}:${u.name}`] === "yes")
        .map((u) => u.name);

      const noUsers = activeUsers()
        .filter((u) => u.type === "springer")
        .filter((u) => state.availability[`${s.id}:${u.name}`] === "no")
        .map((u) => u.name);

      const suggested = getSuggestedSpringerForShift(s);

      const options = yesUsers
        .map(
          (name) =>
            `<option value='${name}' ${name === suggested ? "selected" : ""}>${name}${name === suggested ? " · Vorschlag" : ""}</option>`,
        )
        .join("");

      const canAssign = primaryAbsent && yesUsers.length > 0;

      return `<tr class='border-b'>
        <td class='p-2'>${formatDateWithWeekday(s.date)}</td>
        <td class='p-2'>${s.label}</td>
        <td class='p-2'>${primary}</td>
        <td class='p-2'>
          ${
            primaryAbsent
              ? "<span class='px-2 py-1 rounded bg-emerald-100 text-emerald-800'>abwesend</span>"
              : "<span class='px-2 py-1 rounded bg-slate-100 text-slate-600'>nicht abwesend</span>"
          }
        </td>
        <td class='p-2'>${yesUsers.length ? yesUsers.join(", ") : "-"}</td>
        <td class='p-2'>${noUsers.length ? noUsers.join(", ") : "-"}</td>
        <td class='p-2 font-semibold'>
          ${
            suggested
              ? `<span class='px-2 py-1 rounded bg-blue-100 text-blue-800'>${suggested}</span>`
              : "-"
          }
        </td>
        <td class='p-2'>
          ${
            canAssign
              ? `<select id='opt-${s.id}' class='border rounded p-1'>${options}</select>`
              : `<select class='border rounded p-1 bg-slate-100 text-slate-400' disabled><option>-</option></select>`
          }
        </td>
        <td class='p-2 whitespace-nowrap'>
          ${
            canAssign
              ? `<button class='px-2 py-1 rounded bg-blue-700 text-white mr-2' onclick="assignOptionalShift('${s.id}')">Auswahl einteilen</button>
                 <button class='px-2 py-1 rounded bg-emerald-700 text-white' onclick="assignSuggestedSpringer('${s.id}')">Vorschlag nehmen</button>`
              : `<span class='text-xs text-slate-500'>Nur bei Abwesenheit A/B/C</span>`
          }
        </td>
      </tr>`;
    })
    .join("");

  return `<div class='space-y-4'>
    <div class='border rounded-lg p-3 bg-white'>
      <div class='flex items-center justify-between gap-2 mb-2'>
        <h3 class='text-md font-semibold'>Wochenende</h3>
        ${helpButton("wochenende")}
      </div>
      <p class='text-sm text-slate-500 mb-2'>
        Springer melden „Kann“ oder „Kann nicht“. Der Admin teilt ein. Der Vorschlag bevorzugt verfügbare Springer mit den wenigsten Wochenend-Einsätzen.
      </p>
      <div class='overflow-auto max-h-[60vh]'>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100 sticky top-0'>
            <tr>
              <th class='p-2 text-left'>Datum</th>
              <th class='p-2 text-left'>Schicht</th>
              <th class='p-2 text-left'>A/B/C</th>
              <th class='p-2 text-left'>Status A/B/C</th>
              <th class='p-2 text-left'>Kann</th>
              <th class='p-2 text-left'>Kann nicht</th>
              <th class='p-2 text-left'>Vorschlag</th>
              <th class='p-2 text-left'>Auswahl</th>
              <th class='p-2 text-left'>Einteilen</th>
            </tr>
          </thead>
          <tbody>${optionalRows || '<tr><td class="p-2" colspan="9">Keine Wochenend-Schichten.</td></tr>'}</tbody>
        </table>
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
      <div class='flex items-center justify-between gap-2 mb-2'>
        <h3 class='text-md font-semibold'>Schichttausch</h3>
        ${helpButton("schichttausch")}
      </div>
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

function canAssignUserToShift(name, shiftId, silent = false) {
  const shift = getShiftById(shiftId);
  if (!shift) return true;

  const absType = getCalendarAbsenceType(name, shift.date);
  if (absType) {
    if (!silent) {
      alert(
        `${name} ist am ${formatDateDisplay(shift.date)} als ${absType} gemeldet und kann nicht eingesetzt werden.`,
      );
    }
    return false;
  }

  const restCheck = checkRestTimeRule(name, shiftId);
  if (!restCheck.ok) {
    if (!silent) {
      alert(restCheck.message);
    }
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

async function addVacation() {
  const data = readAdminCalendarForm();
  if (!data) return;

  const payload = {
    user_name: data.user,
    from_date: data.from,
    to_date: data.to,
    created_by_employee_id: currentEmployeeRecord?.id || null,
  };

  const { data: inserted, error } = await supabaseClient
    .from("planner_vacations")
    .insert(payload)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Speichern von Urlaub:", error);
    return alert(`Urlaub konnte nicht gespeichert werden: ${error.message}`);
  }

  state.vacations.push({
    id: inserted.id,
    user: inserted.user_name,
    from: inserted.from_date,
    to: inserted.to_date,
  });

  persist();
  render();
}

async function addSickLeave() {
  const data = readAdminCalendarForm();
  if (!data) return;

  const payload = {
    user_name: data.user,
    from_date: data.from,
    to_date: data.to,
    created_by_employee_id: currentEmployeeRecord?.id || null,
  };

  const { data: inserted, error } = await supabaseClient
    .from("planner_sick_leaves")
    .insert(payload)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Speichern von Krankmeldung:", error);
    return alert(
      `Krankmeldung konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  state.sickLeaves.push({
    id: inserted.id,
    user: inserted.user_name,
    from: inserted.from_date,
    to: inserted.to_date,
  });

  persist();
  render();
}

function getDatesBetween(start, end) {
  const dates = [];
  let current = new Date(start);
  const last = new Date(end);

  while (current <= last) {
    dates.push(current.toISOString().slice(0, 10));
    current.setDate(current.getDate() + 1);
  }

  return dates;
}

async function deleteVacation(id) {
  if (!supabaseReady) return;

  const entry = state.vacations.find((v) => v.id === id);
  if (!entry) return;

  const confirmDelete = confirm(`Urlaub von ${entry.user} wirklich löschen?`);
  if (!confirmDelete) return;

  const replacementError = await deleteReplacementPlanFromSupabase(
    "vacation",
    id,
  );
  if (replacementError) {
    console.error("Fehler beim Löschen der Ersatzplanung:", replacementError);
    return alert(
      `Urlaub konnte nicht gelöscht werden, weil die Ersatzplanung nicht bereinigt werden konnte: ${replacementError.message}`,
    );
  }

  // 🔴 Urlaub löschen
  const { error } = await supabaseClient
    .from("planner_vacations")
    .delete()
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Löschen Urlaub:", error);
    return alert(error.message);
  }

  // 🔴 Zugehörige Abwesenheiten löschen
  const dates = getDatesBetween(entry.from, entry.to);

  for (const d of dates) {
    await supabaseClient
      .from("planner_absences")
      .delete()
      .eq("absence_date", d)
      .eq("user_name", entry.user);

    delete state.absences[`${d}:${entry.user}`];
  }

  // 🔴 lokal entfernen
  clearReplacementPlanForSource("vacation", id);
  state.vacations = state.vacations.filter((v) => v.id !== id);

  persist();
  render();
}

async function deleteSickLeave(id) {
  if (!supabaseReady) return;

  const entry = state.sickLeaves.find((v) => v.id === id);
  if (!entry) return;

  const confirmDelete = confirm(
    `Krankmeldung von ${entry.user} wirklich löschen?`,
  );
  if (!confirmDelete) return;

  const replacementError = await deleteReplacementPlanFromSupabase("sick", id);
  if (replacementError) {
    console.error("Fehler beim Löschen der Ersatzplanung:", replacementError);
    return alert(
      `Krankmeldung konnte nicht gelöscht werden, weil die Ersatzplanung nicht bereinigt werden konnte: ${replacementError.message}`,
    );
  }

  const { error } = await supabaseClient
    .from("planner_sick_leaves")
    .delete()
    .eq("id", id);

  if (error) {
    console.error("Fehler beim Löschen Krankmeldung:", error);
    return alert(error.message);
  }

  const dates = getDatesBetween(entry.from, entry.to);

  for (const d of dates) {
    await supabaseClient
      .from("planner_absences")
      .delete()
      .eq("absence_date", d)
      .eq("user_name", entry.user);

    delete state.absences[`${d}:${entry.user}`];
  }

  clearReplacementPlanForSource("sick", id);
  state.sickLeaves = state.sickLeaves.filter((v) => v.id !== id);

  persist();
  render();
}

async function addSwap() {
  if (!supabaseReady) return;

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

  const { data: inserted, error } = await supabaseClient
    .from("planner_swaps")
    .insert({
      user_a: userA,
      user_b: userB,
      start_date: startDate,
      end_date: endDate,
    })
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Speichern des Schichttausches:", error);
    return alert(
      `Schichttausch konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  state.swaps.push({
    id: inserted.id,
    userA: inserted.user_a,
    userB: inserted.user_b,
    startDate: inserted.start_date,
    endDate: inserted.end_date || null,
  });

  persist();
  render();
}

async function deleteSwap(index) {
  if (!supabaseReady) return;
  if (index < 0 || index >= state.swaps.length) return;

  const swap = state.swaps[index];

  if (swap?.id) {
    const { error } = await supabaseClient
      .from("planner_swaps")
      .delete()
      .eq("id", swap.id);

    if (error) {
      console.error("Fehler beim Löschen des Schichttausches:", error);
      return alert(
        `Schichttausch konnte nicht gelöscht werden: ${error.message}`,
      );
    }
  }

  state.swaps = state.swaps.filter((_, i) => i !== index);
  persist();
  render();
}

async function resetActiveSwaps(silent = false) {
  if (!supabaseReady) return;

  const today = todayIso();
  const y = new Date(`${today}T00:00:00`);
  y.setDate(y.getDate() - 1);
  const yesterday = isoDate(y);

  const activeSwaps = state.swaps.filter((swap) => {
    if (!swap) return false;
    if (swap.startDate >= today) return true;
    return !swap.endDate || swap.endDate >= today;
  });

  const deleteIds = activeSwaps
    .filter((swap) => swap.startDate >= today && swap.id)
    .map((swap) => swap.id);
  const updateIds = activeSwaps
    .filter((swap) => swap.startDate < today && swap.id)
    .map((swap) => swap.id);

  const operations = [];
  if (deleteIds.length)
    operations.push(
      supabaseClient.from("planner_swaps").delete().in("id", deleteIds),
    );
  if (updateIds.length)
    operations.push(
      supabaseClient
        .from("planner_swaps")
        .update({ end_date: yesterday })
        .in("id", updateIds),
    );

  const results = await Promise.all(operations);
  const firstError = results.find((res) => res.error)?.error;

  if (firstError) {
    console.error("Fehler beim Zurücksetzen der Tausche:", firstError);
    return alert(
      `Aktive/zukünftige Tausche konnten nicht zurückgesetzt werden: ${firstError.message}`,
    );
  }

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

async function resetPlanCurrentFuture() {
  if (
    !confirm(
      "Gesamtplan wirklich für aktuelle und zukünftige Schichten auf Ursprung zurücksetzen?",
    )
  ) {
    return;
  }

  if (!supabaseReady) {
    return alert("Supabase ist nicht bereit. Zurücksetzen nicht möglich.");
  }

  const today = todayIso();

  const deletes = await Promise.all([
    supabaseClient
      .from("planner_assignments")
      .delete()
      .gte("shift_date", today),
    supabaseClient.from("planner_absences").delete().gte("absence_date", today),
    supabaseClient.from("planner_vacations").delete().gte("to_date", today),
    supabaseClient.from("planner_sick_leaves").delete().gte("to_date", today),
    supabaseClient
      .from("planner_availability")
      .delete()
      .gte("shift_date", today),
    supabaseClient
      .from("planner_swaps")
      .delete()
      .or(`end_date.is.null,end_date.gte.${today}`),
    supabaseClient
      .from("planner_shift_cancellations")
      .delete()
      .gte("shift_date", today),
    supabaseClient
      .from("planner_saturday_requests")
      .delete()
      .gte("shift_date", today),
    supabaseClient
      .from("planner_absence_replacements")
      .delete()
      .gte("shift_date", today),
  ]);

  const firstError = deletes.find((res) => res.error)?.error;
  if (firstError) {
    console.error("Fehler beim Zurücksetzen des Plans:", firstError);
    return alert(
      `Gesamtplan konnte nicht vollständig zurückgesetzt werden: ${firstError.message}`,
    );
  }

  state.assignments = {};
  state.shiftCancellations = {};
  state.absences = {};
  state.swaps = [];
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
  alert(
    "Gesamtplan wurde für aktuelle und zukünftige Schichten zurückgesetzt.",
  );
}

async function refreshEmployeesFromSupabase() {
  const employees = await loadEmployeesFromSupabase();
  applyEmployeesToState(employees);

  if (
    currentEmployeeRecord?.id &&
    state.employees?.[currentEmployeeRecord.id]
  ) {
    const refreshed = state.employees[currentEmployeeRecord.id];
    currentEmployeeRecord = {
      ...currentEmployeeRecord,
      display_name: refreshed.display_name,
      role: refreshed.role,
      employee_type: refreshed.employee_type,
      slot_code: refreshed.slot_code,
      is_active: refreshed.is_active,
      color_key: refreshed.color_key,
    };
    currentUser = {
      name: refreshed.display_name,
      role: refreshed.role,
    };
  }

  persist();
}

async function clearSlotInSupabase(slot, exceptEmployeeId = null) {
  if (!slot) return null;

  let query = supabaseClient
    .from("employees")
    .update({ slot_code: null })
    .eq("slot_code", slot);

  if (exceptEmployeeId) {
    query = query.neq("id", exceptEmployeeId);
  }

  const { error } = await query;
  return error || null;
}

async function updateSlotAssignment(slot) {
  const select = document.getElementById(`slot-${slot}`);
  if (!select) return;
  const selectedName = select.value;
  const selectedEmployee = activeUsers().find((u) => u.name === selectedName);

  if (supabaseReady && selectedEmployee?.id) {
    const clearError = await clearSlotInSupabase(slot, selectedEmployee.id);
    if (clearError) {
      console.error("Fehler beim Freigeben des Slots:", clearError);
      alert(clearError.message);
      return;
    }

    const { error } = await supabaseClient
      .from("employees")
      .update({ slot_code: slot })
      .eq("id", selectedEmployee.id);

    if (error) {
      console.error("Fehler beim Speichern der Slot-Zuordnung:", error);
      alert(error.message);
      return;
    }

    await refreshEmployeesFromSupabase();
  }

  state.slotAssignments[slot] = select.value;
  persist();
  render();
}

async function saveAllPersonnelChanges() {
  if (!supabaseReady) return;

  const ui = ensurePersonnelPendingState();
  const pendingEmployees = { ...ui.pendingEmployeeEdits };
  const pendingSlots = { ...ui.pendingSlotAssignments };
  const originalUsers = allUsers();
  const originalNameToId = {};
  originalUsers.forEach((user) => {
    if (user?.name && user?.id) originalNameToId[user.name] = user.id;
  });

  const failedEmployeeEdits = {};
  const failedSlotAssignments = {};
  const errors = [];

  for (const [id, changes] of Object.entries(pendingEmployees)) {
    const payload = {};

    if (Object.prototype.hasOwnProperty.call(changes, "display_name")) {
      const displayName = String(changes.display_name || "").trim();
      if (!displayName) {
        failedEmployeeEdits[id] = changes;
        errors.push("Ein Mitarbeitername fehlt.");
        continue;
      }
      payload.display_name = displayName;
    }

    if (Object.prototype.hasOwnProperty.call(changes, "employee_type")) {
      payload.employee_type = changes.employee_type || "springer";
    }
    if (Object.prototype.hasOwnProperty.call(changes, "color_key")) {
      payload.color_key = changes.color_key || "gray";
    }
    if (Object.prototype.hasOwnProperty.call(changes, "role")) {
      payload.role = changes.role || "employee";
    }
    if (Object.prototype.hasOwnProperty.call(changes, "is_active")) {
      payload.is_active = !!changes.is_active;
      if (!payload.is_active) payload.slot_code = null;
    }

    if (!Object.keys(payload).length) continue;

    const { error } = await supabaseClient
      .from("employees")
      .update(payload)
      .eq("id", id);

    if (error) {
      console.error("Fehler beim Speichern des Mitarbeiters:", error);
      failedEmployeeEdits[id] = changes;
      errors.push(error.message);
    }
  }

  await refreshEmployeesFromSupabase();

  for (const [slot, selectedName] of Object.entries(pendingSlots)) {
    const originalId = originalNameToId[selectedName];
    const selectedEmployee =
      activeUsers().find((u) => u.name === selectedName) ||
      activeUsers().find((u) => u.id === originalId);

    if (!selectedEmployee?.id) {
      failedSlotAssignments[slot] = selectedName;
      errors.push(`Kein aktiver Mitarbeiter für Slot ${slot} gefunden.`);
      continue;
    }

    const clearError = await clearSlotInSupabase(slot, selectedEmployee.id);
    if (clearError) {
      console.error("Fehler beim Freigeben des Slots:", clearError);
      failedSlotAssignments[slot] = selectedName;
      errors.push(clearError.message);
      continue;
    }

    const { error } = await supabaseClient
      .from("employees")
      .update({ slot_code: slot })
      .eq("id", selectedEmployee.id);

    if (error) {
      console.error("Fehler beim Speichern der Slot-Zuordnung:", error);
      failedSlotAssignments[slot] = selectedName;
      errors.push(error.message);
      continue;
    }

    state.slotAssignments[slot] = selectedEmployee.name;
  }

  ui.pendingEmployeeEdits = failedEmployeeEdits;
  ui.pendingSlotAssignments = failedSlotAssignments;

  await refreshEmployeesFromSupabase();
  persist();
  render();

  if (errors.length) {
    alert(`Einige Änderungen konnten nicht gespeichert werden:\n${errors.join("\n")}`);
    return;
  }

  alert("Alle Änderungen wurden gespeichert.");
}

async function addEmployee() {
  if (!supabaseReady) return;

  const name = document.getElementById("newEmployeeName")?.value?.trim() || "";
  const type = document.getElementById("newEmployeeType")?.value || "springer";
  const slot = document.getElementById("newEmployeeSlot")?.value || null;
  const colorKey = document.getElementById("newEmployeeColor")?.value || "gray";
  const role = document.getElementById("newEmployeeRole")?.value || "employee";

  if (!name) {
    alert("Name fehlt");
    return;
  }

  if (slot) {
    const clearError = await clearSlotInSupabase(slot);
    if (clearError) {
      console.error("Fehler beim Freigeben des Slots:", clearError);
      alert(clearError.message);
      return;
    }
  }

  const { error } = await supabaseClient.from("employees").insert([
    {
      display_name: name,
      role,
      employee_type: type,
      slot_code: slot,
      is_active: true,
      color_key: colorKey,
    },
  ]);

  if (error) {
    console.error(error);
    alert(error.message);
    return;
  }

  document.getElementById("newEmployeeName").value = "";
  document.getElementById("newEmployeeSlot").value = "";
  await refreshEmployeesFromSupabase();
  render();
}

function fillEmployeeForm(emp) {
  if (!emp?.id) return;

  const nameInput = document.getElementById(`employee-name-${emp.id}`);
  const typeInput = document.getElementById(`employee-type-${emp.id}`);
  const colorInput = document.getElementById(`employee-color-${emp.id}`);
  const roleInput = document.getElementById(`employee-role-${emp.id}`);

  if (nameInput) nameInput.value = emp.display_name || emp.name || "";
  if (typeInput) typeInput.value = emp.employee_type || emp.type || "springer";
  if (colorInput) colorInput.value = emp.color_key || "gray";
  if (roleInput) roleInput.value = emp.role || "employee";
}

async function updateEmployee(id) {
  if (!supabaseReady) return;

  const name =
    document.getElementById(`employee-name-${id}`)?.value?.trim() || "";
  const type =
    document.getElementById(`employee-type-${id}`)?.value || "springer";
  const colorKey =
    document.getElementById(`employee-color-${id}`)?.value || "gray";
  const role =
    document.getElementById(`employee-role-${id}`)?.value || "employee";

  if (!name) {
    alert("Name fehlt");
    return;
  }

  const { error } = await supabaseClient
    .from("employees")
    .update({
      display_name: name,
      employee_type: type,
      color_key: colorKey,
      role,
    })
    .eq("id", id);

  if (error) {
    console.error(error);
    alert(error.message);
    return;
  }

  await refreshEmployeesFromSupabase();
  render();
}

async function deactivateEmployee(id) {
  const employee = state.employees?.[id];
  const name = employee?.display_name || employee?.name || "Mitarbeiter";
  if (!confirm(`${name} wirklich als inaktiv markieren?`)) return;
  if (supabaseReady && id) {
    const { error } = await supabaseClient
      .from("employees")
      .update({ is_active: false, slot_code: null })
      .eq("id", id);

    if (error) {
      console.error("Fehler beim Deaktivieren des Mitarbeiters:", error);
      alert(error.message);
      return;
    }
  }

  state.inactiveUsers[name] = true;
  Object.keys(state.slotAssignments).forEach((slot) => {
    if (state.slotAssignments[slot] === name)
      state.slotAssignments[slot] = DEFAULT_SLOT_ASSIGNMENTS[slot] || null;
  });
  if (supabaseReady) await refreshEmployeesFromSupabase();
  persist();
  render();
}

async function activateEmployee(id) {
  const employee = state.employees?.[id];
  const name = employee?.display_name || employee?.name || id;
  if (supabaseReady && id) {
    const { error } = await supabaseClient
      .from("employees")
      .update({ is_active: true })
      .eq("id", id);

    if (error) {
      console.error("Fehler beim Aktivieren des Mitarbeiters:", error);
      alert(error.message);
      return;
    }
  }

  delete state.inactiveUsers[name];
  if (supabaseReady) await refreshEmployeesFromSupabase();
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

async function addToolMaterial() {
  if (currentUser.role !== "admin") return;

  const name = prompt("Neuen Schneidwerkstoff eingeben:");
  if (!name) return;

  const cleanName = name.trim();
  if (!cleanName) return;

  if (
    toolMaterials.some((m) => m.name.toLowerCase() === cleanName.toLowerCase())
  ) {
    return alert("Dieser Schneidwerkstoff existiert bereits.");
  }

  const nextSort =
    toolMaterials.reduce(
      (max, m) => Math.max(max, Number(m.sort_order || 0)),
      0,
    ) + 10;

  const { data, error } = await supabaseClient
    .from("tool_materials")
    .insert({
      name: cleanName,
      sort_order: nextSort,
      is_active: true,
    })
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Anlegen des Schneidwerkstoffs:", error);
    return alert(
      `Schneidwerkstoff konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  toolMaterials.push(data);
  toolMaterials.sort(
    (a, b) =>
      Number(a.sort_order || 0) - Number(b.sort_order || 0) ||
      String(a.name || "").localeCompare(String(b.name || "")),
  );

  render();
}

async function renameToolMaterial(materialId) {
  if (currentUser.role !== "admin") return;

  const material = toolMaterials.find((m) => m.id === materialId);
  if (!material) return;

  const nextName = prompt("Schneidwerkstoff umbenennen:", material.name || "");
  if (!nextName) return;

  const cleanName = nextName.trim();
  if (!cleanName) return;

  if (
    toolMaterials.some(
      (m) =>
        m.id !== materialId &&
        String(m.name || "").toLowerCase() === cleanName.toLowerCase(),
    )
  ) {
    return alert("Dieser Schneidwerkstoff existiert bereits.");
  }

  const { data, error } = await supabaseClient
    .from("tool_materials")
    .update({
      name: cleanName,
      updated_at: new Date().toISOString(),
    })
    .eq("id", materialId)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Umbenennen des Schneidwerkstoffs:", error);
    return alert(
      `Schneidwerkstoff konnte nicht umbenannt werden: ${error.message}`,
    );
  }

  const index = toolMaterials.findIndex((m) => m.id === materialId);
  if (index !== -1) {
    toolMaterials[index] = data;
  }

  toolMaterials.sort(
    (a, b) =>
      Number(a.sort_order || 0) - Number(b.sort_order || 0) ||
      String(a.name || "").localeCompare(String(b.name || "")),
  );

  render();
}

async function deactivateToolMaterial(materialId) {
  if (currentUser.role !== "admin") return;

  const material = toolMaterials.find((m) => m.id === materialId);
  if (!material) return;

  const linkedTools = state.tools.filter((t) => t.materialId === materialId);
  const message = linkedTools.length
    ? `Schneidwerkstoff "${material.name}" wirklich deaktivieren?\n\nEr ist noch bei ${linkedTools.length} Werkzeug(en) hinterlegt. Bestehende Werkzeuge behalten den Wert, aber der Werkstoff ist künftig nicht mehr auswählbar.`
    : `Schneidwerkstoff "${material.name}" wirklich deaktivieren?`;

  const ok = confirm(message);
  if (!ok) return;

  const { error } = await supabaseClient
    .from("tool_materials")
    .update({
      is_active: false,
      updated_at: new Date().toISOString(),
    })
    .eq("id", materialId);

  if (error) {
    console.error("Fehler beim Deaktivieren des Schneidwerkstoffs:", error);
    return alert(
      `Schneidwerkstoff konnte nicht deaktiviert werden: ${error.message}`,
    );
  }

  toolMaterials = toolMaterials.filter((m) => m.id !== materialId);
  render();
}

function renderToolMaterialsAdmin() {
  const rows = toolMaterials
    .map(
      (m) => `<tr class='border-b'>
        <td class='p-2'>${m.name || "-"}</td>
        <td class='p-2'>${Number(m.sort_order || 0)}</td>
        <td class='p-2 whitespace-nowrap'>
          <button class='px-2 py-1 rounded bg-amber-600 text-white mr-2' onclick="renameToolMaterial('${m.id}')">Bearbeiten</button>
          <button class='px-2 py-1 rounded bg-rose-700 text-white' onclick="deactivateToolMaterial('${m.id}')">Deaktivieren</button>
        </td>
      </tr>`,
    )
    .join("");

  return `<div class='border-2 border-slate-300 rounded-xl p-3 bg-slate-50'>
    <div class='flex items-center justify-between gap-3 flex-wrap mb-3'>
      <div>
        <h3 class='text-lg font-bold mb-1'>Schneidwerkstoffe</h3>
        <p class='text-sm text-slate-500'>Admin kann Schneidwerkstoffe anlegen, umbenennen und deaktivieren.</p>
      </div>
      <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolMaterial()'>Schneidwerkstoff hinzufügen</button>
    </div>

    <div class='overflow-auto border rounded-lg bg-white max-h-[30vh]'>
      <table class='w-full text-sm'>
        <thead class='bg-slate-100 sticky top-0'>
          <tr>
            <th class='p-2 text-left'>Name</th>
            <th class='p-2 text-left'>Sortierung</th>
            <th class='p-2'></th>
          </tr>
        </thead>
        <tbody>${rows || '<tr><td class="p-2" colspan="3">Keine Schneidwerkstoffe vorhanden.</td></tr>'}</tbody>
      </table>
    </div>
  </div>`;
}

function renderToolMasterDataAdmin() {
  const labelRows = getToolLabels()
    .map(
      (name) => `<tr class='border-b'>
        <td class='p-2'>${name}</td>
        <td class='p-2 text-slate-500'>Bezeichnung</td>
      </tr>`,
    )
    .join("");

  const manufacturerRows = getToolManufacturers()
    .map(
      (name) => `<tr class='border-b'>
        <td class='p-2'>${name}</td>
        <td class='p-2 text-slate-500'>Hersteller</td>
      </tr>`,
    )
    .join("");

  const materialRows = toolMaterials
    .map(
      (m) => `<tr class='border-b'>
        <td class='p-2'>${m.name || "-"}</td>
        <td class='p-2 text-slate-500'>Schneidwerkstoff</td>
      </tr>`,
    )
    .join("");

  return `<div class='border-2 border-slate-300 rounded-xl p-3 bg-slate-50'>
    <div class='flex items-center justify-between gap-3 flex-wrap mb-3'>
      <div>
        <h3 class='text-lg font-bold mb-1'>Stammdaten für Werkzeuge</h3>
        <p class='text-sm text-slate-500'>Hier können Bezeichnungen, Hersteller und Schneidwerkstoffe ergänzt und eingesehen werden.</p>
      </div>
      <div class='flex gap-2 flex-wrap'>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolLabel()'>Bezeichnung hinzufügen</button>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolManufacturer()'>Hersteller hinzufügen</button>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='addToolMaterial()'>Schneidwerkstoff hinzufügen</button>
      </div>
    </div>

    <div class='grid md:grid-cols-3 gap-4'>
      <div class='border rounded-lg bg-white overflow-auto max-h-[28vh]'>
        <div class='p-2 font-semibold border-b bg-slate-100'>Bezeichnungen</div>
        <table class='w-full text-sm'>
          <tbody>${labelRows || '<tr><td class="p-2">Keine Bezeichnungen vorhanden.</td></tr>'}</tbody>
        </table>
      </div>

      <div class='border rounded-lg bg-white overflow-auto max-h-[28vh]'>
        <div class='p-2 font-semibold border-b bg-slate-100'>Hersteller</div>
        <table class='w-full text-sm'>
          <tbody>${manufacturerRows || '<tr><td class="p-2">Keine Hersteller vorhanden.</td></tr>'}</tbody>
        </table>
      </div>

      <div class='border rounded-lg bg-white overflow-auto max-h-[28vh]'>
        <div class='p-2 font-semibold border-b bg-slate-100'>Schneidwerkstoffe</div>
        <table class='w-full text-sm'>
          <tbody>${materialRows || '<tr><td class="p-2">Keine Schneidwerkstoffe vorhanden.</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  </div>`;
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
      isThread &&
      labelEl?.value &&
      document.getElementById(`${prefix}ThreadPrefix`)?.value === "MF"
        ? ""
        : "none";
  if (cornerRadiusWrap) cornerRadiusWrap.style.display = isRadius ? "" : "none";
}

function updateThreadPitchVisibility(prefix = "tool") {
  const label = document.getElementById(`${prefix}Label`)?.value || "";
  const threadPrefix =
    document.getElementById(`${prefix}ThreadPrefix`)?.value || "";
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
    materialId: root.getElementById("toolMaterial")?.value || "",
    shelf: root.getElementById("toolShelf")?.value?.trim().toUpperCase(),
    articleNo: root.getElementById("toolArticle")?.value?.trim(),
    holder: root.getElementById("toolHolder")?.value,
    stock: Number(root.getElementById("toolStock")?.value || 0),
    minStock: Number(root.getElementById("toolMinStock")?.value || 0),
    optimalStock: Number(root.getElementById("toolOptimalStock")?.value || 0),
    manufacturer: root.getElementById("toolManufacturer")?.value,
    insertTool: !!root.getElementById("toolInsertTool")?.checked,
    insertEdges: Number(root.getElementById("toolInsertEdges")?.value || 0),
    insertRadius:
      root.getElementById("toolInsertRadius")?.value?.trim() || "",
  };
}

function isValidToolShelf(value) {
  return /^\d{2}[A-Z]$/.test(String(value || "").trim().toUpperCase());
}

function validateToolData(data) {
  const isThreadTool = isThreadToolLabel(data.label);
  const isRadiusTool = isRadiusToolLabel(data.label);

  if (
    !data.tNumber ||
    !data.label ||
    !data.diameter ||
    !isValidToolShelf(data.shelf) ||
    !data.articleNo ||
    !["HSK 100", "HSK 63"].includes(data.holder)
  ) {
    return {
      ok: false,
      message:
        "Bitte Felder korrekt ausfüllen (2-stellige Zahl + Buchstabe für Fach, z. B. 26C oder 02T; Aufnahme HSK 100 oder HSK 63).",
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
    materialId: data.materialId || null,
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
    insertRadius: data.insertTool ? data.insertRadius || "" : "",
  };
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
  const materialOptions = toolMaterials
    .map((m) => `<option value="${m.id}">${m.name}</option>`)
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

    <select id='${prefix}Material' class='border rounded p-2'>
      <option value=''>Schneidwerkstoff wählen</option>
      ${materialOptions}
    </select>

    <input id='${prefix}Shelf' class='border rounded p-2' placeholder='00A' />
    <input id='${prefix}Article' class='border rounded p-2' placeholder='Artikel Nr.' />
    <select id='${prefix}Holder' class='border rounded p-2'>${holderOptions}</select>
    <input id='${prefix}Stock' type='number' class='border rounded p-2' placeholder='Bestand' />
    <input id='${prefix}MinStock' type='number' class='border rounded p-2' placeholder='Mindestbestand' />
    <input id='${prefix}OptimalStock' type='number' class='border rounded p-2' placeholder='Optimale Stückzahl' />
    <select id='${prefix}Manufacturer' class='border rounded p-2'>${manufacturerOptions}</select>
    <label class='flex items-center gap-2 text-sm md:col-span-2'>
      <input id='${prefix}InsertTool' type='checkbox' onchange='toggleInsertToolFieldsById("${prefix}InsertTool","${prefix}InsertEdges","${prefix}InsertRadius")' />
      Wendeplattenwerkzeug
    </label>
    <input id='${prefix}InsertEdges' type='number' class='border rounded p-2 md:col-span-2' placeholder='Anzahl Schneiden' disabled />
    <div id='${prefix}InsertRadiusWrap' class='md:col-span-2' style='display:none;'>
      <input id='${prefix}InsertRadius' class='border rounded p-2 w-full' placeholder='Plattenradius optional, z. B. 0.8' disabled />
    </div>
  </div>`;
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
    material_id: normalized.materialId || null,
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
    insert_radius: normalized.insertRadius || null,
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

function normalizeToolImageNumber(tNumber) {
  return String(tNumber || "")
    .trim()
    .replace(/^T\s*/i, "")
    .trim();
}

function getToolImagePath(tool) {
  const holderFolder =
    tool.holder === "HSK 100"
      ? "HSK100"
      : tool.holder === "HSK 63"
        ? "HSK63"
        : "";

  const imageNumber = normalizeToolImageNumber(tool.tNumber);

  if (!holderFolder || !imageNumber) return "";

  return `img/Depo/${holderFolder}/${imageNumber}.PNG`;
}

function openToolImagePopup(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const imagePath = getToolImagePath(tool);
  if (!imagePath) {
    alert("Für dieses Werkzeug konnte kein Bildpfad erstellt werden.");
    return;
  }

  const imageNumber = normalizeToolImageNumber(tool.tNumber);
  const host = getModalHost();

  host.innerHTML = `
    <div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-3xl max-h-[90vh] overflow-auto p-4">
        <div class="flex items-center justify-between gap-3 mb-3">
          <div>
            <h3 class="text-lg font-bold">Werkzeugbild – T ${escapeHtml(imageNumber)}</h3>
            <p class="text-sm text-slate-500">${escapeHtml(tool.label || "-")} · ${escapeHtml(tool.holder || "-")}</p>
            <p class="text-xs text-slate-400">${escapeHtml(imagePath)}</p>
          </div>
          <button class="px-3 py-1 rounded bg-slate-200" onclick="closeToolImagePopup()">Schließen</button>
        </div>

        <div class="border rounded-lg bg-slate-50 p-3 flex justify-center">
          <img
            src="${imagePath}"
            alt="Werkzeugbild T ${escapeHtml(imageNumber)}"
            class="max-w-full max-h-[70vh] object-contain rounded"
            onerror="this.outerHTML='<div class=&quot;text-sm text-rose-700 bg-rose-50 border border-rose-200 rounded p-3 w-full&quot;>Kein Bild gefunden: ${imagePath}</div>'"
          />
        </div>
      </div>
    </div>
  `;
}

function closeToolImagePopup() {
  getModalHost().innerHTML = "";
}

function renderSimpleQrSvg(text) {
  try {
    const size = 29;
    const quiet = 4;
    const scale = 6;
    const dataCodewords = 55;
    const ecCodewords = 15;
    const modules = Array.from({ length: size }, () => Array(size).fill(false));
    const reserved = Array.from({ length: size }, () => Array(size).fill(false));

    const setModule = (x, y, dark, reserve = true) => {
      if (x < 0 || y < 0 || x >= size || y >= size) return;
      modules[y][x] = !!dark;
      if (reserve) reserved[y][x] = true;
    };

    const addFinder = (x, y) => {
      for (let dy = -1; dy <= 7; dy++) {
        for (let dx = -1; dx <= 7; dx++) {
          const xx = x + dx;
          const yy = y + dy;
          if (xx < 0 || yy < 0 || xx >= size || yy >= size) continue;
          const inFinder = dx >= 0 && dx <= 6 && dy >= 0 && dy <= 6;
          const dark =
            inFinder &&
            (dx === 0 ||
              dx === 6 ||
              dy === 0 ||
              dy === 6 ||
              (dx >= 2 && dx <= 4 && dy >= 2 && dy <= 4));
          setModule(xx, yy, dark);
        }
      }
    };

    addFinder(0, 0);
    addFinder(size - 7, 0);
    addFinder(0, size - 7);

    for (let i = 8; i < size - 8; i++) {
      setModule(i, 6, i % 2 === 0);
      setModule(6, i, i % 2 === 0);
    }

    for (let dy = -2; dy <= 2; dy++) {
      for (let dx = -2; dx <= 2; dx++) {
        const dist = Math.max(Math.abs(dx), Math.abs(dy));
        setModule(22 + dx, 22 + dy, dist === 2 || dist === 0);
      }
    }

    for (let i = 0; i <= 8; i++) {
      reserved[8][i] = true;
      reserved[i][8] = true;
    }
    for (let i = 0; i < 8; i++) {
      reserved[8][size - 1 - i] = true;
      reserved[size - 1 - i][8] = true;
    }
    setModule(8, size - 8, true);

    const bytes = String(text || "")
      .split("")
      .map((ch) => ch.charCodeAt(0));
    if (!bytes.length || bytes.some((byte) => byte > 255)) {
      throw new Error("QR unterstützt nur kurzen ASCII-Text.");
    }
    if (bytes.length > 53) {
      throw new Error("QR-Text ist für diese lokale Minimalversion zu lang.");
    }

    const bits = [];
    const appendBits = (value, length) => {
      for (let i = length - 1; i >= 0; i--) bits.push((value >>> i) & 1);
    };

    appendBits(0b0100, 4);
    appendBits(bytes.length, 8);
    bytes.forEach((byte) => appendBits(byte, 8));

    const capacityBits = dataCodewords * 8;
    const terminator = Math.min(4, capacityBits - bits.length);
    appendBits(0, terminator);
    while (bits.length % 8) bits.push(0);

    const data = [];
    for (let i = 0; i < bits.length; i += 8) {
      data.push(Number.parseInt(bits.slice(i, i + 8).join(""), 2));
    }
    for (let pad = 0xec; data.length < dataCodewords; pad ^= 0xfd) {
      data.push(pad);
    }

    const gfExp = Array(512).fill(0);
    const gfLog = Array(256).fill(0);
    let value = 1;
    for (let i = 0; i < 255; i++) {
      gfExp[i] = value;
      gfLog[value] = i;
      value <<= 1;
      if (value & 0x100) value ^= 0x11d;
    }
    for (let i = 255; i < 512; i++) gfExp[i] = gfExp[i - 255];

    const gfMul = (a, b) => {
      if (!a || !b) return 0;
      return gfExp[gfLog[a] + gfLog[b]];
    };

    const polyMul = (a, b) => {
      const result = Array(a.length + b.length - 1).fill(0);
      a.forEach((av, i) => {
        b.forEach((bv, j) => {
          result[i + j] ^= gfMul(av, bv);
        });
      });
      return result;
    };

    let generator = [1];
    for (let i = 0; i < ecCodewords; i++) {
      generator = polyMul(generator, [1, gfExp[i]]);
    }

    const ec = Array(ecCodewords).fill(0);
    data.forEach((byte) => {
      const factor = byte ^ ec.shift();
      ec.push(0);
      for (let i = 0; i < ecCodewords; i++) {
        ec[i] ^= gfMul(generator[i + 1], factor);
      }
    });

    const codewordBits = [...data, ...ec].flatMap((byte) => {
      const out = [];
      for (let i = 7; i >= 0; i--) out.push((byte >>> i) & 1);
      return out;
    });

    let bitIndex = 0;
    let upward = true;
    for (let right = size - 1; right >= 1; right -= 2) {
      if (right === 6) right--;
      for (let vert = 0; vert < size; vert++) {
        const y = upward ? size - 1 - vert : vert;
        for (let j = 0; j < 2; j++) {
          const x = right - j;
          if (reserved[y][x]) continue;
          const bit = bitIndex < codewordBits.length ? codewordBits[bitIndex++] : 0;
          const mask = (x + y) % 2 === 0;
          setModule(x, y, bit ^ mask, false);
        }
      }
      upward = !upward;
    }

    const formatData = (0b01 << 3) | 0;
    let formatRemainder = formatData << 10;
    for (let i = 14; i >= 10; i--) {
      if ((formatRemainder >>> i) & 1) {
        formatRemainder ^= 0x537 << (i - 10);
      }
    }
    const formatBits = ((formatData << 10) | formatRemainder) ^ 0x5412;
    const formatBit = (i) => ((formatBits >>> i) & 1) === 1;

    for (let i = 0; i <= 5; i++) setModule(8, i, formatBit(i));
    setModule(8, 7, formatBit(6));
    setModule(8, 8, formatBit(7));
    setModule(7, 8, formatBit(8));
    for (let i = 9; i < 15; i++) setModule(14 - i, 8, formatBit(i));
    for (let i = 0; i < 8; i++) setModule(size - 1 - i, 8, formatBit(i));
    for (let i = 8; i < 15; i++) setModule(8, size - 15 + i, formatBit(i));
    setModule(8, size - 8, true);

    const viewSize = (size + quiet * 2) * scale;
    const rects = [];
    modules.forEach((row, y) => {
      row.forEach((dark, x) => {
        if (!dark) return;
        rects.push(
          `<rect x="${(x + quiet) * scale}" y="${(y + quiet) * scale}" width="${scale}" height="${scale}"/>`,
        );
      });
    });

    return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${viewSize} ${viewSize}" width="150" height="150" role="img" aria-label="QR-Code">
      <rect width="100%" height="100%" fill="#fff"/>
      <g fill="#000">${rects.join("")}</g>
    </svg>`;
  } catch (error) {
    console.error("QR-Code konnte nicht erzeugt werden:", error);
    return `<div class="text-sm text-rose-700 bg-rose-50 border border-rose-200 rounded p-3">[QR-Code konnte nicht erzeugt werden]</div>`;
  }
}

function getToolQrPayload(tool) {
  return `TOOL:${tool.id}`;
}

function openToolQrPopup(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const payload = getToolQrPayload(tool);
  const host = getModalHost();
  const toolTitle = `T ${tool.tNumber}`;
  const qrSvg = renderSimpleQrSvg(payload);
  const diameter = tool.diameter || "-";
  const cornerRadius = tool.cornerRadius ? `R ${tool.cornerRadius}` : "-";
  const manufacturer = tool.manufacturer || "-";
  const articleNo = tool.articleNo || "-";
  const shelf = tool.shelf || "-";
  const insertRadiusLine =
    tool.insertTool && tool.insertRadius
      ? `<div class="text-xs text-slate-600">Plattenradius: R ${escapeHtml(tool.insertRadius)}</div>`
      : "";

  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-3">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-2xl max-h-[90vh] overflow-auto p-3">
      <div class="flex items-center justify-between gap-3 mb-2">
        <div>
          <h3 class="text-lg font-bold">QR-Code für ${escapeHtml(toolTitle)}</h3>
          <p class="text-sm text-slate-500">${escapeHtml(tool.label || "-")} · ${escapeHtml(tool.holder || "-")} · Fach ${escapeHtml(tool.shelf || "-")}</p>
        </div>
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeToolImagePopup()">Schließen</button>
      </div>
      <div class="grid md:grid-cols-[1fr,190px] gap-3 mb-3">
        <div class="border rounded-lg bg-slate-50 p-2 space-y-1">
          <div><span class="text-xs text-slate-500">T-Nummer</span><div class="font-semibold">${escapeHtml(toolTitle)}</div></div>
          <div><span class="text-xs text-slate-500">Bezeichnung</span><div>${escapeHtml(tool.label || "-")}</div></div>
          <div><span class="text-xs text-slate-500">Aufnahme</span><div>${escapeHtml(tool.holder || "-")}</div></div>
          <div><span class="text-xs text-slate-500">Lagerfach</span><div>${escapeHtml(tool.shelf || "-")}</div></div>
          <div><span class="text-xs text-slate-500">QR-Inhalt</span><div class="font-mono text-xs break-all bg-white border rounded p-1 mt-1">${escapeHtml(payload)}</div></div>
        </div>
        <div class="border rounded-lg bg-white p-2 flex items-center justify-center text-center min-h-[170px]">
          ${qrSvg}
        </div>
      </div>
      <div class="border rounded-lg bg-white p-3 mb-3">
        <div class="text-xs uppercase tracking-wide text-slate-500 mb-1">Druckbereich</div>
        <div class="border rounded-lg p-3 text-center">
          <div class="text-2xl font-bold">${escapeHtml(toolTitle)}</div>
          <div class="text-sm mt-1">${escapeHtml(tool.label || "-")}</div>
          <div class="text-sm text-slate-700 mt-1">Ø ${escapeHtml(diameter)} · ${escapeHtml(cornerRadius)}</div>
          ${insertRadiusLine}
          <div class="text-xs text-slate-600">Hersteller: ${escapeHtml(manufacturer)}</div>
          <div class="text-xs text-slate-600">Artikel: ${escapeHtml(articleNo)}</div>
          <div class="text-xs text-slate-600">Fach: ${escapeHtml(shelf)}</div>
          <div class="mt-2 flex justify-center">${qrSvg}</div>
          <div class="font-mono text-[10px] break-all mt-2">QR-Inhalt: ${escapeHtml(payload)}</div>
        </div>
      </div>
      <div class="flex justify-end gap-2">
        <button class="px-3 py-2 rounded bg-slate-200" onclick="printToolQrLabel('${tool.id}')">Drucken</button>
        <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="copyToolQrPayload('${tool.id}')">Text kopieren</button>
      </div>
    </div>
  </div>`;
}

function printToolQrLabel(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const payload = getToolQrPayload(tool);
  const qrSvg = renderSimpleQrSvg(payload);
  const diameter = tool.diameter || "-";
  const cornerRadius = tool.cornerRadius ? `R ${tool.cornerRadius}` : "-";
  const manufacturer = tool.manufacturer || "-";
  const articleNo = tool.articleNo || "-";
  const shelf = tool.shelf || "-";
  const insertRadiusLine =
    tool.insertTool && tool.insertRadius
      ? `<div class="meta">Plattenradius: R ${escapeHtml(tool.insertRadius)}</div>`
      : "";
  const printWindow = window.open("", "_blank", "width=420,height=560");
  if (!printWindow) {
    alert("Druckfenster konnte nicht geöffnet werden.");
    return;
  }

  printWindow.document.write(`<!doctype html>
  <html>
    <head>
      <title>QR-Etikett T ${escapeHtml(tool.tNumber)}</title>
      <style>
        body { font-family: Arial, sans-serif; margin: 24px; color: #0f172a; }
        .label { border: 2px solid #0f172a; border-radius: 12px; padding: 16px; text-align: center; max-width: 340px; }
        .tnumber { font-size: 36px; font-weight: 800; margin: 0 0 3px; }
        .name { font-size: 16px; margin-bottom: 4px; }
        .meta { font-size: 12px; color: #475569; line-height: 1.25; }
        .qr-wrap { display: flex; justify-content: center; margin: 10px 0 8px; }
        .payload { font-family: monospace; font-size: 10px; word-break: break-all; }
        svg { width: 150px; height: 150px; }
      </style>
    </head>
    <body>
      <div class="label">
        <div class="tnumber">T ${escapeHtml(tool.tNumber)}</div>
        <div class="name">${escapeHtml(tool.label || "-")}</div>
        <div class="meta">Ø ${escapeHtml(diameter)} · ${escapeHtml(cornerRadius)}</div>
        ${insertRadiusLine}
        <div class="meta">Hersteller: ${escapeHtml(manufacturer)}</div>
        <div class="meta">Artikel: ${escapeHtml(articleNo)}</div>
        <div class="meta">Fach: ${escapeHtml(shelf)}</div>
        <div class="qr-wrap">${qrSvg}</div>
        <div class="payload">QR-Inhalt: ${escapeHtml(payload)}</div>
      </div>
      <script>
        window.onload = function() {
          window.print();
        };
      </script>
    </body>
  </html>`);
  printWindow.document.close();
}

async function copyToolQrPayload(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) return;

  const payload = getToolQrPayload(tool);
  if (!navigator.clipboard?.writeText) {
    alert(payload);
    return;
  }

  await navigator.clipboard.writeText(payload);
  alert("QR-Inhalt wurde kopiert.");
}

function renderToolScanner() {
  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div>
      <h2 class='text-2xl font-bold'>Werkzeug-Scanner</h2>
      <p class='text-sm text-slate-500 mt-1'>Für Werkzeugbereich / Handy-Terminal</p>
    </div>
    <div class='grid md:grid-cols-2 gap-4'>
      <div class='border rounded-xl p-4 bg-slate-50 space-y-4'>
        <div>
          <h3 class='text-xl font-bold'>Werkzeug entnehmen</h3>
          <p class='text-sm text-slate-500 mt-1'>Werkzeugbestand um 1 reduzieren.</p>
        </div>
        <div class='grid gap-3'>
          <button class='px-4 py-4 rounded-lg bg-emerald-700 text-white text-lg font-semibold' onclick='openManualToolWithdraw()'>Manuell entnehmen</button>
          <button class='px-4 py-4 rounded-lg bg-slate-800 text-white text-lg font-semibold' onclick="openQrScannerPlaceholder('withdraw')">QR-Code scannen</button>
        </div>
      </div>
      <div class='border rounded-xl p-4 bg-slate-50 space-y-4'>
        <div>
          <h3 class='text-xl font-bold'>Werkzeug einlagern</h3>
          <p class='text-sm text-slate-500 mt-1'>Werkzeugbestand um eine gewählte Menge erhöhen.</p>
        </div>
        <div class='grid gap-3'>
          <button class='px-4 py-4 rounded-lg bg-blue-700 text-white text-lg font-semibold' onclick='openManualToolRestock()'>Manuell einlagern</button>
          <button class='px-4 py-4 rounded-lg bg-slate-800 text-white text-lg font-semibold' onclick="openQrScannerPlaceholder('restock')">QR-Code scannen</button>
        </div>
      </div>
    </div>
  </div>`;
}

function findToolByTNumberInput(value) {
  const normalized = normalizeToolImageNumber(value);
  if (!normalized) return null;

  return (
    state.tools.find((tool) => {
      return normalizeToolImageNumber(tool.tNumber) === normalized;
    }) || null
  );
}

function formatToolScannerSummary(tool) {
  if (!tool) return "";
  return `T ${tool.tNumber} · ${tool.label || "-"} · ${formatToolSize(tool)} · ${tool.holder || "-"} · Bestand ${tool.stock}`;
}

function updateScannerWithdrawPreview() {
  const input = document.getElementById("scannerWithdrawTNumber");
  const preview = document.getElementById("scannerWithdrawToolPreview");
  if (!preview) return;

  const tool = findToolByTNumberInput(input?.value || "");
  preview.innerHTML = tool
    ? `<div class="text-sm text-emerald-800 bg-emerald-50 border border-emerald-200 rounded p-2">${escapeHtml(formatToolScannerSummary(tool))}</div>`
    : `<div class="text-sm text-slate-500 bg-slate-50 border rounded p-2">Kein Werkzeug gefunden</div>`;
}

function updateScannerRestockPreview() {
  const input = document.getElementById("scannerRestockTNumber");
  const preview = document.getElementById("scannerRestockToolPreview");
  const qtyInput = document.getElementById("scannerRestockQty");
  if (!preview) return;

  const tool = findToolByTNumberInput(input?.value || "");
  preview.innerHTML = tool
    ? `<div class="text-sm text-emerald-800 bg-emerald-50 border border-emerald-200 rounded p-2">${escapeHtml(formatToolScannerSummary(tool))}</div>`
    : `<div class="text-sm text-slate-500 bg-slate-50 border rounded p-2">Kein Werkzeug gefunden</div>`;

  if (qtyInput && tool?.ordered && Number(tool.orderedQty || 0) > 0) {
    qtyInput.value = String(tool.orderedQty);
  } else if (qtyInput && !qtyInput.value) {
    qtyInput.value = "1";
  }
}

function openQrScannerPlaceholder(mode) {
  openQrScanner(mode);
}

async function openQrScanner(mode) {
  stopQrScanner();

  const modeLabel = mode === "restock" ? "Einlagern" : "Entnehmen";
  const host = getModalHost();

  if (!("BarcodeDetector" in window)) {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
        <h3 class="text-lg font-bold mb-2">QR-Code scannen</h3>
        <p class="text-sm text-slate-700 mb-4">QR-Scanner wird von diesem Browser nicht unterstützt. Bitte manuell per T-Nummer buchen.</p>
        <div class="flex justify-end gap-2">
          <button class="px-3 py-2 rounded bg-slate-200" onclick="closeToolImagePopup()">Schließen</button>
          <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="${mode === "restock" ? "openManualToolRestock()" : "openManualToolWithdraw()"}">Manuell buchen</button>
        </div>
      </div>
    </div>`;
    return;
  }

  if (!navigator.mediaDevices?.getUserMedia) {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
        <h3 class="text-lg font-bold mb-2">QR-Code scannen</h3>
        <p class="text-sm text-slate-700 mb-4">Kamera-Zugriff wird von diesem Browser nicht unterstützt. Bitte manuell per T-Nummer buchen.</p>
        <div class="flex justify-end gap-2">
          <button class="px-3 py-2 rounded bg-slate-200" onclick="closeToolImagePopup()">Schließen</button>
          <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="${mode === "restock" ? "openManualToolRestock()" : "openManualToolWithdraw()"}">Manuell buchen</button>
        </div>
      </div>
    </div>`;
    return;
  }

  qrScannerMode = mode;
  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-lg p-4">
      <h3 class="text-lg font-bold mb-2">QR-Code scannen</h3>
      <p class="text-sm text-slate-500 mb-1">Modus: ${modeLabel}</p>
      <video id="qrScannerVideo" class="w-full bg-black rounded-lg mb-3" autoplay muted playsinline></video>
      <div id="qrScannerStatus" class="text-sm text-slate-600 bg-slate-50 border rounded p-2 mb-4">Kamera wird gestartet ...</div>
      <div class="flex justify-end gap-2">
        <button class="px-3 py-2 rounded bg-slate-200" onclick="stopQrScanner(); closeToolImagePopup()">Abbrechen</button>
        <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="stopQrScanner(); ${mode === "restock" ? "openManualToolRestock()" : "openManualToolWithdraw()"}">Manuell buchen</button>
      </div>
    </div>
  </div>`;

  try {
    qrScannerStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: "environment" },
    });
    const video = document.getElementById("qrScannerVideo");
    if (!video) return;
    video.srcObject = qrScannerStream;
    await video.play();
    const status = document.getElementById("qrScannerStatus");
    if (status) {
      status.textContent = "Kamera aktiv. QR-Code vor die Kamera halten.";
    }
    startQrScannerLoop();
  } catch (error) {
    console.error("Fehler beim Starten des QR-Scanners:", error);
    stopQrScanner();
    const status = document.getElementById("qrScannerStatus");
    if (status) {
      status.textContent =
        "Kamera konnte nicht gestartet werden. Bitte manuell per T-Nummer buchen.";
    }
  }
}

async function startQrScannerLoop() {
  const video = document.getElementById("qrScannerVideo");
  const status = document.getElementById("qrScannerStatus");
  if (!video || !("BarcodeDetector" in window)) return;

  const detector = new window.BarcodeDetector({ formats: ["qr_code"] });
  let scanning = false;

  qrScannerTimer = window.setInterval(async () => {
    if (scanning || !qrScannerStream) return;
    if (video.readyState < 2) return;

    scanning = true;
    try {
      const codes = await detector.detect(video);
      const first = codes?.[0];
      const rawValue = first?.rawValue || "";
      if (rawValue) {
        if (status) status.textContent = "QR-Code erkannt.";
        const mode = qrScannerMode;
        stopQrScanner();
        processScannedToolQr(rawValue, mode);
      }
    } catch (error) {
      console.error("Fehler beim Lesen des QR-Codes:", error);
      if (status) status.textContent = "QR-Code konnte nicht gelesen werden.";
    } finally {
      scanning = false;
    }
  }, 450);
}

function processScannedToolQr(rawValue, mode) {
  const value = String(rawValue || "").trim();
  if (!value.startsWith("TOOL:")) {
    alert("Ungültiger QR-Code");
    return;
  }

  const toolId = value.slice("TOOL:".length);
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) {
    alert("Werkzeug nicht gefunden");
    return;
  }

  const modeLabel = mode === "restock" ? "Einlagerung" : "Entnahme";
  const actionLabel =
    mode === "restock" ? "Einlagerung vorbereiten" : "Entnahme bestätigen";
  const actionCall =
    mode === "restock"
      ? `openManualToolRestock('${escapeHtml(String(tool.tNumber || ""))}')`
      : `confirmQrToolWithdraw('${escapeHtml(String(tool.id || ""))}')`;

  getModalHost().innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
      <h3 class="text-lg font-bold mb-2">Werkzeug erkannt</h3>
      <p class="text-sm text-slate-500 mb-3">Modus: ${modeLabel}</p>
      <div class="border rounded-lg bg-slate-50 p-3 space-y-1 text-sm mb-4">
        <div><span class="text-slate-500">T-Nummer:</span> T ${escapeHtml(tool.tNumber)}</div>
        <div><span class="text-slate-500">Bezeichnung:</span> ${escapeHtml(tool.label || "-")}</div>
        <div><span class="text-slate-500">Durchmesser:</span> ${escapeHtml(formatToolSize(tool))}</div>
        <div><span class="text-slate-500">Aufnahme:</span> ${escapeHtml(tool.holder || "-")}</div>
        <div><span class="text-slate-500">Fach:</span> ${escapeHtml(tool.shelf || "-")}</div>
        <div><span class="text-slate-500">Bestand:</span> ${escapeHtml(tool.stock)}</div>
      </div>
      <div class="flex justify-end gap-2">
        <button class="px-3 py-2 rounded bg-slate-200" onclick="closeToolImagePopup()">Abbrechen</button>
        <button class="px-3 py-2 rounded bg-slate-900 text-white" onclick="${actionCall}">${actionLabel}</button>
      </div>
    </div>
  </div>`;
}

async function confirmQrToolWithdraw(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool) {
    alert("Werkzeug nicht gefunden");
    return;
  }

  const oldStock = Number(tool.stock || 0);
  if (oldStock <= 0) {
    alert("Bestand ist 0. Entnahme nicht möglich.");
    return;
  }

  const newStock = oldStock - 1;
  const { error } = await supabaseClient
    .from("tools")
    .update({ stock: newStock })
    .eq("id", tool.id);

  if (error) {
    console.error("Fehler bei QR-Scanner Entnahme:", error);
    alert("QR-Entnahme konnte nicht gespeichert werden.");
    return;
  }

  const tools = await loadToolsFromSupabase();
  const refreshedTool = tools.find((t) => t.id === tool.id);

  if (!refreshedTool || Number(refreshedTool.stock) !== newStock) {
    console.error("QR-Scanner Entnahme konnte nicht verifiziert werden", {
      toolId: tool.id,
      expectedStock: newStock,
      actualStock: refreshedTool?.stock,
    });
    alert("QR-Entnahme konnte nicht gespeichert werden.");
    return;
  }

  state.tools = tools;
  state.toolJournal.unshift({
    id: `journal-${Date.now()}`,
    user: currentUser?.name || "Werkzeug-Scanner",
    at: new Date().toISOString().slice(0, 16).replace("T", " "),
    toolId: refreshedTool.id,
    tNumber: refreshedTool.tNumber,
    qty: 1,
    action: "QR-Scanner Entnahme 1",
  });

  closeToolImagePopup();
  persist();
  render();
}

function stopQrScanner() {
  if (qrScannerTimer) {
    window.clearInterval(qrScannerTimer);
    qrScannerTimer = null;
  }

  if (qrScannerStream) {
    qrScannerStream.getTracks().forEach((track) => track.stop());
    qrScannerStream = null;
  }

  qrScannerMode = null;
}

function openManualToolWithdraw(initialTNumber = "") {
  const host = getModalHost();
  stopQrScanner();

  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-xl p-4">
      <h3 class="text-lg font-bold mb-3">Werkzeug manuell entnehmen</h3>
      <input id="scannerWithdrawTNumber" inputmode="numeric" class="border rounded p-2 w-full mb-2" placeholder="T-Nummer eingeben, z. B. 133" value="${escapeHtml(initialTNumber)}" oninput="updateScannerWithdrawPreview()" />
      <div id="scannerWithdrawToolPreview" class="mb-4">
        <div class="text-sm text-slate-500 bg-slate-50 border rounded p-2">Kein Werkzeug gefunden</div>
      </div>
      <div class="flex justify-end gap-2">
        <button class="px-3 py-2 rounded bg-slate-200" onclick="closeToolImagePopup()">Abbrechen</button>
        <button id="scannerWithdrawSave" class="px-3 py-2 rounded bg-emerald-700 text-white">Entnehmen</button>
      </div>
    </div>
  </div>`;

  host.querySelector("#scannerWithdrawSave")?.addEventListener("click", async () => {
    const input = host.querySelector("#scannerWithdrawTNumber");
    const tool = findToolByTNumberInput(input?.value || "");
    if (!tool) {
      alert("Werkzeug wurde nicht gefunden.");
      return;
    }
    if (Number(tool.stock || 0) <= 0) {
      alert("Nicht genügend Bestand vorhanden.");
      return;
    }

    const oldStock = Number(tool.stock || 0);
    const newStock = oldStock - 1;
    if (newStock < 0) {
      alert("Nicht genügend Bestand vorhanden.");
      return;
    }

    const payload = { stock: newStock };
    const { error } = await supabaseClient
      .from("tools")
      .update(payload)
      .eq("id", tool.id);

    if (error) {
      console.error("Fehler bei Werkzeug-Scanner Entnahme:", error);
      alert(`Entnahme konnte nicht gespeichert werden: ${error.message}`);
      return;
    }

    const { data: refreshed, error: reloadError } = await supabaseClient
      .from("tools")
      .select("*")
      .eq("id", tool.id)
      .maybeSingle();

    if (reloadError || !refreshed) {
      console.error("Scanner stock reload failed", {
        toolId: tool.id,
        reloadError,
      });
      alert(
        "Entnahme konnte nicht geprüft werden. Werkzeug wurde nicht neu geladen.",
      );
      return;
    }

    if (Number(refreshed.stock) !== newStock) {
      console.error("Scanner stock update verification failed", {
        toolId: tool.id,
        expectedStock: newStock,
        actualStock: refreshed.stock,
      });
      alert(
        "Entnahme wurde nicht gespeichert. Bitte RLS/Update-Rechte für tools prüfen.",
      );
      return;
    }

    const refreshedTool = normalizeToolFromDb(refreshed);
    const index = state.tools.findIndex((t) => t.id === refreshedTool.id);
    if (index !== -1) {
      state.tools[index] = refreshedTool;
    }

    state.toolJournal.unshift({
      id: `journal-${Date.now()}`,
      user: currentUser?.name || "Werkzeug-Scanner",
      at: new Date().toISOString().slice(0, 16).replace("T", " "),
      toolId: refreshedTool.id,
      tNumber: refreshedTool.tNumber,
      qty: 1,
      action: "Werkzeug-Scanner Entnahme 1",
    });

    closeToolImagePopup();
    persist();
    render();
  });

  host.querySelector("#scannerWithdrawTNumber")?.focus();
  updateScannerWithdrawPreview();
}

function openManualToolRestock(initialTNumber = "") {
  const host = getModalHost();
  stopQrScanner();

  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-xl p-4">
      <h3 class="text-lg font-bold mb-3">Werkzeug manuell einlagern</h3>
      <input id="scannerRestockTNumber" inputmode="numeric" class="border rounded p-2 w-full mb-2" placeholder="T-Nummer eingeben, z. B. 133" value="${escapeHtml(initialTNumber)}" oninput="updateScannerRestockPreview()" />
      <div id="scannerRestockToolPreview" class="mb-3">
        <div class="text-sm text-slate-500 bg-slate-50 border rounded p-2">Kein Werkzeug gefunden</div>
      </div>
      <input id="scannerRestockQty" type="number" min="1" class="border rounded p-2 w-full mb-4" value="1" />
      <div class="flex justify-end gap-2">
        <button class="px-3 py-2 rounded bg-slate-200" onclick="closeToolImagePopup()">Abbrechen</button>
        <button id="scannerRestockSave" class="px-3 py-2 rounded bg-blue-700 text-white">Einlagern</button>
      </div>
    </div>
  </div>`;

  host.querySelector("#scannerRestockSave")?.addEventListener("click", async () => {
    const input = host.querySelector("#scannerRestockTNumber");
    const qty = Number(host.querySelector("#scannerRestockQty")?.value || 0);
    const tool = findToolByTNumberInput(input?.value || "");
    if (!tool) {
      alert("Werkzeug wurde nicht gefunden.");
      return;
    }
    if (!Number.isFinite(qty) || qty <= 0) {
      alert("Bitte eine gültige Menge eingeben.");
      return;
    }

    const oldStock = Number(tool.stock || 0);
    const newStock = oldStock + qty;
    const payload = { stock: newStock };
    const { error } = await supabaseClient
      .from("tools")
      .update(payload)
      .eq("id", tool.id);

    if (error) {
      console.error("Fehler bei Werkzeug-Scanner Einlagerung:", error);
      alert(`Einlagerung konnte nicht gespeichert werden: ${error.message}`);
      return;
    }

    const { data: refreshed, error: reloadError } = await supabaseClient
      .from("tools")
      .select("*")
      .eq("id", tool.id)
      .maybeSingle();

    if (reloadError || !refreshed) {
      console.error("Scanner stock reload failed", {
        toolId: tool.id,
        reloadError,
      });
      alert(
        "Einlagerung konnte nicht geprüft werden. Werkzeug wurde nicht neu geladen.",
      );
      return;
    }

    if (Number(refreshed.stock) !== newStock) {
      console.error("Scanner stock update verification failed", {
        toolId: tool.id,
        expectedStock: newStock,
        actualStock: refreshed.stock,
      });
      alert(
        "Einlagerung wurde nicht gespeichert. Bitte RLS/Update-Rechte für tools prüfen.",
      );
      return;
    }

    const refreshedTool = normalizeToolFromDb(refreshed);
    const index = state.tools.findIndex((t) => t.id === refreshedTool.id);
    if (index !== -1) {
      state.tools[index] = refreshedTool;
    }

    state.toolJournal.unshift({
      id: `journal-${Date.now()}`,
      user: currentUser?.name || "Werkzeug-Scanner",
      at: new Date().toISOString().slice(0, 16).replace("T", " "),
      toolId: refreshedTool.id,
      tNumber: refreshedTool.tNumber,
      qty,
      action: `Werkzeug-Scanner Einlagerung ${qty}`,
    });

    closeToolImagePopup();
    persist();
    render();
  });

  host.querySelector("#scannerRestockTNumber")?.focus();
  updateScannerRestockPreview();
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

async function editTool(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool || currentUser.role !== "admin") return;

  const data = await editToolCentered(tool);
  if (!data) return;

  const isThreadTool = isThreadToolLabel(data.label);
  const isRadiusTool = isRadiusToolLabel(data.label);

  if (
    !data.label ||
    !data.diameter ||
    !isValidToolShelf(data.shelf) ||
    !data.articleNo ||
    !["HSK 100", "HSK 63"].includes(data.holder)
  ) {
    return alert(
      "Bitte Felder korrekt ausfüllen (2-stellige Zahl + Buchstabe für Fach, z. B. 26C oder 02T; Aufnahme HSK 100 oder HSK 63).",
    );
  }

  if (isThreadTool && !data.threadPrefix) {
    return alert("Bitte Gewindekennung wählen.");
  }

  if (isThreadTool && data.threadPrefix === "MF" && !data.threadPitch) {
    return alert("Bitte bei MF die Steigung (P) angeben.");
  }

  if (isRadiusTool && !data.cornerRadius) {
    return alert("Bitte Schneidenradius eingeben.");
  }

  if (!isThreadTool && !Number.isFinite(Number(data.diameter))) {
    return alert("Bitte gültigen numerischen Durchmesser eingeben.");
  }

  if (
    data.insertTool &&
    (!Number.isFinite(Number(data.insertEdges)) ||
      Number(data.insertEdges) <= 0)
  ) {
    return alert("Bitte Anzahl der Schneiden > 0 eingeben.");
  }

  const payload = {
    label: data.label,
    diameter: String(data.diameter),
    thread_prefix: isThreadTool ? data.threadPrefix || null : null,
    thread_pitch:
      isThreadTool && data.threadPrefix === "MF"
        ? data.threadPitch || null
        : null,
    corner_radius: isRadiusTool ? data.cornerRadius || null : null,
    material_id: data.materialId || null,
    shelf: data.shelf,
    article_no: data.articleNo,
    holder: data.holder,
    stock: Number(data.stock || 0),
    min_stock: Number(data.minStock || 0),
    optimal_stock: Number(data.optimalStock || 0),
    manufacturer: data.manufacturer || null,
    insert_tool: !!data.insertTool,
    insert_edges: data.insertTool ? Number(data.insertEdges || 0) : 0,
    insert_radius: data.insertTool ? data.insertRadius || null : null,
  };

  const { data: updated, error } = await supabaseClient
    .from("tools")
    .update(payload)
    .eq("id", toolId)
    .select()
    .maybeSingle();

  if (error) {
    console.error("Fehler beim Bearbeiten des Werkzeugs:", error);
    return alert(`Werkzeug konnte nicht gespeichert werden: ${error.message}`);
  }

  if (!updated) {
    console.error("Kein Werkzeug nach Update zurückgegeben", {
      toolId,
      payload,
    });
    return alert(
      "Werkzeug konnte nicht gespeichert werden. Datensatz wurde nicht gefunden oder nicht zurückgegeben.",
    );
  }

  const index = state.tools.findIndex((t) => t.id === toolId);
  if (index !== -1) {
    state.tools[index] = normalizeToolFromDb(updated);
  }

  persist();
  render();
}

async function editToolCentered(tool) {
  const host = getModalHost();

  return new Promise((resolve) => {
    host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
      <div class="bg-white rounded-xl shadow-xl w-full max-w-4xl max-h-[90vh] overflow-auto p-5">
        <div class='flex items-center justify-between mb-4 gap-3'>
          <h3 class="text-lg font-bold">Werkzeug bearbeiten – T ${tool.tNumber}</h3>
          <button id="toolEditCloseTop" class="px-3 py-1 rounded bg-slate-200">Schließen</button>
        </div>

        <div class='grid md:grid-cols-2 gap-4'>
          <label class='text-sm font-medium'>
            T-Nummer
            <input id='editTNumber' class='border rounded p-2 w-full mt-1 bg-slate-100' value="${tool.tNumber}" disabled />
          </label>

          <label class='text-sm font-medium'>
            Bezeichnung
            <select id='editLabel' class='border rounded p-2 w-full mt-1' onchange='updateEditToolTypeFields()'>
              ${getToolLabels()
                .map(
                  (l) =>
                    `<option value="${l}" ${tool.label === l ? "selected" : ""}>${l}</option>`,
                )
                .join("")}
            </select>
          </label>

          <label class='text-sm font-medium'>
            Durchmesser
            <input id='editDiameter' class='border rounded p-2 w-full mt-1' placeholder='Durchmesser' value="${tool.diameter ?? ""}" />
          </label>

          <label class='text-sm font-medium'>
            Schneidwerkstoff
            <select id='editMaterial' class='border rounded p-2 w-full mt-1'>
              <option value=''>Schneidwerkstoff wählen</option>
              ${toolMaterials
                .map(
                  (m) =>
                    `<option value="${m.id}" ${tool.materialId === m.id ? "selected" : ""}>${m.name}</option>`,
                )
                .join("")}
            </select>
          </label>

          <div id='editThreadPrefixWrap' style='display:none;'>
            <label class='text-sm font-medium block'>
              Gewindekennung
              <select id='editThreadPrefix' class='border rounded p-2 w-full mt-1' onchange='updateEditThreadPitchVisibility()'>
                <option value=''>Kennung wählen</option>
                <option value='M' ${tool.threadPrefix === "M" ? "selected" : ""}>M</option>
                <option value='MF' ${tool.threadPrefix === "MF" ? "selected" : ""}>MF</option>
                <option value='G' ${tool.threadPrefix === "G" ? "selected" : ""}>G</option>
                <option value='UNF' ${tool.threadPrefix === "UNF" ? "selected" : ""}>UNF</option>
                <option value='UNC' ${tool.threadPrefix === "UNC" ? "selected" : ""}>UNC</option>
                <option value='Mx' ${tool.threadPrefix === "Mx" ? "selected" : ""}>Mx</option>
              </select>
            </label>
          </div>

          <div id='editThreadPitchWrap' style='display:none;'>
            <label class='text-sm font-medium block'>
              Steigung
              <input id='editThreadPitch' class='border rounded p-2 w-full mt-1' placeholder='Steigung P (nur MF)' value="${tool.threadPitch || ""}" />
            </label>
          </div>

          <div id='editCornerRadiusWrap' style='display:none;'>
            <label class='text-sm font-medium block'>
              Schneidenradius
              <input id='editCornerRadius' class='border rounded p-2 w-full mt-1' placeholder='Schneidenradius' value="${tool.cornerRadius || ""}" />
            </label>
          </div>

          <label class='text-sm font-medium'>
            Lagerfach
            <input id='editShelf' class='border rounded p-2 w-full mt-1' placeholder='01S' value="${tool.shelf || ""}" />
          </label>

          <label class='text-sm font-medium'>
            Artikelnummer
            <input id='editArticleNo' class='border rounded p-2 w-full mt-1' placeholder='Artikel Nr.' value="${tool.articleNo || ""}" />
          </label>

          <label class='text-sm font-medium'>
            Aufnahme
            <select id='editHolder' class='border rounded p-2 w-full mt-1'>
              <option value='HSK 100' ${tool.holder === "HSK 100" ? "selected" : ""}>HSK 100</option>
              <option value='HSK 63' ${tool.holder === "HSK 63" ? "selected" : ""}>HSK 63</option>
            </select>
          </label>

          <label class='text-sm font-medium'>
            Hersteller
            <select id='editManufacturer' class='border rounded p-2 w-full mt-1'>
              ${getToolManufacturers()
                .map(
                  (m) =>
                    `<option value="${m}" ${tool.manufacturer === m ? "selected" : ""}>${m}</option>`,
                )
                .join("")}
            </select>
          </label>

          <label class='text-sm font-medium'>
            Bestand
            <input id='editStock' type='number' class='border rounded p-2 w-full mt-1' placeholder='Bestand' value="${tool.stock ?? 0}" />
          </label>

          <label class='text-sm font-medium'>
            Mindestbestand
            <input id='editMinStock' type='number' class='border rounded p-2 w-full mt-1' placeholder='Mindestbestand' value="${tool.minStock ?? 0}" />
          </label>

          <label class='text-sm font-medium'>
            Optimale Stückzahl
            <input id='editOptimalStock' type='number' class='border rounded p-2 w-full mt-1' placeholder='Optimale Stückzahl' value="${tool.optimalStock ?? 0}" />
          </label>

          <div></div>

          <label class='flex items-center gap-2 text-sm md:col-span-2'>
            <input id='editInsertTool' type='checkbox' ${tool.insertTool ? "checked" : ""} onchange='toggleInsertToolFieldsById("editInsertTool","editInsertEdges","editInsertRadius")' />
            Wendeplattenwerkzeug
          </label>

          <label class='text-sm font-medium md:col-span-2'>
            Anzahl Schneiden
            <input id='editInsertEdges' type='number' class='border rounded p-2 w-full mt-1' placeholder='Anzahl Schneiden' value="${tool.insertEdges ?? 0}" ${tool.insertTool ? "" : "disabled"} />
          </label>

          <label id='editInsertRadiusWrap' class='text-sm font-medium md:col-span-2' style='display:${tool.insertTool ? "" : "none"};'>
            Plattenradius optional
            <input id='editInsertRadius' class='border rounded p-2 w-full mt-1' placeholder='Plattenradius optional, z. B. 0.8' value="${escapeHtml(tool.insertRadius || "")}" ${tool.insertTool ? "" : "disabled"} />
          </label>
        </div>

        <div class='flex justify-end gap-2 mt-5'>
          <button id="toolEditCloseBottom" class="px-3 py-2 rounded bg-slate-200">Abbrechen</button>
          <button id="toolEditSave" class="px-3 py-2 rounded bg-slate-900 text-white">Speichern</button>
        </div>
      </div>
    </div>`;

    updateEditToolTypeFields();
    toggleInsertToolFieldsById(
      "editInsertTool",
      "editInsertEdges",
      "editInsertRadius",
    );

    const close = () => {
      host.innerHTML = "";
    };

    host.querySelector("#toolEditCloseTop")?.addEventListener("click", () => {
      close();
      resolve(null);
    });

    host
      .querySelector("#toolEditCloseBottom")
      ?.addEventListener("click", () => {
        close();
        resolve(null);
      });

    host.querySelector("#toolEditSave")?.addEventListener("click", () => {
      const data = {
        label: document.getElementById("editLabel")?.value || "",
        diameter: document.getElementById("editDiameter")?.value?.trim() || "",
        threadPrefix: document.getElementById("editThreadPrefix")?.value || "",
        threadPitch:
          document.getElementById("editThreadPitch")?.value?.trim() || "",
        cornerRadius:
          document.getElementById("editCornerRadius")?.value?.trim() || "",
        materialId: document.getElementById("editMaterial")?.value || "",
        shelf:
          document.getElementById("editShelf")?.value?.trim().toUpperCase() ||
          "",
        articleNo:
          document.getElementById("editArticleNo")?.value?.trim() || "",
        holder: document.getElementById("editHolder")?.value || "",
        manufacturer: document.getElementById("editManufacturer")?.value || "",
        stock: Number(document.getElementById("editStock")?.value || 0),
        minStock: Number(document.getElementById("editMinStock")?.value || 0),
        optimalStock: Number(
          document.getElementById("editOptimalStock")?.value || 0,
        ),
        insertTool: !!document.getElementById("editInsertTool")?.checked,
        insertEdges: Number(
          document.getElementById("editInsertEdges")?.value || 0,
        ),
        insertRadius:
          document.getElementById("editInsertRadius")?.value?.trim() || "",
      };

      close();
      resolve(data);
    });
  });
}

function updateEditToolTypeFields() {
  const label = document.getElementById("editLabel")?.value || "";
  const isThread = isThreadToolLabel(label);
  const isRadius = isRadiusToolLabel(label);

  const threadPrefixWrap = document.getElementById("editThreadPrefixWrap");
  const threadPitchWrap = document.getElementById("editThreadPitchWrap");
  const cornerRadiusWrap = document.getElementById("editCornerRadiusWrap");

  if (threadPrefixWrap) threadPrefixWrap.style.display = isThread ? "" : "none";
  if (cornerRadiusWrap) cornerRadiusWrap.style.display = isRadius ? "" : "none";

  updateEditThreadPitchVisibility();
}

function updateEditThreadPitchVisibility() {
  const label = document.getElementById("editLabel")?.value || "";
  const threadPrefix = document.getElementById("editThreadPrefix")?.value || "";
  const threadPitchWrap = document.getElementById("editThreadPitchWrap");

  const visible = isThreadToolLabel(label) && threadPrefix === "MF";
  if (threadPitchWrap) threadPitchWrap.style.display = visible ? "" : "none";
}

async function editTool(toolId) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool || currentUser.role !== "admin") return;

  const data = await editToolCentered(tool);
  if (!data) return;

  const isThreadTool = isThreadToolLabel(data.label);
  const isRadiusTool = isRadiusToolLabel(data.label);

  if (
    !data.label ||
    !data.diameter ||
    !isValidToolShelf(data.shelf) ||
    !data.articleNo ||
    !["HSK 100", "HSK 63"].includes(data.holder)
  ) {
    return alert(
      "Bitte Felder korrekt ausfüllen (2-stellige Zahl + Buchstabe für Fach, z. B. 26C oder 02T; Aufnahme HSK 100 oder HSK 63).",
    );
  }

  if (isThreadTool && !data.threadPrefix) {
    return alert("Bitte Gewindekennung wählen.");
  }

  if (isThreadTool && data.threadPrefix === "MF" && !data.threadPitch) {
    return alert("Bitte bei MF die Steigung (P) angeben.");
  }

  if (isRadiusTool && !data.cornerRadius) {
    return alert("Bitte Schneidenradius eingeben.");
  }

  if (!isThreadTool && !Number.isFinite(Number(data.diameter))) {
    return alert("Bitte gültigen numerischen Durchmesser eingeben.");
  }

  if (
    data.insertTool &&
    (!Number.isFinite(Number(data.insertEdges)) ||
      Number(data.insertEdges) <= 0)
  ) {
    return alert("Bitte Anzahl der Schneiden > 0 eingeben.");
  }

  const payload = {
    label: data.label,
    diameter: String(data.diameter),
    thread_prefix: isThreadTool ? data.threadPrefix || null : null,
    thread_pitch:
      isThreadTool && data.threadPrefix === "MF"
        ? data.threadPitch || null
        : null,
    corner_radius: isRadiusTool ? data.cornerRadius || null : null,
    material_id: data.materialId || null,
    shelf: data.shelf,
    article_no: data.articleNo,
    holder: data.holder,
    stock: Number(data.stock || 0),
    min_stock: Number(data.minStock || 0),
    optimal_stock: Number(data.optimalStock || 0),
    manufacturer: data.manufacturer || null,
    insert_tool: !!data.insertTool,
    insert_edges: data.insertTool ? Number(data.insertEdges || 0) : 0,
    insert_radius: data.insertTool ? data.insertRadius || null : null,
  };

  const { data: updated, error } = await supabaseClient
    .from("tools")
    .update(payload)
    .eq("id", toolId)
    .select()
    .maybeSingle();

  if (error) {
    console.error("Fehler beim Bearbeiten des Werkzeugs:", error);
    return alert(`Werkzeug konnte nicht gespeichert werden: ${error.message}`);
  }

  if (!updated) {
    console.error("Kein Werkzeug nach Update zurückgegeben", {
      toolId,
      payload,
    });
    return alert(
      "Werkzeug konnte nicht gespeichert werden. Datensatz wurde nicht gefunden oder nicht zurückgegeben.",
    );
  }

  const index = state.tools.findIndex((t) => t.id === toolId);
  if (index !== -1) {
    state.tools[index] = normalizeToolFromDb(updated);
  }

  persist();
  render();
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
      material_id: normalized.materialId || null,
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
      insert_radius: normalized.insertRadius || null,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    };

    const { data: inserted, error } = await supabaseClient
      .from("tools")
      .insert(payload)
      .select()
      .single();

    if (error) {
      console.error("Fehler beim Anlegen des Werkzeugs in Supabase:", error);
      return alert(
        `Werkzeug konnte nicht gespeichert werden: ${error.message}`,
      );
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

async function markToolOrdered(toolId, ordered) {
  const tool = state.tools.find((t) => t.id === toolId);
  if (!tool || currentUser.role !== "admin") return;

  const nextOrderedQty = ordered ? Math.max(1, effectiveOrderQty(tool)) : 0;

  const { data: updated, error } = await supabaseClient
    .from("tools")
    .update({
      ordered: !!ordered,
      ordered_qty: Number(nextOrderedQty || 0),
    })
    .eq("id", toolId)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Markieren als bestellt:", error);
    return alert(
      `Bestellstatus konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  const index = state.tools.findIndex((t) => t.id === toolId);
  if (index !== -1) {
    state.tools[index] = normalizeToolFromDb(updated);
  }

  if (ordered) {
    archiveOrderEvent(
      state.tools[index],
      state.tools[index].orderedQty,
      "mark_ordered",
    );
  }

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

  const nextStock = Number(tool.stock || 0) + Number(add || 0);

  const { data: updated, error } = await supabaseClient
    .from("tools")
    .update({
      stock: nextStock,
      ordered: false,
      ordered_qty: 0,
    })
    .eq("id", toolId)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Einlagern des Werkzeugs:", error);
    return alert(
      `Einlagerung konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  const index = state.tools.findIndex((t) => t.id === toolId);
  if (index !== -1) {
    state.tools[index] = normalizeToolFromDb(updated);
  }

  archiveOrderEvent(state.tools[index], add, "restock");
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

    if (!Number.isFinite(qty) || qty <= 0) {
      return alert("Bitte eine gültige Entnahmemenge eingeben.");
    }

    if (Number(tool.stock || 0) < Number(qty || 0)) {
      return alert("Nicht genügend Bestand vorhanden.");
    }
  }

  let updatedTool = tool;

  if (takeOut && qty > 0) {
    const nextStock = Math.max(0, Number(tool.stock || 0) - Number(qty || 0));

    const { data, error } = await supabaseClient
      .from("tools")
      .update({
        stock: nextStock,
      })
      .eq("id", toolId)
      .select()
      .maybeSingle();

    if (error) {
      console.error("Fehler bei Werkzeug-Entnahme:", error);
      return alert(
        `Entnahme konnte nicht gespeichert werden: ${error.message}`,
      );
    }

    if (!data) {
      console.error("Keine Werkzeugzeile nach Entnahme zurückgegeben.", {
        toolId,
        nextStock,
      });
      return alert(
        "Entnahme wurde nicht bestätigt. Bitte Seite neu laden und erneut versuchen.",
      );
    }

    if (error) {
      console.error("Fehler bei Werkzeug-Entnahme:", error);
      return alert(
        `Entnahme konnte nicht gespeichert werden: ${error.message}`,
      );
    }

    updatedTool = normalizeToolFromDb(data);

    const index = state.tools.findIndex((t) => t.id === toolId);
    if (index !== -1) {
      state.tools[index] = updatedTool;
    }
  }

  state.toolJournal.unshift({
    id: `journal-${Date.now()}`,
    user: currentUser.name,
    at: new Date().toISOString().slice(0, 16).replace("T", " "),
    toolId: updatedTool.id,
    tNumber: updatedTool.tNumber,
    qty,
    action: takeOut
      ? `Werkzeugwechsel + Entnahme ${qty}`
      : "Werkzeugwechsel ohne Entnahme",
  });

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

  const { error } = await supabaseClient
    .from("tools")
    .delete()
    .eq("id", toolId);

  if (error) {
    console.error("Fehler beim Löschen des Werkzeugs:", error);
    return alert(`Werkzeug konnte nicht gelöscht werden: ${error.message}`);
  }

  state.tools = state.tools.filter((t) => t.id !== toolId);
  state.toolJournal = state.toolJournal.filter((j) => j.toolId !== toolId);

  persist();
  render();
}

async function undoToolJournalEntry(entryId) {
  const idx = state.toolJournal.findIndex((j) => j.id === entryId);
  if (idx === -1) return;

  const entry = state.toolJournal[idx];
  const tool = state.tools.find((t) => t.id === entry.toolId);

  if (tool && Number(entry.qty) > 0) {
    const nextStock = Number(tool.stock || 0) + Number(entry.qty || 0);

    const { data, error } = await supabaseClient
      .from("tools")
      .update({
        stock: nextStock,
      })
      .eq("id", tool.id)
      .select()
      .single();

    if (error) {
      console.error(
        "Fehler beim Rückgängig machen des Journal-Eintrags:",
        error,
      );
      return alert(
        `Rückgängig konnte nicht gespeichert werden: ${error.message}`,
      );
    }

    const updatedTool = normalizeToolFromDb(data);
    const toolIndex = state.tools.findIndex((t) => t.id === tool.id);
    if (toolIndex !== -1) {
      state.tools[toolIndex] = updatedTool;
    }
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
    imageStatus: "",
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
  const imageStatus =
    document.getElementById("toolFilterImageStatus")?.value || "";
  state.toolFilters = {
    search,
    label,
    tNumber,
    diameter,
    holder,
    imageStatus,
  };
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
  const filters = state.toolFilters || {
    search: "",
    label: "",
    tNumber: "",
    diameter: "",
    holder: "",
    imageStatus: "",
  };
  const search = (filters.search || "").toLowerCase();
  const filterLabel = filters.label || "";
  const filterT = filters.tNumber || "";
  const filterD = filters.diameter || "";
  const filterHolder = filters.holder || "";
  const imageStatus = filters.imageStatus || "";
  const isAdmin = currentUser?.role === "admin";

  const filterLabelOptions = labels
    .map(
      (l) =>
        `<option value="${l}" ${l === filterLabel ? "selected" : ""}>${l}</option>`,
    )
    .join("");

  const tools = state.tools.filter((t) => {
    const bySearch =
      !search ||
      `${t.tNumber} ${t.label} ${t.diameter} ${t.articleNo} ${t.holder || ""} ${t.manufacturer || ""} ${getToolMaterialNameById(t.materialId)}`
        .toLowerCase()
        .includes(search);
    const byLabel = !filterLabel || t.label === filterLabel;
    const byT = !filterT || String(t.tNumber).includes(filterT);
    const byD = !filterD || String(t.diameter).includes(filterD);
    const byHolder = !filterHolder || t.holder === filterHolder;
    const imagePath = getToolImagePath(t);
    const byImageStatus =
      !imageStatus ||
      (imageStatus === "withPath" && !!imagePath) ||
      (imageStatus === "withoutPath" && !imagePath);
    return bySearch && byLabel && byT && byD && byHolder && byImageStatus;
  });

  const toolRows = tools
    .map((t) => {
      const statusText = t.ordered ? "Bestellt" : "-";
      const imagePath = getToolImagePath(t);
      const imageCell = imagePath
        ? `<button class='border rounded bg-white p-1 hover:bg-slate-50' onclick="openToolImagePopup('${t.id}')" title="Werkzeugbild öffnen">
            <img src="${escapeHtml(imagePath)}" alt="Bild T ${escapeHtml(normalizeToolImageNumber(t.tNumber))}" class="w-12 h-12 object-contain" onerror="this.style.display='none'; this.parentElement.querySelector('[data-tool-image-missing]').classList.remove('hidden')" />
            <span data-tool-image-missing class="hidden text-xs text-rose-700">kein Bild</span>
          </button>`
        : "-";
      const insertRadiusLine =
        t.insertTool && t.insertRadius
          ? `<div class='text-xs text-slate-500'>Plattenradius: R ${escapeHtml(t.insertRadius)}</div>`
          : "";

      return `<tr class='border-b'>
        <td class='p-2'>T ${t.tNumber}</td>
        <td class='p-2'>${imageCell}</td>
        <td class='p-2'>${t.label}</td>
        <td class='p-2'>${formatToolSize(t)}${insertRadiusLine}</td>
        ${isAdmin ? `<td class='p-2'>${getToolMaterialNameById(t.materialId)}</td>` : ""}
        <td class='p-2'>${t.shelf}</td>
        <td class='p-2'>${t.articleNo}</td>
        ${isAdmin ? `<td class='p-2'>${t.holder || "-"}</td>` : ""}
        <td class='p-2'>${t.stock}</td>
        <td class='p-2'>${t.minStock}</td>
        ${isAdmin ? `<td class='p-2'>${t.manufacturer || "-"}</td>` : ""}
        <td class='p-2'>${statusText}</td>
        <td class='p-2 whitespace-nowrap'>
          <button class='px-2 py-1 rounded bg-emerald-700 text-white mr-1' onclick="bookToolChange('${t.id}')">Wechsel</button>
          ${
            isAdmin
              ? `<button class='px-2 py-1 rounded bg-slate-700 text-white mr-1' onclick="openToolQrPopup('${t.id}')">QR</button>
                 <button class='px-2 py-1 rounded bg-amber-600 text-white mr-1' onclick="editTool('${t.id}')">Bearbeiten</button>
                 <button class='px-2 py-1 rounded bg-rose-700 text-white' onclick="deleteTool('${t.id}')">Löschen</button>`
              : ""
          }
        </td>
      </tr>`;
    })
    .join("");

  const todoTools = state.tools.filter((t) => shouldOrderTool(t) && !t.ordered);
  const orderedTools = state.tools.filter((t) => {
    return t.ordered || Number(t.orderedQty || 0) > 0;
  });
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
        <td class='p-2'>${getToolMaterialNameById(t.materialId)}</td>
        <td class='p-2'>${t.articleNo || "-"}</td>
        <td class='p-2 whitespace-nowrap'>
          <button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="markToolOrdered('${t.id}', true)">Bestellt markieren</button>
        </td>
      </tr>`,
    )
    .join("");

  const orderedCards = orderedTools
    .map((t) => {
      const isThreadTool = String(t.label || "").includes("Gewinde");
      const isRadiusTool = String(t.label || "").includes("Radiusfräser");
      const detailLine = isThreadTool
        ? `<div><span class='font-semibold'>Steigung:</span> ${t.threadPitch || t.pitch || "-"}</div>`
        : isRadiusTool
          ? `<div><span class='font-semibold'>Radius:</span> R ${t.cornerRadius || "-"}</div>`
          : "";
      const insertRadiusLine =
        t.insertTool && t.insertRadius
          ? `<div><span class='font-semibold'>Plattenradius:</span> R ${escapeHtml(t.insertRadius)}</div>`
          : "";

      return `<div class='border rounded-lg p-3 bg-emerald-50 space-y-3'>
        <div class='grid md:grid-cols-2 gap-x-6 gap-y-2 text-sm'>
          <div><span class='font-semibold'>T-Nummer:</span> T ${t.tNumber}</div>
          <div><span class='font-semibold'>Bezeichnung:</span> ${t.label}</div>
          <div><span class='font-semibold'>Durchmesser:</span> ${formatToolSize(t)}</div>
          <div><span class='font-semibold'>Schneidwerkstoff:</span> ${getToolMaterialNameById(t.materialId)}</div>
          ${detailLine}
          ${insertRadiusLine}
          <div><span class='font-semibold'>Menge:</span> ${t.orderedQty || effectiveOrderQty(t)}</div>
          <div><span class='font-semibold'>Artikelnummer:</span> ${t.articleNo}</div>
          <div><span class='font-semibold'>Lagerfach:</span> ${t.shelf}</div>
        </div>
        <div class='flex flex-col sm:flex-row gap-2'>
          <button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="markToolOrdered('${t.id}', false)">Bestellt</button>
          <button class='px-2 py-1 rounded bg-slate-700 text-white' onclick="restockTool('${t.id}')">Einlagern</button>
        </div>
      </div>`;
    })
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
        <td class='p-2'>${getToolMaterialNameById(t.materialId)}</td>
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
                        <th class='p-2 text-left'>Werkstoff</th>
                        <th class='p-2 text-left'>Artikelnummer</th>
                        <th class='p-2 text-left'>Bestand</th>
                        <th class='p-2 text-left'>Menge</th>
                      </tr>
                    </thead>
                    <tbody>${selectedManufacturerRows || '<tr><td class="p-2" colspan="7">Keine bestellrelevanten Werkzeuge für diesen Hersteller.</td></tr>'}</tbody>
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
        <td class='p-2'>${getToolMaterialNameById(t.materialId)}</td>
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

  const stockTableHeader = isAdmin
    ? `<tr>
        <th class='p-2 text-left'>T</th>
        <th class='p-2 text-left'>Bild</th>
        <th class='p-2 text-left'>Bezeichnung</th>
        <th class='p-2 text-left'>Ø</th>
        <th class='p-2 text-left'>Werkstoff</th>
        <th class='p-2 text-left'>Fach</th>
        <th class='p-2 text-left'>Artikel Nr.</th>
        <th class='p-2 text-left'>Aufnahme</th>
        <th class='p-2 text-left'>Bestand</th>
        <th class='p-2 text-left'>Min</th>
        <th class='p-2 text-left'>Hersteller</th>
        <th class='p-2 text-left'>Status</th>
        <th class='p-2'></th>
      </tr>`
    : `<tr>
        <th class='p-2 text-left'>T</th>
        <th class='p-2 text-left'>Bild</th>
        <th class='p-2 text-left'>Bezeichnung</th>
        <th class='p-2 text-left'>Ø</th>
        <th class='p-2 text-left'>Fach</th>
        <th class='p-2 text-left'>Artikel Nr.</th>
        <th class='p-2 text-left'>Bestand</th>
        <th class='p-2 text-left'>Min</th>
        <th class='p-2 text-left'>Status</th>
        <th class='p-2'></th>
      </tr>`;

  return `<div class='bg-white rounded-xl shadow p-4 space-y-4'>
    <div class='flex items-center justify-between gap-2'>
      <h2 class='text-xl font-bold mb-1'>Werkzeugverwaltung</h2>
      ${helpButton("werkzeuge")}
    </div>
    <p class='text-sm text-slate-600 mb-2'>Alle Bereiche sind getrennt dargestellt: Stammdaten, Bestand, To-Do, Bestellt und Journal.</p>

    ${
      isAdmin
        ? `<div class='border-2 border-slate-300 rounded-xl p-3 bg-slate-50'>
            <div class='flex items-center justify-between gap-3 flex-wrap'>
              <div>
                <h3 class='text-lg font-bold mb-1'>1) Neues Werkzeug anlegen</h3>
                <p class='text-sm text-slate-500'>Die Dateneingabe öffnet sich in einem separaten Eingabefenster.</p>
              </div>
              <div class='flex gap-2 flex-wrap'>
                <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='openCreateToolModal()'>Neues Werkzeug erfassen</button>
              </div>
            </div>
          </div>
          ${renderToolMasterDataAdmin()}`
        : ""
    }

    <div class='border-2 border-slate-300 rounded-xl p-3'>
      <div class='flex items-center justify-between gap-3 flex-wrap mb-3'>
        <h3 class='text-lg font-bold'>2) Filter & Suche</h3>
        <span class='text-sm text-slate-500'>${tools.length} von ${state.tools.length} Werkzeugen angezeigt</span>
      </div>
      <div class='grid lg:grid-cols-[2fr,1fr,1fr,1fr,1fr,1fr] md:grid-cols-3 gap-3 mb-3'>
        <input id='toolSearch' class='border rounded p-2 md:col-span-2 lg:col-span-1' placeholder='Suche nach T, Bezeichnung, Artikel, Aufnahme, Hersteller...' value='${escapeHtml(filters.search || "")}' />
        <select id='toolFilterLabel' class='border rounded p-2'>
          <option value='' ${filterLabel ? "" : "selected"}>Alle Bezeichnungen</option>${filterLabelOptions}
        </select>
        <input id='toolFilterT' class='border rounded p-2' placeholder='T-Nummer' value='${escapeHtml(filterT)}' />
        <input id='toolFilterD' class='border rounded p-2' placeholder='Durchmesser' value='${escapeHtml(filterD)}' />
        <select id='toolFilterHolder' class='border rounded p-2'>
          <option value='' ${filterHolder ? "" : "selected"}>Alle Aufnahmen</option>
          <option value='HSK 100' ${filterHolder === "HSK 100" ? "selected" : ""}>HSK 100</option>
          <option value='HSK 63' ${filterHolder === "HSK 63" ? "selected" : ""}>HSK 63</option>
        </select>
        <select id='toolFilterImageStatus' class='border rounded p-2'>
          <option value='' ${imageStatus ? "" : "selected"}>Alle Bildstatus</option>
          <option value='withPath' ${imageStatus === "withPath" ? "selected" : ""}>Mit Bildpfad</option>
          <option value='withoutPath' ${imageStatus === "withoutPath" ? "selected" : ""}>Ohne Bildpfad</option>
        </select>
      </div>
      <div class='flex gap-2 flex-wrap'>
        <button class='px-3 py-2 rounded bg-slate-900 text-white' onclick='applyToolFilters()'>Filter anwenden</button>
        <button class='px-3 py-2 rounded bg-slate-700 text-white' onclick='resetToolFilters()'>Filter zurücksetzen</button>
      </div>
      <p class='text-xs text-slate-500 mt-2'>Bildstatus prüft nur, ob aus T-Nummer und Aufnahme ein Bildpfad erstellt werden kann.</p>
    </div>

    <div class='border-2 border-slate-300 rounded-xl p-3'>
      <h3 class='text-lg font-bold mb-2'>3) Werkzeugbestand</h3>
      <div class='overflow-auto max-h-[35vh] border rounded-lg'>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100 sticky top-0'>
            ${stockTableHeader}
          </thead>
          <tbody>${toolRows || `<tr><td class="p-2" colspan="${isAdmin ? 13 : 10}">Keine Werkzeuge.</td></tr>`}</tbody>
        </table>
      </div>
    </div>

    ${
      isAdmin
        ? `<div class='border-2 border-slate-300 rounded-xl p-3'>
            <h3 class='text-lg font-bold mb-2'>4) Admin-Bestandsaktionen</h3>
            <div class='grid md:grid-cols-2 gap-3'>
              <div class='border rounded p-3 bg-white overflow-auto'>
                <h4 class='font-semibold mb-2'>Werkzeug To-Do (bei Mindestbestand oder darunter)</h4>
                <table class='w-full text-sm'>
                  <thead class='bg-slate-100'>
                    <tr>
                      <th class='p-2 text-left'>T</th>
                      <th class='p-2 text-left'>Bezeichnung</th>
                      <th class='p-2 text-left'>Größe</th>
                      <th class='p-2 text-left'>Werkstoff</th>
                      <th class='p-2 text-left'>Artikelnummer</th>
                      <th class='p-2'></th>
                    </tr>
                  </thead>
                  <tbody>${todoRows || '<tr><td class="p-2" colspan="6">Keine offenen To-Dos.</td></tr>'}</tbody>
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
                <thead class='bg-slate-100'>
                  <tr>
                    <th class='p-2 text-left'>Bezeichnung</th>
                    <th class='p-2 text-left'>Größe</th>
                    <th class='p-2 text-left'>Werkstoff</th>
                    <th class='p-2 text-left'>Aktuell optimal</th>
                    <th class='p-2 text-left'>Vorschlag</th>
                    <th class='p-2'></th>
                  </tr>
                </thead>
                <tbody>${suggestionRows || '<tr><td class="p-2" colspan="6">Noch keine aussagekräftigen Vorschläge vorhanden.</td></tr>'}</tbody>
              </table>
            </div>
          </div>`
        : ""
    }

    <div class='border-2 border-slate-300 rounded-xl p-3'>
      <h3 class='text-lg font-bold mb-2'>5) Schichtjournal – Werkzeugwechsel</h3>
      <div class='overflow-auto max-h-[25vh]'>
        <table class='w-full text-sm'>
          <thead class='bg-slate-100 sticky top-0'>
            <tr>
              <th class='p-2 text-left'>Zeit</th>
              <th class='p-2 text-left'>Wer</th>
              <th class='p-2 text-left'>T-Nr</th>
              <th class='p-2 text-left'>Was</th>
              <th class='p-2'></th>
            </tr>
          </thead>
          <tbody>${journalRows || '<tr><td class="p-2" colspan="5">Keine Einträge.</td></tr>'}</tbody>
        </table>
      </div>
    </div>

    ${orderListPopup}
  </div>`;
}

function toggleInsertToolFields() {
  toggleInsertToolFieldsById(
    "toolInsertTool",
    "toolInsertEdges",
    "toolInsertRadius",
  );
}

function toggleInsertToolFieldsById(checkboxId, edgesId, radiusId = "") {
  const checkbox = document.getElementById(checkboxId);
  const edges = document.getElementById(edgesId);
  const radius = radiusId ? document.getElementById(radiusId) : null;
  const radiusWrap = radiusId ? document.getElementById(`${radiusId}Wrap`) : null;
  if (!checkbox || !edges) return;

  const checked = !!checkbox.checked;
  edges.disabled = !checked;
  if (!checked) edges.value = "";

  if (radius) {
    radius.disabled = !checked;
    if (!checked) radius.value = "";
  }
  if (radiusWrap) radiusWrap.style.display = checked ? "" : "none";
}

function taskAssignee(task) {
  return task.assignee || task.assignedTo || task.assigned_to || "";
}

function taskDueDate(task) {
  return task.dueDate || task.due_date || "";
}

function taskCompletedAt(task) {
  return task.completedAt || task.completed_at || null;
}

function normalizeNameForCompare(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function currentUserNameForCompare() {
  return normalizeNameForCompare(
    currentUser?.name ||
      currentEmployeeRecord?.display_name ||
      currentEmployeeRecord?.name ||
      "",
  );
}

function isTaskAssignedToCurrentUser(task) {
  return (
    normalizeNameForCompare(taskAssignee(task)) === currentUserNameForCompare()
  );
}

function renderTodo() {
  const nowIso = new Date().toISOString();
  const isAdmin = currentUser?.role === "admin";
  const allTasks = Array.isArray(state.tasks) ? state.tasks : [];
  console.log("renderTodo allTasks:", allTasks);

  const visibleTasks = isAdmin
    ? allTasks
    : allTasks.filter((task) => isTaskAssignedToCurrentUser(task));

  console.log("To-Do Render Debug:", {
    user: currentUser,
    employeeRecord: currentEmployeeRecord,
    tasks: state.tasks,
    visibleCount: visibleTasks?.length,
  });

  const openTasks = visibleTasks.filter((task) => task.status !== "done");
  const doneTasks = visibleTasks.filter((task) => task.status === "done");

  const openRows = openTasks
    .map((task) => {
      const dueDate = taskDueDate(task);
      const completedAt = taskCompletedAt(task);
      const overdue = dueDate && nowIso > `${dueDate}T23:59:59`;
      return `<tr class='border-b ${overdue ? "bg-rose-50" : ""}'>
      <td class='p-2'>${task.title}</td><td class='p-2'>${taskAssignee(task)}</td><td class='p-2'>${dueDate ? formatDateDisplay(dueDate) : "-"}</td>
      <td class='p-2'>${completedAt ? "Erledigt" : "Offen"}</td>
      <td class='p-2'>
        ${
          isAdmin
            ? `<button class='px-2 py-1 rounded bg-slate-900 text-white mr-1' onclick="deleteTask('${task.id}')">Löschen</button>
          <button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="reassignTaskPrompt('${task.id}')">Neu zuweisen</button>`
            : `<button class='px-2 py-1 rounded bg-emerald-600 text-white' onclick="completeTask('${task.id}')">Erledigt</button>`
        }
      </td>
    </tr>`;
    })
    .join("");

  const doneRows = doneTasks
    .map((task) => {
      const dueDate = taskDueDate(task);
      const completedAt = taskCompletedAt(task);
      return `<tr class='border-b bg-emerald-50'>
      <td class='p-2'>✅ ${task.title}</td><td class='p-2'>${taskAssignee(task)}</td><td class='p-2'>${dueDate ? formatDateDisplay(dueDate) : "-"}</td><td class='p-2'>${completedAt || "-"}</td>
      <td class='p-2'>${isAdmin ? `<button class='px-2 py-1 rounded bg-slate-900 text-white' onclick="deleteTask('${task.id}')">Löschen</button>` : "-"}</td>
    </tr>`;
    })
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

async function createTask() {
  if (!supabaseReady) return;

  const title = document.getElementById("taskTitle")?.value?.trim();
  const assignee = document.getElementById("taskAssignee")?.value;
  const deadline = document.getElementById("taskDeadline")?.value;
  if (!title || !assignee)
    return alert("Bitte Aufgabe und Mitarbeiter angeben.");

  const payload = {
    title,
    description: "",
    assigned_to: assignee,
    due_date: deadline || null,
    status: "open",
    created_by_employee_id: currentEmployeeRecord?.id || null,
  };

  const { data: inserted, error } = await supabaseClient
    .from("planner_tasks")
    .insert(payload)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Speichern der Aufgabe:", error);
    return alert(`Aufgabe konnte nicht gespeichert werden: ${error.message}`);
  }

  state.tasks.push(normalizeTaskFromDb(inserted));
  persist();
  render();
}

async function completeTask(taskId) {
  if (!supabaseReady) return;

  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) return;

  const completedAt = new Date().toISOString();

  const { data: updated, error } = await supabaseClient
    .from("planner_tasks")
    .update({
      status: "done",
      completed_at: completedAt,
      updated_at: completedAt,
    })
    .eq("id", taskId)
    .select()
    .single();

  if (error) {
    console.error("Fehler beim Abschließen der Aufgabe:", error);
    return alert(`Aufgabe konnte nicht abgeschlossen werden: ${error.message}`);
  }

  task.status = "done";
  task.completedAt = updated?.completed_at || completedAt;
  persist();
  render();
}

async function deleteTask(taskId) {
  if (!supabaseReady) return;

  const { error } = await supabaseClient
    .from("planner_tasks")
    .delete()
    .eq("id", taskId);

  if (error) {
    console.error("Fehler beim Löschen der Aufgabe:", error);
    return alert(`Aufgabe konnte nicht gelöscht werden: ${error.message}`);
  }

  state.tasks = state.tasks.filter((t) => t.id !== taskId);
  persist();
  render();
}

async function reassignTaskPrompt(taskId) {
  if (!supabaseReady) return;

  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) return;

  const currentAssignee = taskAssignee(task);
  const options = activeUsers()
    .map((user) => {
      const selected = user.name === currentAssignee ? "selected" : "";
      return `<option value="${escapeHtml(user.name)}" ${selected}>${escapeHtml(user.name)}</option>`;
    })
    .join("");

  const host = getModalHost();
  host.innerHTML = `<div class="fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-xl w-full max-w-md p-4">
      <h3 class="text-lg font-bold mb-3">Aufgabe neu zuweisen</h3>
      <p class="text-sm text-slate-600 mb-3">${escapeHtml(task.title)}</p>
      <select id="taskReassignSelect" class="border rounded p-2 w-full mb-4">${options}</select>
      <div class="flex justify-end gap-2">
        <button class="px-3 py-1 rounded bg-slate-200" onclick="closeTaskReassignModal()">Abbrechen</button>
        <button class="px-3 py-1 rounded bg-slate-900 text-white" onclick="saveTaskReassignment('${taskId}')">Speichern</button>
      </div>
    </div>
  </div>`;
}

function closeTaskReassignModal() {
  getModalHost().innerHTML = "";
}

async function saveTaskReassignment(taskId) {
  if (!supabaseReady) return;

  const task = state.tasks.find((t) => t.id === taskId);
  if (!task) return;

  const selectedName = document.getElementById("taskReassignSelect")?.value;
  if (!selectedName) return alert("Bitte Mitarbeiter auswählen.");

  const updatedAt = new Date().toISOString();

  const { error } = await supabaseClient
    .from("planner_tasks")
    .update({
      assigned_to: selectedName,
      updated_at: updatedAt,
    })
    .eq("id", taskId);

  if (error) {
    console.error("Fehler beim Neuzuweisen der Aufgabe:", error);
    return alert(
      `Aufgabe konnte nicht neu zugewiesen werden: ${error.message}`,
    );
  }

  task.assignee = selectedName;
  task.assignedTo = selectedName;
  closeTaskReassignModal();
  persist();
  render();
}

function maybeTaskReminder() {
  if (!currentUser || currentUser.role !== "employee") return;
  const now = new Date();
  const today = isoDate(now);
  const hourKey = `${today}:${now.getHours()}:${currentUser.name}`;
  if (state[`_taskReminder_${hourKey}`]) return;
  const dueToday = (state.tasks || []).filter(
    (task) =>
      isTaskAssignedToCurrentUser(task) &&
      task.status !== "done" &&
      taskDueDate(task) === today,
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
      <div class='flex items-center gap-2'>
        <button class='px-2 py-1 rounded ${statsViewPeriod === "week" ? "bg-slate-900 text-white" : "bg-slate-100"}' onclick="setStatsView('week')">Woche</button>
        <button class='px-2 py-1 rounded ${statsViewPeriod === "month" ? "bg-slate-900 text-white" : "bg-slate-100"}' onclick="setStatsView('month')">Monat</button>
        <button class='px-2 py-1 rounded ${statsViewPeriod === "year" ? "bg-slate-900 text-white" : "bg-slate-100"}' onclick="setStatsView('year')">Jahr</button>
        ${helpButton("statistik")}
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

async function markAbsent(shiftId, date, userName) {
  if (!supabaseReady) return;

  const absenceKey = `${date}:${userName}`;

  const { error } = await supabaseClient.from("planner_absences").upsert(
    {
      absence_key: absenceKey,
      absence_date: date,
      user_name: userName,
      absence_type: "abwesend",
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "absence_key" },
  );

  if (error) {
    console.error("Fehler bei Abwesenheit:", error);
    return alert(
      `Abwesenheit konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  // lokal weiterführen (für UI)
  state.absences[absenceKey] = true;

  persist();
  render();
}

async function assignOptionalShift(shiftId) {
  if (!supabaseReady) return;

  const shift = getShiftById(shiftId);
  if (!shift) return alert("Schicht konnte nicht gefunden werden.");

  if (!isPrimaryCoreAbsentForShift(shift)) {
    return alert(
      "Einteilung nicht möglich: Der A/B/C-Mitarbeiter ist nicht abwesend gemeldet.",
    );
  }

  const select = document.getElementById(`opt-${shiftId}`);
  if (!select) return;

  const name = select.value;
  if (!name) return alert("Bitte Springer auswählen.");

  if (!isSpringer(name)) {
    return alert("Für diese Funktion dürfen nur Springer eingeteilt werden.");
  }

  if (!canAssignUserToShift(name, shiftId)) return;

  const { error } = await supabaseClient.from("planner_assignments").upsert(
    {
      shift_id: shiftId,
      shift_date: shift.date,
      assigned_user: name,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "shift_id" },
  );

  if (error) {
    console.error("Fehler bei Springer-Einteilung:", error);
    return alert(
      `Einteilung konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = name;

  persist();
  render();
}

async function assignSuggestedSpringer(shiftId) {
  if (!supabaseReady) return;

  const shift = getShiftById(shiftId);
  if (!shift) return alert("Schicht konnte nicht gefunden werden.");

  if (!isPrimaryCoreAbsentForShift(shift)) {
    return alert(
      "Einteilung nicht möglich: Der A/B/C-Mitarbeiter ist nicht abwesend gemeldet.",
    );
  }

  const suggested = getSuggestedSpringerForShift(shift);
  if (!suggested) {
    return alert("Kein verfügbarer Springer gefunden.");
  }

  if (!canAssignUserToShift(suggested, shiftId)) return;

  const { error } = await supabaseClient.from("planner_assignments").upsert(
    {
      shift_id: shiftId,
      shift_date: shift.date,
      assigned_user: suggested,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "shift_id" },
  );

  if (error) {
    console.error("Fehler bei automatischer Springer-Einteilung:", error);
    return alert(
      `Einteilung konnte nicht gespeichert werden: ${error.message}`,
    );
  }

  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = suggested;

  persist();
  render();
}

async function approveSaturdayRequest(shiftId, user) {
  if (!supabaseReady) return;
  if (!canAssignUserToShift(user, shiftId)) return;

  const shift = getShiftById(shiftId);
  if (!shift) return alert("Schicht konnte nicht gefunden werden.");

  const { error } = await supabaseClient.from("planner_assignments").upsert(
    {
      shift_id: shiftId,
      shift_date: shift.date,
      assigned_user: user,
      created_by_employee_id: currentEmployeeRecord?.id || null,
    },
    { onConflict: "shift_id" },
  );

  if (error) {
    console.error("Fehler beim Bestätigen der Samstags-Anfrage:", error);
    return alert(
      `Samstags-Anfrage konnte nicht übernommen werden: ${error.message}`,
    );
  }

  const [deleteRequest, deleteCancellation] = await Promise.all([
    supabaseClient
      .from("planner_saturday_requests")
      .delete()
      .eq("request_key", `${shiftId}:${user}`),
    supabaseClient
      .from("planner_shift_cancellations")
      .delete()
      .eq("shift_id", shiftId),
  ]);

  const firstError = deleteRequest.error || deleteCancellation.error;
  if (firstError) {
    console.error("Fehler beim Abschließen der Samstags-Anfrage:", firstError);
    return alert(
      `Einteilung gespeichert, aber Anfrage/Ausfall konnte nicht bereinigt werden: ${firstError.message}`,
    );
  }

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
window.openHelp = openHelp;
window.closeHelp = closeHelp;
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
window.openAbsenceReplacementPlanner = openAbsenceReplacementPlanner;
window.closeReplacementPlanner = closeReplacementPlanner;
window.setReplacementPlannerChoice = setReplacementPlannerChoice;
window.toggleReplacementDay = toggleReplacementDay;
window.applyReplacementForSelectedDays = applyReplacementForSelectedDays;
window.applyReplacementForWeek = applyReplacementForWeek;
window.applyReplacementForAll = applyReplacementForAll;
window.addSwap = addSwap;
window.deleteSwap = deleteSwap;
window.resetActiveSwaps = resetActiveSwaps;
window.resetManualAssignments = resetManualAssignments;
window.resetPlanCurrentFuture = resetPlanCurrentFuture;
window.updateSlotAssignment = updateSlotAssignment;
window.queueEmployeeEdit = queueEmployeeEdit;
window.queueEmployeeActive = queueEmployeeActive;
window.queueSlotAssignment = queueSlotAssignment;
window.saveAllPersonnelChanges = saveAllPersonnelChanges;
window.addEmployee = addEmployee;
window.updateEmployee = updateEmployee;
window.fillEmployeeForm = fillEmployeeForm;
window.deactivateEmployee = deactivateEmployee;
window.activateEmployee = activateEmployee;
window.requestSaturdayEvening = requestSaturdayEvening;
window.setWeekendAvailability = setWeekendAvailability;
window.createTask = createTask;
window.completeTask = completeTask;
window.deleteTask = deleteTask;
window.reassignTaskPrompt = reassignTaskPrompt;
window.closeTaskReassignModal = closeTaskReassignModal;
window.saveTaskReassignment = saveTaskReassignment;
window.setConflictResolved = setConflictResolved;
window.stopDowntimeTimer = stopDowntimeTimer;
window.applyToolFilters = applyToolFilters;
window.resetToolFilters = resetToolFilters;
window.bookToolChange = bookToolChange;
window.undoToolJournalEntry = undoToolJournalEntry;
window.openToolImagePopup = openToolImagePopup;
window.closeToolImagePopup = closeToolImagePopup;
window.openManualToolWithdraw = openManualToolWithdraw;
window.openManualToolRestock = openManualToolRestock;
window.openQrScannerPlaceholder = openQrScannerPlaceholder;
window.openQrScanner = openQrScanner;
window.stopQrScanner = stopQrScanner;
window.processScannedToolQr = processScannedToolQr;
window.confirmQrToolWithdraw = confirmQrToolWithdraw;
window.updateScannerWithdrawPreview = updateScannerWithdrawPreview;
window.updateScannerRestockPreview = updateScannerRestockPreview;
window.openToolQrPopup = openToolQrPopup;
window.copyToolQrPayload = copyToolQrPayload;
window.printToolQrLabel = printToolQrLabel;
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
window.updateEditToolTypeFields = updateEditToolTypeFields;
window.updateEditThreadPitchVisibility = updateEditThreadPitchVisibility;
window.addToolMaterial = addToolMaterial;
window.renameToolMaterial = renameToolMaterial;
window.deactivateToolMaterial = deactivateToolMaterial;
window.clearManualShift = clearManualShift;
window.assignSuggestedSpringer = assignSuggestedSpringer;
window.deleteManualAbsence = deleteManualAbsence;
