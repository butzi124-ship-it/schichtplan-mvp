const USERS = [
  { name: "Lavdrim", slot: "A", type: "core" },
  { name: "Roger", slot: "B", type: "core" },
  { name: "Dashmir", slot: "C", type: "core" },
  { name: "Thomas", slot: "D", type: "springer" },
  { name: "Musa", slot: "E", type: "springer" },
  { name: "Ardian", slot: "F", type: "springer" },
];

const STORAGE_KEY = "schichtplan_mvp_v1";
const state = loadState();
let currentUser = null;
let currentTab = "schichtplan";

const WEEK_TEMPLATES = [
  makeWeek("A", "B", "C", "A", "B"),
  makeWeek("B", "C", "A", "B", "C"),
  makeWeek("C", "A", "B", "C", "A"),
  makeWeek("A", "B", "C", "A", "B"),
  makeWeek("B", "C", "A", "B", "C"),
  makeWeek("C", "A", "B", "C", "A"),
];
const ROTATION_ANCHOR_MONDAY = "2026-01-05";

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
  };
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return base;
    return { ...base, ...JSON.parse(raw) };
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
  currentUser = null;
  currentTab = "schichtplan";
  document.getElementById("loginBox").classList.remove("hidden");
  document.getElementById("tabs").classList.add("hidden");
  document.getElementById("sessionInfo").textContent = "";
  document.getElementById("view").innerHTML = "";
}

function render() {
  if (!currentUser) return;
  const tabs = ["schichtplan"];
  if (currentUser.role === "admin") tabs.push("planung");
  if (currentUser.role === "employee") tabs.push("verfuegbarkeit", "abschluss");

  const tabsEl = document.getElementById("tabs");
  tabsEl.className = "flex gap-2 flex-wrap";
  tabsEl.innerHTML =
    tabs
      .map(
        (t) =>
          `<button class="px-3 py-2 rounded ${currentTab === t ? "bg-slate-900 text-white" : "bg-white"}" onclick="setTab('${t}')">${labelTab(t)}</button>`,
      )
      .join("") +
    `<button class="px-3 py-2 rounded bg-red-100" onclick="logout()">Abmelden</button>`;

  document.getElementById("sessionInfo").textContent =
    `Angemeldet: ${currentUser.name} (${currentUser.role})`;

  if (!tabs.includes(currentTab)) currentTab = tabs[0];
  const view = document.getElementById("view");
  if (currentTab === "schichtplan") view.innerHTML = renderSchedule();
  if (currentTab === "planung") view.innerHTML = renderPlanning();
  if (currentTab === "verfuegbarkeit") view.innerHTML = renderAvailability();
  if (currentTab === "abschluss") view.innerHTML = renderClosure();
}

function labelTab(tab) {
  return {
    schichtplan: "Schichtplan",
    planung: "Planung (Admin)",
    verfuegbarkeit: "Unklare Schichten",
    abschluss: "Schichtabschluss",
  }[tab];
}

function setTab(tab) {
  currentTab = tab;
  render();
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

function buildShift(date, id, template) {
  const defaultAssigned = chooseDefault(template.options);
  const swappedDefault = defaultAssigned
    ? applySwap(date, defaultAssigned)
    : defaultAssigned;
  const isOptional = template.options.length > 1;
  const manualAssigned = state.assignments[id] || null;
  const absenceKey = `${date}:${swappedDefault}`;
  const absent =
    !!state.absences[absenceKey] || hasCalendarAbsence(swappedDefault, date);
  const canceled = !!state.shiftCancellations[id];
  const assigned = canceled
    ? null
    : !isOptional && !absent
      ? swappedDefault
      : manualAssigned || swappedDefault;
  return {
    id,
    date,
    label: template.label,
    start: template.start,
    end: template.end,
    options: template.options,
    assigned,
    open: absent || !assigned || assigned === "NONE",
  };
}

function chooseDefault(options) {
  const slot = options[0];
  if (slot === "NONE") return null;
  return slotToName(slot);
}

function slotToName(slot) {
  return USERS.find((u) => u.slot === slot)?.name || null;
}

function userByName(name) {
  return USERS.find((u) => u.name === name);
}

function isCoreEmployee(name) {
  return userByName(name)?.type === "core";
}

function slotOfUser(name) {
  return USERS.find((u) => u.name === name)?.slot || "-";
}

const PERSON_COLORS = {
  Lavdrim: "bg-green-100 text-green-900",
  Roger: "bg-violet-100 text-violet-900",
  Dashmir: "bg-rose-100 text-rose-900",
  Thomas: "bg-cyan-100 text-cyan-900",
  Musa: "bg-amber-100 text-amber-900",
  Ardian: "bg-indigo-100 text-indigo-900",
};

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

function inRange(date, from, to) {
  return date >= from && date <= to;
}

function applySwap(date, name) {
  let resolved = name;
  state.swaps.forEach((swap) => {
    if (!swap || date < swap.startDate) return;
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

function resolveAssigned(shiftId, options) {
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
    : manualAssigned || swappedDefault;
}

function assignedDisplay(shiftId, options) {
  const assignedName = resolveAssigned(shiftId, options);
  if (!assignedName) return "-";
  return `${slotOfUser(assignedName)} • ${assignedName}`;
}

function assignedMeta(shiftId, options) {
  const assignedName = resolveAssigned(shiftId, options);
  if (!assignedName) return { label: "-", cls: "bg-slate-100 text-slate-700" };
  return {
    label: `${slotOfUser(assignedName)} • ${assignedName}`,
    cls: PERSON_COLORS[assignedName] || "bg-slate-100 text-slate-800",
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

    // Start-Zeile (Sonntagabend vor Woche)
    if (w === 0) {
      const startSunday = new Date(weekStart);
      startSunday.setDate(startSunday.getDate() - 1);
      const startMeta = assignedMeta(
        `${isoDate(startSunday)}-su-1`,
        template.sunday[1].options,
      );
      body += `<tr class="border-b bg-amber-50">
        <td class="p-2"></td>
        <td class="p-2 font-semibold">Start (Sonntag)</td>
        <td class="p-2 ${startMeta.cls} font-semibold text-center">${startMeta.label}</td>
        <td class="p-2 font-semibold text-center">18:00-24:00</td>
        <td class="p-2 bg-slate-50" colspan="4"></td>
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
          <td class="p-2">${dayName}<div class="text-[11px] text-slate-500">${dateIso}</div></td>
          <td class="p-2 ${m1.cls} font-semibold text-center">${m1.label}</td>
          <td class="p-2 font-semibold text-center bg-white">05:00-11:00</td>
          <td class="p-2 ${m2.cls} font-semibold text-center">${m2.label}</td>
          <td class="p-2 font-semibold text-center bg-white">13:00-19:00</td>
          <td class="p-2 ${m3.cls} font-semibold text-center">${m3.label}</td>
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
          <td class="p-2 font-semibold">${dayName}<div class="text-[11px] text-slate-500">${dateIso}</div></td>
          <td class="p-2 ${m1.cls} font-semibold text-center">${m1.label}</td>
          <td class="p-2 font-semibold text-center bg-white">05:00-11:00</td>
          <td class="p-2 ${m2.cls} font-semibold text-center">${m2.label}</td>
          <td class="p-2 font-semibold text-center bg-white">16:00-22:00</td>
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
          <td class="p-2 font-semibold">${dayName}<div class="text-[11px] text-slate-500">${dateIso}</div></td>
          <td class="p-2 ${m1.cls} font-semibold text-center">${m1.label}</td>
          <td class="p-2 font-semibold text-center bg-white">06:00-12:00</td>
          <td class="p-2 bg-slate-50" colspan="2"></td>
          <td class="p-2 ${m2.cls} font-semibold text-center">${m2.label}</td>
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
  const shifts = generateThreeMonths();
  const visibleShifts =
    currentUser.role === "admin"
      ? shifts
      : isCoreEmployee(currentUser.name)
        ? shifts
        : shifts.filter((s) => s.assigned === currentUser.name);
  const rows = visibleShifts
    .slice(0, 20)
    .map((s) => {
      const canAbsent = currentUser.role === "employee";
      const status = s.open
        ? '<span class="text-red-600 font-semibold">OFFEN</span>'
        : '<span class="text-emerald-700">Besetzt</span>';
      return `<tr class="border-b">
        <td class="p-2">${s.date}</td><td class="p-2">${s.label}</td><td class="p-2">${s.start}–${s.end}</td>
        <td class="p-2">${s.assigned || "-"}</td><td class="p-2">${status}</td>
        <td class="p-2">${canAbsent ? `<button class='px-2 py-1 bg-amber-200 rounded' onclick="markAbsent('${s.id}','${s.date}','${currentUser.name}')">Kann nicht kommen</button>` : ""}</td>
      </tr>`;
    })
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-3'>Schichtplan (automatisch 90 Tage ab aktueller Woche)</h2>
    ${renderOverviewPlan(Math.ceil(90 / 7))}
    <h3 class='text-md font-semibold mt-6 mb-2'>Meine nächsten Einsätze (Aktionen)</h3>
    <div class='overflow-auto max-h-[40vh] border rounded-lg'>
      <table class='w-full text-sm'><thead class='sticky top-0 bg-slate-100'><tr>
      <th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Zeit</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Status</th><th class='p-2'></th>
      </tr></thead><tbody>${rows}</tbody></table>
    </div></div>`;
}

function markAbsent(shiftId, date, employeeName) {
  state.absences[`${date}:${employeeName}`] = true;
  persist();
  render();
}

function renderPlanning() {
  const openShifts = generateThreeMonths().filter((s) => s.open);
  const optionalShifts = generateThreeMonths().filter(
    (s) => s.options.length > 1,
  );
  const rows = openShifts
    .map((s) => {
      const choices = USERS.filter((u) => u.name !== s.assigned)
        .map((u) => `<option value="${u.name}">${u.name}</option>`)
        .join("");
      return `<tr class='border-b'>
      <td class='p-2'>${s.date}</td><td class='p-2'>${s.label}</td><td class='p-2'>${s.start}–${s.end}</td>
      <td class='p-2'>${s.assigned || "-"}</td>
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

  const optionalRows = optionalShifts
    .slice(0, 120)
    .map((s) => {
      const canUsers = USERS.filter(
        (u) => state.availability[`${s.id}:${u.name}`] === "yes",
      ).map((u) => u.name);
      const options = (canUsers.length ? canUsers : USERS.map((u) => u.name))
        .map((name) => `<option value="${name}">${name}</option>`)
        .join("");
      return `<tr class='border-b'>
      <td class='p-2'>${s.date}</td>
      <td class='p-2'>${s.label}</td>
      <td class='p-2'>${s.options.map((o) => (o === "NONE" ? "0" : o)).join("/")}</td>
      <td class='p-2'>${canUsers.length ? canUsers.join(", ") : "-"}</td>
      <td class='p-2'><select id='opt-${s.id}' class='border rounded p-1'>${options}</select></td>
      <td class='p-2'><button class='px-2 py-1 rounded bg-blue-700 text-white' onclick="assignOptionalShift('${s.id}')">Einteilen</button></td>
    </tr>`;
    })
    .join("");

  const absenceRows = Object.entries(state.absences)
    .slice(-100)
    .reverse()
    .map(([key]) => {
      const [date, user] = key.split(":");
      return `<tr class='border-b'><td class='p-2'>${date}</td><td class='p-2'>${user}</td><td class='p-2'>Kann nicht kommen</td></tr>`;
    })
    .join("");

  const monthValue = `${new Date().getFullYear()}-${String(new Date().getMonth() + 1).padStart(2, "0")}`;
  const personOptions = USERS.map(
    (u) => `<option value='${u.name}'>${u.name}</option>`,
  ).join("");
  const vacationRows = state.vacations
    .slice(-20)
    .reverse()
    .map(
      (v) =>
        `<tr class='border-b'><td class='p-2'>${v.user}</td><td class='p-2'>${v.from}</td><td class='p-2'>${v.to}</td></tr>`,
    )
    .join("");
  const sickRows = state.sickLeaves
    .slice(-20)
    .reverse()
    .map(
      (v) =>
        `<tr class='border-b'><td class='p-2'>${v.user}</td><td class='p-2'>${v.from}</td><td class='p-2'>${v.to}</td></tr>`,
    )
    .join("");

  return `<div class='bg-white rounded-xl shadow p-4'>
    <h2 class='text-lg font-semibold mb-3'>Planung (Admin)</h2>
    <p class='text-sm text-slate-500 mb-4'>Struktur: Abwesenheiten → Urlaube/Krankmeldungen → Vertretung und Tausch.</p>

    <div class='grid md:grid-cols-2 gap-4 mb-6'>
      <div class='border rounded-lg p-3 bg-slate-50'>
        <h3 class='font-semibold mb-2'>Meldungen „Kann nicht kommen“</h3>
        <div class='overflow-auto max-h-56'>
          <table class='w-full text-sm'><thead class='bg-white sticky top-0'><tr><th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Typ</th></tr></thead>
          <tbody>${absenceRows || '<tr><td class=\"p-2\" colspan=\"3\">Keine Meldungen.</td></tr>'}</tbody></table>
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
        <div class='flex gap-2'>
          <button class='px-2 py-1 rounded bg-emerald-700 text-white' onclick='addVacation()'>Urlaub speichern</button>
          <button class='px-2 py-1 rounded bg-amber-700 text-white' onclick='addSickLeave()'>Krank speichern</button>
        </div>
      </div>
    </div>

    <div class='grid md:grid-cols-2 gap-4 mb-6'>
      <div class='border rounded-lg p-3'>
        <h4 class='font-semibold mb-2'>Geplante Urlaube</h4>
        <div class='overflow-auto max-h-48'>
          <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Von</th><th class='p-2 text-left'>Bis</th></tr></thead>
          <tbody>${vacationRows || '<tr><td class=\"p-2\" colspan=\"3\">Keine Einträge.</td></tr>'}</tbody></table>
        </div>
      </div>
      <div class='border rounded-lg p-3'>
        <h4 class='font-semibold mb-2'>Krankmeldungen</h4>
        <div class='overflow-auto max-h-48'>
          <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr><th class='p-2 text-left'>Mitarbeiter</th><th class='p-2 text-left'>Von</th><th class='p-2 text-left'>Bis</th></tr></thead>
          <tbody>${sickRows || '<tr><td class=\"p-2\" colspan=\"3\">Keine Einträge.</td></tr>'}</tbody></table>
        </div>
      </div>
    </div>

    <h3 class='text-md font-semibold mb-2'>Offene Schichten / Übernahme</h3>
    <p class='text-sm text-slate-500 mb-2'>Pro Tag kannst du entscheiden: ganze Schicht übernehmen lassen oder Ausfall markieren.</p>
    <div class='overflow-auto max-h-[50vh] mb-6'>
      <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr>
      <th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Zeit</th><th class='p-2 text-left'>Vorher</th><th class='p-2 text-left'>Übernahme durch</th><th class='p-2'>Aktion</th></tr></thead>
      <tbody>${rows || '<tr><td class="p-2" colspan="6">Keine offenen Schichten.</td></tr>'}</tbody></table>
    </div>
    <h3 class='text-md font-semibold mt-6 mb-2'>Unklare Schichten mit \"Kann\"-Meldungen</h3>
    <p class='text-sm text-slate-500 mb-2'>Wenn Mitarbeitende bei einer Schicht \"Kann\" melden, kannst du sie hier direkt einteilen. Die Zuweisung erscheint danach im Schichtplan.</p>
    <div class='overflow-auto max-h-[60vh]'>
      <table class='w-full text-sm'><thead class='bg-slate-100 sticky top-0'><tr>
        <th class='p-2 text-left'>Datum</th><th class='p-2 text-left'>Schicht</th><th class='p-2 text-left'>Option</th><th class='p-2 text-left'>Kann</th><th class='p-2 text-left'>Einteilen</th><th class='p-2'></th>
      </tr></thead>
      <tbody>${optionalRows || '<tr><td class=\"p-2\" colspan=\"6\">Keine unklaren Schichten.</td></tr>'}</tbody></table>
    </div>

    <h3 class='text-md font-semibold mt-6 mb-2'>Mitarbeiter-Tausch ab Datum</h3>
    <div class='grid md:grid-cols-4 gap-2 items-end'>
      <label class='text-sm'>Mitarbeiter 1<select id='swapA' class='border rounded p-1 w-full'>${personOptions}</select></label>
      <label class='text-sm'>Mitarbeiter 2<select id='swapB' class='border rounded p-1 w-full'>${personOptions}</select></label>
      <label class='text-sm'>Gültig ab<input id='swapDate' type='date' value='${todayIso()}' class='border rounded p-1 w-full'/></label>
      <button class='px-2 py-1 rounded bg-purple-700 text-white h-9' onclick='addSwap()'>Tausch speichern</button>
    </div>
  </div>`;
}

function assignShift(shiftId) {
  const select = document.getElementById(`sel-${shiftId}`);
  if (!select) return;
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

function assignOptionalShift(shiftId) {
  const select = document.getElementById(`opt-${shiftId}`);
  if (!select) return;
  delete state.shiftCancellations[shiftId];
  state.assignments[shiftId] = select.value;
  persist();
  render();
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
  state.vacations.push(data);
  persist();
  render();
}

function addSickLeave() {
  const data = readAdminCalendarForm();
  if (!data) return;
  state.sickLeaves.push(data);
  persist();
  render();
}

function addSwap() {
  const userA = document.getElementById("swapA")?.value;
  const userB = document.getElementById("swapB")?.value;
  const startDate = document.getElementById("swapDate")?.value;
  if (!userA || !userB || userA === userB || !startDate) {
    alert("Bitte zwei verschiedene Mitarbeiter und ein Startdatum wählen.");
    return;
  }
  state.swaps.push({ userA, userB, startDate });
  persist();
  render();
}

function renderAvailability() {
  const shifts = generateThreeMonths().filter((s) => s.options.length > 1);
  const rows = shifts
    .slice(0, 80)
    .map((s) => {
      const key = `${s.id}:${currentUser.name}`;
      const val = state.availability[key] || "";
      return `<tr class='border-b'><td class='p-2'>${s.date}</td><td class='p-2'>${s.label}</td><td class='p-2'>${s.options.join(" / ")}</td>
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
      return `<tr class='border-b'><td class='p-2'>${s.date}</td><td class='p-2'>${s.label}</td><td class='p-2'>${s.start}–${s.end}</td><td class='p-2'>${done}</td>
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
    const h = wrapper.querySelector("#uh").value;
    const m = wrapper.querySelector("#um").value;
    state.unmanned[shiftId] = `${h}:${m}`;
    persist();
    wrapper.remove();
    render();
  };
}

function checkbox(id, label) {
  return `<label class='flex items-center gap-2'><input id='${id}' type='checkbox' class='w-4 h-4'/> ${label}</label>`;
}

render();
window.loginAs = loginAs;
window.logout = logout;
window.setTab = setTab;
window.markAbsent = markAbsent;
window.assignShift = assignShift;
window.cancelShift = cancelShift;
window.assignOptionalShift = assignOptionalShift;
window.setAvailability = setAvailability;
window.openChecklist = openChecklist;
window.addVacation = addVacation;
window.addSickLeave = addSickLeave;
window.addSwap = addSwap;
