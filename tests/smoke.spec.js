const { test, expect } = require("@playwright/test");
const path = require("path");
const fs = require("fs");
const { pathToFileURL } = require("url");

const appUrl = pathToFileURL(
  path.resolve(__dirname, "..", "schichtplan_mvp.html"),
).href;

const rowsByTable = {
  employees: [
    {
      id: "employee-admin",
      auth_user_id: "auth-admin",
      display_name: "Admin",
      role: "admin",
      employee_type: "admin",
      slot_code: null,
      is_active: true,
      first_name: "Mido",
      last_name: "Admin",
      personnel_no: "A-1",
      department: "",
      color_key: "slate",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "employee-one",
      auth_user_id: null,
      display_name: "Lavdrim",
      role: "employee",
      employee_type: "core",
      slot_code: "A",
      is_active: true,
      first_name: "Lavdrim",
      last_name: "",
      personnel_no: "100",
      department: "",
      color_key: "blue",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  shift_definitions: [
    {
      id: "shift-early",
      name: "Frühschicht",
      shift_key: "early",
      start_time: "06:00",
      end_time: "14:00",
      requires_time: true,
      active: true,
    },
    {
      id: "shift-day",
      name: "Tagschicht",
      shift_key: "day",
      start_time: null,
      end_time: null,
      requires_time: false,
      active: true,
    },
  ],
  departments: [
    {
      id: "department-one",
      name: "Fertigung",
      code: "FERT",
      leader_employee_id: "employee-admin",
      active: true,
    },
  ],
  production_machines: [
    {
      id: "machine-one",
      name: "Maschine 50",
      machine_code: "M50",
      department_id: "department-one",
      active: true,
    },
    {
      id: "machine-two",
      name: "Maschine 51",
      machine_code: "M51",
      department_id: "department-one",
      active: true,
    },
    {
      id: "machine-over",
      name: "Maschine 52",
      machine_code: "M52",
      department_id: "department-one",
      active: true,
    },
    {
      id: "machine-chain",
      name: "Maschine 53",
      machine_code: "M53",
      department_id: "department-one",
      active: true,
    },
  ],
  production_orders: [
    {
      id: "order-one",
      machine_id: "machine-one",
      department_id: "department-one",
      ba_number: "BA-100",
      article_number: "ART-100",
      ba_quantity: 100,
      target_quantity: 2,
      pallet_count: 2,
      pieces_per_pallet: 50,
      use_chain_logic: false,
      status: "running",
      started_at: "2026-01-01T00:00:00Z",
      completed_at: null,
      created_by_employee_id: "employee-admin",
      updated_by_employee_id: "employee-admin",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "order-two",
      machine_id: "machine-two",
      department_id: "department-one",
      ba_number: "BA-200",
      article_number: "ART-200",
      ba_quantity: 50,
      target_quantity: 50,
      pallet_count: 1,
      pieces_per_pallet: 50,
      use_chain_logic: false,
      status: "running",
      started_at: "2026-01-01T00:00:00Z",
      completed_at: null,
      created_by_employee_id: "employee-admin",
      updated_by_employee_id: "employee-admin",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "order-over",
      machine_id: "machine-over",
      department_id: "department-one",
      ba_number: "BA-300",
      article_number: "ART-300",
      ba_quantity: 10,
      target_quantity: 2,
      pallet_count: 1,
      pieces_per_pallet: 10,
      use_chain_logic: false,
      status: "running",
      started_at: "2026-01-01T00:00:00Z",
      completed_at: null,
      created_by_employee_id: "employee-admin",
      updated_by_employee_id: "employee-admin",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "order-chain",
      machine_id: "machine-chain",
      department_id: "department-one",
      ba_number: "BA-400",
      article_number: "ART-400",
      ba_quantity: 10,
      target_quantity: 2,
      pallet_count: 1,
      pieces_per_pallet: 10,
      use_chain_logic: true,
      status: "running",
      started_at: "2026-01-01T00:00:00Z",
      completed_at: null,
      created_by_employee_id: "employee-admin",
      updated_by_employee_id: "employee-admin",
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  production_order_stations: [
    {
      id: "station-one",
      order_id: "order-one",
      station_no: 1,
      name: "Spannung 1",
      lock_name: false,
      op_number: "10",
      time_status: "ok",
      actual_time_minutes: null,
      scrap_total: 0,
      clarify_total: 0,
      scrap_lifetime: 0,
      clarify_lifetime: 0,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "station-over",
      order_id: "order-over",
      station_no: 1,
      name: "Spannung 1",
      lock_name: false,
      op_number: "10",
      time_status: "ok",
      actual_time_minutes: null,
      scrap_total: 1,
      clarify_total: 0,
      scrap_lifetime: 1,
      clarify_lifetime: 0,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "station-chain-one",
      order_id: "order-chain",
      station_no: 1,
      name: "Spannung 1",
      lock_name: false,
      op_number: "10",
      time_status: "ok",
      actual_time_minutes: null,
      scrap_total: 0,
      clarify_total: 0,
      scrap_lifetime: 0,
      clarify_lifetime: 0,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "station-chain-two",
      order_id: "order-chain",
      station_no: 2,
      name: "Spannung 2",
      lock_name: false,
      op_number: "20",
      time_status: "ok",
      actual_time_minutes: null,
      scrap_total: 0,
      clarify_total: 0,
      scrap_lifetime: 0,
      clarify_lifetime: 0,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  production_order_employees: [
    {
      id: "entry-over",
      order_id: "order-over",
      employee_id: "employee-one",
      role: "worker",
      active: true,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "entry-chain",
      order_id: "order-chain",
      employee_id: "employee-one",
      role: "worker",
      active: true,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  production_checklist_templates: [
    "Stückzahl gezählt?",
    "Auftrag fertig gemeldet?",
    "Lagerort eingetragen?",
    "Hilfsmittel eingeräumt?",
    "Werkzeuge eingeräumt?",
    "Teile eingeölt?",
    "Doku ergänzt?",
    "Excel eingetragen?",
    "Doku-Fehler notiert?",
    "Doku abgelegt?",
  ].map((label, index) => ({
    id: `template-${index + 1}`,
    item_key: `item_${index + 1}`,
    item_label: label,
    active: true,
    sort_order: index + 1,
    created_at: "2026-01-01T00:00:00Z",
    updated_at: "2026-01-01T00:00:00Z",
  })),
  production_order_checklist: [],
  production_order_history: [
    {
      id: "history-completed",
      order_id: "order-one",
      station_id: null,
      employee_id: "employee-admin",
      history_type: "order_completed",
      qty: 0,
      payload: {
        ba_number: "BA-100",
        good_total: 1,
        scrap_total: 1,
        clarify_total: 0,
        remaining: 0,
      },
      note: "Historie Test",
      created_at: "2026-01-01T00:04:00Z",
    },
  ],
  production_station_counts: [
    {
      id: "count-over",
      order_id: "order-over",
      station_id: "station-over",
      order_employee_id: "entry-over",
      good_qty: 2,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "count-chain-one",
      order_id: "order-chain",
      station_id: "station-chain-one",
      order_employee_id: "entry-chain",
      good_qty: 2,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  production_station_events: [
    {
      id: "event-good",
      order_id: "order-one",
      station_id: "station-one",
      order_employee_id: "",
      employee_id: "employee-one",
      event_type: "good",
      qty: 1,
      qa_cause_id: null,
      note: "Gutteil Test",
      created_at: "2026-01-01T00:02:00Z",
    },
    {
      id: "event-scrap",
      order_id: "order-one",
      station_id: "station-one",
      order_employee_id: "",
      employee_id: "employee-admin",
      event_type: "scrap",
      qty: 1,
      qa_cause_id: "qa-cause-human",
      note: "Ausschuss Test",
      created_at: "2026-01-01T00:03:00Z",
    },
  ],
  production_qa_causes: [
    {
      id: "qa-cause-human",
      group_id: "human",
      group_label: "Mensch",
      reason_id: "training",
      reason_label: "Einweisung",
      active: true,
      sort_order: 1,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
    {
      id: "qa-cause-machine",
      group_id: "machine",
      group_label: "Maschine",
      reason_id: "setup",
      reason_label: "Einrichtung",
      active: true,
      sort_order: 2,
      created_at: "2026-01-01T00:00:00Z",
      updated_at: "2026-01-01T00:00:00Z",
    },
  ],
  production_counts: [],
  tool_materials: [],
  tools: [
    {
      id: "tool-one",
      t_number: "133",
      label: "Schaftfräser",
      diameter: "10",
      corner_radius: "",
      article_no: "ART-133",
      holder: "HSK 63",
      shelf: "12N",
      stock: 4,
      min_stock: 1,
      optimal_stock: 4,
      manufacturer: "Test",
      ordered: false,
      ordered_qty: 0,
      insert_tool: false,
      insert_edges: 0,
      insert_radius: null,
      is_borrowed: false,
      borrowed_to: null,
      borrowed_at: null,
      material_id: null,
    },
  ],
  tool_journal: [],
  planner_tasks: [],
  planner_assignments: [],
  planner_absences: [],
  planner_vacations: [],
  planner_sick_leaves: [],
  planner_availability: [],
  planner_swaps: [],
  planner_shift_cancellations: [],
  planner_saturday_requests: [],
  planner_absence_replacements: [],
};

function installSupabaseMock(rowsByTableArg) {
  const rowsByTable = rowsByTableArg;
  window.__SUPABASE_MOCK_ROWS = rowsByTable;
  window.__SUPABASE_WRITE_LOG = [];
  window.__SUPABASE_MOCK_ERRORS = {};
  window.TEST_ADMIN_EMAIL = "admin@example.test";
  window.TEST_EMPLOYEE_EMAIL = "employee@example.test";

  function consumeMockError(table, action) {
    const tableErrors = window.__SUPABASE_MOCK_ERRORS?.[table];
    const error = tableErrors?.[action];
    if (error) {
      delete tableErrors[action];
      return error;
    }
    return null;
  }

  class QueryMock {
    constructor(table) {
      this.table = table;
      this.filters = [];
      this.useSingle = false;
      this.useMaybeSingle = false;
      this.pendingUpdate = null;
      this.pendingDelete = false;
    }

    select() { return this; }
    order() { return this; }
    limit() { return this; }
    insert(payload) {
      window.__SUPABASE_WRITE_LOG.push({ table: this.table, action: "insert", payload });
      const mockError = consumeMockError(this.table, "insert");
      if (mockError) return Promise.resolve({ data: null, error: mockError });
      const rows = Array.isArray(payload) ? payload : [payload];
      const tableRows = rowsByTable[this.table] || [];
      rowsByTable[this.table] = tableRows;
      const created = rows.map((row, index) => ({
        id: row.id || `${this.table}-${Date.now()}-${index}`,
        created_at: row.created_at || new Date().toISOString(),
        updated_at: row.updated_at || new Date().toISOString(),
        ...row,
      }));
      tableRows.push(...created);
      return Promise.resolve({ data: created, error: null });
    }
    update(payload) {
      window.__SUPABASE_WRITE_LOG.push({ table: this.table, action: "update", payload });
      this.pendingUpdate = payload || {};
      return this;
    }
    delete() {
      window.__SUPABASE_WRITE_LOG.push({ table: this.table, action: "delete" });
      this.pendingDelete = true;
      return this;
    }
    upsert(payload) { return this.insert(payload); }
    single() { this.useSingle = true; return this; }
    maybeSingle() { this.useMaybeSingle = true; return this; }
    eq(column, value) { this.filters.push({ column, value }); return this; }

    result() {
      const tableRows = rowsByTable[this.table] || [];
      let data = [...tableRows];
      for (const { column, value } of this.filters) {
        data = data.filter((row) => String(row[column]) === String(value));
      }
      const mockAction = this.pendingUpdate ? "update" : this.pendingDelete ? "delete" : "select";
      const mockError = consumeMockError(this.table, mockAction);
      if (mockError) return { data: null, error: mockError };
      if (this.pendingUpdate) {
        data.forEach((row) => Object.assign(row, this.pendingUpdate));
        return { data, error: null };
      }
      if (this.pendingDelete) {
        rowsByTable[this.table] = tableRows.filter((row) => !data.includes(row));
        return { data, error: null };
      }
      if (this.useSingle || this.useMaybeSingle) {
        return { data: data[0] || null, error: null };
      }
      return { data, error: null };
    }

    then(resolve, reject) {
      const hold = window.__SUPABASE_MOCK_HOLD;
      if (hold?.table === this.table && this.pendingUpdate) {
        return new Promise((release) => {
          hold.release = () => release(this.result());
        }).then(resolve, reject);
      }
      return Promise.resolve(this.result()).then(resolve, reject);
    }
  }

  window.supabase = {
    createClient: () => ({
      from: (table) => new QueryMock(table),
      auth: {
        getSession: async () => ({ data: { session: null }, error: null }),
        getUser: async () => ({ data: { user: null }, error: null }),
        signInWithPassword: async () => ({ data: { user: null }, error: null }),
        signOut: async () => ({ error: null }),
      },
      channel: () => ({
        on() { return this; },
        subscribe(callback) {
          if (typeof callback === "function") callback("SUBSCRIBED");
          return this;
        },
      }),
      removeChannel: async () => null,
    }),
  };
}

test.describe("Humbel app smoke", () => {
  test.beforeEach(async ({ page }, testInfo) => {
    const feedbackTest = testInfo.titlePath.includes("Production counter feedback");
    const navigationTest = testInfo.titlePath.includes("Production order navigation");
    const touchViewTest = testInfo.titlePath.includes("Production touch view");
    const consoleErrors = [];
    page.on("console", (message) => {
      if (message.type() !== "error") return;
      const text = message.text();
      const ignored = [
        "Failed to load resource",
        "Tailwind",
        "cdn.tailwindcss.com",
        "supabase",
        "production_order_stations Supabase-Fehler",
        "production_station_events Supabase-Fehler",
        "production_station_counts Supabase-Fehler",
        "Fehler beim Laden von production_station_events",
        "Fehler beim Laden von production_station_counts",
      ];
      if (!ignored.some((entry) => text.includes(entry))) {
        if ((navigationTest || touchViewTest) && testInfo.title.includes("load error") && text.includes("Fehler beim Laden von production_orders")) return;
        consoleErrors.push(text);
      }
    });
    page.on("pageerror", (error) => {
      consoleErrors.push(error.message);
    });
    page.consoleErrors = consoleErrors;
    await page.route("**/cdn.tailwindcss.com**", (route) => {
      // Generated from the actual app classes with Tailwind 3.4.17; no CDN needed in CI.
      const css = feedbackTest || navigationTest || touchViewTest
        ? fs.readFileSync(path.join(__dirname, "fixtures", "app-tailwind.css"), "utf8")
        : "";
      route.fulfill({ contentType: "application/javascript", body: css
        ? `const sheet = document.createElement('style'); sheet.textContent = ${JSON.stringify(css)}; document.head.appendChild(sheet);`
        : "" });
    });
    await page.route("**/@supabase/supabase-js@2**", (route) => {
      route.fulfill({ contentType: "application/javascript", body: "" });
    });
    await page.route("**/*.supabase.co/**", (route) => route.abort());
    const rows = JSON.parse(JSON.stringify(rowsByTable));
    if (feedbackTest || navigationTest || touchViewTest) {
      rows.production_orders.find((row) => row.id === "order-one").target_quantity = 50;
      Object.assign(rows.production_order_stations[0], { scrap_total: 2, clarify_total: 2 });
      rows.production_order_employees.push({
        id: "counter-worker", order_id: "order-one", employee_id: "employee-one", active: true, role: "worker",
      });
      rows.production_station_counts.push({
        id: "counter-count", order_id: "order-one", station_id: "station-one",
        order_employee_id: "counter-worker", good_qty: 2,
      });
      rows.production_order_checklist = rows.production_checklist_templates.map((item) => ({
        id: `counter-${item.id}`, order_id: "order-one", item_key: item.item_key,
        item_label: item.item_label, sort_order: item.sort_order, checked: false,
      }));
    }
    if (navigationTest || touchViewTest) prepareProductionNavigationRows(rows);
    if (touchViewTest) prepareProductionTouchRows(rows);
    if (!touchViewTest) {
      // The original 80 regression tests exercise the explicitly chosen full overview.
      await page.addInitScript(() => localStorage.setItem("humbel.productionCounterView", "overview"));
    }
    await page.addInitScript(installSupabaseMock, rows);
  });

  test.afterEach(async ({ page }) => {
    expect(page.consoleErrors).toEqual([]);
  });

  test("loads, logs in locally, and navigates modules and subtabs", async ({ page }) => {
    await page.goto(appUrl);

    await expect(page.getByRole("heading", { name: "Anmeldung" })).toBeVisible();
    await page.getByRole("button", { name: "Mido/Admin" }).click();

    for (const tabName of [
      "Start / Dashboard",
      "Personalverwaltung",
      "Produktion",
      "Werkzeugverwaltung",
      "Lagerverwaltung",
      "Scanner",
      "Auswertung",
      "Admin / System",
    ]) {
      await expect(page.getByRole("button", { name: tabName })).toBeVisible();
    }

    await clickMain(page, "Personalverwaltung");
    await clickSubtabs(page, [
      "Mitarbeiter",
      "Schichtmodell",
      "Schichtplanung",
      "Einstellungen",
    ]);

    await clickMain(page, "Produktion");
    await clickSubtabs(page, [
      "Abteilungen",
      "Maschinen",
      "Aufträge / BA",
      "Stückzahl",
      "Spannungen",
      "Ausschuss / Abklärung",
      "Protokoll",
      "Einstellungen",
    ]);

    await clickSubtab(page, "Protokoll");
    await expectViewHeading(page, "Produktionsprotokoll");
    await expect(page.locator("#view")).toContainText("Gutteil");
    await expect(page.locator("#view")).toContainText("Ausschuss");
    await expect(page.locator("#view")).toContainText("Mensch: Einweisung");
    await expect(page.locator("#view")).toContainText("Auftrag abgeschlossen");
    await expect(page.locator("#view")).toContainText("Gutteile 1");
    await page.getByLabel("Auftrag / BA").selectOption("order-two");
    await expect(page.locator("#view")).toContainText("Noch keine Protokolleinträge vorhanden.");

    await clickSubtab(page, "Aufträge / BA");
    await expectViewHeading(page, "Neuen Auftrag / BA anlegen");
    await expectViewHeading(page, "Laufende Aufträge");
    await expectViewHeading(page, "Abgeschlossene / abgebrochene Aufträge");
    await clickSubtab(page, "Stückzahl");
    await expectViewHeading(page, "Zähler-Dashboard");
    await expectViewHeading(page, "Maschinenübersicht");
    await expect(page.locator("#view")).toContainText("Maschine 50");
    await expect(page.locator("#view")).toContainText("BA-100");
    await page.locator("#view").getByRole("button", { name: /Maschine 50/ }).click();
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(page.locator("#view").getByRole("tab", { name: /BA BA-100/ })).toBeVisible();
    await expect(page.locator("#view")).toContainText("Rest");
    await expect(page.locator("#view")).toContainText("Gesamt gezählt");
    await expect(page.locator("#view")).toContainText("Mitarbeiter am Auftrag");
    await page.locator("#view").getByLabel("Mitarbeiter hinzufügen").selectOption("employee-one");
    await page.locator("#view").getByRole("button", { name: "Hinzufügen", exact: true }).click();
    await expect(page.locator("#view")).toContainText("Lavdrim");
    await expect(page.locator("#view")).toContainText("PN 100");
    await expect(page.locator("#view")).toContainText("Spannung 1");
    await expect(page.locator("#view")).toContainText("Auftrag abschließen");
    await expect(page.locator("#view")).not.toContainText("Stückzahl gezählt?");
    await page.locator("#view").getByRole("button", { name: "Auftrag fertig melden" }).click();
    await expect(page.getByRole("heading", { name: "Auftrag fertig melden" })).toBeVisible();
    await expect(page.locator("body")).toContainText("Abschluss-Checkliste");
    await expect(page.locator("#view")).toContainText("Stückzahl gezählt?");
    await expect(page.locator("#view")).toContainText("Auftrag fertig gemeldet?");
    await expect(page.locator("#view")).toContainText("Lagerort eingetragen?");
    await expect(page.locator("#view")).toContainText("Hilfsmittel eingeräumt?");
    await expect(page.locator("#view")).toContainText("Werkzeuge eingeräumt?");
    await expect(page.locator("#view")).toContainText("Teile eingeölt?");
    await expect(page.locator("#view")).toContainText("Doku ergänzt?");
    await expect(page.locator("#view")).toContainText("Excel eingetragen?");
    await expect(page.locator("#view")).toContainText("Doku-Fehler notiert?");
    await expect(page.locator("#view")).toContainText("Doku abgelegt?");
    await expect(page.locator("body")).not.toContainText("Abschluss-Checkliste wird vorbereitet.");
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_order_checklist" &&
            entry.action === "insert" &&
            entry.payload?.length === 10 &&
            entry.payload?.some?.((item) => item.order_id === "order-one" && item.item_key === "item_1"),
        ),
      ))
      .toBeTruthy();
    await page.getByLabel("Stückzahl gezählt?").check();
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_order_checklist" &&
            entry.action === "update" &&
            entry.payload?.checked === true &&
            !!entry.payload?.checked_at &&
            entry.payload?.checked_by_employee_id === "employee-admin",
        ),
      ))
      .toBeTruthy();
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_order_history" &&
            entry.action === "insert" &&
            entry.payload?.[0]?.history_type === "checklist",
        ),
      ))
      .toBeTruthy();
    await expect(page.getByRole("button", { name: "Endgültig fertig melden" })).toBeDisabled();
    await page.evaluate(() => window.completeProductionOrderWithChecklist("order-one"));
    await expect(page.locator("#view")).toContainText("Alle Checklistenpunkte müssen erledigt sein.");
    for (const label of [
      "Auftrag fertig gemeldet?",
      "Lagerort eingetragen?",
      "Hilfsmittel eingeräumt?",
      "Werkzeuge eingeräumt?",
      "Teile eingeölt?",
      "Doku ergänzt?",
      "Excel eingetragen?",
      "Doku-Fehler notiert?",
      "Doku abgelegt?",
    ]) {
      await page.getByLabel(label).check();
    }
    await expect(page.getByRole("button", { name: "Endgültig fertig melden" })).toBeDisabled();
    await page.evaluate(() => window.completeProductionOrderWithChecklist("order-one"));
    await expect(page.locator("#view")).toContainText("Auftrag hat noch Restmenge. Abschluss nicht möglich.");
    await page.getByRole("button", { name: "Abbrechen" }).click();
    const stationOne = page.locator("#view").locator("article").filter({ hasText: "Spannung 1" });
    await expect(stationOne).toContainText("Lavdrim");
    const stationInsertsBeforeSave = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) => entry.table === "production_order_stations" && entry.action === "insert",
      ).length,
    );
    await stationOne.getByText("Details bearbeiten").click();
    await stationOne.getByLabel("OP-Nummer").fill("20");
    await stationOne.getByRole("button", { name: "Speichern" }).click();
    await expect(stationOne).toContainText("OP 20");
    await expect(page.locator("#view")).toContainText("Spannung wurde gespeichert.");
    const stationInsertsAfterSave = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) => entry.table === "production_order_stations" && entry.action === "insert",
      ).length,
    );
    expect(stationInsertsAfterSave).toBe(stationInsertsBeforeSave);
    await expect(stationOne).toContainText(/Gutteile\s+0/);
    await stationOne.scrollIntoViewIfNeeded();
    await page.evaluate(() => window.scrollBy(0, -80));
    const stationTopBeforeGoodCount = await stationOne.evaluate((element) => element.getBoundingClientRect().top);
    await stationOne.getByRole("button", { name: /Gutteile \+1/ }).click();
    await expect
      .poll(() => stationOne.evaluate((element, before) => Math.abs(element.getBoundingClientRect().top - before) <= 30, stationTopBeforeGoodCount))
      .toBeTruthy();
    await expect.poll(() => page.evaluate(() => window.scrollY > 20)).toBeTruthy();
    await expect(stationOne).toContainText(/Gutteile\s+1/);
    await expect(page.locator("#productionCounterAnnouncement")).toHaveText("Gutteil wurde gezählt.");
    await expect(page.locator("#view")).not.toContainText("Protokolleintrag konnte nicht geschrieben werden");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(page.locator("#view").getByRole("tab", { name: /BA BA-100/ })).toBeVisible();
    await expect(stationOne).toContainText("Lavdrim");
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_station_counts" &&
            ["insert", "update"].includes(entry.action),
        ),
      ))
      .toBeTruthy();
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_station_events" &&
            entry.action === "insert" &&
            entry.payload?.[0]?.event_type === "good" &&
            entry.payload?.[0]?.qty === 1,
        ),
      ))
      .toBeTruthy();
    await expect(page.locator("#view")).toContainText("Gutteile gesamt");
    await expect(page.locator("#view")).toContainText("Fertige Gutteile");
    await expect(page.locator("#view")).toContainText("Rest");
    await expect(page.locator("#view")).toContainText("Gesamt gezählt");
    await stationOne.getByRole("button", { name: /Gutteile \+1/ }).click();
    await expect(stationOne).toContainText(/Gutteile\s+2/);
    await expect(stationOne).toContainText("Ziel erreicht");
    await expect(stationOne.getByRole("button", { name: /Gutteile \+1/ })).toBeDisabled();
    const goodEventsBeforeBlocked = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) =>
          entry.table === "production_station_events" &&
          entry.action === "insert" &&
          entry.payload?.[0]?.event_type === "good",
      ).length,
    );
    const countWritesBeforeBlocked = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) => entry.table === "production_station_counts",
      ).length,
    );
    await page.evaluate(() => {
      const entry = window.__SUPABASE_MOCK_ROWS.production_order_employees.find(
        (item) => item.order_id === "order-one",
      );
      return window.adjustProductionStationGoodQty("station-one", entry.id, 1);
    });
    await expect(page.locator("#view")).toContainText("Ziel erreicht - +1 gesperrt");
    await expect(stationOne).toContainText(/Gutteile\s+2/);
    const goodEventsAfterBlocked = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) =>
          entry.table === "production_station_events" &&
          entry.action === "insert" &&
          entry.payload?.[0]?.event_type === "good",
      ).length,
    );
    const countWritesAfterBlocked = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) => entry.table === "production_station_counts",
      ).length,
    );
    expect(goodEventsAfterBlocked).toBe(goodEventsBeforeBlocked);
    expect(countWritesAfterBlocked).toBe(countWritesBeforeBlocked);
    await page.evaluate(() => {
      const entry = window.__SUPABASE_MOCK_ROWS.production_order_employees.find(
        (item) => item.order_id === "order-one",
      );
      return window.adjustProductionStationGoodQty("station-one", entry.id, -1);
    });
    await expect(stationOne).toContainText(/Gutteile\s+1/);
    await expect(page.locator("#productionCounterAnnouncement")).toHaveText("Gutteil-Korrektur wurde gespeichert.");
    await page.evaluate(() => {
      const entry = window.__SUPABASE_MOCK_ROWS.production_order_employees.find(
        (item) => item.order_id === "order-one",
      );
      return window.adjustProductionStationGoodQty("station-one", entry.id, -1);
    });
    await expect(stationOne).toContainText(/Gutteile\s+0/);
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_station_events" &&
            entry.action === "insert" &&
            entry.payload?.[0]?.event_type === "good_correction" &&
            entry.payload?.[0]?.qty === -1,
        ),
      ))
      .toBeTruthy();
    await page.evaluate(() => {
      const entry = window.__SUPABASE_MOCK_ROWS.production_order_employees.find(
        (item) => item.order_id === "order-one",
      );
      return window.adjustProductionStationGoodQty("station-one", entry.id, -1);
    });
    await expect(page.locator("#view")).toContainText("Gutmenge ist bereits 0");
    await expect(stationOne).toContainText(/Gutteile\s+0/);
    await page.evaluate(() => {
      window.__SUPABASE_MOCK_ERRORS.production_station_counts = {
        select: {
          code: "42501",
          message: "new row violates row-level security policy for table \"production_station_counts\"",
        },
      };
    });
    await stationOne.getByRole("button", { name: /Gutteile \+1/ }).click();
    await expect(stationOne).toContainText(/Gutteile\s+1/);
    await expect(page.locator("#view")).toContainText("Gutteil gespeichert, Zähler konnte nicht neu geladen werden.");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(stationOne).toContainText("Spannung 1");
    await page.evaluate(() => {
      window.__SUPABASE_MOCK_ERRORS.production_station_events = {
        select: {
          code: "42501",
          message: "new row violates row-level security policy for table \"production_station_events\"",
        },
      };
    });
    await stationOne.getByRole("button", { name: /Gutteile -1/ }).click();
    await expect(stationOne).toContainText(/Gutteile\s+0/);
    await expect(page.locator("#view")).toContainText("Gutteil gespeichert, Protokoll konnte nicht neu geladen werden.");
    await expect(page.locator("#view")).not.toContainText("Protokolleintrag konnte nicht geschrieben werden");
    await page.evaluate(() => {
      window.__SUPABASE_MOCK_ERRORS.production_station_events = {
        insert: {
          code: "42501",
          message: "new row violates row-level security policy for table \"production_station_events\"",
        },
      };
    });
    await stationOne.getByRole("button", { name: /Gutteile \+1/ }).click();
    await expect(stationOne).toContainText(/Gutteile\s+1/);
    await expect(page.locator("#view")).toContainText("Gutteil gespeichert, Protokolleintrag konnte nicht geschrieben werden.");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(stationOne).toContainText(/Ausschuss\s+0/);
    await stationOne.getByRole("button", { name: "Ausschuss +1" }).click();
    await expect(page.getByRole("heading", { name: "6M-Ursache für Ausschuss" })).toBeVisible();
    await expect(page.locator("body")).toContainText("Mensch");
    await page.getByRole("button", { name: "Abbrechen" }).first().click();
    await expect(stationOne).toContainText(/Ausschuss\s+0/);
    const scrapAnchor = page.locator('[data-scroll-anchor="amount:station-one:scrap"]').first();
    await scrapAnchor.scrollIntoViewIfNeeded();
    await page.evaluate(() => window.scrollBy(0, -80));
    const scrapTopBeforeCount = await scrapAnchor.evaluate((element) => element.getBoundingClientRect().top);
    await stationOne.getByRole("button", { name: "Ausschuss +1" }).click();
    await page.getByRole("button", { name: "Einweisung" }).click();
    await page.getByLabel("Notiz").fill("Testnotiz Ausschuss");
    await page.getByRole("button", { name: "Ursache speichern" }).click();
    await expect
      .poll(() => scrapAnchor.evaluate((element, before) => Math.abs(element.getBoundingClientRect().top - before) <= 30, scrapTopBeforeCount))
      .toBeTruthy();
    await expect(stationOne).toContainText(/Ausschuss\s+1/);
    await expect(stationOne).toContainText("Ziel erreicht");
    await expect(stationOne.getByRole("button", { name: /Gutteile \+1/ })).toBeDisabled();
    await expect(stationOne.getByRole("button", { name: "Ausschuss +1" })).toBeDisabled();
    await expect(stationOne.getByRole("button", { name: "In Abklärung +1" })).toBeDisabled();
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_station_events" &&
            entry.action === "insert" &&
            entry.payload?.[0]?.event_type === "scrap" &&
            entry.payload?.[0]?.qa_cause_id === "qa-cause-human" &&
            entry.payload?.[0]?.note === "Testnotiz Ausschuss",
        ),
      ))
      .toBeTruthy();
    const amountWritesBeforeBlocked = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) =>
          ["production_order_stations", "production_station_events"].includes(entry.table) &&
          ["update", "insert"].includes(entry.action),
      ).length,
    );
    await page.evaluate(() =>
      window.adjustProductionStationAmount("station-one", "scrap", 1, {
        qaCauseId: "qa-cause-human",
        note: "blocked scrap",
      }),
    );
    await expect(page.locator("#view")).toContainText("Ziel erreicht - +1 gesperrt");
    await page.evaluate(() =>
      window.adjustProductionStationAmount("station-one", "clarify", 1, {
        qaCauseId: "qa-cause-machine",
        note: "blocked clarify",
      }),
    );
    await expect(page.locator("#view")).toContainText("Ziel erreicht - +1 gesperrt");
    await expect(stationOne).toContainText(/Ausschuss\s+1/);
    await expect(stationOne).toContainText(/In Abklärung\s+0/);
    const amountWritesAfterBlocked = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) =>
          ["production_order_stations", "production_station_events"].includes(entry.table) &&
          ["update", "insert"].includes(entry.action),
      ).length,
    );
    expect(amountWritesAfterBlocked).toBe(amountWritesBeforeBlocked);
    await stationOne.getByRole("button", { name: "Ausschuss -1" }).click();
    await expect(stationOne).toContainText(/Ausschuss\s+0/);
    await expect(page.locator("#productionCounterAnnouncement")).toHaveText("Ausschuss-Korrektur wurde gespeichert.");
    await expect(page.locator("#view")).not.toContainText("Ursache/Protokoll konnte nicht geschrieben werden");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(page.locator("#view").getByRole("tab", { name: /BA BA-100/ })).toBeVisible();
    await expect(stationOne).toContainText("Spannung 1");
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_station_events" &&
            entry.action === "insert" &&
            entry.payload?.[0]?.event_type === "scrap_correction" &&
            entry.payload?.[0]?.qty === -1 &&
            entry.payload?.[0]?.qa_cause_id === null,
        ),
      ))
      .toBeTruthy();
    await page.evaluate(() => window.adjustProductionStationAmount("station-one", "scrap", -1));
    await expect(page.locator("#view")).toContainText("Ausschuss ist bereits 0");
    await expect(stationOne).toContainText(/Ausschuss\s+0/);
    await expect(stationOne).toContainText(/In Abklärung\s+0/);
    await stationOne.getByRole("button", { name: "In Abklärung +1" }).click();
    await expect(page.getByRole("heading", { name: "6M-Ursache für In Abklärung" })).toBeVisible();
    await page.getByRole("button", { name: "Einrichtung" }).click();
    await page.getByRole("button", { name: "Ursache speichern" }).click();
    await expect(stationOne).toContainText(/In Abklärung\s+1/);
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_station_events" &&
            entry.action === "insert" &&
            entry.payload?.[0]?.event_type === "clarify" &&
            entry.payload?.[0]?.qa_cause_id === "qa-cause-machine",
        ),
      ))
      .toBeTruthy();
    await stationOne.getByRole("button", { name: "In Abklärung -1" }).click();
    await expect(stationOne).toContainText(/In Abklärung\s+0/);
    await expect(page.locator("#productionCounterAnnouncement")).toHaveText("Abklär-Korrektur wurde gespeichert.");
    await expect(page.locator("#view")).not.toContainText("Ursache/Protokoll konnte nicht geschrieben werden");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_station_events" &&
            entry.action === "insert" &&
            entry.payload?.[0]?.event_type === "clarify_correction" &&
            entry.payload?.[0]?.qty === -1 &&
            entry.payload?.[0]?.qa_cause_id === null,
        ),
      ))
      .toBeTruthy();
    await page.evaluate(() => window.adjustProductionStationAmount("station-one", "clarify", -1));
    await expect(page.locator("#view")).toContainText("Abklärung ist bereits 0");
    await expect(stationOne).toContainText(/In Abklärung\s+0/);
    await stationOne.getByRole("button", { name: "Ausschuss +1" }).click();
    await page.getByRole("button", { name: "Einweisung" }).click();
    await page.getByRole("button", { name: "Ursache speichern" }).click();
    await expect(stationOne).toContainText(/Ausschuss\s+1/);
    await page.evaluate(() => {
      window.__SUPABASE_MOCK_ERRORS.production_station_events = {
        select: {
          code: "42501",
          message: "new row violates row-level security policy for table \"production_station_events\"",
        },
      };
    });
    await stationOne.getByRole("button", { name: "Ausschuss -1" }).click();
    await expect(stationOne).toContainText(/Ausschuss\s+0/);
    await expect(page.locator("#view")).toContainText("Menge gespeichert, Protokoll konnte nicht neu geladen werden.");
    await expect(page.locator("#view")).not.toContainText("Ursache/Protokoll konnte nicht geschrieben werden");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(stationOne).toContainText("Spannung 1");
    await stationOne.getByRole("button", { name: "In Abklärung +1" }).click();
    await page.getByRole("button", { name: "Einrichtung" }).click();
    await page.getByRole("button", { name: "Ursache speichern" }).click();
    await expect(stationOne).toContainText(/In Abklärung\s+1/);
    await page.evaluate(() => {
      window.__SUPABASE_MOCK_ERRORS.production_station_events = {
        insert: {
          code: "42501",
          message: "new row violates row-level security policy for table \"production_station_events\"",
        },
      };
    });
    await stationOne.getByRole("button", { name: "In Abklärung -1" }).click();
    await expect(stationOne).toContainText(/In Abklärung\s+0/);
    await expect(page.locator("#view")).toContainText("Menge gespeichert, Ursache/Protokoll konnte nicht geschrieben werden.");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(page.locator("#view")).toContainText("Ausschuss Auftrag");
    await expect(page.locator("#view")).toContainText("In Abklärung Auftrag");
    await expect(stationOne).toContainText("Gutteile");
    await expect(stationOne).toContainText("Ausschuss");
    await expect(stationOne).toContainText("In Abklärung");
    await page.locator("#view").getByRole("button", { name: /Spannung/ }).filter({ hasText: /hinzuf/ }).click();
    await expect(page.locator("#view")).toContainText("Spannung 2");
    const stationTwo = page.locator("#view").locator("article").filter({ hasText: "Spannung 2" });
    await expect(stationTwo).toContainText("Lavdrim");
    await stationTwo.getByRole("button", { name: /Gutteile \+1/ }).click();
    await expect(stationTwo).toContainText(/Gutteile\s+1/);
    await expect(stationTwo.getByRole("button", { name: /Gutteile \+1/ })).toBeDisabled();
    await expect(page.locator("#view")).toContainText("Rest");
    await expect(page.locator("#view")).toContainText("Gesamt gezählt");
    await expect(page.locator("#view")).toContainText("0");
    await page.locator("#view").getByRole("button", { name: "Auftrag fertig melden" }).click();
    await expect(page.getByRole("button", { name: "Endgültig fertig melden" })).toBeEnabled();
    page.once("dialog", (dialog) => dialog.accept());
    await page.getByRole("button", { name: "Endgültig fertig melden" }).click();
    await expect(page.locator("#view")).toContainText("Auftrag wurde fertig gemeldet.");
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_orders" &&
            entry.action === "update" &&
            entry.payload?.status === "completed" &&
            !!entry.payload?.completed_at,
        ),
      ))
      .toBeTruthy();
    await expect
      .poll(() => page.evaluate(() =>
        (window.__SUPABASE_WRITE_LOG || []).some(
          (entry) =>
            entry.table === "production_order_history" &&
            entry.action === "insert" &&
            entry.payload?.[0]?.history_type === "order_completed",
        ),
      ))
      .toBeTruthy();
    await expect(page.locator("#view")).not.toContainText("BA-100");
    await expect
      .poll(() => page.evaluate(() => window.__SUPABASE_WRITE_LOG || []))
      .not.toContainEqual(expect.objectContaining({ table: "production_counts" }));
    const stationCountWrites = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter((entry) => entry.table === "production_station_counts"),
    );
    expect(stationCountWrites.every((entry) => ["insert", "update"].includes(entry.action))).toBeTruthy();
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(page.locator("#view")).toContainText("Keine laufenden oder pausierten Aufträge auf dieser Maschine.");
    await page.getByRole("button", { name: "Zurück zur Maschinenübersicht" }).click();
    await expectViewHeading(page, "Maschinenübersicht");
    await page.evaluate(() => {
      window.__SUPABASE_MOCK_ERRORS.production_order_stations = {
        insert: {
          code: "42501",
          message: "new row violates row-level security policy for table \"production_order_stations\"",
        },
      };
    });
    await page.locator("#view").getByRole("button", { name: /Maschine 51/ }).click();
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(page.locator("#view")).toContainText("Maschine 51");
    await expect(page.locator("#view")).toContainText("Erste Spannung konnte nicht angelegt werden");
    await expect(page.locator("#view")).toContainText("Noch keine Spannung geladen");
    await page.getByRole("button", { name: "Zurück zur Maschinenübersicht" }).click();
    await expectViewHeading(page, "Maschinenübersicht");

    await clickMain(page, "Werkzeugverwaltung");
    await clickSubtab(page, "Werkzeugliste");
    await expectViewHeading(page, "Werkzeugbestand");
    await expectNoViewHeading(page, "Schichtjournal");
    await clickSubtab(page, "Bestand");
    await expectViewHeading(page, "Werkzeugbestand");
    await expectNoViewHeading(page, "Schichtjournal");
    await clickSubtab(page, "Entnahme / Einlagerung");
    await expectViewHeading(page, "Werkzeug-Scanner");
    await clickSubtab(page, "Nachbestellen");
    await expectViewHeading(page, "Nachbestellen prüfen");
    await expectNoViewHeading(page, "Schichtjournal – Werkzeugwechsel");
    await expectNoViewHeading(page, "Lagerfach-/Regalansicht");
    await clickSubtab(page, "Journal");
    await expectViewHeading(page, "Schichtjournal – Werkzeugwechsel");
    await expectNoViewHeading(page, "Nachbestellen prüfen");
    await expectNoViewHeading(page, "Lagerfach-/Regalansicht");
    await clickSubtab(page, "QR-Etiketten");
    await expectViewHeading(page, "QR-Etiketten");
    await expectNoViewHeading(page, "Werkzeugbestand");
    await clickSubtab(page, "Einstellungen");

    await clickMain(page, "Lagerverwaltung");
    await clickSubtab(page, "Lagerübersicht");
    await expectViewHeading(page, "Lagerfach-/Regalansicht");
    await expectNoViewHeading(page, "Fachsuche");
    await clickSubtab(page, "Fachsuche");
    await expectViewHeading(page, "Fachsuche");
    await expectNoViewHeading(page, "Lagerfach-/Regalansicht");
    await clickSubtab(page, "Umlagern");
    await expectViewHeading(page, "Umlagern");
    await clickSubtab(page, "Lagerfach-QR");
    await expectViewHeading(page, "Lagerfach-QR");
    await clickSubtab(page, "Inventur");
    await expect(page.locator("#view")).toContainText("Inventur bleibt aktuell deaktiviert");
    await clickSubtab(page, "Einstellungen");

    await clickMain(page, "Scanner");
    await expect(page.getByRole("heading", { name: "Scanner", exact: true })).toBeVisible();

    await clickMain(page, "Auswertung");
    await expect(page.getByRole("heading", { name: "Auswertung" })).toBeVisible();

    await clickMain(page, "Admin / System");
    await expect(page.getByRole("heading", { name: "Admin / System" })).toBeVisible();
  });

  test("blocks production overcounts and preserves scroll in target quantity flows", async ({ page }) => {
    await page.goto(appUrl);
    await page.getByRole("button", { name: "Mido/Admin" }).click();
    await clickMain(page, "Produktion");
    await clickSubtab(page, "Stückzahl");

    await page.locator("#view").getByRole("button", { name: /Maschine 52/ }).click();
    await expectViewHeading(page, "Produktionsvorschau");
    const overStation = page.locator("#view").locator("article").filter({ hasText: "Spannung 1" });
    await expect(overStation).toContainText("Über Ziel - bitte korrigieren");
    await expect(overStation.getByRole("button", { name: /Gutteile \+1/ })).toBeDisabled();
    await expect(overStation.getByRole("button", { name: "Ausschuss +1" })).toBeDisabled();
    await expect(overStation.getByRole("button", { name: "In Abklärung +1" })).toBeDisabled();

    const blockedWritesBefore = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) =>
          ["production_order_stations", "production_station_counts", "production_station_events"].includes(entry.table) &&
          ["update", "insert"].includes(entry.action),
      ).length,
    );
    await page.evaluate(() => window.openProductionQaCauseModal("station-over", "scrap"));
    await expect(page.locator("#view")).toContainText("Ziel erreicht - +1 gesperrt");
    await expect(page.getByRole("heading", { name: "6M-Ursache für Ausschuss" })).toHaveCount(0);
    await page.evaluate(() => window.openProductionQaCauseModal("station-over", "clarify"));
    await expect(page.locator("#view")).toContainText("Ziel erreicht - +1 gesperrt");
    await expect(page.getByRole("heading", { name: "6M-Ursache für In Abklärung" })).toHaveCount(0);
    const blockedWritesAfter = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) =>
          ["production_order_stations", "production_station_counts", "production_station_events"].includes(entry.table) &&
          ["update", "insert"].includes(entry.action),
      ).length,
    );
    expect(blockedWritesAfter).toBe(blockedWritesBefore);

    await overStation.scrollIntoViewIfNeeded();
    await page.evaluate(() => window.scrollBy(0, -80));
    const stationTopBeforeScrapCorrection = await overStation.evaluate((element) => element.getBoundingClientRect().top);
    await overStation.getByRole("button", { name: "Ausschuss -1" }).click();
    await expect
      .poll(() => overStation.evaluate((element, before) => Math.abs(element.getBoundingClientRect().top - before) <= 30, stationTopBeforeScrapCorrection))
      .toBeTruthy();
    await expect(overStation).toContainText(/Ausschuss\s+0/);
    await expect(overStation).not.toContainText("Über Ziel - bitte korrigieren");
    await expect(overStation.getByRole("button", { name: /Gutteile \+1/ })).toBeDisabled();
    await overStation.getByRole("button", { name: /Gutteile -1/ }).click();
    await expect(overStation).toContainText(/Gutteile\s+1/);
    await expect(overStation.getByRole("button", { name: /Gutteile \+1/ })).toBeEnabled();

    await page.getByRole("button", { name: "Zurück zur Maschinenübersicht" }).click();
    await page.locator("#view").getByRole("button", { name: /Maschine 53/ }).click();
    await expectViewHeading(page, "Produktionsvorschau");
    const chainStationOne = page.locator("#view").locator("article").filter({ hasText: "Spannung 1" });
    const chainStationTwo = page.locator("#view").locator("article").filter({ hasText: "Spannung 2" });
    await expect(chainStationOne).toContainText("Ziel erreicht");
    await expect(chainStationOne.getByRole("button", { name: /Gutteile \+1/ })).toBeDisabled();
    await page.evaluate(() =>
      window.adjustProductionStationAmount("station-chain-one", "clarify", 1, {
        qaCauseId: "qa-cause-machine",
      }),
    );
    await expect(page.locator("#view")).toContainText("Ziel erreicht - +1 gesperrt");
    await expect(chainStationTwo.getByRole("button", { name: /Gutteile \+1/ })).toBeEnabled();
    await chainStationTwo.getByRole("button", { name: /Gutteile \+1/ }).click();
    await expect(chainStationTwo).toContainText(/Gutteile\s+1/);
  });

  test.describe("Production order navigation", () => {
    test.beforeEach(async ({ page }, testInfo) => {
      page.counterTouch = !!(testInfo.project.use.hasTouch || testInfo.project.use.isMobile);
      await page.goto(appUrl);
      await page.getByRole("button", { name: "Mido/Admin" }).click();
      await clickMain(page, "Produktion");
      await clickSubtab(page, "Aufträge / BA");
    });

    test.afterEach(async ({ page }) => {
      const writes = await page.evaluate(() => window.__SUPABASE_WRITE_LOG);
      expect(writes.some((entry) => ["production_counts", "production_station_counts", "production_station_events"].includes(entry.table))).toBe(false);
    });

    for (const [id, name] of [["nav-second", "BA BA-nav-second öffnen"], ["nav-third", "Artikel 4345433 öffnen"]]) {
      test(`direct list link ${name} opens the exact order`, async ({ page }) => {
        await pressCounterButton(page, page.locator(`tr[data-production-order-id="${id}"]`).getByRole("button", { name, exact: true }));
        await expect(page.getByRole("button", { name: "Stückzahl", exact: true })).toHaveClass(/humbel-subtab-active/);
        await expectViewHeading(page, "Produktionsvorschau");
        await expect(page.getByRole("tabpanel")).toHaveAttribute("data-order-id", id);
        await expect(page.locator('[data-production-machine-id="machine-one"]')).toBeVisible();
        await expect(page.getByRole("tab", { selected: true })).toHaveAttribute("data-order-id", id);
        await expect(page.getByRole("tablist", { name: "Aktive Aufträge" }).getByRole("tab")).toHaveCount(3);
        expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
      });
    }

    test("article tabs isolate counts employees stations checklist and protocol without mutations", async ({ page }) => {
      await page.evaluate(() => selectProductionCounterOrder("nav-second", true));
      await page.evaluate(() => { window.__navigationSentinel = "same-document"; });
      const before = await page.evaluate(() => JSON.stringify(window.__SUPABASE_MOCK_ROWS));
      const tabs = page.getByRole("tablist", { name: "Aktive Aufträge" });
      for (const id of ["order-one", "nav-third", "nav-second"]) {
        await pressCounterButton(page, tabs.locator(`[data-order-id="${id}"]`));
        expect(await page.evaluate(() => window.__navigationSentinel)).toBe("same-document");
        const expected = await page.evaluate((id) => ({
          ba: window.__SUPABASE_MOCK_ROWS.production_orders.find((order) => order.id === id).ba_number,
          qty: window.__SUPABASE_MOCK_ROWS.production_station_counts.find((count) => count.order_id === id).good_qty,
        }), id);
        const panel = page.getByRole("tabpanel");
        await expect(panel).toHaveAttribute("data-order-id", id);
        await expect(panel).toContainText(`Team ${expected.ba}`);
        await expect(panel).toContainText(`Spannung ${expected.ba}`);
        await expect(panel.locator('div[data-scroll-anchor^="good:"] [data-counter-value]')).toContainText(String(expected.qty));
        await page.evaluate(() => render());
        await expect(page.getByRole("tab", { selected: true })).toHaveAttribute("data-order-id", id);
        await pressCounterButton(page, panel.getByRole("button", { name: "Auftrag fertig melden", exact: true }));
        await expect(page.getByRole("heading", { name: "Auftrag fertig melden" })).toBeVisible();
        await expect(page.getByRole("checkbox", { name: `${expected.ba} Checkliste 1 erledigt`, exact: true })).toBeChecked();
        await page.getByRole("button", { name: "Abbrechen", exact: true }).click();
        await pressCounterButton(page, panel.getByRole("button", { name: "Protokoll anzeigen" }));
        await expect(page.getByRole("combobox", { name: /^Auftrag \/ BA/ })).toHaveValue(id);
        await expect(page.locator("#view")).toContainText(`Protokoll ${expected.ba}`);
        await page.evaluate((id) => selectProductionCounterOrder(id, true), id);
      }
      expect(await page.evaluate(() => JSON.stringify(window.__SUPABASE_MOCK_ROWS))).toBe(before);
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
      await expect(tabs).not.toContainText("BA-200");
    });

    test("identical articles remain distinct and a missing article falls back to BA", async ({ page }) => {
      await page.evaluate(() => selectProductionCounterOrder("order-one", true));
      const tabs = page.getByRole("tablist");
      await expect(tabs.locator('[data-order-id="order-one"]')).toHaveText(/4345433\s+BA BA-100/);
      await expect(tabs.locator('[data-order-id="nav-third"]')).toHaveText(/4345433\s+BA BA-nav-third/);
      await page.evaluate(async () => {
        window.__SUPABASE_MOCK_ROWS.production_orders.find((row) => row.id === "nav-second").article_number = "";
        await refreshProductionOrdersFromSupabase();
        render();
      });
      await expect(tabs.locator('[data-order-id="nav-second"] > span').first()).toHaveText("BA-nav-second");
      await expect(tabs).not.toContainText("BA-nav-closed");
      await expect(tabs).not.toContainText("BA-nav-cancelled");
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });

    test("required first initialization runs once and machine reselect preserves the chosen BA", async ({ page }) => {
      await page.evaluate(async () => {
        window.__SUPABASE_MOCK_ROWS.production_orders.push({
          ...window.__SUPABASE_MOCK_ROWS.production_orders[0], id: "nav-fresh", ba_number: "BA-FRESH", article_number: "FRESH",
        });
        await refreshProductionOrdersFromSupabase();
        await selectProductionCounterOrder("nav-fresh", true);
        await selectProductionCounterMachine("machine-one");
      });
      await expect(page.getByRole("tabpanel")).toHaveAttribute("data-order-id", "nav-fresh");
      await page.evaluate(async () => {
        await selectProductionCounterOrder("order-one");
        await selectProductionCounterOrder("nav-fresh");
      });
      const writes = await page.evaluate(() => window.__SUPABASE_WRITE_LOG);
      expect(writes.filter((entry) => entry.table === "production_order_stations" && entry.action === "insert")).toHaveLength(1);
      expect(writes.filter((entry) => entry.table === "production_order_checklist" && entry.action === "insert")).toHaveLength(1);
      expect(writes).toHaveLength(2);
    });

    for (const status of ["completed", "cancelled"]) {
      test(`${status} order leaves active tabs and keeps machine open`, async ({ page }) => {
        await page.evaluate(async () => {
          window.__SUPABASE_MOCK_ROWS.production_orders.find((row) => row.id === "nav-second").target_quantity = 0;
          await refreshProductionOrdersFromSupabase();
          await selectProductionCounterOrder("nav-second", true);
        });
        if (status === "completed") {
          page.once("dialog", (dialog) => dialog.accept());
          await page.evaluate(() => completeProductionOrderWithChecklist("nav-second"));
        } else await page.evaluate(() => updateProductionOrderStatus("nav-second", "cancelled"));
        await expect(page.locator('[data-production-machine-id="machine-one"]')).toBeVisible();
        await expect(page.getByRole("tablist").getByRole("tab")).toHaveCount(2);
        await expect(page.getByRole("tablist").locator('[data-order-id="nav-second"]')).toHaveCount(0);
        await expect(page.getByRole("tabpanel")).toHaveAttribute("data-order-id", "order-one");
        await page.evaluate(() => selectProductionCounterOrder("nav-second", true));
        await expect(page.getByRole("tabpanel")).toHaveAttribute("data-order-id", "order-one");
      });
    }

    test("last inactive order shows empty machine view and terminal orders cannot open", async ({ page }) => {
      await expect(page.locator('tr[data-production-order-id="nav-closed"]').getByRole("button")).toHaveCount(0);
      await expect(page.locator('tr[data-production-order-id="nav-cancelled"]').getByRole("button")).toHaveCount(0);
      await page.evaluate(async () => {
        await selectProductionCounterOrder("order-one", true);
        for (const id of ["order-one", "nav-second", "nav-third"]) await updateProductionOrderStatus(id, "cancelled");
      });
      await expect(page.locator('[data-production-machine-id="machine-one"]')).toBeVisible();
      await expect(page.getByRole("tabpanel")).toContainText("Keine laufenden oder pausierten Aufträge");
      await expect(page.getByRole("tablist").getByRole("tab")).toHaveCount(0);
    });

    test("fallback to a fresh active BA initializes its station and checklist", async ({ page }) => {
      await page.evaluate(async () => {
        const rows = window.__SUPABASE_MOCK_ROWS;
        for (const row of rows.production_orders) {
          if (row.machine_id === "machine-one" && row.id !== "nav-second") row.status = "completed";
        }
        rows.production_orders.push({ ...rows.production_orders.find((row) => row.id === "nav-second"),
          id: "nav-fallback", ba_number: "BA-FALLBACK", article_number: "FALLBACK" });
        await refreshProductionOrdersFromSupabase();
        await selectProductionCounterOrder("nav-second", true);
        await updateProductionOrderStatus("nav-second", "cancelled");
      });
      await expect(page.getByRole("tabpanel")).toHaveAttribute("data-order-id", "nav-fallback");
      await expect(page.getByRole("tabpanel").locator("article")).toHaveCount(1);
      expect(await page.evaluate(() => window.__SUPABASE_MOCK_ROWS.production_order_checklist
        .filter((item) => item.order_id === "nav-fallback").length)).toBe(10);
      const writes = await page.evaluate(() => window.__SUPABASE_WRITE_LOG);
      expect(writes.filter((entry) => entry.table === "production_order_stations" && entry.action === "insert")).toHaveLength(1);
    });

    test("department admin opens only own machines and orders without global rights", async ({ page }) => {
      await page.evaluate(() => {
        currentUser = { name: "Leitung", role: "department_admin" };
        currentEmployeeRecord = state.employeesList.find((row) => row.id === "leader-test");
        render();
      });
      await expect(page.locator('tr[data-production-order-id="order-two"]')).toHaveCount(0);
      await page.evaluate(() => selectProductionCounterOrder("nav-third", true));
      await expect(page.getByRole("tabpanel")).toHaveAttribute("data-order-id", "nav-third");
      await page.evaluate(async () => {
        await selectProductionCounterOrder("order-two", true);
        await selectProductionCounterMachine("machine-two");
      });
      await expect(page.getByRole("tabpanel")).toHaveAttribute("data-order-id", "nav-third");
      expect(await page.evaluate(() => canManageProduction())).toBe(false);
      await expect(page.locator("#tabs").getByRole("button", { name: /Personalverwaltung|Admin \/ System/ })).toHaveCount(0);
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });

    for (const role of ["employee", "tool_scanner"]) {
      test(`${role} cannot see or invoke production order navigation`, async ({ page }) => {
        const before = await page.evaluate(() => [state.ui.productionCounterSelectedMachineId, state.ui.productionCounterActiveOrderId]);
        await page.evaluate(async (role) => {
          currentUser = { name: role, role };
          currentEmployeeRecord = null;
          currentTab = "dashboard";
          render();
          await selectProductionCounterOrder("nav-second", true);
          await selectProductionCounterMachine("machine-one");
        }, role);
        await expect(page.locator("#tabs").getByRole("button", { name: "Produktion", exact: true })).toHaveCount(0);
        await expect(page.getByRole("tablist", { name: "Aktive Aufträge" })).toHaveCount(0);
        expect(await page.evaluate(() => [state.ui.productionCounterSelectedMachineId, state.ui.productionCounterActiveOrderId])).toEqual(before);
        expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
      });
    }

    test("touch tabs scroll internally without page overflow or vertical jumps", async ({ page }, testInfo) => {
      await page.evaluate(async () => {
        const rows = window.__SUPABASE_MOCK_ROWS;
        for (let i = 0; i < 10; i++) {
          const id = `many-${i}`;
          rows.production_orders.push({ ...rows.production_orders[0], id, ba_number: `BA-${id}`,
            article_number: `LANGER-ARTIKEL-4345433-${i}` });
          rows.production_order_stations.push({ ...rows.production_order_stations[0], id: `station-${id}`, order_id: id });
          rows.production_order_checklist.push(...rows.production_order_checklist.filter((item) => item.order_id === "order-one")
            .map((item) => ({ ...item, id: `${id}-${item.id}`, order_id: id })));
        }
        await refreshProductionOrdersFromSupabase();
        await refreshProductionOrderStationsFromSupabase();
        await refreshProductionOrderChecklistFromSupabase();
        await selectProductionCounterOrder("order-one", true);
      });
      const strip = page.getByRole("tablist", { name: "Aktive Aufträge" });
      const target = strip.locator('[data-order-id="many-9"]');
      if (page.counterTouch) await target.tap({ trial: true });
      else await target.click({ trial: true });
      const before = await strip.boundingBox();
      const yBefore = await page.evaluate(() => window.scrollY);
      const tabBox = await target.boundingBox();
      expect(tabBox.width).toBeGreaterThanOrEqual(44);
      expect(tabBox.height).toBeGreaterThanOrEqual(44);
      expect(await strip.evaluate((el) => el.scrollWidth > el.clientWidth)).toBe(true);
      await pressCounterButton(page, target);
      await expect(page.getByRole("tabpanel")).toHaveAttribute("data-order-id", "many-9");
      await expect.poll(async () => Math.abs((await strip.boundingBox()).y - before.y)).toBeLessThanOrEqual(30);
      expect(Math.abs(await page.evaluate(() => window.scrollY) - yBefore)).toBeLessThanOrEqual(30);
      expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
      expect(await strip.evaluate((el) => el.scrollLeft)).toBeGreaterThan(0);
      await testInfo.attach("article-tabs", { body: await page.screenshot(), contentType: "image/png" });
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });

    test("load error retains selected order quantities and all existing production data", async ({ page }) => {
      await page.evaluate(() => selectProductionCounterOrder("nav-second", true));
      const before = await page.evaluate(() => JSON.stringify({
        orders: state.productionOrders, counts: state.productionStationCounts, stations: state.productionOrderStations,
      }));
      await page.evaluate(async () => {
        window.__SUPABASE_MOCK_ERRORS.production_orders = { select: { code: "42501", message: "Aufträge nicht lesbar" } };
        await refreshProductionOrdersFromSupabase();
        renderPreservingScroll("production-order-tabs");
      });
      await expect(page.getByRole("alert")).toContainText("Aufträge nicht lesbar");
      await expect(page.getByRole("tabpanel")).toHaveAttribute("data-order-id", "nav-second");
      expect(await page.evaluate(() => JSON.stringify({
        orders: state.productionOrders, counts: state.productionStationCounts, stations: state.productionOrderStations,
      }))).toBe(before);
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });
  });

  test.describe("Production touch view", () => {
    test.beforeEach(async ({ page }, testInfo) => {
      page.counterTouch = !!(testInfo.project.use.hasTouch || testInfo.project.use.isMobile);
      await page.addInitScript(installProductionAnimationCapture);
      await page.goto(appUrl);
      await page.getByRole("button", { name: "Mido/Admin" }).click();
      await clickMain(page, "Produktion");
      await clickSubtab(page, "Aufträge / BA");
      await page.locator('tr[data-production-order-id="order-one"]').getByRole("button", { name: "BA BA-100 öffnen", exact: true }).click();
      if (!testInfo.title.includes("responsive defaults")) {
        await page.getByRole("group", { name: "Zähleransicht" }).getByRole("button", { name: "Produktionsansicht", exact: true }).click();
      }
    });

    test.afterEach(async ({ page }) => {
      const writes = await page.evaluate(() => window.__SUPABASE_WRITE_LOG);
      expect(writes.every((entry) => [
        "production_station_counts", "production_station_events", "production_order_stations", "production_order_employees",
      ].includes(entry.table))).toBe(true);
    });

    test("responsive defaults and explicit preference survive navigation and reload", async ({ page }, testInfo) => {
      const productionDefault = testInfo.project.name !== "chromium";
      const toggle = page.getByRole("group", { name: "Zähleransicht" });
      await expect(toggle.getByRole("button", { name: "Produktionsansicht", exact: true })).toHaveAttribute("aria-pressed", String(productionDefault));
      expect(await page.evaluate(() => localStorage.getItem("humbel.productionCounterView"))).toBe(null);
      await toggle.getByRole("button", { name: productionDefault ? "Übersicht" : "Produktionsansicht", exact: true }).click();
      const chosen = productionDefault ? "overview" : "production";
      await page.evaluate(() => selectProductionCounterOrder("nav-second"));
      await clickMain(page, "Start / Dashboard");
      await clickMain(page, "Produktion");
      await expect(toggle.getByRole("button", { name: chosen === "production" ? "Produktionsansicht" : "Übersicht", exact: true })).toHaveAttribute("aria-pressed", "true");
      await expect(page.locator("#productionActiveOrderPanel")).toHaveAttribute("data-order-id", "nav-second");
      await page.reload();
      await page.getByRole("button", { name: "Mido/Admin" }).click();
      await clickMain(page, "Produktion");
      await expect(toggle.getByRole("button", { name: chosen === "production" ? "Produktionsansicht" : "Übersicht", exact: true })).toHaveAttribute("aria-pressed", "true");
      expect(await page.evaluate(() => localStorage.getItem("humbel.productionCounterView"))).toBe(chosen);
      const stored = await page.evaluate(() => Object.values(localStorage).join(" "));
      expect(stored).not.toMatch(/"(?:good_qty|scrap_total|clarify_total|productionStationCounts)"\s*:/);
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });

    test("overview stays complete while touch view contains only counting controls", async ({ page }) => {
      const before = await page.evaluate(() => JSON.stringify(window.__SUPABASE_MOCK_ROWS));
      const panel = page.locator("#productionActiveOrderPanel");
      await expect(page.locator("#versionFooter")).toBeHidden();
      await expect(panel.locator("article")).toHaveCount(1);
      await expect(panel).not.toContainText(/Details bearbeiten|Gutteile gesamt|Restmenge|Checkliste|Protokoll anzeigen|OP-Nummer/);
      await expect(panel.getByRole("button", { name: "Auftrag abschließen", exact: true })).toBeVisible();
      await page.getByRole("group", { name: "Zähleransicht" }).getByRole("button", { name: "Übersicht", exact: true }).click();
      await expectViewHeading(page, "Produktionsvorschau");
      await expect(page.locator("#versionFooter")).toBeVisible();
      await expect(panel.locator("article")).toHaveCount(4);
      await expect(panel).toContainText("Gutteile gesamt");
      await expect(panel).toContainText("Mitarbeiter am Auftrag");
      await expect(panel).toContainText("Checkliste");
      await expect(panel.getByRole("button", { name: "Protokoll anzeigen" })).toBeVisible();
      await expect(panel.getByRole("button", { name: "Auftrag fertig melden", exact: true })).toBeVisible();
      await expect(panel.getByLabel("OP-Nummer")).toHaveCount(4);
      await page.getByRole("group", { name: "Zähleransicht" }).getByRole("button", { name: "Produktionsansicht", exact: true }).click();
      expect(await page.evaluate(() => JSON.stringify(window.__SUPABASE_MOCK_ROWS))).toBe(before);
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });

    test("station and article tabs preserve unique selections and separate values", async ({ page }) => {
      await expect(page.locator(".production-touch-heading")).toContainText("Maschine 50");
      const stationTabs = page.getByRole("tablist", { name: "Spannungsauswahl" });
      await expect(stationTabs.getByRole("tab")).toHaveCount(4);
      for (let n = 1; n <= 4; n++) {
        await pressCounterButton(page, stationTabs.getByRole("tab", { name: `Spannung ${n}`, exact: true }));
        const id = n === 1 ? "station-one" : `touch-station-${n}`;
        await expect(page.locator("#productionSelectedStation")).toHaveAttribute("data-station-id", id);
        await expect(page.locator(`div[data-scroll-anchor="good:${id}:counter-worker"] [data-counter-value]`)).toContainText(String(n === 1 ? 2 : n + 2));
        await expect(page.locator(".production-touch-view article")).toHaveCount(1);
      }
      await page.evaluate(() => selectProductionCounterStation("order-one", "station-nav-second"));
      await expect(page.locator("#productionSelectedStation")).toHaveAttribute("data-station-id", "touch-station-4");
      await pressCounterButton(page, page.getByRole("tablist", { name: "Aktive Aufträge" }).locator('[data-order-id="nav-second"]'));
      await expect(page.locator("#productionActiveOrderPanel")).toHaveAttribute("data-order-id", "nav-second");
      await expect(page.locator(".production-touch-view")).not.toContainText("Lavdrim Muster");
      await pressCounterButton(page, page.getByRole("tablist", { name: "Aktive Aufträge" }).locator('[data-order-id="order-one"]'));
      await expect(page.locator("#productionSelectedStation")).toHaveAttribute("data-station-id", "touch-station-4");
      for (const name of ["Übersicht", "Produktionsansicht"]) {
        await page.getByRole("group", { name: "Zähleransicht" }).getByRole("button", { name, exact: true }).click();
      }
      await expect(page.locator("#productionSelectedStation")).toHaveAttribute("data-station-id", "touch-station-4");
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });

    test("multiple employees count separately using existing good plus and minus actions", async ({ page }) => {
      const first = counterArea(page, "good");
      const second = page.locator('div[data-scroll-anchor="good:station-one:touch-worker"]');
      await pressCounterButton(page, first.getByRole("button", { name: /Gutteile \+1/ }));
      await expect(first.locator("[data-counter-value]")).toContainText("3");
      await expect(second.locator("[data-counter-value]")).toContainText("3");
      await pressCounterButton(page, second.getByRole("button", { name: /Gutteile \+1/ }));
      await expect(second.locator("[data-counter-value]")).toContainText("4");
      await pressCounterButton(page, first.getByRole("button", { name: /Gutteile -1/ }));
      await expect(first.locator("[data-counter-value]")).toContainText("2");
      await expect(second.locator("[data-counter-value]")).toContainText("4");
      const events = await page.evaluate(() => window.__SUPABASE_WRITE_LOG.filter((entry) => entry.table === "production_station_events"));
      expect(events).toHaveLength(3);
      expect(events.map((entry) => entry.payload[0].order_employee_id)).toEqual(["counter-worker", "touch-worker", "counter-worker"]);
    });

    test("employee add disclosure reuses assignment and displays a new separate counter", async ({ page }) => {
      const employees = page.getByRole("region", { name: "Mitarbeiter am Auftrag" });
      await employees.locator("summary").click();
      await employees.getByLabel("Mitarbeiter hinzufügen").selectOption("employee-extra");
      await pressCounterButton(page, employees.getByRole("button", { name: "Hinzufügen", exact: true }));
      await expect(page.locator('div[data-scroll-anchor^="good:station-one:"]').filter({ hasText: "Anna Neu" })).toContainText("0");
      const writes = await page.evaluate(() => window.__SUPABASE_WRITE_LOG);
      expect(writes).toHaveLength(1);
      expect(writes[0].table).toBe("production_order_employees");
      expect(writes[0].payload[0]).toMatchObject({ order_id: "order-one", employee_id: "employee-extra", active: true });
    });

    for (const [type, color] of [["scrap", "rgb(254, 202, 202)"], ["clarify", "rgb(254, 215, 170)"]]) {
      test(`${type} plus requires 6M and correction uses the same colored feedback`, async ({ page }) => {
        await bookCounter(page, type, 1);
        await expect(counterArea(page, type).locator("[data-counter-value]")).toHaveText("3");
        await expect.poll(() => page.evaluate(() => window.__counterAnimations.length)).toBe(1);
        expect(await page.evaluate(() => getComputedStyle(window.__counterAnimations[0].effect.target).backgroundColor)).toBe(color);
        const events = await page.evaluate(() => window.__SUPABASE_WRITE_LOG.filter((entry) => entry.table === "production_station_events"));
        expect(events[0].payload[0]).toMatchObject({ order_id: "order-one", station_id: "station-one", qa_cause_id: "qa-cause-human" });
        await bookCounter(page, type, -1);
        await expect(counterArea(page, type).locator("[data-counter-value]")).toHaveText("2");
        await expect.poll(() => page.evaluate(() => window.__counterAnimations.length)).toBe(2);
        expect(await page.evaluate(() => getComputedStyle(window.__counterAnimations[1].effect.target).backgroundColor)).toBe(color);
      });
    }

    test("cancelled 6M dialog restores the counter anchor without booking", async ({ page }) => {
      const area = counterArea(page, "scrap");
      const button = area.getByRole("button", { name: "Ausschuss +1", exact: true });
      if (page.counterTouch) await button.tap({ trial: true });
      else await button.click({ trial: true });
      const before = await area.boundingBox();
      await pressCounterButton(page, button);
      await page.getByRole("button", { name: "Abbrechen", exact: true }).first().click();
      await expect(page.getByRole("heading", { name: "6M-Ursache für Ausschuss" })).toHaveCount(0);
      await expect.poll(async () => Math.abs((await area.boundingBox()).y - before.y)).toBeLessThanOrEqual(30);
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
      await expect(area.locator("[data-counter-value]")).toHaveText("2");
    });

    test("invalid or vanished 6M cause cannot be saved", async ({ page }) => {
      await pressCounterButton(page, counterArea(page, "clarify").getByRole("button", { name: "In Abklärung +1" }));
      await page.evaluate(() => setProductionQaCauseSelection("missing-cause"));
      await page.getByRole("button", { name: "Ursache speichern" }).click();
      await expect(page.locator("#view")).toContainText("Bitte eine 6M-Ursache auswählen.");
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });

    test("target limit blocks all positive amounts while correction remains available", async ({ page }) => {
      await page.evaluate(async () => {
        window.__SUPABASE_MOCK_ROWS.production_orders[0].target_quantity = 33;
        await refreshProductionOrdersFromSupabase();
        renderPreservingScroll("station:station-one");
      });
      await expect(page.locator(".production-touch-station")).toContainText("Ziel erreicht");
      for (const type of ["good", "scrap", "clarify"]) {
        await expect(counterArea(page, type).getByRole("button", { name: /\+1/ })).toBeDisabled();
      }
      await page.evaluate(async () => {
        await adjustProductionStationGoodQty("station-one", "counter-worker", 1);
        for (const type of ["scrap", "clarify"]) openProductionQaCauseModal("station-one", type);
      });
      await expect(page.getByRole("heading", { name: /6M-Ursache für/ })).toHaveCount(0);
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
      await bookCounter(page, "good", -1);
      await expect(counterArea(page, "good").getByRole("button", { name: /\+1/ })).toBeEnabled();
    });

    test("target is rechecked when the 6M dialog is saved", async ({ page }) => {
      await pressCounterButton(page, counterArea(page, "scrap").getByRole("button", { name: "Ausschuss +1" }));
      await page.evaluate(async () => {
        window.__SUPABASE_MOCK_ROWS.production_orders[0].target_quantity = 33;
        await refreshProductionOrdersFromSupabase();
      });
      await page.getByRole("button", { name: "Ursache speichern" }).click();
      await expect(counterArea(page, "scrap")).toContainText("Ziel erreicht - +1 gesperrt");
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });

    test("rapid good presses while saving produce one mutation and event", async ({ page }) => {
      await page.evaluate(() => {
        window.__SUPABASE_MOCK_HOLD = { table: "production_station_counts" };
        window.__pendingCounter = adjustProductionStationGoodQty("station-one", "counter-worker", 1);
        void adjustProductionStationGoodQty("station-one", "counter-worker", 1);
        void adjustProductionStationGoodQty("station-one", "counter-worker", 1);
      });
      await expect.poll(() => page.evaluate(() => !!window.__SUPABASE_MOCK_HOLD.release)).toBe(true);
      await expect(counterArea(page, "good").getByRole("button", { name: /\+1/ })).toBeDisabled();
      await page.evaluate(async () => {
        window.__SUPABASE_MOCK_HOLD.release();
        delete window.__SUPABASE_MOCK_HOLD;
        await window.__pendingCounter;
      });
      await expect(counterArea(page, "good").locator("[data-counter-value]")).toContainText("3");
      const writes = await page.evaluate(() => window.__SUPABASE_WRITE_LOG);
      expect(writes.filter((entry) => entry.table === "production_station_counts")).toHaveLength(1);
      expect(writes.filter((entry) => entry.table === "production_station_events")).toHaveLength(1);
    });

    test("touch targets animation height and scroll stay stable without horizontal overflow", async ({ page }, testInfo) => {
      const area = counterArea(page, "good");
      const button = area.getByRole("button", { name: /\+1/ });
      if (page.counterTouch) await button.tap({ trial: true });
      else await button.click({ trial: true });
      const before = await area.boundingBox();
      const height = (await page.locator(".production-touch-station").boundingBox()).height;
      const scroll = await page.evaluate(() => window.scrollY);
      for (const type of ["good", "scrap", "clarify"]) {
        const box = await counterArea(page, type).getByRole("button", { name: /\+1/ }).boundingBox();
        expect(box.width).toBeGreaterThanOrEqual(60);
        expect(box.height).toBeGreaterThanOrEqual(60);
      }
      await bookCounter(page, "good", 1);
      await expect.poll(() => page.evaluate(() => window.__counterAnimations.length)).toBe(1);
      const after = await area.boundingBox();
      expect(Math.abs(after.height - before.height)).toBeLessThanOrEqual(1);
      expect(Math.abs((await page.locator(".production-touch-station").boundingBox()).height - height)).toBeLessThanOrEqual(1);
      expect(Math.abs(after.y - before.y)).toBeLessThanOrEqual(30);
      if (scroll > 100) expect(await page.evaluate(() => window.scrollY)).toBeGreaterThan(50);
      expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
      await expect(area).not.toContainText(/wurde gezählt|gespeichert/);
      expect(await page.evaluate(() => getComputedStyle(window.__counterAnimations[0].effect.target).backgroundColor)).toBe("rgb(187, 247, 208)");
      await testInfo.attach("production-touch-view", { body: await page.locator(".production-touch-view").screenshot(), contentType: "image/png" });
      await page.evaluate(async () => { const animation = window.__counterAnimations[0]; animation.play(); await animation.finished; });
    });

    test("complete order opens mandatory checklist and cannot bypass unmet conditions", async ({ page }) => {
      await openProductionCompletionByPointer(page);
      await expect(page.getByRole("heading", { name: "Auftrag fertig melden", exact: true })).toBeVisible();
      await expect(page.getByRole("checkbox")).toHaveCount(10);
      await expect(page.getByRole("button", { name: "Endgültig fertig melden" })).toBeDisabled();
      await page.evaluate(() => completeProductionOrderWithChecklist("order-one"));
      await expect(page.getByRole("alert")).toContainText("Auftrag hat noch Restmenge.");
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });

    test("completion touch surface remains clear after scrolling with larger header typography", async ({ page }, testInfo) => {
      // Linux font metrics exposed the overflow in CI; enlarged type reproduces it on Windows too.
      await page.locator("h1").evaluate(el => { el.style.fontSize = "28px"; });
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.evaluate(() => window.scrollTo(0, document.documentElement.scrollHeight));
      expect(await page.evaluate(() => scrollY)).toBeGreaterThan(0);
      await openProductionCompletionByPointer(page);
      await expect(page.getByRole("heading", { name: "Auftrag fertig melden", exact: true })).toBeVisible();
      await expect(page.getByRole("checkbox")).toHaveCount(10);
      await expect(page.getByRole("button", { name: "Endgültig fertig melden" })).toBeDisabled();
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
      await pressCounterButton(page, page.getByRole("button", { name: "Abbrechen", exact: true }));
      await expect(page.getByRole("heading", { name: "Auftrag fertig melden", exact: true })).toHaveCount(0);
      await testInfo.attach("completion-touch-surface", { body: await page.screenshot(), contentType: "image/png" });
      // The existing controls must still work after opening and cancelling completion.
      for (const type of ["good", "scrap", "clarify"]) await bookCounter(page, type, 1);
      await pressCounterButton(page, page.getByRole("tablist", { name: "Spannungsauswahl" }).getByRole("tab", { name: "Spannung 4", exact: true }));
      await expect(page.locator("#productionSelectedStation")).toHaveAttribute("data-station-id", "touch-station-4");
      await pressCounterButton(page, page.getByRole("tablist", { name: "Aktive Aufträge" }).locator('[data-order-id="nav-second"]'));
      await expect(page.locator("#productionActiveOrderPanel")).toHaveAttribute("data-order-id", "nav-second");
    });

    test("completion checklist blocks finishing even when no remaining quantity is open", async ({ page }) => {
      await page.evaluate(async () => {
        window.__SUPABASE_MOCK_ROWS.production_orders.find(row => row.id === "order-one").target_quantity = 1;
        window.__SUPABASE_MOCK_ROWS.production_order_checklist.filter(item => item.order_id === "order-one").forEach(item => { item.checked = false; });
        await refreshProductionOrdersFromSupabase();
        await refreshProductionOrderChecklistFromSupabase();
        renderPreservingScroll();
      });
      expect(await page.evaluate(() => calculatePreparedRemainingQuantity(getProductionOrderById("order-one")))).toBe(0);
      await openProductionCompletionByPointer(page);
      await expect(page.getByRole("checkbox")).toHaveCount(10);
      await expect(page.getByRole("checkbox", { checked: false })).toHaveCount(10);
      await expect(page.getByRole("button", { name: "Endgültig fertig melden" })).toBeDisabled();
      await page.evaluate(() => completeProductionOrderWithChecklist("order-one"));
      await expect(page.getByRole("alert")).toContainText("Alle Checklistenpunkte müssen erledigt sein.");
      expect(await page.evaluate(() => getProductionOrderById("order-one").status)).toBe("running");
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    });

    test("department admin keeps own department access and cannot navigate to foreign work", async ({ page }) => {
      await page.evaluate(() => {
        currentUser = { name: "Leitung", role: "department_admin" };
        currentEmployeeRecord = state.employeesList.find((employee) => employee.id === "leader-test");
        render();
      });
      await expect(counterArea(page, "good").getByRole("button", { name: /\+1/ })).toBeEnabled();
      await page.evaluate(() => selectProductionCounterOrder("order-two", true));
      await expect(page.locator("#productionActiveOrderPanel")).toHaveAttribute("data-order-id", "order-one");
      expect(await page.evaluate(() => canManageProduction())).toBe(false);
    });

    for (const role of ["employee", "tool_scanner"]) {
      test(`${role} cannot enter the production touch view`, async ({ page }) => {
        await page.evaluate(async (role) => {
          currentUser = { name: role, role };
          currentEmployeeRecord = null;
          render();
          setProductionCounterView("overview");
          await selectProductionCounterOrder("order-one", true);
        }, role);
        await expect(page.locator(".production-touch-view")).toHaveCount(0);
        await expect(page.locator("#tabs").getByRole("button", { name: "Produktion", exact: true })).toHaveCount(0);
        expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
      });
    }

    test("failed count write preserves values and shows a technical error without feedback", async ({ page }) => {
      await page.evaluate(() => {
        window.__SUPABASE_MOCK_ERRORS.production_station_counts = { update: { code: "42501", message: "Buchung nicht erlaubt" } };
      });
      await bookCounter(page, "good", 1);
      await expect(page.getByRole("alert")).toContainText("Buchung nicht erlaubt");
      await expect(counterArea(page, "good").locator("[data-counter-value]")).toContainText("2");
      expect(await page.evaluate(() => window.__counterAnimations.length)).toBe(0);
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG.some((entry) => entry.table === "production_station_events"))).toBe(false);
    });

    test("load error preserves all existing production data and the active station", async ({ page }) => {
      await page.evaluate(() => selectProductionCounterStation("order-one", "touch-station-3"));
      const before = await page.evaluate(() => JSON.stringify({ orders: state.productionOrders, stations: state.productionOrderStations, counts: state.productionStationCounts }));
      await page.evaluate(async () => {
        window.__SUPABASE_MOCK_ERRORS.production_orders = { select: { message: "Netzwerk nicht erreichbar" } };
        await refreshProductionOrdersFromSupabase();
        renderPreservingScroll("production-station-tabs");
      });
      await expect(page.getByRole("alert")).toContainText("Netzwerk nicht erreichbar");
      await expect(page.locator("#productionSelectedStation")).toHaveAttribute("data-station-id", "touch-station-3");
      expect(await page.evaluate(() => JSON.stringify({ orders: state.productionOrders, stations: state.productionOrderStations, counts: state.productionStationCounts }))).toBe(before);
    });
  });

  test.describe("Production counter feedback", () => {
    test.beforeEach(async ({ page }, testInfo) => {
      page.counterTouch = !!testInfo.project.use.isMobile;
      await page.addInitScript(installProductionAnimationCapture);
      await page.goto(appUrl);
      await page.getByRole("button", { name: "Mido/Admin" }).click();
      await clickMain(page, "Produktion");
      await clickSubtab(page, "Stückzahl");
      await page.getByRole("button", { name: /Maschine 50/ }).click();
      await expectViewHeading(page, "Produktionsvorschau");
    });

    test.afterEach(async ({ page }) => {
      const writes = await page.evaluate(() => window.__SUPABASE_WRITE_LOG);
      expect(writes.some((entry) => entry.table === "production_counts")).toBe(false);
      expect(writes.filter((entry) => ![
        "production_station_counts", "production_station_events", "production_order_stations",
      ].includes(entry.table))).toEqual([]);
    });

    for (const [type, color] of [
      ["good", "rgb(187, 247, 208)"], ["scrap", "rgb(254, 202, 202)"], ["clarify", "rgb(254, 215, 170)"],
    ]) {
      for (const delta of [1, -1]) {
        test(`${type} ${delta > 0 ? "+1" : "-1"} confirms with color and stable layout`, async ({ page }, testInfo) => {
          const area = counterArea(page, type);
          const station = page.locator('article[data-scroll-anchor="station:station-one"]');
          await area.scrollIntoViewIfNeeded();
          if (page.counterTouch) {
            await area.getByRole("button", { name: delta > 0 ? /\+1/ : /-1/ }).tap({ trial: true });
          }
          const before = await area.boundingBox();
          const stationHeight = (await station.boundingBox()).height;
          const scrollBefore = await page.evaluate(() => window.scrollY);
          await bookCounter(page, type, delta);
          await expect(area.locator("[data-counter-value]")).toContainText(String(2 + delta));
          await expect.poll(() => page.evaluate(() => window.__counterAnimations.length)).toBe(1);
          const effect = await page.evaluate(() => {
            const animation = window.__counterAnimations[0];
            return {
              color: getComputedStyle(animation.effect.target).backgroundColor,
              duration: animation.effect.getTiming().duration,
              count: animation.effect.target.getAnimations().length,
            };
          });
          expect(effect).toEqual({ color, duration: 600, count: 1 });
          const after = await area.boundingBox();
          expect(Math.abs(after.height - before.height)).toBeLessThanOrEqual(1);
          expect(Math.abs((await station.boundingBox()).height - stationHeight)).toBeLessThanOrEqual(1);
          expect(Math.abs(after.y - before.y)).toBeLessThanOrEqual(30);
          if (scrollBefore > 100) expect(await page.evaluate(() => window.scrollY)).toBeGreaterThan(50);
          expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
          await expect(area).not.toContainText(/wurde gezählt|Korrektur wurde gespeichert/);
          await expect(page.locator("#productionCounterAnnouncement")).toHaveText(/gezählt|gespeichert/);
          expect((await page.locator("#productionCounterAnnouncement").boundingBox()).height).toBe(1);
          if (type === "good" && delta === 1) {
            await testInfo.attach("counter-feedback", { body: await page.screenshot(), contentType: "image/png" });
          }
          await page.evaluate(async () => {
            const animation = window.__counterAnimations[0];
            animation.play();
            await animation.finished;
          });
          expect(await area.locator("[data-counter-value]").evaluate((el) => el.getAnimations().length)).toBe(0);
        });
      }

      test(`${type} failed database write never flashes success`, async ({ page }) => {
        const table = type === "good" ? "production_station_counts" : "production_order_stations";
        await page.evaluate((table) => {
          window.__SUPABASE_MOCK_ERRORS[table] = { update: { code: "42501", message: "Buchung durch RLS abgelehnt" } };
        }, table);
        const area = counterArea(page, type);
        const height = (await area.boundingBox()).height;
        await bookCounter(page, type, 1);
        await expect(page.getByRole("alert")).toContainText("Buchung durch RLS abgelehnt");
        await expect(area.locator("[data-counter-value]")).toContainText("2");
        expect((await area.boundingBox()).height).toBe(height);
        expect(await page.evaluate(() => window.__counterAnimations.length)).toBe(0);
        expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG.filter((entry) =>
          entry.table === "production_station_events").length)).toBe(0);
      });

      test(`${type} rapid duplicate calls produce one booking and one event`, async ({ page }) => {
        const table = type === "good" ? "production_station_counts" : "production_order_stations";
        await page.evaluate(({ type, table }) => {
          window.__SUPABASE_MOCK_HOLD = { table };
          const book = () => type === "good"
            ? window.adjustProductionStationGoodQty("station-one", "counter-worker", 1)
            : window.adjustProductionStationAmount("station-one", type, 1, { qaCauseId: "qa-cause-human" });
          window.__pendingCounter = book();
          void book();
          void book();
        }, { type, table });
        await expect.poll(() => page.evaluate(() => !!window.__SUPABASE_MOCK_HOLD.release)).toBe(true);
        await expect(counterArea(page, type).getByRole("button", { name: /\+1/ })).toBeDisabled();
        await expect.poll(() => page.evaluate(() => window.__counterAnimations.length)).toBe(0);
        await page.evaluate(async () => {
          window.__SUPABASE_MOCK_HOLD.release();
          delete window.__SUPABASE_MOCK_HOLD;
          await window.__pendingCounter;
        });
        await expect(counterArea(page, type).locator("[data-counter-value]")).toContainText("3");
        const writes = await page.evaluate(() => window.__SUPABASE_WRITE_LOG);
        expect(writes.filter((entry) => entry.table === table && entry.action === "update")).toHaveLength(1);
        expect(writes.filter((entry) => entry.table === "production_station_events")).toHaveLength(1);
      });
    }

    test("target block prevents every mutation and keeps correction available", async ({ page }) => {
      await page.evaluate(async () => {
        window.__SUPABASE_MOCK_ROWS.production_orders.find((row) => row.id === "order-one").target_quantity = 6;
        await refreshProductionOrdersFromSupabase();
        renderPreservingScroll("station:station-one");
      });
      const station = page.locator('article[data-scroll-anchor="station:station-one"]');
      await expect(station).toContainText("Ziel erreicht");
      const before = await page.evaluate(() => window.__SUPABASE_WRITE_LOG.length);
      await page.evaluate(async () => {
        await adjustProductionStationGoodQty("station-one", "counter-worker", 1);
        for (const type of ["scrap", "clarify"]) {
          openProductionQaCauseModal("station-one", type);
          await adjustProductionStationAmount("station-one", type, 1, { qaCauseId: "qa-cause-human" });
        }
      });
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG.length)).toBe(before);
      expect(await page.evaluate(() => window.__counterAnimations.length)).toBe(0);
      await expect(page.getByRole("heading", { name: /6M-Ursache für/ })).toHaveCount(0);
      for (const type of ["good", "scrap", "clarify"]) {
        await expect(counterArea(page, type).getByRole("button", { name: /\+1/ })).toBeDisabled();
        await expect(counterArea(page, type).getByRole("button", { name: /-1/ })).toBeEnabled();
      }
    });

    test("five station metrics keep their values and shortened labels", async ({ page }) => {
      await expect(page.getByRole("tablist").getByRole("tab")).toHaveCount(1);
      await expect(page.getByRole("tab", { selected: true })).toHaveText(/ART-100\s+BA BA-100/);
      const chips = page.locator('article[data-scroll-anchor="station:station-one"] > div').nth(1).locator(":scope > div");
      await expect(chips).toHaveCount(5);
      const expected = [["Gutteile", "2"], ["Ausschuss", "2"], ["In Abklärung", "2"], ["Gesamt", "6"], ["Ziel", "50"]];
      for (let i = 0; i < expected.length; i++) {
        await expect(chips.nth(i).locator("div").nth(0)).toHaveText(expected[i][0]);
        await expect(chips.nth(i).locator("div").nth(1)).toHaveText(expected[i][1]);
      }
      await expect(page.locator('article[data-scroll-anchor="station:station-one"]')).not.toContainText(/Ausschuss gesamt|In Abklärung gesamt/);
    });

    test("target reached while 6M dialog is open blocks saving without feedback or writes", async ({ page }) => {
      await pressCounterButton(page, counterArea(page, "scrap").getByRole("button", { name: "Ausschuss +1" }));
      await expect(page.getByRole("heading", { name: "6M-Ursache für Ausschuss" })).toBeVisible();
      await page.evaluate(async () => {
        window.__SUPABASE_MOCK_ROWS.production_orders.find((row) => row.id === "order-one").target_quantity = 6;
        await refreshProductionOrdersFromSupabase();
      });
      const before = await page.evaluate(() => window.__SUPABASE_WRITE_LOG.length);
      await pressCounterButton(page, page.getByRole("button", { name: "Ursache speichern" }));
      await expect(page.getByRole("heading", { name: "6M-Ursache für Ausschuss" })).toHaveCount(0);
      await expect(counterArea(page, "scrap")).toContainText("Ziel erreicht - +1 gesperrt");
      expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG.length)).toBe(before);
      expect(await page.evaluate(() => window.__counterAnimations.length)).toBe(0);
    });

    test("failed event write keeps the technical error and avoids a success animation", async ({ page }) => {
      await page.evaluate(() => {
        window.__SUPABASE_MOCK_ERRORS.production_station_events = {
          insert: { code: "42501", message: "Protokoll durch RLS abgelehnt" },
        };
      });
      await bookCounter(page, "good", 1);
      await expect(page.getByRole("alert")).toContainText("Gutteil gespeichert, Protokolleintrag konnte nicht geschrieben werden.");
      await expect(counterArea(page, "good").locator("[data-counter-value]")).toContainText("3");
      expect(await page.evaluate(() => window.__counterAnimations.length)).toBe(0);
    });

    test("reduced motion announces success without animation", async ({ page }) => {
      await page.emulateMedia({ reducedMotion: "reduce" });
      await bookCounter(page, "good", 1);
      await expect(counterArea(page, "good").locator("[data-counter-value]")).toContainText("3");
      await expect(page.locator("#productionCounterAnnouncement")).toHaveText("Gutteil wurde gezählt.");
      expect(await page.evaluate(() => window.__counterAnimations.length)).toBe(0);
    });

    test("consecutive successful bookings restart the effect after full render", async ({ page }) => {
      await bookCounter(page, "good", 1);
      await expect.poll(() => page.evaluate(() => window.__counterAnimations.length)).toBe(1);
      await bookCounter(page, "good", 1);
      await expect.poll(() => page.evaluate(() => window.__counterAnimations.length)).toBe(2);
      await expect(counterArea(page, "good").locator("[data-counter-value]")).toContainText("4");
      expect(await counterArea(page, "good").locator("[data-counter-value]").evaluate((el) => el.getAnimations().length)).toBe(1);
    });
  });
});

async function openProductionCompletionByPointer(page) {
  const button = page.getByRole("button", { name: "Auftrag abschließen", exact: true });
  await button.scrollIntoViewIfNeeded();
  const geometry = await button.evaluate(el => {
    const rect = el.getBoundingClientRect();
    const station = document.getElementById("productionSelectedStation").getBoundingClientRect();
    const hits = [];
    for (const x of [0.1, 0.5, 0.9]) {
      for (const y of [0.1, 0.5, 0.9]) {
        const hit = document.elementFromPoint(rect.left + rect.width * x, rect.top + rect.height * y);
        hits.push(hit === el || el.contains(hit));
      }
    }
    return { hits, stationBottom: station.bottom, button: rect.toJSON(), layoutWidth: innerWidth,
      scrollWidth: document.documentElement.scrollWidth, visualWidth: visualViewport.width };
  });
  const width = page.viewportSize().width;
  expect(geometry.layoutWidth).toBe(width);
  expect(geometry.scrollWidth).toBeLessThanOrEqual(width);
  expect(geometry.visualWidth).toBeCloseTo(width, 0);
  expect(geometry.stationBottom).toBeLessThanOrEqual(geometry.button.top);
  expect(geometry.button.left).toBeGreaterThanOrEqual(0);
  expect(geometry.button.right).toBeLessThanOrEqual(width);
  expect(geometry.button.height).toBeGreaterThanOrEqual(52);
  expect(geometry.hits).toEqual(Array(9).fill(true));
  await page.evaluate(() => {
    window.__completionPointer = null;
    document.addEventListener("pointerdown", event => {
      window.__completionPointer = { trusted: event.isTrusted, type: event.pointerType,
        completionButton: !!event.target.closest(".production-touch-complete") };
    }, { capture: true, once: true });
  });
  await pressCounterButton(page, button);
  expect(await page.evaluate(() => window.__completionPointer)).toEqual({
    trusted: true, type: page.counterTouch ? "touch" : "mouse", completionButton: true,
  });
}

async function pressCounterButton(page, button) {
  if (page.counterTouch) await button.tap();
  else await button.click();
}

function installProductionAnimationCapture() {
  window.__counterAnimations = [];
  const animate = Element.prototype.animate;
  Element.prototype.animate = function (frames, options) {
    const animation = animate.call(this, frames, options);
    if (this.matches("[data-counter-value]")) {
      // Keep the real WAAPI effect inspectable; tests resume and await its finish.
      animation.pause();
      animation.currentTime = 0;
      window.__counterAnimations.push(animation);
    }
    return animation;
  };
}

function prepareProductionTouchRows(rows) {
  rows.employees.push(
    { ...rows.employees[1], id: "employee-two", first_name: "Roger", last_name: "Lang", display_name: "Roger Lang", personnel_no: "200" },
    { ...rows.employees[1], id: "employee-extra", first_name: "Anna", last_name: "Neu", display_name: "Anna Neu", personnel_no: "300" },
  );
  rows.production_order_employees.find((entry) => entry.id === "counter-worker").employee_name = "Lavdrim Muster";
  rows.production_order_employees.push({ id: "touch-worker", order_id: "order-one", employee_id: "employee-two",
    employee_name: "Roger Lang", personnel_no: "200", role: "worker", active: true });
  for (let n = 2; n <= 4; n++) {
    const id = `touch-station-${n}`;
    rows.production_order_stations.push({ ...rows.production_order_stations[0], id, station_no: n,
      name: `Aufspannung ${n}`, scrap_total: n, clarify_total: 0 });
    rows.production_station_counts.push({ id: `touch-count-${n}`, station_id: id, order_id: "order-one",
      order_employee_id: "counter-worker", good_qty: n + 2 });
  }
  rows.production_station_counts.push({ id: "touch-second-count", station_id: "station-one", order_id: "order-one",
    order_employee_id: "touch-worker", good_qty: 3 });
}

function prepareProductionNavigationRows(rows) {
  const first = rows.production_orders.find((order) => order.id === "order-one");
  first.article_number = "4345433";
  rows.departments.push({ id: "department-private", name: "Fremde Abteilung", code: "PRIV", active: true });
  rows.production_machines.find((machine) => machine.id === "machine-two").department_id = "department-private";
  rows.production_orders.find((order) => order.id === "order-two").department_id = "department-private";
  rows.employees.push({ id: "leader-test", display_name: "Leitung", role: "department_admin", is_active: true, auth_user_id: "leader-auth" });
  rows.departments[0].leader_employee_id = "leader-test";
  for (const [id, article, status] of [
    ["nav-second", "4712001", "running"], ["nav-third", "4345433", "paused"],
    ["nav-closed", "ART-CLOSED", "completed"], ["nav-cancelled", "ART-CANCELLED", "cancelled"],
  ]) {
    rows.production_orders.push({ ...first, id, ba_number: `BA-${id}`, article_number: article, status });
  }
  const activeOrders = rows.production_orders.filter((order) => ["running", "paused"].includes(order.status));
  for (const [index, order] of activeOrders.entries()) {
    let station = rows.production_order_stations.find((entry) => entry.order_id === order.id);
    if (!station) {
      station = { ...rows.production_order_stations[0], id: `station-${order.id}`, order_id: order.id };
      rows.production_order_stations.push(station);
    }
    station.name = `Spannung ${order.ba_number}`;
    const existingEmployee = rows.production_order_employees.find((entry) => entry.order_id === order.id);
    const employee = existingEmployee || {
      id: `worker-${order.id}`, order_id: order.id, employee_id: "employee-one", role: "worker", active: true,
    };
    employee.employee_name = `Team ${order.ba_number}`;
    if (!existingEmployee) rows.production_order_employees.push(employee);
    rows.production_station_counts = rows.production_station_counts.filter((count) => count.order_id !== order.id);
    rows.production_station_counts.push({ id: `count-${order.id}`, order_id: order.id,
      station_id: station.id, order_employee_id: employee.id, good_qty: index + 2 });
    rows.production_order_checklist = rows.production_order_checklist.filter((item) => item.order_id !== order.id);
    rows.production_order_checklist.push(...rows.production_checklist_templates.map((item) => ({
      id: `${order.id}-${item.id}`, order_id: order.id, item_key: item.item_key,
      item_label: `${order.ba_number} Checkliste ${item.sort_order}`, sort_order: item.sort_order, checked: true,
    })));
    rows.production_station_events.push({ id: `event-${order.id}`, order_id: order.id, station_id: station.id,
      employee_id: "employee-one", event_type: "good", qty: 1, note: `Protokoll ${order.ba_number}` });
  }
}

function counterArea(page, type) {
  const key = type === "good" ? "good:station-one:counter-worker" : `amount:station-one:${type}`;
  return page.locator(`div[data-scroll-anchor="${key}"]`).first();
}

async function bookCounter(page, type, delta) {
  const label = { good: "Gutteile", scrap: "Ausschuss", clarify: "In Abklärung" }[type];
  await pressCounterButton(page, counterArea(page, type).getByRole("button", {
    name: new RegExp(`^${label} ${delta > 0 ? "\\+1" : "-1"}`),
  }));
  if (type !== "good" && delta > 0) {
    await expect(page.getByRole("heading", { name: `6M-Ursache für ${label}` })).toBeVisible();
    await pressCounterButton(page, page.getByRole("button", { name: "Einweisung" }));
    await pressCounterButton(page, page.getByRole("button", { name: "Ursache speichern" }));
    await expect(page.getByRole("heading", { name: `6M-Ursache für ${label}` })).toHaveCount(0);
  }
}

async function clickMain(page, name) {
  await page.locator("#tabs").getByRole("button", { name }).click();
  await expect(page.locator("#view")).toContainText(name);
}

async function clickSubtabs(page, names) {
  for (const name of names) {
    await clickSubtab(page, name);
  }
}

async function clickSubtab(page, name) {
  const button = page.locator("#view").getByRole("button", { name }).first();
  await expect(button).toBeVisible();
  await button.click();
  await expect(button).toHaveClass(/humbel-subtab-active/);
}

async function expectViewHeading(page, name) {
  await expect(page.locator("#view").getByRole("heading", { name, exact: true })).toBeVisible();
}

async function expectNoViewHeading(page, name) {
  await expect(page.locator("#view").getByRole("heading", { name, exact: true })).toHaveCount(0);
}
