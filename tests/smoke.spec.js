const { test, expect } = require("@playwright/test");
const path = require("path");
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
  test.beforeEach(async ({ page }) => {
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
        "Fehler beim Laden von production_station_events",
        "Fehler beim Laden von production_station_counts",
      ];
      if (!ignored.some((entry) => text.includes(entry))) {
        consoleErrors.push(text);
      }
    });
    page.on("pageerror", (error) => {
      consoleErrors.push(error.message);
    });
    page.consoleErrors = consoleErrors;
    await page.route("**/cdn.tailwindcss.com**", (route) => {
      route.fulfill({ contentType: "application/javascript", body: "" });
    });
    await page.route("**/@supabase/supabase-js@2**", (route) => {
      route.fulfill({ contentType: "application/javascript", body: "" });
    });
    await page.addInitScript(installSupabaseMock, JSON.parse(JSON.stringify(rowsByTable)));
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
    await expect(page.locator("#view").getByRole("button", { name: /BA BA-100/ })).toBeVisible();
    await expect(page.locator("#view")).toContainText("Restmenge vorbereitet");
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
    await stationOne.getByRole("button", { name: /Gutteile \+1/ }).scrollIntoViewIfNeeded();
    const scrollBeforeGoodCount = await page.evaluate(() => window.scrollY);
    await stationOne.getByRole("button", { name: /Gutteile \+1/ }).click();
    await expect
      .poll(() => page.evaluate((before) => Math.abs(window.scrollY - before) <= 8, scrollBeforeGoodCount))
      .toBeTruthy();
    await expect(stationOne).toContainText(/Gutteile\s+1/);
    await expect(page.locator("#view")).toContainText("Gutteil wurde gezählt.");
    await expect(page.locator("#view")).not.toContainText("Protokolleintrag konnte nicht geschrieben werden");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(page.locator("#view").getByRole("button", { name: /BA BA-100/ })).toBeVisible();
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
    await expect(page.locator("#view")).toContainText("Restmenge vorbereitet");
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
    await expect(page.locator("#view")).toContainText("Zielstückzahl ist erreicht.");
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
    await stationOne.getByRole("button", { name: /Gutteile -1/ }).click();
    await expect(stationOne).toContainText(/Gutteile\s+1/);
    await expect(page.locator("#view")).toContainText("Gutteil-Korrektur wurde gespeichert.");
    await stationOne.getByRole("button", { name: /Gutteile -1/ }).click();
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
    await stationOne.getByRole("button", { name: /Gutteile -1/ }).click();
    await expect(page.locator("#view")).toContainText("Gutmenge kann nicht unter 0 fallen.");
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
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+0/);
    await stationOne.getByRole("button", { name: "Ausschuss +1" }).click();
    await expect(page.getByRole("heading", { name: "6M-Ursache für Ausschuss" })).toBeVisible();
    await expect(page.locator("body")).toContainText("Mensch");
    await page.getByRole("button", { name: "Abbrechen" }).first().click();
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+0/);
    await stationOne.getByRole("button", { name: "Ausschuss +1" }).scrollIntoViewIfNeeded();
    const scrollBeforeScrapCount = await page.evaluate(() => window.scrollY);
    await stationOne.getByRole("button", { name: "Ausschuss +1" }).click();
    await page.getByRole("button", { name: "Einweisung" }).click();
    await page.getByLabel("Notiz").fill("Testnotiz Ausschuss");
    await page.getByRole("button", { name: "Ursache speichern" }).click();
    await expect
      .poll(() => page.evaluate((before) => Math.abs(window.scrollY - before) <= 8, scrollBeforeScrapCount))
      .toBeTruthy();
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+1/);
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
    await expect(page.locator("#view")).toContainText("Zielstückzahl ist erreicht.");
    await page.evaluate(() =>
      window.adjustProductionStationAmount("station-one", "clarify", 1, {
        qaCauseId: "qa-cause-machine",
        note: "blocked clarify",
      }),
    );
    await expect(page.locator("#view")).toContainText("Zielstückzahl ist erreicht.");
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+1/);
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+0/);
    const amountWritesAfterBlocked = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) =>
          ["production_order_stations", "production_station_events"].includes(entry.table) &&
          ["update", "insert"].includes(entry.action),
      ).length,
    );
    expect(amountWritesAfterBlocked).toBe(amountWritesBeforeBlocked);
    await stationOne.getByRole("button", { name: "Ausschuss -1" }).click();
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+0/);
    await expect(page.locator("#view")).toContainText("Ausschuss-Korrektur wurde gespeichert.");
    await expect(page.locator("#view")).not.toContainText("Ursache/Protokoll konnte nicht geschrieben werden");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(page.locator("#view").getByRole("button", { name: /BA BA-100/ })).toBeVisible();
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
    await expect(page.locator("#view")).toContainText("Ausschuss kann nicht unter 0 fallen.");
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+0/);
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+0/);
    await stationOne.getByRole("button", { name: "In Abklärung +1" }).click();
    await expect(page.getByRole("heading", { name: "6M-Ursache für In Abklärung" })).toBeVisible();
    await page.getByRole("button", { name: "Einrichtung" }).click();
    await page.getByRole("button", { name: "Ursache speichern" }).click();
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+1/);
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
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+0/);
    await expect(page.locator("#view")).toContainText("Abklär-Korrektur wurde gespeichert.");
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
    await expect(page.locator("#view")).toContainText("Abklärmenge kann nicht unter 0 fallen.");
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+0/);
    await stationOne.getByRole("button", { name: "Ausschuss +1" }).click();
    await page.getByRole("button", { name: "Einweisung" }).click();
    await page.getByRole("button", { name: "Ursache speichern" }).click();
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+1/);
    await page.evaluate(() => {
      window.__SUPABASE_MOCK_ERRORS.production_station_events = {
        select: {
          code: "42501",
          message: "new row violates row-level security policy for table \"production_station_events\"",
        },
      };
    });
    await stationOne.getByRole("button", { name: "Ausschuss -1" }).click();
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+0/);
    await expect(page.locator("#view")).toContainText("Menge gespeichert, Protokoll konnte nicht neu geladen werden.");
    await expect(page.locator("#view")).not.toContainText("Ursache/Protokoll konnte nicht geschrieben werden");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(stationOne).toContainText("Spannung 1");
    await stationOne.getByRole("button", { name: "In Abklärung +1" }).click();
    await page.getByRole("button", { name: "Einrichtung" }).click();
    await page.getByRole("button", { name: "Ursache speichern" }).click();
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+1/);
    await page.evaluate(() => {
      window.__SUPABASE_MOCK_ERRORS.production_station_events = {
        insert: {
          code: "42501",
          message: "new row violates row-level security policy for table \"production_station_events\"",
        },
      };
    });
    await stationOne.getByRole("button", { name: "In Abklärung -1" }).click();
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+0/);
    await expect(page.locator("#view")).toContainText("Menge gespeichert, Ursache/Protokoll konnte nicht geschrieben werden.");
    await expectViewHeading(page, "Produktionsvorschau");
    await expect(page.locator("#view")).toContainText("Ausschuss Auftrag");
    await expect(page.locator("#view")).toContainText("In Abklärung Auftrag");
    await expect(page.locator("#view")).toContainText("Ursache wird beim +1 erfasst.");
    await page.locator("#view").getByRole("button", { name: /Spannung/ }).filter({ hasText: /hinzuf/ }).click();
    await expect(page.locator("#view")).toContainText("Spannung 2");
    const stationTwo = page.locator("#view").locator("article").filter({ hasText: "Spannung 2" });
    await expect(stationTwo).toContainText("Lavdrim");
    await stationTwo.getByRole("button", { name: /Gutteile \+1/ }).click();
    await expect(stationTwo).toContainText(/Gutteile\s+1/);
    await expect(stationTwo.getByRole("button", { name: /Gutteile \+1/ })).toBeDisabled();
    await expect(page.locator("#view")).toContainText("Restmenge vorbereitet");
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
    await expect(overStation).toContainText("Istmenge liegt über Zielstückzahl. Bitte per -1 korrigieren.");
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
    await expect(page.locator("#view")).toContainText("Zielstückzahl ist erreicht.");
    await expect(page.getByRole("heading", { name: "6M-Ursache für Ausschuss" })).toHaveCount(0);
    await page.evaluate(() => window.openProductionQaCauseModal("station-over", "clarify"));
    await expect(page.locator("#view")).toContainText("Zielstückzahl ist erreicht.");
    await expect(page.getByRole("heading", { name: "6M-Ursache für In Abklärung" })).toHaveCount(0);
    const blockedWritesAfter = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter(
        (entry) =>
          ["production_order_stations", "production_station_counts", "production_station_events"].includes(entry.table) &&
          ["update", "insert"].includes(entry.action),
      ).length,
    );
    expect(blockedWritesAfter).toBe(blockedWritesBefore);

    await overStation.getByRole("button", { name: "Ausschuss -1" }).scrollIntoViewIfNeeded();
    const scrollBeforeScrapCorrection = await page.evaluate(() => window.scrollY);
    await overStation.getByRole("button", { name: "Ausschuss -1" }).click();
    await expect
      .poll(() => page.evaluate((before) => Math.abs(window.scrollY - before) <= 8, scrollBeforeScrapCorrection))
      .toBeTruthy();
    await expect(overStation).toContainText(/Ausschuss gesamt\s+0/);
    await expect(overStation).not.toContainText("Istmenge liegt über Zielstückzahl. Bitte per -1 korrigieren.");
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
    await expect(page.locator("#view")).toContainText("Zielstückzahl für diese Spannung ist erreicht.");
    await expect(chainStationTwo.getByRole("button", { name: /Gutteile \+1/ })).toBeEnabled();
    await chainStationTwo.getByRole("button", { name: /Gutteile \+1/ }).click();
    await expect(chainStationTwo).toContainText(/Gutteile\s+1/);
  });
});

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
