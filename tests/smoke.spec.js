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
  ],
  production_orders: [
    {
      id: "order-one",
      machine_id: "machine-one",
      department_id: "department-one",
      ba_number: "BA-100",
      article_number: "ART-100",
      ba_quantity: 100,
      target_quantity: 100,
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
  ],
  production_order_employees: [],
  production_station_counts: [],
  production_station_events: [],
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
  window.__SUPABASE_WRITE_LOG = [];
  window.TEST_ADMIN_EMAIL = "admin@example.test";
  window.TEST_EMPLOYEE_EMAIL = "employee@example.test";

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
      window.__SUPABASE_WRITE_LOG.push({ table: this.table, action: "insert" });
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
      window.__SUPABASE_WRITE_LOG.push({ table: this.table, action: "update" });
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
    await page.addInitScript(installSupabaseMock, rowsByTable);
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
    const stationOne = page.locator("#view").locator("article").filter({ hasText: "Spannung 1" });
    await expect(stationOne).toContainText("Lavdrim");
    await expect(stationOne).toContainText(/Gutteile\s+0/);
    await stationOne.getByRole("button", { name: /Gutteile \+1/ }).click();
    await expect(stationOne).toContainText(/Gutteile\s+1/);
    await expect(page.locator("#view")).toContainText("Gutteile Auftrag");
    await expect(page.locator("#view")).toContainText("Restmenge vorbereitet");
    await expect(page.locator("#view")).toContainText("99");
    await stationOne.getByRole("button", { name: /Gutteile -1/ }).click();
    await expect(stationOne).toContainText(/Gutteile\s+0/);
    await stationOne.getByRole("button", { name: /Gutteile -1/ }).click();
    await expect(page.locator("#view")).toContainText("Gutmenge kann nicht unter 0 fallen.");
    await expect(stationOne).toContainText(/Gutteile\s+0/);
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+0/);
    await stationOne.getByRole("button", { name: "Ausschuss +1" }).click();
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+1/);
    await stationOne.getByRole("button", { name: "Ausschuss -1" }).click();
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+0/);
    await stationOne.getByRole("button", { name: "Ausschuss -1" }).click();
    await expect(page.locator("#view")).toContainText("Ausschuss kann nicht unter 0 fallen.");
    await expect(stationOne).toContainText(/Ausschuss gesamt\s+0/);
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+0/);
    await stationOne.getByRole("button", { name: "In Abklärung +1" }).click();
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+1/);
    await stationOne.getByRole("button", { name: "In Abklärung -1" }).click();
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+0/);
    await stationOne.getByRole("button", { name: "In Abklärung -1" }).click();
    await expect(page.locator("#view")).toContainText("Abklärmenge kann nicht unter 0 fallen.");
    await expect(stationOne).toContainText(/In Abklärung gesamt\s+0/);
    await expect(page.locator("#view")).toContainText("Ausschuss Auftrag");
    await expect(page.locator("#view")).toContainText("In Abklärung Auftrag");
    await expect(page.locator("#view")).toContainText("Ursachen werden im nächsten Schritt ergänzt.");
    await page.locator("#view").getByRole("button", { name: /Spannung/ }).filter({ hasText: /hinzuf/ }).click();
    await expect(page.locator("#view")).toContainText("Spannung 2");
    await expect
      .poll(() => page.evaluate(() => window.__SUPABASE_WRITE_LOG || []))
      .not.toContainEqual(expect.objectContaining({ table: "production_counts" }));
    const stationCountWrites = await page.evaluate(() =>
      (window.__SUPABASE_WRITE_LOG || []).filter((entry) => entry.table === "production_station_counts"),
    );
    expect(stationCountWrites.every((entry) => ["insert", "update"].includes(entry.action))).toBeTruthy();
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
