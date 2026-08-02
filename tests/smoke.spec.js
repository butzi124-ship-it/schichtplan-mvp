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
  window.TEST_ADMIN_EMAIL = "admin@example.test";
  window.TEST_EMPLOYEE_EMAIL = "employee@example.test";

  class QueryMock {
    constructor(table) {
      this.table = table;
      this.filters = [];
      this.useSingle = false;
      this.useMaybeSingle = false;
    }

    select() { return this; }
    order() { return this; }
    limit() { return this; }
    insert() { return Promise.resolve({ data: [], error: null }); }
    update() { return this; }
    delete() { return this; }
    upsert() { return this; }
    single() { this.useSingle = true; return this; }
    maybeSingle() { this.useMaybeSingle = true; return this; }
    eq(column, value) { this.filters.push({ column, value }); return this; }

    result() {
      let data = [...(rowsByTable[this.table] || [])];
      for (const { column, value } of this.filters) {
        data = data.filter((row) => String(row[column]) === String(value));
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
