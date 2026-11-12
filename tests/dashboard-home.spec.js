const { test, expect } = require("@playwright/test");
const { prepareMockPage, appUrl } = require("./helpers/mock-page");

test.describe("Dashboard home", () => {
  test.beforeEach(async ({ page }) => {
    await prepareMockPage(page);
    await page.goto(appUrl);
    await page.getByRole("button", { name: "Mido/Admin", exact: true }).click();
    await expect(page.locator(".h-home-hero")).toBeVisible();
  });

  test("responsive approved overview and existing dashboard navigation remain reachable", async ({ page }, info) => {
    await expect(page.locator(".h-home-tile")).toHaveCount(5);
    await expect(page.locator(".h-home-panel")).toHaveCount(6);
    await expect(page.locator(".h-home-machine").first()).toContainText("Maschine 50");
    await expect(page.locator(".h-home-machine small").first()).toHaveText("Status nicht verfügbar");
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
    const tile = await page.locator("button.h-home-tile").first().boundingBox();
    expect(tile.height).toBeGreaterThanOrEqual(44);
    await page.screenshot({ path: info.outputPath("dashboard.png"), fullPage: true });
    for (const name of ["Meine Schichten", "To-Do", "Alt-Schichtplan", "Übersicht"]) {
      await page.locator("#view").getByRole("button", { name, exact: true }).click();
      await expect(page.locator("#view").getByRole("button", { name, exact: true })).toBeVisible();
    }
    await expect(page.locator(".h-home-hero")).toBeVisible();
    await page.getByRole("button", { name: "Fertigung Abteilungen · Maschinen · Aufträge" }).click();
    await expect(page.locator("body")).not.toHaveClass(/humbel-home-mode/);
    await expect(page.locator("#tabs").getByRole("button", { name: "Produktion", exact: true })).toBeVisible();
    await page.locator("#tabs").getByRole("button", { name: "Start / Dashboard", exact: true }).click();
    await expect(page.locator(".h-home-hero")).toBeVisible();
    expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
  });

  test("employee sees own tasks and existing tools and shifts without production data", async ({ page }) => {
    await page.evaluate(() => {
      state.tasks = [
        { title: "Eigene offene Aufgabe", assignedTo: "Lavdrim", status: "open" },
        { title: "Fremde Aufgabe", assignedTo: "Roger", status: "open" },
        { title: "Erledigte Aufgabe", assignedTo: "Lavdrim", status: "done" },
      ];
      loginAs("Lavdrim");
    });
    await expect(page.locator(".h-home").first()).toContainText("Eigene offene Aufgabe");
    await expect(page.locator(".h-home").first()).not.toContainText("Fremde Aufgabe");
    await expect(page.locator(".h-home").first()).not.toContainText("Erledigte Aufgabe");
    await expect(page.locator(".h-home-machine")).toHaveCount(0);
    await expect(page.locator("#tabs").getByRole("button", { name: /Produktion|Personalverwaltung|Admin \/ System/ })).toHaveCount(0);
    await page.getByRole("button", { name: "Schichtplan Schichten · Planung", exact: true }).click();
    await expect(page.locator("#view")).toContainText("Meine Schichten");
    await page.locator("#view").getByRole("button", { name: "Übersicht", exact: true }).click();
    await page.getByRole("button", { name: "Werkzeuge Bestand · Journal", exact: true }).click();
    await expect(page.locator("body")).not.toHaveClass(/humbel-home-mode/);
    await expect(page.locator("#view")).toContainText("Werkzeugverwaltung");
    expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
  });

  test("department admin retains own department filters in read-only summaries", async ({ page }) => {
    await page.evaluate(() => {
      currentUser = { name: "Leitung", role: "department_admin" };
      currentEmployeeRecord = { id: "home-leader" };
      state.departments = [{ id: "own", leader_employee_id: "home-leader" }, { id: "other", leader_employee_id: "other-leader" }];
      state.productionMachines = [{ id: "own-machine", name: "Eigene Maschine", department_id: "own" }, { id: "other-machine", name: "Fremde Maschine", department_id: "other" }];
      state.productionCounts = [{ department_id: "own" }, { department_id: "other" }];
      render();
    });
    await expect(page.locator(".h-home-machine")).toHaveCount(1);
    await expect(page.locator(".h-home-machine")).toContainText("Eigene Maschine");
    await expect(page.locator(".h-home").first()).not.toContainText("Fremde Maschine");
    await expect(page.locator(".h-home-panel").first()).toContainText("Sichtbare Stückzahleinträge1");
    await expect(page.locator("#tabs").getByRole("button", { name: /Personalverwaltung|Admin \/ System/ })).toHaveCount(0);
    expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
  });

  test("scanner tab resolution removes dashboard mode and cannot open production", async ({ page }) => {
    await page.evaluate(() => {
      currentUser = { name: "Scanner", role: "tool_scanner" };
      currentEmployeeRecord = null;
      setTab("dashboard");
    });
    await expect(page.locator(".h-home-hero")).toHaveCount(0);
    await expect(page.locator("body")).not.toHaveClass(/humbel-home-mode/);
    await expect(page.locator("#tabs").getByRole("button")).toHaveCount(2);
    await page.evaluate(() => setTab("produktion"));
    await expect(page.locator("#tabs").getByRole("button", { name: "Produktion", exact: true })).toHaveCount(0);
    expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
  });

  test("missing presentation module preserves legacy dashboard and all subtabs", async ({ page }) => {
    await page.evaluate(() => { window.HumbelHome = undefined; render(); });
    await expect(page.locator("#view").getByRole("heading", { name: "Start / Dashboard", exact: true })).toBeVisible();
    for (const name of ["Übersicht", "Meine Schichten", "To-Do", "Alt-Schichtplan"]) {
      await expect(page.locator("#view").getByRole("button", { name, exact: true })).toBeVisible();
    }
    expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
  });
});
