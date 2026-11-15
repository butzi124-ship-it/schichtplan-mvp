const fs = require("node:fs");
const assert = require("node:assert/strict");
const { test } = require("node:test");
const { appBusiness, hash } = require("../../scripts/qa-policy");
const { reverseDashboardPresentation, verifyDashboardFiles, reverseDashboardFile } = require("../../scripts/dashboard-presentation-policy");
const reviewed = require("./dashboard-presentation-patches.json");
const app = fs.readFileSync("schichtplan_mvp.js", "utf8");

test("dashboard exceptions are four exact reversible presentation hooks", () => {
  assert.equal(reviewed.patches.length, 4);
  const restored = reverseDashboardPresentation(app);
  for (const { before, after } of reviewed.patches) {
    assert.equal(app.split(after).length - 1, 1);
    assert(restored.includes(before));
    assert.throws(() => reverseDashboardPresentation(app.replace(after, after + " /* mutation */" + after)));
    assert.throws(() => reverseDashboardPresentation(app.replace(after, "")));
    assert.throws(() => reverseDashboardPresentation(app.replace(after, after.replace(/\S/, "!"))));
  }
});

test("dashboard normalization preserves unauthorized business edits for the original hash guard", () => {
  const normalized = appBusiness(reverseDashboardPresentation(app));
  const changed = app.replace('const INVENTORY_MODE_ENABLED = false;', 'const INVENTORY_MODE_ENABLED = true;');
  assert.notEqual(changed, app);
  assert.notEqual(hash(appBusiness(reverseDashboardPresentation(changed))), hash(normalized));
  assert.equal(appBusiness(reverseDashboardPresentation(app.replace('"0.15.4"', '"0.15.5"'))), normalized);
});

test("reviewed dashboard assets reject changed or missing code, CSS and HTML", () => {
  verifyDashboardFiles(file => fs.readFileSync(file));
  assert.deepEqual(Object.keys(reviewed.files).sort(), ["css/home-dashboard.css", "js/dashboard-home.js", "schichtplan_mvp.html"]);
  for (const changed of Object.keys(reviewed.files)) {
    assert.throws(() => verifyDashboardFiles(file => file === changed ? Buffer.from("mutation") : fs.readFileSync(file)));
    assert.throws(() => verifyDashboardFiles(file => {
      if (file === changed) throw new Error("missing reviewed source");
      return fs.readFileSync(file);
    }));
  }
});

test("both HTML asset insertions are exact and unique; unrelated HTML is preserved", () => {
  const html = fs.readFileSync("schichtplan_mvp.html", "utf8");
  const restored = reverseDashboardFile("schichtplan_mvp.html", html);
  for (const { before, after } of reviewed.filePatches) {
    assert(restored.includes(before));
    assert.throws(() => reverseDashboardFile("schichtplan_mvp.html", html.replace(after, "")));
    assert.throws(() => reverseDashboardFile("schichtplan_mvp.html", html.replace(after, after + after)));
  }
  assert.notEqual(reverseDashboardFile("schichtplan_mvp.html", html + "unreviewed"), restored);
});

test("dashboard renderer escapes real data and never fabricates machine states or progress", () => {
  const vm = require("node:vm");
  const context = { window: {} };
  vm.runInNewContext(fs.readFileSync("js/dashboard-home.js", "utf8"), context);
  const model = { tabs: ["dashboard"], machines: [{ name: '<img src=x onerror="bad()">' }],
    tasks: [{ title: "<script>bad()</script>", status: "<open>" }],
    shifts: [{ label: "<early>", date: "2026-11-12", start: "06:00", end: "14:00" }] };
  const html = context.window.HumbelHome.render(model);
  assert(!html.includes("<img"));
  assert(!html.includes("<script>"));
  assert(html.includes("&lt;early&gt;"));
  assert(html.includes("Status nicht verfügbar"));
  assert(html.includes("Kein verlässlich berechneter Tagesfortschritt verfügbar."));
  assert(!html.includes("onclick="));
  assert(!context.window.HumbelHome.render({ tabs: [] }).includes("h-home-machine\""));
});
