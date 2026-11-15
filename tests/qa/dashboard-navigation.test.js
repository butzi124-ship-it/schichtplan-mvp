const fs = require("node:fs");
const vm = require("node:vm");
const assert = require("node:assert/strict");
const { test } = require("node:test");
const { verifyDashboardFiles } = require("../../scripts/dashboard-presentation-policy");
const reviewed = require("./dashboard-presentation-patches.json");

test("dashboard retains its visible navigation title and unambiguous personnel shortcut", () => {
  const context = { window: {} };
  vm.runInNewContext(fs.readFileSync("js/dashboard-home.js", "utf8"), context);
  const html = context.window.HumbelHome.render({ tabs: ["dashboard", "personalverwaltung"] });
  assert(html.includes("<h1>Start / Dashboard</h1>"));
  const buttons = html.match(/<button\b[^>]*>[\s\S]*?<\/button>/g) || [];
  const personnel = buttons.filter(button => button.includes("setTab('personalverwaltung')"));
  assert.equal(personnel.length, 1);
  assert(personnel[0].includes("<strong>Mitarbeiter</strong>"));
  assert(personnel[0].includes("<small>Stammdaten · Team</small>"));
  assert(buttons.every(button => !button.includes("Personalverwaltung")));
});

test("both reviewed text corrections retain the old renderer hash and reject missing or changed text", () => {
  assert.equal(reviewed.assetPatches.length, 2);
  assert.equal(reviewed.files["js/dashboard-home.js"], "48cc78191e0866a279826f168bebe26b54081b8f681d5a2f372a2136d9ba93e2");
  verifyDashboardFiles(file => fs.readFileSync(file));
  for (const change of reviewed.assetPatches) {
    for (const replacement of [change.before, change.after + change.after, "unreviewed text"]) {
      assert.throws(() => verifyDashboardFiles(file => {
        const content = fs.readFileSync(file);
        return file === change.file ? content.toString("utf8").replace(change.after, replacement) : content;
      }));
    }
  }
});
