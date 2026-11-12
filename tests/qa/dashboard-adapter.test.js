const fs = require("node:fs");
const vm = require("node:vm");
const assert = require("node:assert/strict");
const { test } = require("node:test");
const source = fs.readFileSync("schichtplan_mvp.js", "utf8");
const names = ["renderDashboard", "getManagedDepartmentIds", "getVisibleProductionMachines",
  "getVisibleProductionCounts", "taskAssignee", "normalizeNameForCompare",
  "currentUserNameForCompare", "isTaskAssignedToCurrentUser", "renderModuleSubTabs", "renderModulePlaceholder"];
// Execute the actual application adapter/selectors, with read-only synthetic inputs.
const functions = names.map(name => {
  const start = source.indexOf(`function ${name}(`);
  assert(start >= 0);
  return source.slice(start, source.indexOf("\nfunction ", start + 1));
}).join("\n");

function fixture(role) {
  let model;
  const context = {
    currentUser: { name: "Lavdrim", role }, currentEmployeeRecord: { id: "leader" },
    state: { dashboardSubTab: "overview", departments: [{ id: "own", leader_employee_id: "leader" }, { id: "other", leader_employee_id: "other" }],
      productionMachines: [{ name: "Own", department_id: "own" }, { name: "Other", department_id: "other" }],
      productionCounts: [{ department_id: "own" }, { department_id: "other" }],
      tasks: [{ title: "own", assignee: " Lavdrim ", status: "open" }, { title: "other", assignee: "Roger", status: "open" }, { title: "done", assignee: "Lavdrim", status: "done" }] },
    window: { HumbelHome: { render(value) { model = JSON.parse(JSON.stringify(value)); return "HOME"; } } },
    articleMaster: { hasAccess: () => true },
    generateThreeMonths: () => [{ assigned: "Lavdrim", label: "own" }, { assigned: "Roger", label: "other" }],
    DASHBOARD_SUBTABS: [{ id: "overview", label: "Übersicht" }, { id: "myShifts", label: "Meine Schichten" }, { id: "todo", label: "To-Do" }, { id: "oldSchedule", label: "Alt-Schichtplan" }],
    escapeHtml: value => String(value), renderMyShifts: () => "EXISTING_SHIFTS", renderTodo: () => "EXISTING_TODO",
  };
  vm.createContext(context);
  vm.runInContext(functions, context);
  return { context, model: () => model };
}

test("actual dashboard adapter preserves admin/department/employee selectors and source data", () => {
  for (const role of ["admin", "department_admin", "employee"]) {
    const { context, model } = fixture(role);
    const before = JSON.stringify(context.state);
    const tabs = role === "admin" ? ["dashboard", "produktion", "personalverwaltung"] : role === "department_admin" ? ["dashboard", "produktion"] : ["dashboard", "werkzeugverwaltung"];
    const rendered = context.renderDashboard(tabs);
    assert.deepEqual(model().machines.map(m => m.name), role === "admin" ? ["Own", "Other"] : role === "department_admin" ? ["Own"] : []);
    assert.equal(model().counts.length, role === "admin" ? 2 : role === "department_admin" ? 1 : 0);
    assert.deepEqual(model().tasks.map(t => t.title), role === "admin" ? ["own", "other"] : ["own"]);
    assert.deepEqual(model().shifts.map(s => s.label), role === "employee" ? ["own"] : []);
    assert.equal(model().articleEnabled, false);
    assert.equal(JSON.stringify(context.state), before);
    for (const subtab of context.DASHBOARD_SUBTABS) assert(rendered.includes(subtab.label));
  }
});

test("existing dashboard subtabs and fallback still use the original render paths", () => {
  const { context } = fixture("employee");
  context.state.dashboardSubTab = "myShifts";
  assert(context.renderDashboard([]).includes("EXISTING_SHIFTS"));
  context.state.dashboardSubTab = "todo";
  assert(context.renderDashboard([]).includes("EXISTING_TODO"));
  context.state.dashboardSubTab = "oldSchedule";
  assert(context.renderDashboard([]).includes("Alter Schichtplan bleibt Alt / deaktiviert"));
  context.state.dashboardSubTab = "overview";
  context.window.HumbelHome = undefined;
  assert(context.renderDashboard([]).includes("Start / Dashboard"));
});
