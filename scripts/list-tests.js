const path = require("path");
const { spawnSync } = require("child_process");
function listTests(config = "playwright.config.js") {
  const cli = path.join(path.dirname(require.resolve("playwright/package.json")), "cli.js");
  const r = spawnSync(process.execPath, [cli, "test", "-c", config, "--list", "--reporter=json"], { encoding: "utf8", maxBuffer: 8 * 1024 * 1024 });
  if (r.status !== 0) throw new Error(r.stderr || r.stdout);
  const results = [];
  function walk(suites, parents = []) {
    for (const suite of suites) {
      for (const spec of suite.specs || []) for (const run of spec.tests) results.push({
        project: run.projectName, title: [...parents, suite.title, spec.title].join(" > "),
      });
      walk(suite.suites || [], [...parents, suite.title]);
    }
  }
  walk(JSON.parse(r.stdout).suites);
  return results;
}
module.exports = { listTests };
if (require.main === module) process.stdout.write(JSON.stringify(listTests()));
