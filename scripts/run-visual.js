const { spawnSync } = require("child_process");
const path = require("path");
if (process.platform !== "linux" || process.env.HUMBEL_VISUAL_ENV !== "playwright-1.62.1-noble") {
  console.error("BLOCKED: visual comparisons require the pinned mcr.microsoft.com/playwright:v1.62.1-noble CI container. Windows references are not accepted.");
  process.exit(2);
}
if (process.argv.length > 2) {
  console.error("Snapshot update/override arguments are forbidden in the normal visual check.");
  process.exit(2);
}
const cli = path.join(path.dirname(require.resolve("playwright/package.json")), "cli.js");
const r = spawnSync(process.execPath, [cli, "test", "-c", "playwright.visual.config.js", "--update-snapshots=none"], { stdio: "inherit" });
process.exit(r.status ?? 1);
