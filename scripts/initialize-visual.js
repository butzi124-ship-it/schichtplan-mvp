const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
if (process.platform !== "linux" || process.env.HUMBEL_VISUAL_ENV !== "playwright-1.62.1-noble") throw new Error("Initial references require the pinned Linux CI container");
const root = "tests/visual-snapshots";
if (fs.existsSync(root) && fs.readdirSync(root, { recursive: true }).some(f => f.endsWith(".png"))) throw new Error("Refusing to overwrite existing references");
const cli = path.join(path.dirname(require.resolve("playwright/package.json")), "cli.js");
// Playwright 1.62 intentionally soft-fails missing references while writing the first images.
const initial = spawnSync(process.execPath, [cli, "test", "-c", "playwright.visual.config.js", "--update-snapshots=missing"], { stdio: "inherit" });
const report = JSON.parse(fs.readFileSync(path.join(process.env.QA_ARTIFACT_DIR, "summary.json"), "utf8"));
const expectedMissing = report.failures.length === 33 && report.failures.every(f => f.attempts.every(a =>
  a.errors.length === 1 && a.errors[0].startsWith("Error: A snapshot doesn't exist at ")));
if (initial.status !== 0 && !expectedMissing) throw new Error("Initial generation found real failures; inspect artifacts and fix them. No references may be approved.");
const compare = spawnSync(process.execPath, [cli, "test", "-c", "playwright.visual.config.js", "--update-snapshots=none"], { stdio: "inherit" });
if (compare.status !== 0) process.exit(compare.status ?? 1);
require("./baseline-metadata");
