const { defineConfig, devices } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests",
  testMatch: "smoke.spec.js",
  grep: /@visual/,
  forbidOnly: true,
  retries: 0,
  workers: 1,
  timeout: 30000,
  updateSnapshots: "none",
  snapshotPathTemplate: "{testDir}/visual-snapshots/{projectName}/{arg}{ext}",
  expect: { toHaveScreenshot: { animations: "disabled", caret: "hide", maxDiffPixels: 0, threshold: 0.2 } },
  reporter: [["html", { outputFolder: "visual-report", open: "never" }], ["list"], ["./scripts/qa-reporter.js"]],
  outputDir: "test-results/visual",
  use: { trace: "retain-on-failure", screenshot: "only-on-failure", serviceWorkers: "block",
    locale: "de-CH", timezoneId: "Europe/Berlin", reducedMotion: "reduce", colorScheme: "light" },
  projects: [
    { name: "visual-desktop", use: { ...devices["Desktop Chrome"] } },
    { name: "visual-mobile", use: { ...devices["Pixel 5"], viewport: { width: 390, height: 844 } } },
    { name: "visual-tablet", use: { ...devices["Desktop Chrome"], viewport: { width: 1024, height: 768 }, hasTouch: true } },
  ],
});
