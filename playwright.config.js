const { defineConfig, devices } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests",
  testMatch: "smoke.spec.js",
  grepInvert: /@visual/,
  forbidOnly: !!process.env.CI,
  failOnFlakyTests: !!process.env.CI,
  timeout: 30 * 1000,
  retries: process.env.CI ? 1 : 0,
  expect: {
    timeout: 5000,
  },
  reporter: [["html", { outputFolder: process.env.QA_REPORT_DIR || "playwright-report", open: "never" }], ["list"], ["./scripts/qa-reporter.js"], ["json", { outputFile: "test-results/results.json" }]],
  outputDir: process.env.QA_OUTPUT_DIR || "test-results/functional",
  use: {
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
    serviceWorkers: "block",
    locale: "de-CH",
    timezoneId: "Europe/Berlin",
  },
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] },
    },
    {
      name: "mobile-chromium",
      grep: /Production counter feedback|Production order navigation|Production touch view|Sequential production stations/,
      use: { ...devices["Pixel 5"], viewport: { width: 390, height: 844 } },
    },
    {
      name: "tablet-chromium",
      grep: /Production order navigation|Production touch view|Sequential production stations/,
      use: { ...devices["Desktop Chrome"], viewport: { width: 1024, height: 768 }, hasTouch: true },
    },
    {
      name: "firefox",
      grep: /Production order navigation|Production touch view|Sequential production stations/,
      use: { ...devices["Desktop Firefox"] },
    },
    {
      name: "webkit",
      grep: /Production order navigation|Production touch view|Sequential production stations/,
      use: { ...devices["Desktop Safari"] },
    },
  ],
});
