const { defineConfig, devices } = require("@playwright/test");

module.exports = defineConfig({
  testDir: "./tests",
  timeout: 30 * 1000,
  retries: process.env.CI ? 1 : 0,
  expect: {
    timeout: 5000,
  },
  reporter: [["html", { outputFolder: "playwright-report", open: "never" }], ["list"]],
  use: {
    trace: "retain-on-failure",
    screenshot: "only-on-failure",
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
  ],
});
