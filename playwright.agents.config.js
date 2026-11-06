const { defineConfig, devices } = require("@playwright/test");
module.exports = defineConfig({
  testDir: "./tests", testMatch: "seed.spec.ts", forbidOnly: true,
  use: { ...devices["Desktop Chrome"], serviceWorkers: "block", trace: "retain-on-failure" },
  reporter: [["list"]],
});
