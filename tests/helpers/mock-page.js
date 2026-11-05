const fs = require("fs");
const path = require("path");
const { pathToFileURL } = require("url");
const { rowsByTable, installSupabaseMock } = require("./supabase-mock");
const appUrl = pathToFileURL(path.resolve(__dirname, "../../schichtplan_mvp.html")).href;
async function prepareMockPage(page) {
  await page.route("**/*", route => new URL(route.request().url()).protocol === "file:" ? route.continue() : route.abort());
  await page.route("**/cdn.tailwindcss.com**", route => route.fulfill({ contentType: "application/javascript",
    body: `const s=document.createElement('style');s.textContent=${JSON.stringify(fs.readFileSync(path.join(__dirname, "../fixtures/app-tailwind.css"), "utf8"))};document.head.appendChild(s);` }));
  await page.route("**/@supabase/supabase-js@2**", route => route.fulfill({ contentType: "application/javascript", body: "" }));
  await page.addInitScript(installSupabaseMock, JSON.parse(JSON.stringify(rowsByTable)));
}
module.exports = { prepareMockPage, appUrl };
