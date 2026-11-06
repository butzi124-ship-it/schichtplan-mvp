const fs = require("fs");
function localTestUrl(value) {
  try {
    const url = new URL(value);
    return ["http:", "https:"].includes(url.protocol) && ["127.0.0.1", "localhost", "[::1]"].includes(url.hostname)
      && !url.username && !url.password && url.pathname === "/";
  } catch { return false; }
}
function prerequisites(env = process.env) {
  const missing = [];
  if (!localTestUrl(env.SUPABASE_TEST_URL)) missing.push("An explicit loopback-only SUPABASE_TEST_URL; all hosted production URLs are rejected");
  if (!fs.existsSync("supabase/config.toml")) missing.push("A reviewed local Supabase configuration and Docker runtime");
  if (!fs.existsSync("supabase/migrations")) missing.push("Authoritative, restorable schema, grants, functions, triggers and RLS migrations (catalog inventory alone is insufficient)");
  missing.push("Reviewed isolated auth users, fixtures and real RLS/session/realtime/concurrency integration tests");
  return missing;
}
if (require.main === module) {
  console.error(JSON.stringify({ status: "BLOCKED_NOT_TESTED", productionWrites: false, missing: prerequisites() }, null, 2));
  process.exitCode = 2;
}
module.exports = { localTestUrl, prerequisites };
