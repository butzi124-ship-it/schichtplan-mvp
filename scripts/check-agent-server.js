const { spawn } = require("child_process");
const path = require("path");
const cli = path.join(path.dirname(require.resolve("playwright/package.json")), "cli.js");
const child = spawn(process.execPath, [cli, "run-test-mcp-server", "--headless", "--config", "playwright.agents.config.js"], { stdio: ["pipe", "pipe", "inherit"] });
let buffer = "";
let finished = false;
const timer = setTimeout(() => finish(new Error("MCP handshake timed out")), 25000);
function finish(error) {
  if (finished) return;
  finished = true;
  clearTimeout(timer);
  child.kill();
  if (error) { console.error(error.message); process.exitCode = 1; }
  else console.log("Official Playwright MCP: planner/generator/healer tools available; no app or DB writes executed.");
}
function send(message) { child.stdin.write(JSON.stringify(message) + "\n"); }
child.on("error", finish);
child.stdout.on("data", chunk => {
  buffer += chunk.toString();
  let end;
  while ((end = buffer.indexOf("\n")) >= 0) {
    const line = buffer.slice(0, end); buffer = buffer.slice(end + 1);
    if (!line.trim()) continue;
    try {
      const message = JSON.parse(line);
      if (message.error) return finish(new Error(JSON.stringify(message.error)));
      if (message.id === 1) {
        send({ jsonrpc: "2.0", method: "notifications/initialized" });
        send({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} });
      }
      if (message.id === 2) {
        const names = message.result.tools.map(tool => tool.name);
        const missing = ["planner_setup_page", "generator_setup_page", "test_run", "test_debug"].filter(name => !names.includes(name));
        finish(missing.length ? new Error(`Missing tools: ${missing.join(", ")}`) : null);
      }
    } catch (error) { finish(error); }
  }
});
send({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "humbel-qa-check", version: "0.5.17" } } });
