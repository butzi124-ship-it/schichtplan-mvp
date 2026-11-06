const crypto = require("crypto");
function hash(value) { return crypto.createHash("sha256").update(value).digest("hex"); }
function appBusiness(source) {
  return source.replace(/\r\n/g, "\n").replace(/const APP_VERSION = "[^"]+";/, "const APP_VERSION = VERSION;")
    .replace(/const VERSION_LOG = \[[\s\S]*?\n\];/, "const VERSION_LOG = LOG;");
}
function sourceViolations(source) {
  return [/\b(?:test|describe)\.(?:skip|fixme|only|fail)\s*\(/g, /force\s*:\s*true/g]
    .flatMap(pattern => [...source.matchAll(pattern)].map(m => m[0]));
}
module.exports = { hash, appBusiness, sourceViolations };
