// Required, explicit extension points. Each workflow mode uses a fresh isolated stack.
const assert = require('node:assert/strict');
const modes = Object.freeze({
  diagnostic: null,
  hardening: './hardening-extension',
  'ba-op': './initial-responsibility-extension',
});
async function run(context, mode, load = require) {
  assert.ok(Object.hasOwn(modes, mode), 'Explicit native acceptance mode required');
  context.results.acceptanceMode = mode;
  context.results.baselineSecurityStatus = context.results.securityStatus;
  if (modes[mode]) await load(modes[mode]).run(context);
  if (mode === 'hardening') context.results.hardeningCandidateStatus = context.results.securityStatus;
  // The legacy findings remain open, even when the isolated candidate passes.
  context.results.securityStatus = context.results.baselineSecurityStatus;
}
module.exports = {run, modes};
