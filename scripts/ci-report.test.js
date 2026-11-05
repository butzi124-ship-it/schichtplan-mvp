const { test } = require('node:test');
const assert = require('node:assert/strict');
const { renderReport } = require('./ci-report');
const source = 'const APP_VERSION = "0.5.17"; const VERSION_LOG = [{version: "0.5.17", date: "2026-11-05 06:34"';
test('reports nested device results, retries, skips and tested commit without hiding failures', () => {
  const report = { suites: [{ suites: [{ specs: [{ tests: [
    { projectName: 'chromium', status: 'expected' },
    { projectName: 'chromium', status: 'unexpected' },
    { projectName: 'mobile-chromium', status: 'flaky' },
    { projectName: 'tablet-chromium', status: 'skipped' },
  ] }] }] }] };
  const output = renderReport(report, source, { TESTED_SHA: 'merge-sha', PR_HEAD_SHA: 'head-sha' });
  assert.match(output, /chromium \| 1 \| 1 \| 0 \| 0/);
  assert.match(output, /mobile-chromium \| 0 \| 0 \| 1 \| 0/);
  assert.match(output, /tablet-chromium \| 0 \| 0 \| 0 \| 1/);
  assert.match(output, /merge-sha/);
  assert.match(output, /head-sha/);
  assert.match(output, /0\.5\.17 \/ 2026-11-05 06:34 Europe\/Berlin/);
});
test('missing report is explicitly unavailable, never reported as passed', () => {
  assert.match(renderReport(null, source, {}), /Keine Browser-Ergebnisse verfügbar/);
});
