const fs = require('node:fs');

function renderReport(report, source, env = process.env) {
  const version = source.match(/const APP_VERSION = "([^"]+)"/)?.[1] || 'unbekannt';
  const log = source.match(/const VERSION_LOG = \[\s*\{\s*version: "([^"]+)",\s*date: "([^"]+)"/);
  const rows = new Map();
  function visit(suite) {
    for (const spec of suite.specs || []) {
      for (const test of spec.tests || []) {
        const row = rows.get(test.projectName) || { expected: 0, unexpected: 0, flaky: 0, skipped: 0 };
        row[test.status] = (row[test.status] || 0) + 1;
        rows.set(test.projectName, row);
      }
    }
    for (const child of suite.suites || []) visit(child);
  }
  for (const suite of report?.suites || []) visit(suite);
  return [
    '## CI-Prüfbericht',
    `- APP_VERSION: ${version}`,
    `- VERSION_LOG: ${log ? `${log[1]} / ${log[2]} Europe/Berlin` : 'nicht gefunden'}`,
    `- Commit (getesteter Checkout): ${env.TESTED_SHA || 'lokal'}`,
    `- PR-Head: ${env.PR_HEAD_SHA || 'kein PR'}`,
    `- CI: ${env.GITHUB_SERVER_URL || 'https://github.com'}/${env.GITHUB_REPOSITORY || ''}/actions/runs/${env.GITHUB_RUN_ID || ''}`,
    '',
    '| Projekt / Gerät | Bestanden | Fehlgeschlagen | Instabil | Übersprungen |',
    '| --- | ---: | ---: | ---: | ---: |',
    ...[...rows].map(([name, r]) => `| ${name} | ${r.expected} | ${r.unexpected} | ${r.flaky} | ${r.skipped} |`),
    ...(rows.size ? [] : ['| Keine Browser-Ergebnisse verfügbar | – | – | – | – |']),
    '',
    'Änderungen und Risiken: siehe PR-Beschreibung. Browser-Tests verwenden Testdaten und Supabase-Mocks; echte SQL/RLS-/Parallelzugriffe sind damit nicht nachgewiesen.',
    'Dieser Bericht erteilt keine Merge- oder Deployment-Freigabe. Maßgeblich ist der Statuscheck e2e samt Review und verifiziertem Branch-Schutz.',
    '',
  ].join('\n');
}

if (require.main === module) {
  const path = 'test-results/results.json';
  const report = fs.existsSync(path) ? JSON.parse(fs.readFileSync(path, 'utf8')) : null;
  const output = renderReport(report, fs.readFileSync('schichtplan_mvp.js', 'utf8'));
  fs.mkdirSync('test-results', { recursive: true });
  fs.writeFileSync('test-results/ci-summary.md', output);
  if (process.env.GITHUB_STEP_SUMMARY) fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, output);
  else process.stdout.write(output);
}
module.exports = { renderReport };
