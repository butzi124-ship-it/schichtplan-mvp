// Review copies only. Never writes into tests/visual-snapshots or changes tolerances.
const fs = require('node:fs'), path = require('node:path'), crypto = require('node:crypto');
const scenes = ['login', 'production-overview', 'touch-production', 'article-tabs', 'station-tabs', 'good-counter', 'scrap-clarification', '6m-dialog', 'mandatory-checklist', 'completion-ready', 'tool-management'];
const root = 'test-results/visual', out = 'artifacts/visual/reference-proposals';
const manifest = { approvalRequired: true, referencesChanged: false, testedSha: process.env.GITHUB_SHA || 'local', version: fs.readFileSync('schichtplan_mvp.js', 'utf8').match(/const APP_VERSION = "([^"]+)"/)[1], proposals: [], missing: [] };
fs.mkdirSync(out, { recursive: true });
for (const project of ['visual-desktop', 'visual-tablet', 'visual-mobile']) {
  for (const scene of scenes) {
    const directory = fs.existsSync(root) && fs.readdirSync(root).find(name => name.endsWith('-' + project) && fs.existsSync(path.join(root, name, scene + '-actual.png')));
    if (!directory) { manifest.missing.push(project + '/' + scene + '.png'); continue; }
    const source = path.join(root, directory, scene + '-actual.png');
    const target = path.join(out, project, scene + '.png');
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.copyFileSync(source, target);
    manifest.proposals.push({ file: project + '/' + scene + '.png', source, sha256: crypto.createHash('sha256').update(fs.readFileSync(target)).digest('hex') });
  }
}
fs.writeFileSync(path.join(out, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
fs.writeFileSync(path.join(out, 'README.txt'), 'UNAPPROVED PROPOSALS. Review these images and obtain separate owner approval before replacing any reference. This is not visual acceptance.\n');
console.log(`Review-only proposals: ${manifest.proposals.length}/33; missing: ${manifest.missing.length}; references untouched.`);
if (manifest.missing.length) process.exitCode = 1;
