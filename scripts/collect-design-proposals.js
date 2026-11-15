// Review copies only. Never writes into tests/visual-snapshots or changes tolerances.
const fs = require('node:fs'), path = require('node:path'), crypto = require('node:crypto');
const scenes = ['login', 'production-overview', 'touch-production', 'article-tabs', 'station-tabs', 'good-counter', 'scrap-clarification', '6m-dialog', 'mandatory-checklist', 'completion-ready', 'tool-management'];
const projects = ['visual-desktop', 'visual-tablet', 'visual-mobile'];
const root = 'test-results/visual', out = 'artifacts/visual/reference-proposals';
const run = process.env.DESIGN_CAPTURE_RUN || 'single';
if (!['first', 'second', 'single'].includes(run)) throw Error('Unknown capture run');
const runOut = run === 'single' ? out : path.join(out, run);
const digest = value => crypto.createHash('sha256').update(value).digest('hex');
const manifest = { approvalRequired: true, referencesChanged: false, testedSha: process.env.GITHUB_SHA || 'local', version: fs.readFileSync('schichtplan_mvp.js', 'utf8').match(/const APP_VERSION = "([^"]+)"/)[1], run, proposals: [], missing: [] };
fs.mkdirSync(runOut, { recursive: true });
for (const project of projects) {
  for (const scene of scenes) {
    const key = project + '/' + scene + '.png';
    const directories = fs.existsSync(root) ? fs.readdirSync(root).filter(name => name.endsWith('-' + project) && fs.existsSync(path.join(root, name, scene + '-actual.png'))) : [];
    // Ambiguous or absent artifacts cannot be silently accepted as a current capture.
    if (directories.length !== 1) { manifest.missing.push(key); continue; }
    const source = path.join(root, directories[0], scene + '-actual.png');
    const target = path.join(runOut, project, scene + '.png');
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.copyFileSync(source, target);
    const bytes = fs.readFileSync(target);
    const png = require('pngjs').PNG.sync.read(bytes);
    const proposal = { file: key, source, sha256: digest(bytes), width: png.width, height: png.height, pixelsSha256: digest(png.data) };
    const reference = path.join('tests/visual-snapshots', key);
    const diff = path.join(root, directories[0], scene + '-diff.png');
    for (const [kind, from] of [['reference', reference], ['diff', diff]]) {
      if (!fs.existsSync(from)) { manifest.missing.push(project + '/' + scene + '-' + kind + '.png'); continue; }
      const to = path.join(runOut, project, scene + '-' + kind + '.png');
      fs.copyFileSync(from, to);
      proposal[kind] = path.relative(runOut, to);
      proposal[kind + 'Sha256'] = digest(fs.readFileSync(to));
    }
    manifest.proposals.push(proposal);
  }
}
fs.writeFileSync(path.join(runOut, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
fs.writeFileSync(path.join(runOut, 'README.txt'), 'UNAPPROVED PROPOSALS. Review these images and obtain separate owner approval before replacing any reference. This is not visual acceptance.\n');
console.log(`Review-only proposals (${run}): ${manifest.proposals.length}/33; missing: ${manifest.missing.length}; references untouched.`);
if (manifest.missing.length) process.exitCode = 1;
if (run === 'second') {
  const firstPath = path.join(out, 'first/manifest.json');
  const first = fs.existsSync(firstPath) ? JSON.parse(fs.readFileSync(firstPath, 'utf8')) : null;
  const comparisons = manifest.proposals.map(current => {
    const previous = first?.proposals.find(p => p.file === current.file);
    return { file: current.file, identicalPixels: !!previous && previous.width === current.width && previous.height === current.height && previous.pixelsSha256 === current.pixelsSha256,
      identicalBytes: previous?.sha256 === current.sha256, referenceUnchanged: previous?.referenceSha256 === current.referenceSha256 };
  });
  const stable = !!first && first.testedSha === manifest.testedSha && first.version === manifest.version && first.missing.length === 0 && manifest.missing.length === 0 && first.proposals.length === 33 && comparisons.length === 33 && comparisons.every(c => c.identicalPixels && c.referenceUnchanged);
  const stability = { testedSha: manifest.testedSha, version: manifest.version, image: 'mcr.microsoft.com/playwright:v1.62.1-noble', approvalRequired: true, referencesChanged: false, status: stable ? 'STABLE_33_OF_33' : 'UNSTABLE_OR_INCOMPLETE', comparisons };
  fs.writeFileSync(path.join(out, 'stability.json'), JSON.stringify(stability, null, 2) + '\n');
  fs.writeFileSync(path.join(out, 'manifest.json'), JSON.stringify({ ...manifest, captureDirectory: 'second', stability: stability.status }, null, 2) + '\n');
  const sections = manifest.proposals.map(p => {
    const scene = p.file.replace(/\.png$/, '');
    const compare = comparisons.find(c => c.file === p.file);
    const images = [['Bisherige Referenz', 'second/' + p.reference], ['Aktuell – Lauf 1', 'first/' + p.file], ['Aktuell – Lauf 2', 'second/' + p.file], ['Differenz zur Referenz', 'second/' + p.diff]];
    return `<section><h2>${scene}</h2><p>Pixel zwischen Lauf 1 und 2: ${compare.identicalPixels ? 'identisch' : 'ABWEICHEND / FEHLEND'}</p><div class="images">${images.map(([label,src]) => `<figure><figcaption>${label}</figcaption><a href="${src}"><img loading="lazy" src="${src}" alt="${scene}: ${label}"></a></figure>`).join('')}</div></section>`;
  });
  fs.writeFileSync(path.join(out, 'index.html'), `<!doctype html><html lang="de"><meta charset="utf-8"><title>FABOS – nicht freigegebene Bildprüfung</title><style>body{font:16px system-ui;margin:24px;background:#f3f7fc;color:#102033}section{background:white;padding:16px;margin:24px 0;border:1px solid #dce5f1}.images{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}figure{margin:0}img{max-width:100%;height:auto;border:1px solid #aaa}figcaption{font-weight:bold;margin:8px 0}@media(max-width:900px){.images{grid-template-columns:repeat(2,minmax(0,1fr))}}</style><h1>Bildprüfung – keine Referenzfreigabe</h1><p>Version ${manifest.version}; Checkout ${manifest.testedSha}; ${stability.status}.</p><p>33 Szenenbilder pro Lauf. Alte Referenzen sind unverändert. Einzelbilder zum Vergrößern anklicken. Die Bildabweichungen sind nicht automatisch genehmigt.</p>${sections.join('\n')}</html>\n`);
  console.log(`Capture stability: ${stability.status}; ${comparisons.filter(c => c.identicalPixels).length}/33 identical pixel images.`);
  if (!stable) process.exitCode = 1;
}
