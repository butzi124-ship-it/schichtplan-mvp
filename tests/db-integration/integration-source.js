// Assembly/provenance only. All SQL, Edge and application sources come from the
// integration checkout. The pinned #15 checkout supplies only the native harness.
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const assert = require('node:assert/strict');
const {execFileSync} = require('node:child_process');
const roots = ['supabase', 'js', 'tests/db-integration'];
const fixture = 'tests/fixtures/forecast-clock.json';
const manifestPath = 'artifacts/db-integration/integration-source.json';
const hash = value => crypto.createHash('sha256').update(value).digest('hex');
function inventory(root) {
  const files = {};
  function visit(relative) {
    const full = path.join(root, relative), stat = fs.lstatSync(full);
    assert.ok(!stat.isSymbolicLink(), 'Symlink forbidden in acceptance source: ' + relative);
    if (stat.isDirectory()) for (const name of fs.readdirSync(full).sort()) visit(relative + '/' + name);
    else { assert.ok(stat.isFile()); files[relative] = hash(fs.readFileSync(full)); }
  }
  for (const relative of [...roots, fixture]) visit(relative);
  return files;
}
function verify(root = process.cwd()) {
  const manifest = JSON.parse(fs.readFileSync(path.join(root, manifestPath), 'utf8'));
  assert.match(manifest.commit, /^[a-f0-9]{40}$/);
  assert.deepEqual(inventory(root), manifest.files, 'Integration source mismatch: replacement/addition/deletion');
  return manifest;
}
function assemble(source, target) {
  assert.notEqual(fs.realpathSync(source), fs.realpathSync(target));
  // CI must assemble a clean, concrete checkout (PR merge SHA or push SHA).
  const commit = execFileSync('git', ['rev-parse', 'HEAD'], {cwd:source, encoding:'utf8'}).trim();
  execFileSync('git', ['diff', '--exit-code', 'HEAD', '--', ...roots, fixture], {cwd:source});
  assert.equal(execFileSync('git', ['ls-files', '--others', '--', ...roots, fixture], {cwd:source,encoding:'utf8'}), '', 'Untracked acceptance inputs (including ignored files)');
  const files = inventory(source);
  for (const relative of roots) {
    fs.rmSync(path.join(target, relative), {recursive:true,force:true});
    fs.cpSync(path.join(source, relative), path.join(target, relative), {recursive:true});
  }
  fs.mkdirSync(path.join(target, 'tests/fixtures'), {recursive:true});
  fs.copyFileSync(path.join(source, fixture), path.join(target, fixture));
  fs.mkdirSync(path.join(target, 'artifacts/db-integration'), {recursive:true});
  fs.writeFileSync(path.join(target, manifestPath), JSON.stringify({commit, files, harness:'25e9a6914e38fc791d9ecad13d92a45e30d19bb3', sqlClock:'anchored-elapsed; raw SQL hashes retained'}, null, 2));
  verify(target);
  console.log('Integration input manifest verified for checkout ' + commit);
}
// Exact adapter for the original optional hook. All baseline assertions survive.
const hook = " if(fs.existsSync('tests/db-integration/security-extension.js'))await require('../../tests/db-integration/security-extension').run({db,s,check,results});";
const dispatch = " await require('../../tests/db-integration/integration-dispatch').run({db,s,check,results},process.argv[2]);";
function prepareRunner(root) {
  const file = path.join(root, 'scripts/db-integration/run.js');
  let source = fs.readFileSync(file, 'utf8');
  // This is #15's exact source after the existing, tested clock-only adapter.
  const original = source.replace(" results.testClock=await require('../../tests/db-integration/forecast-clock').install(db,config());\n let expected=31;", ' let expected=31;');
  assert.equal(hash(original), '0015fd8cedad4437ff614aa3ec2abba861f6eed6549bbaa512875535dd65d052', 'Unknown native runner');
  assert.equal(source.split(hook).length, 2);
  const before = 'async function main(){const db=await connect();try{';
  const after = "async function main(){results.integrationSource=require('../../tests/db-integration/integration-source').verify();assert.ok(Object.hasOwn(require('../../tests/db-integration/integration-dispatch').modes,process.argv[2]),'Explicit native acceptance mode required');const db=await connect();try{";
  assert.equal(source.split(before).length, 2);
  source = source.replace(hook, dispatch).replace(before, after);
  // Prove the adapter can be reversed byte-for-byte; no assertions removed.
  assert.equal(source.replace(dispatch, hook).replace(after, before), fs.readFileSync(file, 'utf8'));
  fs.writeFileSync(file, source);
}
module.exports = {assemble, inventory, verify, prepareRunner, hook, dispatch};
if (require.main === module) {
  if (process.argv[2] === 'assemble') assemble(path.resolve(process.argv[3]), path.resolve(process.argv[4]));
  else { assert.equal(process.argv[2], 'prepare-runner'); prepareRunner(process.cwd()); }
}
