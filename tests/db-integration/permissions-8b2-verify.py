"""Bounded local QA for 8B2. No network/DB connection setup, rerun or dispatch."""
import datetime, json, os, pathlib, signal, subprocess, time, shutil, sys
ROOT=pathlib.Path(__file__).resolve().parents[2]
OUT=pathlib.Path('/tmp/humbel-8b2-verification');OUT.mkdir(exist_ok=True)
started=time.monotonic();OVERALL=1800;results=[];ACTIVE=None
def save():
 (OUT/'results.json').write_text(json.dumps(dict(startedAt=start_iso,overallLimitSeconds=OVERALL,checks=results),indent=2))
def interrupted(signum,frame):
 child=ACTIVE
 if child is not None and child.poll() is None:
  try:os.killpg(child.pid,signal.SIGTERM)
  except ProcessLookupError:pass
  try:child.wait(timeout=5)
  except subprocess.TimeoutExpired:
   try:os.killpg(child.pid,signal.SIGKILL)
   except ProcessLookupError:pass
   child.wait(timeout=5)
 (OUT/'active-process.json').unlink(missing_ok=True)
 results.append(dict(command='supervisor interruption',status='BLOCKED_NOT_TESTED',reason='Controlled interruption; active process group terminated, no test success synthesized',signal=signum));save()
 raise SystemExit(128+signum)
def run(label,args,limit):
 global ACTIVE
 remaining=OVERALL-(time.monotonic()-started)
 if remaining<=0:
  results.append(dict(command=label,status='BLOCKED_NOT_TESTED',reason='Overall 1800s QA deadline exhausted; not started',timeoutSeconds=limit));save();return
 cap=min(limit,max(1,int(remaining)));log=OUT/(label.replace(':','-').replace(' ','_').replace('/','_')+'.log');t=time.monotonic()
 with log.open('w') as output:
  p=subprocess.Popen(args,cwd=ROOT,stdout=output,stderr=subprocess.STDOUT,start_new_session=True);ACTIVE=p
  (OUT/'active-process.json').write_text(json.dumps(dict(command=label,pid=p.pid,timeoutSeconds=cap,startedAt=datetime.datetime.now(datetime.timezone.utc).isoformat())))
  timeout=False
  try:p.wait(timeout=cap)
  except subprocess.TimeoutExpired:
   timeout=True
   try:os.killpg(p.pid,signal.SIGTERM)
   except ProcessLookupError:pass
   try:p.wait(timeout=5)
   except subprocess.TimeoutExpired:
    try:os.killpg(p.pid,signal.SIGKILL)
    except ProcessLookupError:pass
    p.wait(timeout=5)
 (OUT/'active-process.json').unlink(missing_ok=True);ACTIVE=None
 row=dict(command=label,status='BLOCKED_NOT_TESTED' if timeout else ('PASS' if p.returncode==0 else 'FAIL'),exitCode=p.returncode,timeoutSeconds=cap,seconds=round(time.monotonic()-t,2),log=str(log))
 if timeout:row['reason']='Fixed process timeout; complete process group terminated, no automatic retry. Partial assertions not counted as completed.'
 results.append(row);save();print(json.dumps(row),flush=True)
def blocked(command,reason):
 results.append(dict(command=command,status='BLOCKED_NOT_TESTED',reason=reason,started=False));save()
start_iso=datetime.datetime.now(datetime.timezone.utc).isoformat()
signal.signal(signal.SIGTERM,interrupted);signal.signal(signal.SIGINT,interrupted)
if not (ROOT/'node_modules/@playwright/test').exists():
 blocked('dependency preflight','Run bounded npm ci --ignore-scripts first');raise SystemExit(2)
package=json.loads((ROOT/'package.json').read_text());(OUT/'script-definitions.json').write_text(json.dumps(package['scripts'],indent=2))
for f in ['schichtplan_mvp.js','tests/smoke.spec.js','tests/db-integration/permissions-8b2-review.js','tests/security/permissions-8b2-fixture.js','tests/security/permissions-8b2.test.js']:
 run('syntax '+f,['node','--check',f],30)
for script,cap in [('check:syntax',120),('test:ci',60),('test:qa',120),('qa:check',120),('test:integration',180),('test:security',900),('test:article-db',180),('test:operations',180),('test:readiness',180),('test:forecast',180),('test:ba-entry',180),('test:routing',180),('test:setup',180),('test:quantity',180),('test:lifecycle',180),('qa:agent-server',60)]:
 assert script in package['scripts'];run('npm run '+script,['npm','run',script],cap)
# No doomed browser matrix is launched when an engine is absent. No test definitions,
# filters, skips, expectations or references are modified by this supervisor.
try:
 probe=subprocess.run(['node','-e',"const fs=require('node:fs'),p=require('@playwright/test');process.stdout.write(JSON.stringify(Object.fromEntries(['chromium','firefox','webkit'].map(k=>[k,fs.existsSync(p[k].executablePath())]))))"],cwd=ROOT,capture_output=True,text=True,timeout=30)
 engines=json.loads(probe.stdout) if probe.returncode==0 else {}
except (subprocess.TimeoutExpired,json.JSONDecodeError):engines={}
(OUT/'browser-preflight.json').write_text(json.dumps(engines))
if not engines or not all(engines.values()):
 for name in ['test:e2e','test:agents']:blocked('npm run '+name,'Browser prerequisite absent: '+str(engines)+'; not launched, no application assertions passed')
else:
 run('npm run test:e2e',['npm','run','test:e2e'],600);run('npm run test:agents',['npm','run','test:agents'],120)
if os.environ.get('HUMBEL_VISUAL_ENV')!='playwright-1.62.1-noble':blocked('npm run test:visual','Required pinned Playwright container/environment unavailable; not launched, no references changed')
else:run('npm run test:visual',['npm','run','test:visual'],300)
run('npm run test:db:preflight',['npm','run','test:db:preflight'],30)
if results[-1]['command']=='npm run test:db:preflight' and results[-1].get('exitCode')==2:
 results[-1]['status']='BLOCKED_NOT_TESTED';results[-1]['reason']='Existing DB prerequisite guard rejected unavailable approved local runtime/auth/REST fixtures';save()
blocked('native 8B2 PostgreSQL/GoTrue/PostgREST/JWT/concurrency','No approved original-restore native target/adapters supplied; no production target probed. docker='+str(bool(shutil.which('docker')))+', psql='+str(bool(shutil.which('psql'))))
run('git diff --check',['git','diff','--check'],30)
raise SystemExit(1 if any(r['status']=='FAIL' for r in results) else (2 if any(r['status']=='BLOCKED_NOT_TESTED' for r in results) else 0))
