from pathlib import Path
import datetime as dt, hashlib, json, os, stat, subprocess
ROOT=Path('/home/zhuhe/litchi-performance-0959'); REC=ROOT/'receipts/layout-probe-summary-v1.json'
def sha(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for b in iter(lambda:f.read(1<<20),b''):h.update(b)
 return h.hexdigest()
def desc(p):
 p=Path(p).resolve(); st=p.stat(); return {'path':str(p),'bytes':st.st_size,'mode':stat.S_IMODE(st.st_mode),'executable':bool(st.st_mode&0o111),'sha256':sha(p)}
def load(name): return json.loads((ROOT/'receipts'/name).read_text())
v1t=load('layout-probe-terminal-v1.json'); v1r=load('layout-probe-result-v1.json'); v2t=load('layout-probe-terminal-v2.json'); v2r=load('layout-probe-result-v2.json')
status=subprocess.check_output(['git','-C','/home/zhuhe/code/litchi','status','--short'],text=True).splitlines()
summary={
 'schema':'litchi.performance.0959.layout-probe-summary.v1','created_utc':dt.datetime.now(dt.timezone.utc).isoformat(timespec='microseconds'),'status':'passed',
 'purpose':'test-only type layout observation; no integrated representation or production candidate change',
 'source':{'root':str(ROOT/'source-baseline'),'manifest':desc(ROOT/'source-manifest-probe-v2.json'),'baseline_receipt':desc(ROOT/'receipts/source-baseline-v2.json'),'probe_patch':desc(ROOT/'review/cell-layout-probe-v2.patch'),'cell_sha256':sha(ROOT/'source-baseline/crates/litchi-xlsx/src/cell.rs'),'cargo_lock_present':(ROOT/'source-baseline/Cargo.lock').exists()},
 'checkout':{'head':subprocess.check_output(['git','-C','/home/zhuhe/code/litchi','rev-parse','HEAD'],text=True).strip(),'status':status,'production_unchanged':True},
 'tooling':{'runner':desc(ROOT/'tooling/run-layout-probe-0959-v2.py'),'rustc':'rustc 1.95.0 (59807616e 2026-04-14)','cargo':'cargo 1.95.0 (f2d3ce0bd 2026-03-21)','cargo_command':['/home/zhuhe/.cargo/bin/cargo','test','--manifest-path',str(ROOT/'source-baseline/Cargo.toml'),'--locked','--offline','-p','litchi-xlsx','--lib','layout_probe_0959','--','--nocapture'],'target':str(ROOT/'target-v1'),'tmp':str(ROOT/'tmp/layout-probe')},
 'failed_v1_preserved':{'terminal':desc(ROOT/'receipts/layout-probe-terminal-v1.json'),'result':desc(ROOT/'receipts/layout-probe-result-v1.json'),'stdout':desc(ROOT/'review/layout-probe.stdout'),'stderr':desc(ROOT/'review/layout-probe.stderr'),'reason':'workspace deny lint rejected unnecessary std::mem qualification in test-only probe'},
 'passed_v2': {
  'start': desc(ROOT / 'receipts/layout-probe-start-v2.json'),
  'terminal': desc(ROOT / 'receipts/layout-probe-terminal-v2.json'),
  'result': desc(ROOT / 'receipts/layout-probe-result-v2.json'),
  'stdout': desc(ROOT / 'review/layout-probe-v2.stdout'),
  'stderr': desc(ROOT / 'review/layout-probe-v2.stderr'),
  'lock_cleanup': desc(ROOT / 'receipts/layout-probe-lock-cleanup-v2.json'),
  'returncode': v2t['run']['returncode'],
  'rows': v2r['rows'],
 },
 'layout_rows':v2r['rows'],'no_workload':True,'no_shared_target':True,'process_probe_after_capture':[],
}
with REC.open('x') as f: json.dump(summary,f,indent=2,sort_keys=True); f.write('\n'); f.flush(); os.fsync(f.fileno())
print(REC,REC.stat().st_size,sha(REC))
