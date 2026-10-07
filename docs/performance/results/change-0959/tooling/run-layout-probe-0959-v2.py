#!/usr/bin/env python3
"""Bounded 0959 layout probe runner; no production workload or source edit."""
from __future__ import annotations
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import signal
import shutil
import subprocess
import time
from typing import Any

ROOT=Path('/home/zhuhe/litchi-performance-0959')
SOURCE=ROOT/'source-baseline'
TARGET=ROOT/'target-v1'
TMP=ROOT/'tmp/layout-probe'
REVIEW=ROOT/'review'
RECEIPTS=ROOT/'receipts'
MANIFEST=ROOT/'source-manifest-probe-v2.json'
CARGO=Path('/home/zhuhe/.cargo/bin/cargo')
CHECKOUT=Path('/home/zhuhe/code/litchi')
CARGO_LOCK=SOURCE/'Cargo.lock'
ROOT_LOCK=CHECKOUT/'Cargo.lock'
ENV={
 'CARGO_HOME':'/home/zhuhe/.cargo','CARGO_TARGET_DIR':str(TARGET),
 'CARGO_BUILD_JOBS':'2','CARGO_INCREMENTAL':'0','CARGO_NET_OFFLINE':'true',
 'HOME':'/home/zhuhe','LANG':'C','LC_ALL':'C','PATH':'/home/zhuhe/.cargo/bin:/usr/bin:/bin',
 'RUSTUP_TOOLCHAIN':'1.95.0','RUST_BACKTRACE':'0','TZ':'UTC','TMP':str(TMP),'TEMP':str(TMP),'TMPDIR':str(TMP),
}
ARGV=[str(CARGO),'test','--manifest-path',str(SOURCE/'Cargo.toml'),'--locked','--offline','-p','litchi-xlsx','--lib','layout_probe_0959','--','--nocapture']
EXPECTED_INTEGRATED={'Number','Value','Cell','Stored','Formula','Content','Option<Number>','Option<Value>','Option<Cell>','Option<Stored>','Option<Formula>','Option<Content>'}
EXPECTED_ALTERNATIVE={'Inline15','Inline23','BoxArc(Box+Arc<str>)','Option<Inline15>','Option<Inline23>','Option<BoxArc>'}

def now(): return dt.datetime.now(dt.timezone.utc).isoformat(timespec='microseconds')
def sha(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for b in iter(lambda:f.read(1<<20),b''): h.update(b)
 return h.hexdigest()
def desc(p):
 p=p.resolve(); st=p.stat(); return {'path':str(p),'bytes':st.st_size,'mode':st.st_mode&0o777,'executable':bool(st.st_mode&0o111),'sha256':sha(p)}
def write_once(p,v):
 if p.exists() or p.is_symlink(): raise RuntimeError(f'refusing overwrite: {p}')
 p.parent.mkdir(parents=True,exist_ok=True)
 with p.open('x',encoding='utf-8') as f:
  json.dump(v,f,indent=2,sort_keys=True); f.write('\n'); f.flush(); os.fsync(f.fileno())

def ancestors():
 s={os.getpid()}; cur=os.getpid()
 while True:
  try: parent=int(Path(f'/proc/{cur}/stat').read_text().split()[3])
  except (OSError,ValueError,IndexError): break
  if parent<=1 or parent in s: break
  s.add(parent);cur=parent
 return s

def process_probe():
 bad=[]; skip=ancestors()
 for e in Path('/proc').iterdir():
  if not e.name.isdigit() or int(e.name) in skip: continue
  try: comm=(e/'comm').read_text().strip(); args=(e/'cmdline').read_bytes().replace(b'\0',b' ').decode(errors='replace').strip()
  except OSError: continue
  if comm in {'cargo','rustc','rustdoc'} or any(x in args for x in ('layout_probe_0959','litchi-xlsx')):
   bad.append({'pid':int(e.name),'comm':comm,'cmdline':args})
 return bad

def clear(pgid):
 if pgid is None:return True
 try: os.killpg(pgid,0)
 except ProcessLookupError:return True
 except PermissionError:return False
 return False

def stop(pgid,proc):
 actions=[]
 if pgid is not None:
  try: os.killpg(pgid,signal.SIGTERM);actions.append('SIGTERM')
  except ProcessLookupError: pass
 try: proc.wait(timeout=15)
 except subprocess.TimeoutExpired:
  if pgid is not None:
   try: os.killpg(pgid,signal.SIGKILL);actions.append('SIGKILL')
   except ProcessLookupError: pass
  proc.wait()
 end=time.monotonic()+30
 while time.monotonic()<end and not clear(pgid): time.sleep(.05)
 return {'actions':actions,'process_group_clear':clear(pgid),'reaped':True}

if __name__=='__main__':
 for p in (SOURCE,ROOT_LOCK,MANIFEST,CARGO):
  if not p.exists(): raise SystemExit(f'missing required input: {p}')
 if CARGO_LOCK.exists(): raise SystemExit(f'source lock already present: {CARGO_LOCK}')
 if process_probe(): raise SystemExit(f'competing cargo/rust process: {process_probe()}')
 TMP.mkdir(parents=True,exist_ok=True); TARGET.mkdir(parents=True,exist_ok=True); REVIEW.mkdir(exist_ok=True); RECEIPTS.mkdir(exist_ok=True)
 lock_desc=desc(ROOT_LOCK); shutil.copy2(ROOT_LOCK,CARGO_LOCK); lock_injected=desc(CARGO_LOCK)
 start_path=RECEIPTS/'layout-probe-start-v2.json'; terminal_path=RECEIPTS/'layout-probe-terminal-v2.json'; result_path=RECEIPTS/'layout-probe-result-v2.json'; cleanup_path=RECEIPTS/'layout-probe-lock-cleanup-v2.json'
 start={'schema':'litchi.performance.0959.layout-probe-start.v2','status':'started','created_utc':now(),'argv':ARGV,'cwd':str(SOURCE),'environment':ENV,'source_manifest':desc(MANIFEST),'source_cell':desc(SOURCE/'crates/litchi-xlsx/src/cell.rs'),'checkout_head':subprocess.check_output(['git','-C',str(CHECKOUT),'rev-parse','HEAD'],text=True).strip(),'cargo':desc(CARGO),'root_lock':lock_desc,'injected_lock':lock_injected,'target':str(TARGET),'tmp':str(TMP),'deadline_seconds':1800,'no_workload':True,'process_group':True}
 write_once(start_path,start)
 stdout=REVIEW/'layout-probe-v2.stdout'; stderr=REVIEW/'layout-probe-v2.stderr'
 proc=None; pgid=None; run={'started_utc':now(),'launched_utc':None,'terminal_utc':None,'pid':None,'process_group_id':None,'returncode':None,'timed_out':False,'launch_error':None,'cleanup':{'actions':[],'process_group_clear':True,'reaped':False}}
 try:
  with stdout.open('x') as out, stderr.open('x') as err:
   try:
    run['started_utc']=now(); proc=subprocess.Popen(ARGV,cwd=SOURCE,env=ENV,stdin=subprocess.DEVNULL,stdout=out,stderr=err,start_new_session=True); pgid=proc.pid; run['pid']=proc.pid; run['process_group_id']=pgid; run['launched_utc']=now()
    try: run['returncode']=proc.wait(timeout=1800)
    except subprocess.TimeoutExpired:
     run['timed_out']=True; run['cleanup']=stop(pgid,proc); run['returncode']=proc.returncode
    if not run['cleanup']['reaped']:
     run['cleanup']={'actions':[],'process_group_clear':clear(pgid),'reaped':True}; run['returncode']=proc.wait()
   except BaseException as e:
    run['launch_error']=f'{type(e).__name__}: {e}'
    if proc is not None: run['cleanup']=stop(pgid,proc); run['returncode']=proc.returncode
 except BaseException as e:
  run['launch_error']=run['launch_error'] or f'{type(e).__name__}: {e}'
  if proc is not None and not run['cleanup']['reaped']: run['cleanup']=stop(pgid,proc); run['returncode']=proc.returncode
 run['terminal_utc']=now()
 terminal={'schema':'litchi.performance.0959.layout-probe-terminal.v2','status':'terminal','run':run,'stdout':desc(stdout) if stdout.exists() else None,'stderr':desc(stderr) if stderr.exists() else None,'source_manifest':desc(MANIFEST),'target':str(TARGET),'terminal_before_output_parse':True,'no_workload':True}
 write_once(terminal_path,terminal)
 # Only now inspect probe output, after the durable process terminal.
 lines=stdout.read_text(errors='replace').splitlines() if stdout.exists() else []
 rows=[]
 for line in lines:
  if not line.startswith('layout_probe_0959 '): continue
  fields={}
  for token in line.split()[1:]:
   if '=' in token:
    k,v=token.split('=',1);fields[k]=v
  if fields.get('size','').isdigit() and fields.get('align','').isdigit(): rows.append({'raw':line,'category':fields.get('category'),'type':fields.get('type'),'size':int(fields['size']),'align':int(fields['align'])})
 integrated={x['type'] for x in rows if x['category']=='integrated'}; alternatives={x['type'] for x in rows if x['category']=='alternative_nonintegrated'}
 accepted=(run['returncode']==0 and not run['timed_out'] and run['cleanup']['process_group_clear'] and run['cleanup']['reaped'] and integrated==EXPECTED_INTEGRATED and alternatives==EXPECTED_ALTERNATIVE and all(x['size']>0 and x['align']>0 for x in rows))
 result={'schema':'litchi.performance.0959.layout-probe-result.v2','status':'pass' if accepted else 'fail','terminal':desc(terminal_path),'stdout':desc(stdout) if stdout.exists() else None,'stderr':desc(stderr) if stderr.exists() else None,'integrated_types':sorted(integrated),'alternative_types':sorted(alternatives),'rows':rows,'expected_integrated_types':sorted(EXPECTED_INTEGRATED),'expected_alternative_types':sorted(EXPECTED_ALTERNATIVE),'no_production_candidate':True,'no_workload':True}
 # Remove only the temporary lock injected for this test; retain its descriptor in cleanup receipt.
 lock_after=desc(CARGO_LOCK); CARGO_LOCK.unlink(); cleanup={'schema':'litchi.performance.0959.layout-probe-lock-cleanup.v2','status':'removed_test_only_lock','path':str(CARGO_LOCK),'descriptor':lock_after,'original_checkout_lock':lock_desc,'remaining_source_lock':not CARGO_LOCK.exists()}
 write_once(cleanup_path,cleanup)
 result['lock_cleanup']=desc(cleanup_path); write_once(result_path,result)
 print(json.dumps({'status':result['status'],'returncode':run['returncode'],'rows':len(rows),'integrated':len(integrated),'alternatives':len(alternatives),'terminal':str(terminal_path),'result':str(result_path)},sort_keys=True))
 raise SystemExit(0 if accepted else 1)
