#!/usr/bin/env python3
"""Run heaptrack_print offline against the sealed 0958 profiles only."""
from __future__ import annotations
import datetime as dt, hashlib, json, os, signal, subprocess, sys, time
from pathlib import Path

ROOT=Path('/home/zhuhe/litchi-performance-0958')
MANIFEST=ROOT/'capture-manifest-v1.json'
OUTPUT=ROOT/'heaptrack-exports-v1'
TOOL=Path('/usr/bin/heaptrack_print')
EXPECTED=('h00-medium-percent-r1','h01-dense-percent-r1','h02-medium-one-edit-r1','h03-medium-one-edit-r2','h04-dense-percent-r2','h05-medium-percent-r2')

def now(): return dt.datetime.now(dt.timezone.utc).isoformat(timespec='microseconds')
def sha(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for block in iter(lambda:f.read(1024*1024),b''): h.update(block)
 return h.hexdigest()
def desc(p):
 if not p.is_file() or p.is_symlink(): raise RuntimeError(f'not regular: {p}')
 s=p.stat(); return {'path':str(p.resolve()),'bytes':s.st_size,'mode':s.st_mode&0o777,'sha256':sha(p)}
def write(p,v):
 if p.exists() or p.is_symlink(): raise RuntimeError(f'refusing overwrite: {p}')
 p.parent.mkdir(parents=True,exist_ok=True)
 with p.open('xb') as f:
  f.write((json.dumps(v,indent=2,sort_keys=True)+'\n').encode()); f.flush(); os.fsync(f.fileno())
def load():
 if not MANIFEST.is_file() or MANIFEST.is_symlink(): raise RuntimeError('missing sealed capture manifest')
 m=json.loads(MANIFEST.read_text())
 if m.get('schema')!='litchi.performance.0958.capture-manifest.v1': raise RuntimeError('capture manifest schema')
 rows={}
 for item in m.get('files',[]):
  if isinstance(item,dict) and isinstance(item.get('path'),str): rows[item['path']]=item
 profiles=[]
 for name in EXPECTED:
  rel=f'capture-v1/children/{name}/heaptrack.data.gz.zst'
  old=rows.get(rel)
  if not old: raise RuntimeError(f'missing manifest profile {rel}')
  p=ROOT/rel; actual=desc(p)
  if actual['sha256']!=old.get('sha256') or actual['bytes']!=old.get('bytes'): raise RuntimeError(f'profile descriptor changed: {rel}')
  profiles.append((name,p,actual))
 return m,profiles

def main():
 if OUTPUT.exists() or OUTPUT.is_symlink(): raise RuntimeError(f'output root exists: {OUTPUT}')
 tool=desc(TOOL); _m,profiles=load(); OUTPUT.mkdir()
 rows=[]
 for name,profile,pdesc in profiles:
  out=OUTPUT/name; out.mkdir()
  hist=out/'allocation-size.histogram'; folded=out/'allocation-stacks.folded'; stdout=out/'stdout.txt'; stderr=out/'stderr.txt'
  argv=[str(TOOL),'-f',str(profile),'-m','1','-n','1000','-s','1000','-a','-p','-T','-l','--disable-builtin-suppressions','--disable-embedded-suppressions','-H',str(hist),'-F',str(folded),'--flamegraph-cost-type','allocations']
  start={'schema':'litchi.performance.0958.heaptrack-export-child.v1','status':'started','id':name,'profile':pdesc,'tool':tool,'argv':argv,'cwd':str(ROOT),'environment':{'PATH':'/usr/bin:/bin','LANG':'C','LC_ALL':'C','RUST_BACKTRACE':'0','TZ':'UTC'},'deadline_seconds':900,'workload_started':False,'terminal_before_parse':True,'started_utc':now()}
  write(out/'start.json',start)
  rc=None; timed=False; error=None; began=time.monotonic()
  with stdout.open('xb') as out_stream, stderr.open('xb') as err_stream:
   try:
    proc=subprocess.Popen(argv,cwd=ROOT,env=start['environment'],stdin=subprocess.DEVNULL,stdout=out_stream,stderr=err_stream,start_new_session=True)
    try: rc=proc.wait(timeout=900)
    except subprocess.TimeoutExpired:
     timed=True
     try: os.killpg(proc.pid,signal.SIGTERM)
     except ProcessLookupError: pass
     try: proc.wait(timeout=15)
     except subprocess.TimeoutExpired:
      try: os.killpg(proc.pid,signal.SIGKILL)
      except ProcessLookupError: pass
      proc.wait()
     error='timeout'
   except Exception as exc: error=f'{type(exc).__name__}: {exc}'
  terminal={'schema':'litchi.performance.0958.heaptrack-export-child.v1','status':'pass' if rc==0 and not timed and error is None else 'fail','id':name,'profile':pdesc,'tool':tool,'argv':argv,'cwd':str(ROOT),'returncode':rc,'timed_out':timed,'error':error,'elapsed_seconds':time.monotonic()-began,'terminal_before_parse':True,'workload_started':False,'stdout':desc(stdout),'stderr':desc(stderr),'histogram':desc(hist) if hist.exists() else None,'folded':desc(folded) if folded.exists() else None,'terminal_utc':now()}
  write(out/'terminal.json',terminal); print(name,terminal['status'],flush=True)
  rows.append({'id':name,'status':terminal['status'],'terminal':desc(out/'terminal.json'),'histogram':terminal['histogram'],'folded':terminal['folded']})
 index={'schema':'litchi.performance.0958.heaptrack-exports.v1','status':'pass' if all(x['status']=='pass' for x in rows) else 'fail','tool':tool,'manifest':desc(MANIFEST),'rows':rows,'workload_started':False,'created_utc':now()}
 write(OUTPUT/'exports.json',index)
 print(json.dumps({'status':index['status'],'output':str(OUTPUT)},sort_keys=True))
 return 0 if index['status']=='pass' else 1
if __name__=='__main__':
 try: raise SystemExit(main())
 except Exception as exc: print(f'0958 export refused/failed: {exc}',file=sys.stderr); raise SystemExit(1)
