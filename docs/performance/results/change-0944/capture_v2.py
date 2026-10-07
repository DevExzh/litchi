#!/usr/bin/env python3
"""Write-once, fail-closed Callgrind capture for the frozen 0944 matrix."""
from __future__ import annotations
import datetime, hashlib, importlib.util, json, os, pathlib, signal, subprocess, sys, time
from typing import Any
PACKET=pathlib.Path(__file__).resolve().parent.parent
PLAN_DIR=PACKET/'docs/performance/results/change-0944'
PLAN_PATH=PLAN_DIR/'plan.json'
SOURCE=PACKET/'source-build'
RUNTIME_ENV={'PATH':'/usr/bin:/bin','LANG':'C','LC_ALL':'C','TZ':'UTC','RUST_BACKTRACE':'0'}
DEADLINE=900
class EvidenceError(ValueError): pass
def now(): return datetime.datetime.now(datetime.timezone.utc).isoformat()
def reject_pairs(pairs):
 d={}
 for k,v in pairs:
  if k in d: raise EvidenceError(f'duplicate JSON key: {k}')
  d[k]=v
 return d
def load(path): return json.loads(path.read_text(encoding='utf-8'),object_pairs_hook=reject_pairs)
def write(path,obj,exclusive=False):
 path.parent.mkdir(parents=True,exist_ok=True)
 mode='x' if exclusive else 'w'
 with path.open(mode,encoding='utf-8') as f:
  json.dump(obj,f,indent=2,sort_keys=True);f.write(chr(10));f.flush();os.fsync(f.fileno())
 d=os.open(path.parent,os.O_DIRECTORY);os.fsync(d);os.close(d)
def sha(path):
 if not path.is_file() or path.is_symlink(): raise EvidenceError(f'missing regular file: {path}')
 h=hashlib.sha256();n=0
 with path.open('rb') as f:
  for b in iter(lambda:f.read(1<<20),b''): h.update(b);n+=len(b)
 return n,h.hexdigest()
def rel(path): return str(path.resolve().relative_to(PACKET.resolve()))
def desc(path):
 n,h=sha(path);return {'path':rel(path),'bytes':n,'sha256':h}
def module(path,name):
 spec=importlib.util.spec_from_file_location(name,path)
 if spec is None or spec.loader is None: raise EvidenceError(f'cannot load {path}')
 m=importlib.util.module_from_spec(spec);sys.modules[name]=m;spec.loader.exec_module(m);return m
def require(c,m):
 if not c: raise EvidenceError(m)
def process_group_clear(pgid):
 for _ in range(100):
  try: os.killpg(pgid,0)
  except ProcessLookupError: return True
  except PermissionError: return False
  time.sleep(.05)
 return False
def child_command(plan,row,folder):
 binary=(PACKET/pathlib.Path(*pathlib.PurePosixPath(plan['binary']['path']).parts)).resolve()
 callgrind=folder/'callgrind.out';report=folder/'report.json'
 return ['/usr/bin/taskset','-c',str(plan['cpu']),'/usr/bin/valgrind',*plan['options'],'--callgrind-out-file='+str(callgrind),str(binary),'--warmup',str(plan['warmup']),'--samples',str(plan['samples']),'--case',plan['case'],'--xlsx-cell-crud-shape',row['shape'],'--json',str(report)]
def artifact_list(folder):
 out=[]
 for p in sorted(folder.iterdir()):
  if p.name=='terminal.json': continue
  if p.is_symlink() or not p.is_file(): raise EvidenceError(f'unexpected capture entry: {p}')
  out.append(desc(p))
 return out
def main():
 if len(sys.argv)!=3 or sys.argv[1] != '--index': raise SystemExit('usage: capture_v2.py --index INDEX')
 try: index=int(sys.argv[2])
 except ValueError: raise SystemExit('index must be integer')
 plan=load(PLAN_PATH)
 # The custody checker validates all frozen manifests/tools before any child starts.
 source_checker=module(PACKET/'tools/source_check_v2.py','source_check_v2_runtime')
 source_result=source_checker.check(PLAN_PATH)
 schedule=plan['schedule']; require(0<=index<len(schedule),'index outside frozen schedule')
 row=schedule[index]; require(type(row.get('index')) is int and row['index']==index,'schedule index mismatch')
 if index>0:
  admission=PLAN_DIR/'first-child-admission.json'; require(admission.is_file(),'first-child admission is required before later child')
  a=load(admission); require(a.get('status')=='pass','first-child admission failed')
  prior=schedule[index-1]; pf=PLAN_DIR/'capture'/f"{prior['index']:02d}-{prior['shape']}"
  prev_terminal=load(pf/'terminal.json'); require(prev_terminal.get('schema')=='litchi.xlsx-publication-profile-child-terminal-v2' and prev_terminal.get('returncode')==0 and prev_terminal.get('timed_out') is False and prev_terminal.get('process_group_clear') is True and prev_terminal.get('terminal_before_parse') is True,'prior terminal gate missing')
  require(prev_terminal.get('plan')==desc(PLAN_PATH) and prev_terminal.get('binary')==plan['binary'],'prior terminal custody differs')
  sem=load(pf/'semantics.json'); require(sem.get('status')=='pass','prior semantic gate missing')
  raw=load(pf/'raw.json'); require(raw.get('status')=='pass' and raw.get('index')==prior['index'],'prior raw gate missing')
  cmd=load(pf/'command.json'); require(cmd.get('terminal_receipt')==desc(pf/'terminal.json') and cmd.get('semantics')==desc(pf/'semantics.json') and cmd.get('raw')==desc(pf/'raw.json'),'prior command custody differs')
 folder=PLAN_DIR/'capture'/f"{index:02d}-{row['shape']}"; folder.mkdir(parents=True,exist_ok=False)
 argv=child_command(plan,row,folder); started=now(); runner=pathlib.Path(__file__).resolve()
 start={'schema':'litchi.xlsx-publication-profile-child-start-v2','index':index,'row':row,'started_utc':started,'runner':desc(runner),'plan':desc(PLAN_PATH),'source_check':source_result,'argv':argv,'cwd':'source-build','cwd_absolute':str(SOURCE),'environment':RUNTIME_ENV,'binary':plan['binary'],'source_manifest':plan['source_manifest'],'process_group':True,'start_new_session':True,'deadline_seconds':DEADLINE,'terminal_receipt':'terminal.json','artifact_paths':{'stdout':rel(folder/'stdout.txt'),'stderr':rel(folder/'stderr.txt'),'callgrind':rel(folder/'callgrind.out'),'report':rel(folder/'report.json')}}
 write(folder/'start.json',start,exclusive=True)
 stdout=folder/'stdout.txt';stderr=folder/'stderr.txt'
 pid=None; pgid=None; rc=-999; timed_out=False; launch_error=None
 with stdout.open('xb') as out, stderr.open('xb') as err:
  try:
   child=subprocess.Popen(argv,cwd=SOURCE,env=RUNTIME_ENV,stdout=out,stderr=err,start_new_session=True)
   pid=child.pid; pgid=os.getpgid(pid); deadline=time.monotonic()+DEADLINE
   while True:
    rc=child.poll()
    if rc is not None: break
    if time.monotonic()>=deadline:
     timed_out=True
     try: os.killpg(pgid,signal.SIGTERM)
     except ProcessLookupError: pass
     try: child.wait(timeout=10)
     except subprocess.TimeoutExpired:
      try: os.killpg(pgid,signal.SIGKILL)
      except ProcessLookupError: pass
      child.wait()
     rc=child.returncode; break
    time.sleep(.2)
  except Exception as error:
   launch_error=repr(error)
 clear=True if pgid is None else process_group_clear(pgid)
 artifact_error=None
 try: artifacts=artifact_list(folder)
 except Exception as error: artifact_error=repr(error); artifacts=[]
 terminal={'schema':'litchi.xlsx-publication-profile-child-terminal-v2','index':index,'row':row,'started_utc':started,'terminal_utc':now(),'pid':pid,'process_group_id':pgid,'returncode':rc,'timed_out':timed_out,'launch_error':launch_error,'artifact_error':artifact_error,'process_group_clear':clear,'terminal_before_parse':True,'argv':argv,'cwd':'source-build','environment':RUNTIME_ENV,'plan':desc(PLAN_PATH),'runner':desc(runner),'binary':plan['binary'],'artifacts':artifacts}
 write(folder/'terminal.json',terminal,exclusive=True)
 require(rc==0 and not timed_out and clear and launch_error is None and artifact_error is None,'child terminal failed, launch failed, process group remained, or artifact set was invalid')
 # Terminal receipt is durable before any report/raw/semantic parse.
 report=load(folder/'report.json'); semantic=module(PACKET/'tools/semantic_v2.py','semantic_v2_runtime').validate(plan,report,row['shape'])
 write(folder/'semantics.json',semantic,exclusive=True)
 rawmod=module(PACKET/'tools/raw_analyzer_v2.py','raw_analyzer_v2_runtime')
 raw=rawmod.analyze(PLAN_PATH,folder,row);write(folder/'raw.json',raw,exclusive=True)
 command=dict(terminal);command['schema']='litchi.xlsx-publication-profile-command-v2';command['terminal_receipt']=desc(folder/'terminal.json');command['semantics']=desc(folder/'semantics.json');command['raw']=desc(folder/'raw.json');command['artifacts']=artifact_list(folder);write(folder/'command.json',command,exclusive=True)
 if index==0:
  admission={'schema':'litchi.xlsx-publication-profile-first-child-admission-v2','status':'pass','index':0,'row':row,'plan':desc(PLAN_PATH),'runner':desc(runner),'binary':plan['binary'],'source_manifest':plan['source_manifest'],'terminal':desc(folder/'terminal.json'),'report':desc(folder/'report.json'),'semantics':desc(folder/'semantics.json'),'raw':desc(folder/'raw.json'),'validation':{'semantic_gate':True,'raw_callgrind_gate':True,'terminal_before_parse':True,'no_retry_or_replacement':True,'one_measured_owner_call':True,'lifecycle_calls_retained':True}}
  write(PLAN_DIR/'first-child-admission.json',admission,exclusive=True)
 print(json.dumps({'status':'pass','index':index,'shape':row['shape'],'folder':rel(folder),'terminal':rel(folder/'terminal.json'),'report':rel(folder/'report.json'),'raw':rel(folder/'raw.json')},sort_keys=True))
 return 0
if __name__=='__main__':
 try: raise SystemExit(main())
 except Exception as e:
  print(f'capture_v2: {e}',file=sys.stderr);raise
