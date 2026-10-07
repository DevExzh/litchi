"""Offline heaptrack exports after all capture terminals; no workloads."""
from pathlib import Path
import datetime,hashlib,json,os,signal,subprocess,time
P=Path(__file__).resolve().parent.parent
O=P/'heaptrack-exports-v2'
def desc(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for b in iter(lambda:f.read(1048576),b''):h.update(b)
 return {'path':str(p),'bytes':p.stat().st_size,'sha256':h.hexdigest()}
def write(p,d):
 with p.open('x') as f:json.dump(d,f,indent=2,sort_keys=True);f.write('\n');f.flush();os.fsync(f.fileno())
def now():return datetime.datetime.now(datetime.timezone.utc).isoformat()
assert not O.exists();O.mkdir()
index=json.loads((P/'receipts/capture-artifact-index-v1.json').read_text())
bound={r['path']:r for r in index['artifacts']}
tool=Path('/usr/bin/heaptrack_print');tool_desc=desc(tool)
rows=[]
for folder in sorted((P/'capture-v1/children').iterdir()):
 terminal=json.loads((folder/'terminal.json').read_text());assert terminal['status']=='pass'
 raw=folder/terminal['artifacts']['profile_actual_name'];observed=desc(raw);old=bound[raw.relative_to(P).as_posix()];assert observed['sha256']==old['sha256'] and observed['bytes']==old['bytes']
 out=O/folder.name;out.mkdir()
 argv=[str(tool),'-f',str(raw),'-m','1','-n','10','-s','10','--print-allocators=1','--print-peaks=0','--print-temporary=0','--print-leaks=0','-H',str(out/'allocation-size.histogram'),'-F',str(out/'allocation-stacks.folded'),'--flamegraph-cost-type','allocations']
 r={'schema':'litchi.performance.0949.heaptrack-export.v1','profile':observed,'tool':tool_desc,'helper':desc(Path(__file__)),'argv':argv,'cwd':str(P),'started_utc':now(),'deadline_seconds':600,'workload_launched':False}
 write(out/'start.json',r)
 start=time.monotonic();code=None;error=None
 with (out/'stdout.txt').open('xb') as stdout,(out/'stderr.txt').open('xb') as stderr:
  proc=subprocess.Popen(argv,cwd=P,stdout=stdout,stderr=stderr,start_new_session=True)
  try:code=proc.wait(timeout=600)
  except subprocess.TimeoutExpired:
   os.killpg(proc.pid,signal.SIGKILL);code=proc.wait();error='timeout'
 r.update({'terminal_utc':now(),'elapsed_seconds':time.monotonic()-start,'exit_code':code,'error':error})
 write(out/'terminal.json',r)
 assert code==0 and error is None,r
 artifacts=[desc(x) for x in sorted(out.iterdir()) if x.is_file()]
 rows.append({'child':folder.name,'status':'pass','artifacts':artifacts})
 print(folder.name,'exported',flush=True)
write(O/'exports.json',{'schema':'litchi.performance.0949.heaptrack-exports.v1','status':'pass','rows':rows,'workload_launched':False})
