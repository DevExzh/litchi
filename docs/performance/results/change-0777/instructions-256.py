"""Separate marginal user-instruction observations; never native latency."""
from pathlib import Path
import json,os,subprocess,time
import capture as c
P=c.PACKET;D=c.CAPTURE

def write(path,value):c.json_write(path,value)
def sha(path):return c.sha256(path)
def run(cmd,prefix):
 start=time.time();log=prefix.with_suffix('.log')
 with log.open('w') as f:r=subprocess.run(cmd,cwd=c.AFTER,stdout=f,stderr=subprocess.STDOUT)
 return {'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'log':log.name,'log_sha256':sha(log)}

if __name__=='__main__':
 complete=c.json_read(D/'complete.json');assert complete['complete']
 out=P/'instructions-256';assert not out.exists();out.mkdir();bins=c.json_read(D/'binaries.json')
 refs={'before':c.BASE_REF,'after':complete['candidate']};roots={'before':c.BEFORE,'after':c.AFTER}
 inputs={leg:c.source_manifest(root,refs[leg]) for leg,root in roots.items()}
 for leg in bins:
  for b in bins[leg].values():assert sha(Path(b['path']))==b['sha256']
 csv=out/'qualification.csv';cmd=['perf','stat','-x',';','--no-big-num','-e','instructions:u','-o',str(csv),'--','taskset','-c',str(c.CPU),'/usr/bin/true']
 row=run(cmd,out/'qualification');row.update({'csv':csv.name,'csv_sha256':sha(csv) if csv.exists() else None});write(out/'qualification.json',row)
 supported=csv.exists() and any(len(parts:=line.split(';'))>2 and parts[2]=='instructions:u' and parts[0].strip().isdigit() for line in csv.read_text().splitlines())
 if row['exit']!=0 or not supported:
  write(out/'complete.json',{'supported':False,'reason':'perf qualification failed; inspect retained command/log/csv. No instruction result claimed.'})
  raise SystemExit(0)
 cases=[('opc_relationship_declarations',256)]
 rows=[]
 for case,n in cases:
  for repeat in range(2):
   for leg,count in [('before',3),('after',3),('after',23),('before',23)]:
    key=f'{case}-{n}-{repeat}-{leg}-{count}';prefix=out/key;csv=prefix.with_suffix('.csv');report=prefix.with_suffix('.json')
    binary=bins[leg]['mce-stream-probe' if n is None else 'attribute_checks'];assert sha(Path(binary['path']))==binary['sha256']
    args=[binary['path'],'adversarial' if n is None else 'probe','--case',case,'--samples',str(count),'--warmup','2','--json',str(report)]
    if n is not None:args+=['--n',str(n)]
    cmd=['perf','stat','-x',';','--no-big-num','-e','instructions:u,cycles:u,branches:u,branch-misses:u','-o',str(csv),'--','taskset','-c',str(c.CPU),*args]
    row=run(cmd,prefix);row.update({'case':case,'n':n,'repeat':repeat,'leg':leg,'samples':count,'binary_sha256':binary['sha256'],'csv':csv.name,'csv_sha256':sha(csv),'report':report.name,'report_sha256':sha(report) if report.exists() else None});rows.append(row);write(out/'runs.json',rows);assert row['exit']==0,row
   print(key+' instruction pair done',flush=True)
 iterator_rows=[];binary=bins['after']['attribute_checks_equivalence'];assert sha(Path(binary['path']))==binary['sha256']
 for repeat in range(0):
  for mode,count in [('quick-xml',3),('checked',3),('checked',23),('quick-xml',23)]:
   prefix=out/f'iterator-{mode}-{repeat}-{count}';csv=prefix.with_suffix('.csv')
   cmd=['perf','stat','-x',';','--no-big-num','-e','instructions:u,cycles:u,branches:u,branch-misses:u','-o',str(csv),'--','taskset','-c',str(c.CPU),binary['path'],'--bench',mode,'--rounds',str(count),'test-data/ooxml']
   row=run(cmd,prefix);row.update({'mode':mode,'repeat':repeat,'rounds':count,'binary_sha256':binary['sha256'],'csv':csv.name,'csv_sha256':sha(csv)});iterator_rows.append(row);write(out/'iterator-runs.json',iterator_rows);assert row['exit']==0,row
 for leg in inputs:assert c.source_manifest(roots[leg],refs[leg])==inputs[leg]
 assert c.fixture_manifest()==c.json_read(D/'fixtures-before.json')
 write(out/'complete.json',{'supported':True,'source_unchanged':True,'fixtures_unchanged':True,'runs':len(rows),'iterator_runs':len(iterator_rows),'scope':'Whole-process user counters; two repeats of 3 versus 23 measured operations, same two warmups. Difference/20 is a marginal estimate, not exact timed-region or function attribution. Instrumented elapsed times are excluded from native statistics.'})
