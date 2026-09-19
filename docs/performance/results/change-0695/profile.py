#!/usr/bin/env python3
"""Isolated MCE hardware counters and full-sequence attribution profile."""
import hashlib,json,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
build=json.loads((P/'build.json').read_text());binary=Path(build['binary'])
assert sha(binary)==build['binary_sha256']
out=P/'profile';out.mkdir(exist_ok=True)
scratch=ROOT.parent/'litchi-0695-profile';scratch.mkdir(exist_ok=True)
rows=[]
def run(name,command):
 output=out/(name+'.stdout');error=out/(name+'.stderr')
 start=time.monotonic()
 with output.open('w') as stdout,error.open('w') as stderr:
  r=subprocess.run(command,stdout=stdout,stderr=stderr)
 rows.append(dict(name=name,command=command,exit_code=r.returncode,seconds=time.monotonic()-start,stdout_sha256=sha(output),stderr_sha256=sha(error),binary_sha256=sha(binary)))
 (P/'profile-runs.json').write_text(json.dumps(rows,indent=2)+'\n')
 print(name,r.returncode,flush=True)
 return r.returncode
for repeat in range(3):
 for case in (['all','presentation','slides'] if repeat%2==0 else ['slides','presentation','all']):
  for count in ([10,210] if repeat%2==0 else [210,10]):
   run(f'{case}-{repeat}-{count}',['perf','stat','-x','\t','-e','cycles,instructions,branches,branch-misses,cache-misses,page-faults,task-clock','--','taskset','-c','12',str(binary),str(P/'sequences'/(case+'.txt')),'0',str(count),'10'])
data=scratch/'all.data'
if run('record',['perf','record','-F','997','-g','--call-graph','dwarf,16384','-o',str(data),'--','taskset','-c','12',str(binary),str(P/'sequences/all.txt'),'10','1000','1'])==0:
 for name,flag,limit in [('self','--no-children','0.1'),('inclusive','--children','1')]:
  run(name,['perf','report','--stdio',flag,'--percent-limit',limit,'-i',str(data)])
 (P/'profile-data.json').write_text(json.dumps(dict(path=str(data),sha256=sha(data)),indent=2)+'\n')
