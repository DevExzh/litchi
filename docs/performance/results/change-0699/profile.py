#!/usr/bin/env python3
"""Diagnostic call graphs and repeated counter slopes, separate from native timers."""
import hashlib,json,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
BIN=ROOT.parent/'litchi-0699-bin'
SCRATCH=ROOT.parent/'litchi-0699-profile'
SCRATCH.mkdir(exist_ok=True)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
records=[]
def run(name,phase,command):
 out=P/'profiles';out.mkdir(exist_ok=True)
 start=time.monotonic()
 with (out/(name+'.stdout')).open('w') as stdout,(out/(name+'.stderr')).open('w') as stderr:
  r=subprocess.run(command,cwd=ROOT,stdout=stdout,stderr=stderr)
 records.append(dict(name=name,phase=phase,command=command,exit_code=r.returncode,seconds=time.monotonic()-start,binary_sha256=sha(BIN/phase),stdout_sha256=sha(out/(name+'.stdout')),stderr_sha256=sha(out/(name+'.stderr'))))
 (P/'profile-runs.json').write_text(json.dumps(records,indent=2)+'\n')
 print(name,r.returncode,flush=True)
 return r.returncode
for phase in ['baseline','candidate']:
 binary=BIN/phase
 assert sha(binary)==json.loads((P/f'build-{phase}.json').read_text())['binary_sha256']
 for case,n in [('early-name-error',100000),('late-root-error-mce',20000)]:
  name=phase+'-'+case;data=SCRATCH/(name+'.data')
  assert not data.exists(), 'stale profile data'
  rc=run(name+'-record',phase,['perf','record','-F','997','-g','--call-graph','dwarf,16384','-o',str(data),'--','taskset','-c','12',str(binary),'profile',case,str(n)])
  assert rc==0, 'perf record failed'
  if rc==0:
   for mode in ['self','inclusive']:
    assert run(name+'-'+mode,phase,['perf','report','--stdio','--no-children' if mode=='self' else '--children','--percent-limit','0','-i',str(data)])==0, 'perf report failed'
for repeat in range(3):
 for phase in (['baseline','candidate'] if repeat%2==0 else ['candidate','baseline']):
  for n in [1000,11000]:
   assert run(f'{phase}-counter-{repeat}-{n}',phase,['perf','stat','-x','\t','-e','cycles,instructions,branches,branch-misses,cache-misses,page-faults,task-clock','--','taskset','-c','12',str(BIN/phase),'profile','early-name-error',str(n)])==0, 'perf stat failed'
(P/'raw-profile-hashes.json').write_text(json.dumps({f.name:sha(f) for f in SCRATCH.iterdir() if f.is_file()},indent=2)+'\n')
