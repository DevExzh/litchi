#!/usr/bin/env python3
"""Balanced single-case process legs followed by the unchanged matrix control."""
import hashlib,json,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
BIN=ROOT.parent/'litchi-0699-bin'
SCHEDULE=['baseline','candidate','candidate','baseline','candidate','baseline','baseline','candidate','baseline','candidate','candidate','baseline']
CASES=['early-name-error','small-valid','late-root-error-mce']
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
records=[]
for mode,schedule in [('case',SCHEDULE),('matrix',SCHEDULE[:4])]:
 for leg,phase in enumerate(schedule):
  cases=(CASES if leg%2==0 else CASES[::-1]) if mode=='case' else [None]
  for case in cases:
   binary=BIN/phase
   build=json.loads((P/f'build-{phase}.json').read_text())
   assert sha(binary)==build['binary_sha256']
   name=f'{mode}-{case or "all"}-{leg:02}-{phase}'
   out=P/'timing';out.mkdir(exist_ok=True)
   command=['taskset','-c','12',str(binary),mode]+([case] if case else [])+['300','10']
   started=time.monotonic()
   with (out/(name+'.tsv')).open('w') as stdout,(out/(name+'.stderr')).open('w') as stderr:
    r=subprocess.run(command,cwd=ROOT,stdout=stdout,stderr=stderr)
   records.append(dict(mode=mode,case=case,leg=leg,phase=phase,command=command,exit_code=r.returncode,seconds=time.monotonic()-started,binary_sha256=sha(binary),stdout=str((out/(name+'.tsv')).relative_to(P)),stdout_sha256=sha(out/(name+'.tsv')),stderr_sha256=sha(out/(name+'.stderr'))))
   (P/'runs.json').write_text(json.dumps(records,indent=2)+'\n')
   assert r.returncode==0,(out/(name+'.stderr')).read_text()
   print(name,flush=True)
