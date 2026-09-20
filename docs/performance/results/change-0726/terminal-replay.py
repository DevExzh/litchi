#!/usr/bin/env python3
"""Replay terminal evidence after cleanup while captured candidate source is installed."""
import hashlib,json,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
def h(b):return hashlib.sha256(b).hexdigest()
original=(P/'analysis.json').read_bytes();md=(P/'analysis.md').read_bytes()
before=json.loads(original);before.pop('generated_utc')
expected=0 if before['passed'] else 1
disposition='retained' if before['passed'] else 'rejected'
commands=[(['python3','-B',str(P/'analyze.py'),'--offline'],expected,'analysis-final'),(['python3','-B',str(P/'audit.py'),'--allow-rejected'],expected,'audit-final'),(['python3','-B',str(P/'audit-corpus.py')],0,'corpus-final'),(['python3','-B',str(P/'negative-checks.py')],0,'negative-final')]
rows=[]
try:
 for cmd,expected,name in commands:
  t=time.monotonic()
  with (P/(name+'.log')).open('w') as f:r=subprocess.run(cmd,cwd=ROOT,stdout=f,stderr=subprocess.STDOUT)
  row=dict(command=cmd,exit_code=r.returncode,expected_exit_code=expected,seconds=time.monotonic()-t,log=name+'.log');rows.append(row)
  assert r.returncode==expected,row
  if name=='analysis-final':
   after=json.loads((P/'analysis.json').read_text());after.pop('generated_utc');assert before==after
   assert (P/'analysis.md').read_bytes()==md
  if name=='audit-final':
   raw=(P/(name+'.log')).read_text();assert raw.startswith('PASS ' if before['passed'] else 'REJECTED ')
   report=json.loads(raw.split(' ',1)[1]);assert report['hard_gates']==json.loads((P/'audit-initial.log').read_text().split(' ',1)[1])['hard_gates']
  print(name,r.returncode,flush=True)
finally:
 (P/'analysis.json').write_bytes(original);(P/'analysis.md').write_bytes(md)
 (P/'terminal-replay.json').write_text(json.dumps(dict(commands=rows,analysis_bytes_preserved=h(original),expected_disposition=disposition),indent=2)+'\n')
