#!/usr/bin/env python3
"""Replay the sealed baseline inputs after owned build-root cleanup."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
outputs=['analysis.json','analysis.md','report-summary.json','audit.json','negative-checks.json']
def hashes():return {n:hashlib.sha256((P/n).read_bytes()).hexdigest() for n in outputs}
before=hashes();commands=[['python3','-B',str(P/n)] for n in ['source-guard.py','analyze.py','report-tables.py','audit.py','negative-checks.py']]
commands += [['python3','-B','tools/check_report_claim_classification.py','--registry','docs/performance/report-claim-classification-v1.json','--repo-root','.'],['python3','-B','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--mode','structural'],['python3','-B','tools/non_iwork_gate.py','verify']]
rows=[]
for i,cmd in enumerate(commands):
 log=P/f'terminal-{i}.log'
 with log.open('w') as f:r=subprocess.run(cmd,cwd=ROOT,stdout=f,stderr=subprocess.STDOUT)
 rows.append(dict(command=cmd,exit_code=r.returncode,log=log.name,sha256=hashlib.sha256(log.read_bytes()).hexdigest()));assert r.returncode==0,rows[-1]
assert hashes()==before
assert all(not Path(p).exists() for p in json.loads((P/'cleanup.json').read_text())['roots'])
(P/'terminal-checks.json').write_text(json.dumps(dict(commands=rows,exact_replay_sha256=before,cleanup_verified=True),indent=2)+'\n');print('PASS eight terminal checks; exact offline replay')
