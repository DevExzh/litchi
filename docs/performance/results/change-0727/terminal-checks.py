#!/usr/bin/env python3
"""Replay complete diagnostic evidence after cleanup, requiring exact outputs."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
outputs=['analysis.json','analysis.md','report-summary.json','audit.json','negative-checks.json']
before={n:hashlib.sha256((P/n).read_bytes()).hexdigest() for n in outputs}
cmds=[['python3','-B',str(P/n)] for n in ['source-guard.py','analyze.py','report-tables.py','audit.py','negative-checks.py']]
cmds += [['python3','-B','tools/check_report_claim_classification.py','--registry','docs/performance/report-claim-classification-v1.json','--repo-root','.'],['python3','-B','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--mode','structural'],['python3','-B','tools/non_iwork_gate.py','verify']]
rows=[]
for i,cmd in enumerate(cmds):
 expected=1 if i==3 else 0
 if i==3:cmd.append('--allow-rejected')
 log=P/f'terminal-{i}.log'
 with log.open('w') as f:r=subprocess.run(cmd,cwd=ROOT,stdout=f,stderr=subprocess.STDOUT)
 rows.append(dict(command=cmd,exit_code=r.returncode,expected_exit_code=expected,log=log.name));assert r.returncode==expected,rows[-1]
after={n:hashlib.sha256((P/n).read_bytes()).hexdigest() for n in outputs};assert before==after
assert all(not Path(p).exists() for p in json.loads((P/'cleanup.json').read_text())['roots'])
(P/'terminal-checks.json').write_text(json.dumps(dict(commands=rows,exact_replay_sha256=after,cleanup_verified=True),indent=2)+'\n');print('PASS eight terminal checks; exact offline replay')
