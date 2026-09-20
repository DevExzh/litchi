#!/usr/bin/env python3
"""Replay the complete diagnostic evidence after cleanup, preserving exact results."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
outputs=['analysis.json','analysis.md','audit.json','negative-checks.json']
before={n:hashlib.sha256((P/n).read_bytes()).hexdigest() for n in outputs}
cmds=[['python3','-B',str(P/n)] for n in ('source-guard.py','analyze.py','audit.py','negative-checks.py')]
cmds += [['python3','-B','tools/check_report_claim_classification.py','--registry','docs/performance/report-claim-classification-v1.json','--repo-root','.'],['python3','-B','tools/check_perf_claims.py','--registry','docs/performance/claim-registry-v1.json','--repo-root','.','--mode','structural'],['python3','-B','tools/non_iwork_gate.py','verify']]
rows=[]
for i,cmd in enumerate(cmds):
 log=P/f'terminal-{i}.log'
 with log.open('w') as f:r=subprocess.run(cmd,cwd=ROOT,stdout=f,stderr=subprocess.STDOUT)
 rows.append(dict(command=cmd,exit_code=r.returncode,log=log.name));assert r.returncode==0,rows[-1]
after={n:hashlib.sha256((P/n).read_bytes()).hexdigest() for n in outputs};assert before==after
assert json.loads((P/'analysis.json').read_text())['stages']==json.loads((P/'audit.json').read_text())['stages']
assert all(not Path(p).exists() for p in json.loads((P/'cleanup.json').read_text())['roots'])
(P/'terminal-checks.json').write_text(json.dumps(dict(commands=rows,exact_replay_sha256=after,independent_statistics_equal=True,cleanup_verified=True),indent=2)+'\n');print('PASS seven terminal checks; exact offline replay and independent statistics')
