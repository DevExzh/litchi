#!/usr/bin/env python3
"""Bind the terminal documentation classification check to the final reports."""
import hashlib
import json
from pathlib import Path
import subprocess
import time
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def main():
    assert not (P/'final-report-gate.json').exists()
    command=['python3','tools/check_report_claim_classification.py','--registry','docs/performance/report-claim-classification-v1.json','--repo-root','.']
    names=['docs/performance/change-0711.md','docs/performance/REPORT.md',*[f'docs/performance/{n}.md' for n in ['BASELINE','HOTSPOTS','GOAL_AUDIT']],'docs/performance/results/change-0711/README.md']
    before={n:sha(ROOT/n) for n in names};start=time.monotonic()
    with (P/'final-report-gate.log').open('w') as log:r=subprocess.run(command,cwd=ROOT,stdout=log,stderr=subprocess.STDOUT)
    assert before=={n:sha(ROOT/n) for n in names}
    (P/'final-report-gate.json').write_text(json.dumps(dict(command=command,exit_code=r.returncode,seconds=time.monotonic()-start,log_sha256=sha(P/'final-report-gate.log'),docs=before),indent=2)+'\n')
    assert r.returncode==0
if __name__=='__main__':main()
