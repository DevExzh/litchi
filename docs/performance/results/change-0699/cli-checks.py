#!/usr/bin/env python3
"""Verify bounded diagnostic CLI modes on both frozen binaries."""
import hashlib,json,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
cases=[(['case','early-name-error','1','1'],0),(['profile','early-name-error','1'],0),(['case','unknown','1','1'],1),(['case','early-name-error','0','1'],1),(['case','early-name-error','100001','1'],1),(['case','early-name-error','1','0'],1),(['case','early-name-error','1','10001'],1),(['profile','early-name-error','0'],1),(['profile','early-name-error','100001'],1),(['profile','unknown','1'],1)]
records=[]
for phase in ['baseline','candidate']:
 binary=ROOT.parent/'litchi-0699-bin'/phase
 for args,expected in cases:
  r=subprocess.run([str(binary),*args],capture_output=True,text=True)
  assert r.returncode==expected,(args,r.returncode,r.stderr)
  if expected==0:assert 'all_iterations_passed\ttrue' in r.stdout
  records.append(dict(phase=phase,command=[str(binary),*args],exit_code=r.returncode,expected_exit_code=expected,stdout=r.stdout,stderr=r.stderr,binary_sha256=hashlib.sha256(binary.read_bytes()).hexdigest()))
(P/'cli-checks.json').write_text(json.dumps(records,indent=2)+'\n')
print(len(records),'CLI checks pass')
