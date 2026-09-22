#!/usr/bin/env python3
"""Replay the finalized evidence after scratch removal, without native execution."""
import hashlib,json,os,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent
assert json.loads((P/'cleanup.json').read_text())['removed'] is True
assert all(not Path(r).exists() for r in json.loads((P/'cleanup.json').read_text())['roots'])
rows=[]
for name in ['source-guard.py','analyze.py','audit.py']:
 r=subprocess.run([sys.executable,str(P/name)],env=dict(os.environ,PYTHONDONTWRITEBYTECODE='1'),capture_output=True,text=True)
 rows.append(dict(command=[sys.executable,str(P/name)],exit_code=r.returncode,stdout=r.stdout,stderr=r.stderr));assert r.returncode==0,r.stdout+r.stderr
(P/'terminal.json').write_text(json.dumps(dict(status='passed',runs=rows,cleanup_sha256=hashlib.sha256((P/'cleanup.json').read_bytes()).hexdigest()),indent=2)+'\n')
print('PASS post-cleanup analysis and independent audit')
