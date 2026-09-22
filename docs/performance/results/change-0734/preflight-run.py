#!/usr/bin/env python3
"""Retain every synthetic integration attempt, without overwriting failures."""
import hashlib,json,os,shutil,subprocess,sys
from pathlib import Path
P=Path(__file__).resolve().parent
i=0
while (P/f'preflight-attempt-{i}').exists():i+=1
out=P/f'preflight-attempt-{i}';out.mkdir()
for name in ['analyze.py','audit.py','preflight.py','freeze.json']:
 shutil.copy2(P/name,out/name)
with (out/'output.log').open('wb') as f:r=subprocess.run([sys.executable,str(P/'preflight.py')],stdout=f,stderr=subprocess.STDOUT,env=dict(os.environ,PYTHONDONTWRITEBYTECODE='1'))
(out/'receipt.json').write_text(json.dumps(dict(exit_code=r.returncode,files={f.name:hashlib.sha256(f.read_bytes()).hexdigest() for f in out.iterdir() if f.is_file()}),indent=2)+'\n')
print((out/'output.log').read_text());sys.exit(r.returncode)
