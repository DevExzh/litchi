#!/usr/bin/env python3
"""Remove only this batch's two owned roots after retaining exact binary witnesses."""
import hashlib,json,shutil
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
assert not (P/'cleanup.json').exists()
binaries=json.loads((P/'build.json').read_text())['binaries']+[json.loads((P/'feature-witness.json').read_text())['binary']]
for r in binaries:
 p=Path(r['path']);assert p.is_file() and not p.is_symlink();assert p.stat().st_size==r['bytes'] and hashlib.sha256(p.read_bytes()).hexdigest()==r['sha256']
roots=[ROOT.parent/'litchi-target-0731',ROOT.parent/'litchi-0731-bin']
for p in roots:assert p.is_dir() and not p.is_symlink()
x=dict(binaries=binaries,roots=list(map(str,roots)),removed=False);(P/'cleanup.json').write_text(json.dumps(x,indent=2)+'\n')
for p in roots:shutil.rmtree(p)
for p in P.rglob('__pycache__'):shutil.rmtree(p)
x['removed']=all(not p.exists() for p in roots);assert x['removed'];(P/'cleanup.json').write_text(json.dumps(x,indent=2)+'\n');print('PASS owned roots removed; three executable witnesses retained')
