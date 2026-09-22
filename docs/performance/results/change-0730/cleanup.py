#!/usr/bin/env python3
"""Remove only the two owned roots after binding both release executables."""
import hashlib,json,shutil
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];roots=[ROOT.parent/'litchi-target-0730',ROOT.parent/'litchi-0730-bin']
assert not (P/'cleanup.json').exists();identities=[]
for b in [b for v in ['baseline','candidate','retention'] for b in json.loads((P/(v+'-builds.json')).read_text())['binaries']]:
 p=Path(b['path']);assert p.is_file() and not p.is_symlink();assert hashlib.sha256(p.read_bytes()).hexdigest()==b['sha256'] and p.stat().st_size==b['bytes'];identities.append(b)
for root in roots:assert root.is_dir() and not root.is_symlink()
r=dict(roots=[str(x) for x in roots],identities=identities,removed=False);(P/'cleanup.json').write_text(json.dumps(r,indent=2)+'\n')
for root in roots:shutil.rmtree(root)
for root in P.rglob('__pycache__'):shutil.rmtree(root)
r['removed']=all(not x.exists() for x in roots);assert r['removed'];(P/'cleanup.json').write_text(json.dumps(r,indent=2)+'\n');print('PASS two owned roots removed, six executable witnesses retained')
