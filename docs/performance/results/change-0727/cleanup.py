#!/usr/bin/env python3
"""Delete only the two owned diagnostic roots and preserve all binary identities."""
import hashlib,json,shutil
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3]
roots=[ROOT.parent/'litchi-target-0727',ROOT.parent/'litchi-0727-bin'];assert not (P/'cleanup.json').exists()
identities=[]
for b in json.loads((P/'builds.json').read_text()):
 path=Path(b['binary']);assert path.is_file() and not path.is_symlink()
 digest=hashlib.sha256(path.read_bytes()).hexdigest();assert digest==b['binary_sha256'] and path.stat().st_size==b['bytes']
 identities.append(dict(path=str(path),sha256=digest,bytes=b['bytes']))
for root in roots:assert root.is_dir() and not root.is_symlink()
receipt=dict(roots=[str(r) for r in roots],identities=identities,removed=False)
(P/'cleanup.json').write_text(json.dumps(receipt,indent=2)+'\n')
for root in roots:shutil.rmtree(root)
for path in P.rglob('__pycache__'):shutil.rmtree(path)
receipt['removed']=all(not r.exists() for r in roots);assert receipt['removed'];(P/'cleanup.json').write_text(json.dumps(receipt,indent=2)+'\n');print('PASS two owned roots removed; two binary witnesses retained')
