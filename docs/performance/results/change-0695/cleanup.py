#!/usr/bin/env python3
"""Remove exactly the two batch-owned sibling scratch directories."""
import hashlib,json,shutil
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
build=json.loads((P/'build.json').read_text())
assert sha(ROOT/'Cargo.lock')==build['root_lock_sha256']
assert sha(Path(build['binary']))==build['binary_sha256']
profile=json.loads((P/'profile-data.json').read_text())
assert sha(Path(profile['path']))==profile['sha256']
rows=[]
for path in [ROOT.parent/'litchi-target-0695',ROOT.parent/'litchi-0695-profile']:
 assert path.is_dir() and not path.is_symlink(),path
 shutil.rmtree(path)
 rows.append(dict(path=str(path),removed=not path.exists()))
(P/'cleanup.json').write_text(json.dumps(dict(removed=rows,binary_sha256=build['binary_sha256'],profile_data_sha256=profile['sha256'],root_lock_sha256=sha(ROOT/'Cargo.lock')),indent=2)+'\n')
print('Two owned scratch directories removed; workspace Cargo.lock preserved.')
