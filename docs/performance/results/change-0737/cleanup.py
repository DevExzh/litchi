#!/usr/bin/env python3
"""Remove only the batch-owned build target and exact captured binaries."""
import shutil
from contract import P, ROOT, read, sha, write
from run import guard

guard()
assert read(P/'analysis.json')['status']=='passed'
assert read(P/'audit.json')['status']=='passed'
binaries = [row for variant in ('legacy','controls')
            for row in read(P/f'{variant}-build.json')['binaries'].values()]
from pathlib import Path
for row in binaries:
    assert sha(Path(row['path']))==row['sha256']
roots = [ROOT.parent/'litchi-target-0737',ROOT.parent/'litchi-0737-bin']
assert {str(p) for p in roots[1].iterdir()}=={row['path'] for row in binaries}
for root in roots:
    assert root.is_dir() and not root.is_symlink()
    shutil.rmtree(root)
write(P/'cleanup.json',dict(removed=True,roots=[str(p) for p in roots],binaries=binaries))
guard()
print('PASS exact four-binary cleanup and restored-source guard')
