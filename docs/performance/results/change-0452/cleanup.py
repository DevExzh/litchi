#!/usr/bin/env python3
"""Remove only the two hashed lifecycle executables after portable precleanup replay."""
import hashlib,json,shutil
from pathlib import Path
ROOT=Path(__file__).resolve().parent;TASK=Path('/tmp/litchi-goal-0452-pptx-capture')
assert TASK.parent==Path('/tmp') and not TASK.is_symlink()
assert json.loads((ROOT/'checks/precleanup.json').read_text())['status']=='pass'
binaries={name:json.loads((ROOT/(name+'-build.json')).read_text())['binary'] for name in ['baseline','candidate']}
assert {p.name for p in TASK.iterdir()}==set(binaries)
for name,row in binaries.items():
    p=TASK/name;assert str(p)==row['path'] and p.is_file() and not p.is_symlink()
    with p.open('rb') as f:assert hashlib.file_digest(f,'sha256').hexdigest()==row['sha256']
    assert p.stat().st_size==row['bytes']
shutil.rmtree(TASK)
v={'status':'pass','task':str(TASK),'temporary_directory_absent':not TASK.exists(),'binaries':binaries,'bytes_removed':sum(r['bytes'] for r in binaries.values())}
with (ROOT/'checks/binary-cleanup.json').open('x') as f:f.write(json.dumps(v,indent=2)+'\n')
print(json.dumps({'status':'pass','bytes_removed':v['bytes_removed']}))
