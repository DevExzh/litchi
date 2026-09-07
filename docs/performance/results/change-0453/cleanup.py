#!/usr/bin/env python3
"""Delete only inventoried baseline drafts and final compared executables."""
import hashlib,json,shutil
from pathlib import Path
ROOT=Path(__file__).resolve().parent;TASK=Path('/tmp/litchi-goal-0453-pptx-payload')
assert TASK.parent==Path('/tmp') and not TASK.is_symlink()
assert json.loads((ROOT/'checks/precleanup.json').read_text())['status']=='pass'
binaries={}
for name in ['draft-baseline-build.json','draft-baseline-r2-build.json','baseline-build.json','candidate-build.json']:
 for row in json.loads((ROOT/name).read_text())['binaries'].values():
  p=Path(row['path']);assert p.parent==TASK and p.name not in binaries;binaries[p.name]=row
assert {p.name for p in TASK.iterdir()}==set(binaries)
for name,row in binaries.items():
 p=TASK/name;assert p.is_file() and not p.is_symlink()
 with p.open('rb') as f:assert hashlib.file_digest(f,'sha256').hexdigest()==row['sha256']
 assert p.stat().st_size==row['bytes']
shutil.rmtree(TASK)
v={'status':'pass','task':str(TASK),'temporary_directory_absent':not TASK.exists(),'binaries':binaries,'files_removed':len(binaries),'bytes_removed':sum(r['bytes'] for r in binaries.values())}
with (ROOT/'checks/binary-cleanup.json').open('x') as f:f.write(json.dumps(v,indent=2)+'\n')
print(json.dumps({'status':'pass','files_removed':v['files_removed'],'bytes_removed':v['bytes_removed']}))
