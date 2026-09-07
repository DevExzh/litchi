#!/usr/bin/env python3
"""Inventory and remove only this experiment's private task directory."""
import hashlib,json,shutil
from pathlib import Path
ROOT=Path(__file__).resolve().parent;TASK=Path('/tmp/litchi-goal-0456')
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
pre=ROOT/'precleanup.json';assert pre.is_file();proof=json.loads(pre.read_text());assert proof['status']=='pass' and proof['exit_code']==0
assert proof['verifier_sha256']==sha(ROOT/'formal/verify.py')
assert TASK.is_dir() and not TASK.is_symlink()
assert not any(p.is_symlink() for p in TASK.rglob('*'))
rows=[{'path':str(p.relative_to(TASK)),'bytes':p.stat().st_size,'sha256':sha(p)} for p in sorted(TASK.rglob('*')) if p.is_file()]
manifest=ROOT/'temporary-artifacts.json'
with manifest.open('x') as f:f.write(json.dumps({'task':str(TASK),'files':len(rows),'bytes':sum(r['bytes'] for r in rows),'artifacts':rows},indent=2)+'\n')
shutil.rmtree(TASK)
record={'status':'pass','task':str(TASK),'temporary_directory_absent':not TASK.exists(),'files_removed':len(rows),'bytes_removed':sum(r['bytes'] for r in rows),'inventory':{'path':manifest.name,'bytes':manifest.stat().st_size,'sha256':sha(manifest)},'precleanup_sha256':sha(pre),'driver_sha256':sha(Path(__file__))}
with (ROOT/'cleanup.json').open('x') as f:f.write(json.dumps(record,indent=2)+'\n')
print(json.dumps({k:v for k,v in record.items() if k!='inventory'}))
