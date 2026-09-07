#!/usr/bin/env python3
"""Rebuild old production against the exact final harness, then restore reviewed candidate."""
import hashlib,json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[3]
files=json.loads((ROOT/'source-files.json').read_text());production=files[:2]
assert not (ROOT/'checks/baseline-restoration.json').exists()
for name in files:
 p=ROOT/'candidate'/('staged-'+Path(name).name+'.txt');assert not p.exists();p.write_bytes((REPO/name).read_bytes())
# Old production comes from the initial matching-harness baseline source copies.
for name in production:
 (REPO/name).write_bytes((ROOT/'candidate'/('before-'+Path(name).name+'.txt')).read_bytes())
(ROOT/'baseline-build.json').rename(ROOT/'draft-baseline-r2-build.json')
for p in (ROOT/'candidate').glob('before-*.txt'):p.rename(p.with_name('draft-r2-'+p.name))
r=subprocess.run([sys.executable,'-B',str(ROOT/'build.py'),'baseline'])
# subprocess has terminated, so no owned Rust build remains live during restore.
for name in production:(REPO/name).write_bytes((ROOT/'candidate'/('staged-'+Path(name).name+'.txt')).read_bytes())
rows={name:hashlib.sha256((REPO/name).read_bytes()).hexdigest() for name in files}
assert all((REPO/name).read_bytes()==(ROOT/'candidate'/('staged-'+Path(name).name+'.txt')).read_bytes() for name in files)
(ROOT/'checks/baseline-restoration.json').write_text(json.dumps({'status':'pass' if r.returncode==0 else 'failed','build_exit_code':r.returncode,'restored':rows},indent=2)+'\n')
raise SystemExit(r.returncode)
