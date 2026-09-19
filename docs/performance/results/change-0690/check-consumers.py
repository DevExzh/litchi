#!/usr/bin/env python3
"""Legacy DOC/PPT consumer compatibility checks; shared CFB directory lookup changes."""
import hashlib,json,os,subprocess,time
from pathlib import Path
P=Path(__file__).resolve().parent;ROOT=P.parents[3];cmd=['cargo','test','-p','litchi-doc','-p','litchi-ppt','--all-features','--locked','--','--test-threads=1']
hashes={str(f.relative_to(ROOT)):hashlib.sha256(f.read_bytes()).hexdigest() for owner in ['litchi-cfb','litchi-doc','litchi-ppt'] for f in sorted((ROOT/'crates'/owner).rglob('*.rs'))};start=time.monotonic()
with (P/'consumer-tests.log').open('w') as log:r=subprocess.run(cmd,cwd=ROOT,env=dict(os.environ,CARGO_BUILD_JOBS='2'),stdout=log,stderr=subprocess.STDOUT)
assert all(hashlib.sha256((ROOT/n).read_bytes()).hexdigest()==h for n,h in hashes.items())
(P/'consumer-tests.json').write_text(json.dumps(dict(command=cmd,source_sha256=hashes,exit_code=r.returncode,seconds=time.monotonic()-start),indent=2)+'\n');print('consumer tests',r.returncode,flush=True);raise SystemExit(r.returncode)
