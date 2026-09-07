#!/usr/bin/env python3
"""Retain a completed build before any source mutation or further build."""
import hashlib,json,shutil,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[4]
kind=sys.argv[1];assert kind in ['baseline','candidate']
r=json.loads((ROOT/'checks'/(kind+'-build.json')).read_text());assert r['status']=='pass' and r['source_unchanged']
binaries={}
for instrument,name in [('normal','litchi-perf-baseline'),('allocator','litchi-perf-baseline-alloc')]+([('external','pptx_external_cross_copy')] if kind=='candidate' else []):
    p=Path('/tmp/litchi-goal-0456')/('formal-'+kind+'-'+instrument);assert not p.exists()
    shutil.copyfile(REPO/'tools/perf-baseline/target/release'/name,p);p.chmod(0o700)
    binaries[instrument]={'path':str(p),'sha256':hashlib.sha256(p.read_bytes()).hexdigest(),'bytes':p.stat().st_size}
with (ROOT/(kind+'-build.json')).open('x') as f:f.write(json.dumps({'revision':r['revision'],'source_manifest':r['source_after'],'build_receipt':'checks/'+kind+'-build.json','binaries':binaries},indent=2)+'\n')
print(json.dumps({'status':'pass','build':kind}))
