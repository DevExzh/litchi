#!/usr/bin/env python3
"""Build and retain both instrumented and ordinary binaries for one source epoch."""
import hashlib,json,shutil,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[3];TASK=Path('/tmp/litchi-goal-0453-pptx-payload')
role=sys.argv[1];assert role in ['baseline','candidate']
version='-r3' if role=='baseline' else ''
argv=['cargo','build','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--features','allocator-metrics','--bin','litchi-perf-baseline','--bin','litchi-perf-baseline-alloc']
subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag',role+'-build'+version,'--',*argv],check=True)
r=json.loads((ROOT/'checks'/f'{role}-build{version}.json').read_text());assert r['status']=='pass'
binaries={}
for label,suffix in [('normal',''),('alloc','-alloc')]:
 p=TASK/(role+version+suffix);assert not p.exists();shutil.copyfile(REPO/'tools/perf-baseline/target/release'/('litchi-perf-baseline'+suffix),p);p.chmod(0o700)
 with p.open('rb') as f:sha=hashlib.file_digest(f,'sha256').hexdigest()
 binaries[label]={'path':str(p),'bytes':p.stat().st_size,'sha256':sha}
with (ROOT/(role+'-build.json')).open('x') as f:f.write(json.dumps({'revision':r['revision'],'source_manifest':r['source_after'],'binaries':binaries},indent=2)+'\n')
for name in json.loads((ROOT/'source-files.json').read_text()):
 (ROOT/'candidate'/(( 'before-' if role=='baseline' else 'after-')+Path(name).name+'.txt')).write_bytes((REPO/name).read_bytes())
print(json.dumps({'status':'pass','role':role,'binaries':binaries}))
