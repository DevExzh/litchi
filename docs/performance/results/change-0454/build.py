#!/usr/bin/env python3
"""Build final ordinary lifecycle and external fixture executables."""
import hashlib,json,shutil,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent;REPO=ROOT.parents[3];TASK=Path('/tmp/litchi-goal-0454-native-copy')
argv=['cargo','build','--locked','--release','--manifest-path','tools/perf-baseline/Cargo.toml','--features','allocator-metrics','--bin','litchi-perf-baseline','--bin','pptx_external_cross_copy','--bin','pptx_native_copy_probe']
subprocess.run([sys.executable,'-B',str(ROOT/'check.py'),'--tag','final-candidate-build-r2','--',*argv],check=True)
r=json.loads((ROOT/'checks/final-candidate-build-r2.json').read_text());assert r['status']=='pass'
binaries={}
for label,name in [('normal','litchi-perf-baseline'),('external','pptx_external_cross_copy'),('probe','pptx_native_copy_probe')]:
 p=TASK/('candidate-'+label);assert not p.exists();shutil.copyfile(REPO/'tools/perf-baseline/target/release'/name,p);p.chmod(0o700)
 with p.open('rb') as f:digest=hashlib.file_digest(f,'sha256').hexdigest()
 binaries[label]={'path':str(p),'bytes':p.stat().st_size,'sha256':digest}
with (ROOT/'candidate-build.json').open('x') as f:f.write(json.dumps({'revision':r['revision'],'source_manifest':r['source_after'],'binaries':binaries},indent=2)+'\n')
for name in json.loads((ROOT/'source-files.json').read_text()):
 (ROOT/'candidate'/('after-'+Path(name).name+'.txt')).write_bytes((REPO/name).read_bytes())
print(json.dumps({'status':'pass','binaries':binaries}))
