#!/usr/bin/env python3
import datetime, hashlib, json, subprocess, sys
from pathlib import Path
p=Path(__file__).resolve().parent
root=p.parent
stage=sys.argv[1]
d=Path('/tmp/litchi-goal-0510/target/release/deps')
out=Path('/tmp/litchi-goal-0510')/stage/'export-guardrail'
out.parent.mkdir(exist_ok=True)
sha=lambda f:hashlib.sha256(f.read_bytes()).hexdigest()
cmd=['rustc','--edition=2024','-O','-C','debuginfo=0',str(p/'probe.rs'),'-L','dependency='+str(d),'-o',str(out)]
deps={}
for crate in ['litchi_odt','litchi_core','sha2','soapberry_zip']:
 if stage=='before':
  matches=list(d.glob('lib'+crate+'-*.rlib'))
  if crate=='sha2':matches=[f for f in matches if 'sha2-0.11.0/' in f.with_name(f.name.removeprefix('lib')).with_suffix('.d').read_text()]
  assert len(matches)==1,(crate,matches)
  path=matches[0]
 else:
  original=json.loads((p/'before-build.json').read_text())['dependency_artifact_sha256']
  path=next(Path(f) for f in original if Path(f).name.startswith('lib'+crate+'-'))
 cmd += ['--extern',crate+'='+str(path)]
 deps[str(path)]=sha(path)
with (p/f'{stage}-build.log').open('x') as log:r=subprocess.run(cmd,stdout=log,stderr=subprocess.STDOUT)
manifest=root/'before/source-manifest.json' if stage=='before' else root/'source-manifest.json'
v={'command':cmd,'exit_code':r.returncode,'recorded_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'dependency_artifact_sha256':deps,'probe_source_sha256':sha(p/'probe.rs'),'library_source_manifest_sha256':sha(manifest)}
if r.returncode==0:v['binary_sha256']=sha(out)
(p/f'{stage}-build.json').write_text(json.dumps(v,indent=2)+'\n')
print(stage,r.returncode);assert r.returncode==0
