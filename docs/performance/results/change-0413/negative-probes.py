#!/usr/bin/env python3
"""Mutation-specific rejection checks against private copies of 0413 evidence."""
import argparse,json,os,shutil,subprocess,sys,tempfile
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--root',type=Path,required=True);p.add_argument('--repo-root',type=Path,required=True);p.add_argument('--verifier',type=Path,required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
def load(p):return json.loads(p.read_text() if p.exists() else subprocess.check_output(['zstd','-q','-dc',str(p)+'.zst']))
results=[]
for name,marker in [('duplicate-order','sample order is not a complete permutation'),('opaque-read','frozen locality counter mismatch: opaque_payload_read_bytes'),('plain-observer','plain/eager result fabricated source observer'),('bad-oracle','output oracle mismatch'),('bad-build','journal binary hash mismatch'),('missing-build','missing candidate build identity sidecar'),('observer-version','observer scope mismatch')]:
 with tempfile.TemporaryDirectory(prefix='litchi-goal-0413-negative-') as td:
  root=Path(td)/'bundle';shutil.copytree(a.root,root)
  if name=='missing-build':(root/'checks/candidate-build-identity.json').unlink()
  elif name=='bad-build':
   path=root/'checks/candidate-build-identity.json';r=load(path);r['binaries']['litchi-perf-baseline']['sha256']='0'*64;path.write_text(json.dumps(r))
  else:
   path=root/'capture/normal-1.json';r=load(path)
   if name=='duplicate-order':r['results'][0]['elapsed_ns']['sample_order'][1]=r['results'][0]['elapsed_ns']['sample_order'][0]
   elif name=='bad-oracle':r['results'][0]['output_sha256']='0'*64
   elif name=='plain-observer':next(v for v in r['results'] if v['case']=='xls_owned_source_open')['source']={}
   elif name=='observer-version':next(v for v in r['results'] if v['case']=='xls_source_backed_open')['source']['xls']['source_counter_scope']='v1'
   else:next(v for v in r['results'] if v['case']=='xls_source_backed_open')['source']['xls']['opaque_payload_read_bytes'][0]=1
   path.write_text(json.dumps(r));Path(str(path)+".zst").unlink(missing_ok=True)
  env=os.environ.copy();env['PYTHONDONTWRITEBYTECODE']='1'
  run=subprocess.run([sys.executable,str(a.verifier.resolve()),'--root',str(root/'capture'),'--repo-root',str(a.repo_root.resolve()),'--output',str(Path(td)/'result.json')],env=env,capture_output=True,text=True)
  assert run.returncode==2 and marker in run.stderr,(name,run.returncode,run.stderr)
  results.append(dict(mutation=name,exit_code=run.returncode,expected_diagnostic=marker,status='passed'))
a.output.write_text(json.dumps(dict(status='passed',probes=results),indent=2)+'\n');print('Passed 7 mutation-specific capture rejection probes')
