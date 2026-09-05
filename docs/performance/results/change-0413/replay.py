#!/usr/bin/env python3
"""Replay retained 0413 evidence without benchmark executables or live worktrees."""
import argparse,json,os,subprocess,sys,tempfile
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--repo-root',type=Path,required=True);p.add_argument('--output-dir',type=Path,required=True);a=p.parse_args()
root=Path(__file__).resolve().parent;repo=a.repo_root.resolve();out=a.output_dir.resolve();out.mkdir(parents=True,exist_ok=True)
env=os.environ.copy();env['PYTHONDONTWRITEBYTECODE']='1';journal=[]
def run(label,args):
 argv=[sys.executable,*map(str,args)]
 with (out/(label+'.log')).open('wb') as log:r=subprocess.run(argv,env=env,stdout=log,stderr=subprocess.STDOUT)
 journal.append(dict(label=label,argv=argv,exit_code=r.returncode));(out/'commands.json').write_text(json.dumps(journal,indent=2)+'\n');print(label,r.returncode,flush=True);r.check_returncode()
run('inventory',[root/'verify-artifacts.py'])
run('capture',[root/'verify-capture.py','--root',root/'capture','--repo-root',repo,'--output',out/'capture.json'])
run('schema',[root/'schema-check.py','--root',root/'capture','--repo-root',repo,'--output',out/'schema.json'])
run('recheck-schema',[root/'recheck-schema.py','--root',root/'guard-recheck','--repo-root',repo,'--output',out/'recheck-schema.json'])
run('resources',[root/'verify-resources.py','--root',root/'capture','--recheck',root/'guard-recheck','--output',out/'resources.json'])
parser=repo/'docs/performance/results/change-0412/attribution/attribute.py'
run('profile',[root/'profile/attribute.py','--capture',root/'capture','--parser',parser,'--repo',repo,'--control-repo',repo,'--candidate-repo',repo,'--output',out/'profile.json'])
run('negative',[root/'negative-probes.py','--root',root,'--repo-root',repo,'--verifier',root/'verify-capture.py','--output',out/'negative.json'])
run('profile-negative',[root/'profile/negative-probes.py','--capture',root/'capture','--verifier',root/'profile/attribute.py','--parser',parser,'--repo',repo,'--output',out/'profile-negative.json'])
for name in ['latency','guards','guard-recheck-latency']:
 package=root/name;manifest=json.loads(next(package.glob('*manifest.json')).read_text())
 with tempfile.TemporaryDirectory(prefix='litchi-goal-0413-abba-replay-') as td:
  paths={}
  for item in manifest['artifacts']:
   raw=Path(td)/(item['role']+'.json');raw.write_bytes(subprocess.check_output(['zstd','-q','-dc',str(package/item['path'])]));paths[item['role']]=raw
  dest=out/(name+'-summary.json')
  run(name,[repo/'tools/perf_abba_summary.py',*[paths[k] for k in ['a1','b1','b2','a2']],'--json-out',dest])
  assert json.loads(dest.read_text())==json.loads((package/manifest['summary']['path']).read_text()),name+' summary mismatch'
print('All retained evidence and ABBA summaries replayed')
