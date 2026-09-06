#!/usr/bin/env python3
"""Check local drivers and their real pilot/build bindings before formal capture."""
import hashlib,importlib.util,json,subprocess,sys
from pathlib import Path
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[3]
def module(name,path):
 spec=importlib.util.spec_from_file_location(name,path);value=importlib.util.module_from_spec(spec);spec.loader.exec_module(value);return value
files=sorted(ROOT.glob('*.py'))+sorted((ROOT/'oracle').glob('*.py'))
for path in files:compile(path.read_text(),str(path),'exec')
for name in ['capture.py','profile.py','derive.py','lifecycle.py','replay.py','portable-probes.py','cleanup.py','build-descriptors.py','save-binaries.py','pilot.py']:
 subprocess.run([sys.executable,'-B',str(ROOT/name),'--help'],stdout=subprocess.DEVNULL,check=True)
cap=module('capture439',ROOT/'capture.py');oracle=module('oracle439',ROOT/'oracle/verify-report.py');custody=module('custody439',ROOT/'check.py')
identities=json.loads((ROOT/'after/binary-copies.json').read_text());build=json.loads((ROOT/'checks/routing-build.json').read_text())
assert custody.sources()==build['source_before']==build['source_after']
for shape,count in oracle.SHAPES.items():assert oracle._semantic_digest(count)==oracle.EXPECTED_SOURCE_SEMANTIC[shape]
for mode in ['normal','allocator']:
 for shape in ['tiny','medium','large']:
  report=ROOT/f'pilots/after/aligned/{mode}-{shape}.json'
  cap.verify_report_identity(report,identities[mode],build['revision'])
  oracle.validate_report(report,mode,shape,samples=3,warmups=1)
assert not subprocess.check_output(['git','diff','--name-only','--','crates'],cwd=REPO).strip()
goal_sha=hashlib.sha256((REPO/'docs/GOAL.md').read_bytes()).hexdigest();assert goal_sha=='bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1'
assert subprocess.check_output(['git','rev-parse','HEAD:docs/adr'],cwd=REPO,text=True).strip()=='c950b6c8be822561b498d7bbe87c460873dcbf49'
print(json.dumps({'status':'pass','python_files':len(files),'aligned_pilot_checks':6,'production_files_changed':0,'goal_sha256':goal_sha}))
