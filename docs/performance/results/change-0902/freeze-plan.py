"""Bind the predeclared protocol to completed build artifacts exactly once."""
import hashlib,json,pathlib
p=pathlib.Path(__file__).resolve().parent;root=p.parents[3];base=p.parent/'change-0894'
def d(path):
 path=path.resolve();return {'path':str(path),'bytes':path.stat().st_size,'sha256':hashlib.sha256(path.read_bytes()).hexdigest()}
plan=json.loads((p/'plan-template.json').read_text())
plan['source'].update(baseline_manifest=d(base/'source.json'),candidate_manifest=d(p/'candidate-source-v2.json'),candidate_files=['crates/xml-minifier/src/audit.rs','crates/xml-minifier/tests/source_replacement.rs'])
for role in ['baseline','candidate']:
 for kind in ['native','allocator']:
  binary=(base if role=='baseline' else p)/'binaries'/(role+'-'+kind)
  receipt=base/(kind+'-build.json') if role=='baseline' else p/('candidate-'+kind+'-build-run.json')
  r=json.loads(receipt.read_text());assert r['exit_code']==0
  if role=='candidate':assert r['inputs_unchanged'] and r['inputs_before']==r['inputs_after']
  plan['roles'][role][kind]={**d(binary),'build_receipt':d(receipt)}
plan['semantic_validator']=d(p.parent/'change-0705'/'analyze.py')
plan['method']=d(p/'METHOD.md');plan['environment_receipt']=d(p/'environment.json')
plan['helpers']={f:d(p/f) for f in ['run.py','analyze.py','command.py','quality-and-build.py','freeze-plan.py']}
plan['quality_lock']=d(p/'quality-Cargo.lock')
assert plan['cpu']==12 and plan['source']['candidate_files']==['crates/xml-minifier/src/audit.rs','crates/xml-minifier/tests/source_replacement.rs']
with (p/'plan.json').open('x') as f:json.dump(plan,f,indent=2,sort_keys=True);f.write('\n')
print('frozen',d(p/'plan.json'))
