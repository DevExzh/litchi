#!/usr/bin/env python3
"""Exercise actual analysis with temporary raw mutations; never edit captures."""
import contextlib,hashlib,importlib.util,io,json,shutil,tempfile
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('replication',P/'analyze.py');mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
checks=[]
with tempfile.TemporaryDirectory(prefix='litchi-0727-negative-') as temp:
 q=Path(temp);shutil.copytree(P/'captures',q/'captures')
 for name in ['plan.json','freeze.json','cleanup.json']:
  if (P/name).exists():shutil.copy2(P/name,q/name)
 mod.P=q;manifest=q/'captures/manifest.json';original=manifest.read_bytes();plan=json.loads((q/'plan.json').read_text())
 def run(label,expected):
  okay=True
  try:
   with contextlib.redirect_stdout(io.StringIO()):mod.main()
  except (AssertionError,ValueError,KeyError,StopIteration):okay=False
  assert okay==expected,label;checks.append(dict(name=label,accepted=okay,expected=expected))
 run('complete capture accepted',True)
 m=json.loads(original);target=q/'captures'/m['runs'][0]['output'];raw=target.read_bytes()
 target.write_bytes(raw+b' ');run('raw hash mismatch rejected',False);target.write_bytes(raw)
 x=json.loads(raw);x['records'][0]['queries'][0]['outcome']['status']='missing';target.write_text(json.dumps(x));m['runs'][0]['sha256']=sha(target);manifest.write_text(json.dumps(m));run('outcome corruption with matching hash rejected',False);target.write_bytes(raw);manifest.write_bytes(original)
 m=json.loads(original);m['runs'][1]=m['runs'][0];manifest.write_text(json.dumps(m));run('duplicate process rejected',False);manifest.write_bytes(original)
 m=json.loads(original);m['runs'][0]['command'][-1]='99';manifest.write_text(json.dumps(m));run('command sample count corruption rejected',False);manifest.write_bytes(original)
 m=json.loads(original);m['runs'][0]['phase']='candidate';manifest.write_text(json.dumps(m));run('phase relabel rejected',False);manifest.write_bytes(original)
 m=json.loads(original);m['bindings_end']={};manifest.write_text(json.dumps(m));run('end binding removed rejected',False);manifest.write_bytes(original)
 for label,metric,a,b,want in [('warm exception','q8',100,108,True),('warm failure','q8',100,120,False),('workflow no exception','open-plus-eight',100,108,False)]:
  got=mod.gate(metric,a,b,plan)['pass_'];assert got==want;checks.append(dict(name=label,accepted=got,expected=want))
receipt=dict(analyzer_sha256=sha(P/'analyze.py'),script_sha256=sha(Path(__file__)),checks=checks)
(P/'negative-checks.json').write_text(json.dumps(receipt,indent=2)+'\n');print('PASS',len(checks),'actual-analyzer controls')
