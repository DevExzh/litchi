#!/usr/bin/env python3
"""Exercise the frozen analyzer on temporary raw copies, never replacing captures."""
import contextlib,hashlib,importlib.util,io,json,shutil,tempfile
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('ablation',P/'analyze.py');mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
checks=[]
with tempfile.TemporaryDirectory(prefix='litchi-0724-neg-') as temp:
 q=Path(temp);shutil.copytree(P/'captures',q/'captures')
 for name in ['plan.json','freeze.json']:shutil.copy2(P/name,q/name)
 mod.P=q
 manifest=q/'captures/manifest.json';original=manifest.read_bytes()
 def run(label,expected):
  okay=True
  try:
   with contextlib.redirect_stdout(io.StringIO()):mod.main()
  except (AssertionError,ValueError,KeyError):okay=False
  assert okay==expected,label;checks.append(dict(name=label,accepted=okay,expected=expected))
 run('positive complete capture',True)
 m=json.loads(original);target=q/'captures'/m['runs'][0]['output'];raw=target.read_bytes()
 target.write_bytes(raw+b' ');run('raw hash mismatch',False);target.write_bytes(raw)
 x=json.loads(raw);x['records'][0]['queries'][0]['outcome']['status']='missing';target.write_text(json.dumps(x));m['runs'][0]['sha256']=sha(target);manifest.write_text(json.dumps(m));run('outcome altered with matching raw hash',False);target.write_bytes(raw);manifest.write_bytes(original)
 m=json.loads(original);m['runs'][1]=m['runs'][0];manifest.write_text(json.dumps(m));run('duplicate run identity',False);manifest.write_bytes(original)
 m=json.loads(original);r=next(r for r in m['runs'] if r['lane']=='repeat');target=q/'captures'/r['output'];raw=target.read_bytes();target.write_bytes(raw.replace(b'found=50000',b'found=49999'));r['sha256']=sha(target);manifest.write_text(json.dumps(m));run('repeat count altered with matching raw hash',False);target.write_bytes(raw);manifest.write_bytes(original)
 m=json.loads(original);m['bindings_end']={};manifest.write_text(json.dumps(m));run('end binding removed',False)
receipt=dict(analyzer_sha256=sha(P/'analyze.py'),script_sha256=sha(Path(__file__)),checks=checks)
(P/'negative-checks.json').write_text(json.dumps(receipt,indent=2)+'\n');print('PASS',len(checks),'actual-analyzer controls')
