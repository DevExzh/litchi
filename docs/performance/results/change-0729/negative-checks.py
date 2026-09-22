#!/usr/bin/env python3
"""Exercise actual analyzer acceptance without changing original captures."""
import contextlib,hashlib,importlib.util,io,json,shutil,tempfile
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('attribution',P/'analyze.py');mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
checks=[]
with tempfile.TemporaryDirectory(prefix='litchi-0729-negative-') as temp:
 q=Path(temp);shutil.copytree(P/'captures',q/'captures')
 for name in ['plan.json','cases.json','oracle-contract.json','freeze.json','cleanup.json']:
  if (P/name).exists():shutil.copy2(P/name,q/name)
 mod.P=q;manifest=q/'captures/manifest.json';original=manifest.read_bytes()
 def run(label,expected):
  okay=True
  try:
   with contextlib.redirect_stdout(io.StringIO()):mod.main()
  except (AssertionError,ValueError,KeyError,StopIteration):okay=False
  assert okay==expected,label;checks.append(dict(name=label,accepted=okay,expected=expected))
 run('complete capture accepted',True)
 m=json.loads(original);target=q/'captures'/m['runs'][0]['output'];raw=target.read_bytes();target.write_bytes(raw+b' ');run('raw hash mismatch rejected',False);target.write_bytes(raw)
 def mutate_raw(label,route,fn):
  m=json.loads(original);r=next(r for r in m['runs'] if r['route']==route);target=q/'captures'/r['output'];raw=target.read_bytes();x=json.loads(raw);fn(x);target.write_text(json.dumps(x));r['sha256']=sha(target);manifest.write_text(json.dumps(m));run(label,False);target.write_bytes(raw);manifest.write_bytes(original)
 mutate_raw('false semantic oracle rejected','ordinary-opaque',lambda x:x['samples'][0]['oracle'].update(semantic_reopen_ok=False))
 mutate_raw('missing corruption control rejected','ordinary-opaque',lambda x:x['oracle_controls'].pop())
 mutate_raw('source identity corruption rejected','ordinary-opaque',lambda x:x.update(source_sha256='0'*64))
 mutate_raw('outer residual corruption rejected','ordinary-split',lambda x:x['samples'][0]['split'].update(whole_residual_ns=-1))
 mutate_raw('false event balance rejected','profiled-clock',lambda x:x['samples'][0]['diagnostics']['commit'].update(balanced=False))
 mutate_raw('event outcome corruption rejected','profiled-clock',lambda x:x['samples'][0]['diagnostics']['commit']['events'][1].update(outcome='error'))
 mutate_raw('event order corruption rejected','profiled-clock',lambda x:x['samples'][0]['diagnostics']['commit']['events'][0].update(phase='Patch'))
 mutate_raw('span arithmetic corruption rejected','profiled-clock',lambda x:x['samples'][0]['diagnostics']['commit']['spans'][0].update(duration_ns=0))
 mutate_raw('observer clock on empty route rejected','profiled-empty',lambda x:x['samples'][0].update(observer_clock_control_ns=1))
 for label,fn in [
  ('duplicate process rejected',lambda m:m['runs'].__setitem__(1,m['runs'][0])),
  ('command count corruption rejected',lambda m:m['runs'][0]['command'].__setitem__(-1,'99')),
  ('missing end binding rejected',lambda m:m.update(bindings_end={})),
  ('missing process rejected',lambda m:m['runs'].pop()),
 ]:
  m=json.loads(original);fn(m);manifest.write_text(json.dumps(m));run(label,False);manifest.write_bytes(original)
receipt=dict(analyzer_sha256=sha(P/'analyze.py'),script_sha256=sha(Path(__file__)),checks=checks)
(P/'negative-checks.json').write_text(json.dumps(receipt,indent=2)+'\n');print('PASS',len(checks),'actual-analyzer controls')
