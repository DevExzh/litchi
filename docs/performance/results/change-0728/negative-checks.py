#!/usr/bin/env python3
"""Exercise actual analysis with temporary mutations; leave original evidence intact."""
import contextlib,hashlib,importlib.util,io,json,shutil,tempfile
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('baseline',P/'analyze.py');mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
checks=[]
with tempfile.TemporaryDirectory(prefix='litchi-0728-negative-') as temp:
 q=Path(temp);shutil.copytree(P/'captures',q/'captures')
 for name in ['plan.json','oracle-contract.json','cases.json','freeze.json','cleanup.json']:
  if (P/name).exists():shutil.copy2(P/name,q/name)
 mod.P=q;manifest=q/'captures/manifest.json';original=manifest.read_bytes()
 def run(label,expected):
  okay=True
  try:
   with contextlib.redirect_stdout(io.StringIO()):mod.main()
  except (AssertionError,ValueError,KeyError,StopIteration):okay=False
  assert okay==expected,label;checks.append(dict(name=label,accepted=okay,expected=expected))
 run('complete capture accepted',True)
 m=json.loads(original);target=q/'captures'/m['runs'][0]['output'];raw=target.read_bytes()
 target.write_bytes(raw+b' ');run('raw hash mismatch rejected',False);target.write_bytes(raw)
 def mutate_raw(label,fn):
  x=json.loads(raw);fn(x);target.write_text(json.dumps(x));m=json.loads(original);m['runs'][0]['sha256']=sha(target);manifest.write_text(json.dumps(m));run(label,False);target.write_bytes(raw);manifest.write_bytes(original)
 mutate_raw('false semantic oracle despite aggregate true rejected',lambda x:x['samples'][0]['oracle'].update(semantic_reopen_ok=False))
 mutate_raw('physical-only length proof rejected',lambda x:x['changed_length_proof'].update(logical_stream_length_change_proven=False))
 mutate_raw('missing corruption control rejected',lambda x:x['oracle_controls'].pop())
 mutate_raw('semantic witness corruption rejected',lambda x:x['samples'][0]['oracle'].update(semantic_witness={}))
 mutate_raw('format length witness removed rejected',lambda x:x['changed_length_proof'].update(format_specific_semantic_length_proven=False))
 mutate_raw('metadata inventory claim removed rejected',lambda x:x.update(directory_metadata_fields=[]))
 mutate_raw('source identity corruption rejected',lambda x:x.update(source_sha256='0'*64))
 mutate_raw('policy scope relabel rejected',lambda x:x.update(policy_applied=True))
 mutate_raw('stream corruption rejected',lambda x:x['samples'][0]['output_inventory']['streams'][0].update(sha256='0'*64))
 for label,fn in [
  ('duplicate process rejected',lambda m:m['runs'].__setitem__(1,m['runs'][0])),
  ('command count corruption rejected',lambda m:m['runs'][0]['command'].__setitem__(-1,'99')),
  ('missing end binding rejected',lambda m:m.update(bindings_end={})),
  ('missing process rejected',lambda m:m['runs'].pop()),
 ]:
  m=json.loads(original);fn(m);manifest.write_text(json.dumps(m));run(label,False);manifest.write_bytes(original)
receipt=dict(analyzer_sha256=sha(P/'analyze.py'),script_sha256=sha(Path(__file__)),checks=checks)
(P/'negative-checks.json').write_text(json.dumps(receipt,indent=2)+'\n');print('PASS',len(checks),'actual-analyzer controls')
