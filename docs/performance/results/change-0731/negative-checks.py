#!/usr/bin/env python3
"""Actual-analyzer mutations in an isolated copy; original evidence stays intact."""
import contextlib,hashlib,importlib.util,io,json,shutil,tempfile
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('ppt0731',P/'analyze.py');a=importlib.util.module_from_spec(spec);spec.loader.exec_module(a)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,x):p.write_text(json.dumps(x,indent=2)+'\n')
def main():
 checks=[]
 with tempfile.TemporaryDirectory(prefix='litchi-0731-negative-') as t:
  q=Path(t)/'packet';shutil.copytree(P,q,ignore=shutil.ignore_patterns('__pycache__'));a.P=q
  f=json.loads((q/'freeze.json').read_text());write(q/'freeze.json',{k.replace(str(P),str(q),1):v for k,v in f.items()})
  m=json.loads((q/'captures/manifest.json').read_text());m['freeze_sha256']=sha(q/'freeze.json')
  for r in m['runs']:r['command']=[v.replace(str(P),str(q),1) for v in r['command']]
  write(q/'captures/manifest.json',m)
  original=(q/'captures/manifest.json').read_bytes();report=q/m['runs'][0]['output'];original_report=report.read_bytes()
  def run(name,expected):
   accepted=True
   try:
    with contextlib.redirect_stdout(io.StringIO()):a.main()
   except (AssertionError,KeyError,ValueError,OSError):accepted=False
   assert accepted==expected,name;checks.append(dict(name=name,accepted=accepted,expected=expected))
  run('complete packet accepted',True)
  for name,mutate in [('wrong output identity',lambda x:x['samples'][0].__setitem__('output_sha256','0'*64)),('false semantic oracle',lambda x:x['samples'][0]['oracle'].__setitem__('semantic_reopen_ok',False)),('missing negative control',lambda x:x['oracle_controls'].pop()),('zero native time',lambda x:x['samples'][0]['phase_ns'].__setitem__('whole_ns',0))]:
   x=json.loads(original_report);mutate(x);write(report,x);changed=json.loads(original);changed['runs'][0]['sha256']=sha(report);write(q/'captures/manifest.json',changed);run(name,False);report.write_bytes(original_report);(q/'captures/manifest.json').write_bytes(original)
  for name,mutate in [('missing process',lambda x:x['runs'].pop()),('wrong command',lambda x:x['runs'][0]['command'].__setitem__(-1,'99')),('wrong profile hash',lambda x:x['runs'][2].__setitem__('callgrind_sha256','0'*64))]:
   x=json.loads(original);mutate(x);write(q/'captures/manifest.json',x);run(name,False);(q/'captures/manifest.json').write_bytes(original)
  cg=q/'captures/callgrind-0.callgrind';old=cg.read_bytes();text=old.decode();import re
  text=re.sub(r'^summary: (\d+)',lambda m:'summary: '+str(int(m[1])+1),text,flags=re.M);cg.write_text(text);x=json.loads(original);x['runs'][2]['callgrind_sha256']=sha(cg);write(q/'captures/manifest.json',x);run('profile arithmetic with updated digest',False);cg.write_bytes(old);(q/'captures/manifest.json').write_bytes(original)
  run('restored packet accepted',True)
 (P/'negative-checks.json').write_text(json.dumps(dict(analyzer_sha256=sha(P/'analyze.py'),script_sha256=sha(Path(__file__)),checks=checks),indent=2)+'\n');print('PASS',len(checks),'actual-analyzer controls')
if __name__=='__main__':main()
