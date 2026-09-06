#!/usr/bin/env python3
"""Compare retained standalone Clippy debt; reject diagnostics in the new module."""
import argparse,collections,gzip,json,re
from pathlib import Path
ROOT=Path(__file__).resolve().parent

def raw(name):
 p=(ROOT/'lint-baseline'/'candidate-harness-strict-fixed.log') if name=='before-harness-strict' else ROOT/'checks'/(name+'.log')
 return p.read_text() if p.exists() else gzip.decompress(Path(str(p)+'.gz').read_bytes()).decode()

def diagnostics(name):
 text=raw(name); items=[]
 for block in re.split(r'(?m)(?=^error(?:\[[^]]+\])?:)',text):
  if not block.startswith('error'):continue
  heading=block.splitlines()[0]
  if heading.startswith('error: could not compile') or heading.startswith('error: aborting'):continue
  location=re.search(r'(?m)^\s*--> ([^:\n]+):[0-9]+:[0-9]+',block)
  path=location.group(1) if location else '<no-location>'
  heading=re.sub(r'\([0-9]+/[0-9]+\)','(count/limit)',heading)
  items.append((path,heading))
 return collections.Counter(items)

def derive():
 before=diagnostics('before-harness-strict');after=diagnostics('candidate-harness-strict')
 assert before,'missing baseline debt'
 new=after-before
 assert not new,dict(new)
 assert not any('opc_part_add.rs' in path for path,_ in after)
 return {'change':446,'status':'pass','before_rendered_diagnostics':sum(before.values()),'after_rendered_diagnostics':sum(after.values()),'new_diagnostics':0,'before':[{'path':p,'message':m,'count':n} for (p,m),n in sorted(before.items())],'after':[{'path':p,'message':m,'count':n} for (p,m),n in sorted(after.items())],'scope':'existing standalone harness Clippy debt; no diagnostic admitted in the new Part-addition module; numeric function-line counts normalized while diagnostic multiplicities remain checked'}

if __name__=='__main__':
 parser=argparse.ArgumentParser();parser.add_argument('--check',action='store_true');args=parser.parse_args();value=derive();target=ROOT/'lint-comparison.json'
 if args.check:assert json.loads(target.read_text())==value
 else:
  with target.open('x') as stream:stream.write(json.dumps(value,indent=2)+'\n')
 print(json.dumps({'status':value['status'],'new_diagnostics':value['new_diagnostics']}))
