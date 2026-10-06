#!/usr/bin/env python3
import collections,hashlib,json,re
from pathlib import Path
R=Path('/home/zhuhe/litchi-performance-0942')
manifest=json.loads((R/'manifests/cfb-profile-manifest-v5.json').read_text())
header=re.compile(r'^\S+\s+\d+/\d+\s+\d+\.\d+:\s+(\d+)\s+cycles:u:\s*$')
frame=re.compile(r'^\s*[0-9a-f]+\s+(.+)\s+\((.+)\)\s*$')
marker='doc_retention_probe_0942_cfb_profile::profiled_measured_case'
rows=[]
paths=sorted((R/'capture/formal-v5/children').glob('*/receipt.json'))+sorted((R/'preflight/real-v5').glob('*/receipt.json'))
assert len(paths)==17
for p in paths:
 receipt=json.loads(p.read_text());assert receipt['binary']==manifest['build']['binary']
 assert receipt['environment']==manifest['environment']
 for key in ['record_terminal','script_terminal']:
  t=json.loads((R/receipt[key]['path']).read_text())
  assert t['exit_code']==0 and t['timed_out'] is False
  field='record_argv' if key=='record_terminal' else 'script_argv'
  assert t[field]==receipt['record_command' if key=='record_terminal' else 'script_command']
 script=R/receipt['perf_script']['path'];raw=script.read_bytes()
 assert hashlib.sha256(raw).hexdigest()==receipt['perf_script']['sha256']
 totals=collections.Counter();owners=collections.Counter();owners_n=collections.Counter()
 for block in raw.decode().strip().split('\n\n'):
  lines=block.splitlines();m=header.fullmatch(lines[0]);assert m,lines[0]
  weight=int(m[1]);assert weight>0
  stack=[]
  for line in lines[1:]:
   if not line.strip():continue
   f=frame.fullmatch(line);assert f,line
   stack.append((f[1].strip(),f[2]))
  totals['all_samples']+=1;totals['all_period']+=weight
  positions=[i for i,(s,d) in enumerate(stack) if s==marker]
  if not positions:continue
  assert len(positions)==1
  i=positions[0];assert stack[i][1]==manifest['build']['binary']['path']
  totals['marker_samples']+=1;totals['marker_period']+=weight
  owner='other'
  for sym,dso in stack[:i]:
   for needle,key in [('litchi_cfb','cfb'),('litchi_ole_common','ole-common'),('litchi_doc','doc'),('litchi_core','core')]:
    if needle in sym:owner=key;break
   if owner!='other':break
  owners[owner]+=weight;owners_n[owner]+=1
 q=receipt['profile'];assert totals['all_samples']==q['sample_accounting']['all']['samples']
 assert totals['all_period']==q['sample_accounting']['all']['period']
 assert totals['marker_samples']==q['marker']['samples'] and totals['marker_period']==q['marker']['period']
 for owner,value in q['nearest_physical_owner_buckets'].items():
  assert value=={'period':owners[owner],'samples':owners_n[owner]},(p,owner,value)
 rows.append({'id':receipt['id'],**totals,'owner_periods':dict(owners),'raw_sha256':hashlib.sha256(raw).hexdigest()})
out={'schema':'litchi.performance.0942.root-independent-count.v1','status':'passed','script_uses_capture_parser':False,'binary_receipt_and_terminal_argv_checked':True,'children':rows}
p=R/'receipts/root-independent-count-v1.json'
with p.open('x') as f:json.dump(out,f,indent=2,sort_keys=True);f.write('\n')
print('Independent counts, owner buckets, binary receipts and terminal argv agree for',len(rows),'children')
