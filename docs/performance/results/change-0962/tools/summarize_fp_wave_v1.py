#!/usr/bin/env python3
"""Additive exclusive symbol/source summaries for the complete frozen FP wave."""
import collections,hashlib,json,re
from pathlib import Path
if not __debug__:raise RuntimeError('optimized Python is unsupported')
ROOT=Path('/home/zhuhe/litchi-performance-0962')
def descriptor(p):
 return {'path':str(p),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
def groups(rows,den):
 return [{'key':k,'samples':v[0],'period':v[1],'share_of_qualified_period':v[1]/den} for k,v in sorted(rows.items(),key=lambda x:(-x[1][1],x[0]))]
def add(d,k,r):
 d[k][0]+=r['samples'];d[k][1]+=r['period']
def conserve(rows,n,p):
 assert sum(r[0] for r in rows.values())==n
 assert sum(r[1] for r in rows.values())==p
terminal=json.loads((ROOT/'fp-wave-v1/capture-terminal.json').read_text());assert terminal['status']=='capture_validated' and len(terminal['children'])==6
independent_path=ROOT/'review/fp-wave-v1-independent-cpu-summary-v2.json'
independent=json.loads(independent_path.read_text())['children']
children=[]
for row in terminal['children']:
 d=ROOT/'fp-wave-v1'/f"{row['index']:02d}-{row['case']}";a=d/'offline-decode-v2/analyzer.json';v=json.loads(a.read_text());q=v['denominators']['qualified_ancestor']
 assert v['qualification_gate']['admitted'] and v['qualification_gate']['lost_records']==0 and v['pairing']['exact_pairing_pass']
 other=next(x for x in independent if x['index']==row['index'])
 assert v['whole_process']['samples']==other['sample_count'] and v['whole_process']['period']==other['period_sum']
 assert other['lost_record_lines']==0 and set(map(int,other['cpu_counts'])).issubset({12,13,14,15})
 n,p=q['samples'],q['period'];physical=collections.defaultdict(lambda:[0,0]);source=collections.defaultdict(lambda:[0,0])
 for r in v['qualified_attribution']['exclusive_physical_enclosing_leaf']:
  m=re.fullmatch(r'0x[0-9a-f]+ (.*?) \[([^\]]+)\](?: \((.*):([0-9]+)\))?',r['key']);assert m,r['key']
  add(physical,m[1]+' ['+m[2]+']',r)
 for r in v['qualified_attribution']['exclusive_deepest_self_symbol']:
  m=re.search(r' \((.*):([0-9]+)\)$',r['key'])
  # Basename-only or absent locations cannot be assigned to a source owner.
  key=m[1] if m and m[1].startswith('/') and int(m[2])>0 else '[source-unresolved]'
  add(source,key,r)
 conserve(physical,n,p);conserve(source,n,p)
 children.append({'index':row['index'],'case':row['case'],'analysis':descriptor(a),'whole_process':v['whole_process'],'qualified':q,'reasons':v['reasons'],'physical_enclosing_functions':groups(physical,p),'deepest_source_files':groups(source,p)})
cases=[]
for case in sorted({x['case'] for x in children}):
 cs=[x for x in children if x['case']==case];assert len(cs)==2
 tables={}
 for table in ['physical_enclosing_functions','deepest_source_files']:
  keys={r['key'] for c in cs for r in c[table]};rows=[]
  for key in keys:
   vals=[next((r['share_of_qualified_period'] for r in c[table] if r['key']==key),0.0) for c in cs]
   rows.append({'key':key,'per_capture_shares':vals,'minimum':min(vals),'maximum':max(vals),'mean':sum(vals)/2})
  tables[table]=sorted(rows,key=lambda r:(-r['mean'],r['key']))
 cases.append({'case':case,'children':[x['index'] for x in cs],**tables})
result={'schema':'litchi.performance.0962.fp-wave-summary.v1','status':'passed_conservation','scope':'Six diagnostic-build captures; sampled period shares within strict stage ancestry. No latency/speedup/production/CI claim. Physical enclosing functions are exclusive self including inline callees, not inclusive descendant costs. Deepest-source unresolved remains in denominator. Repeats describe consistency only.','independent_cpu_summary':descriptor(independent_path),'capture_terminal':descriptor(ROOT/'fp-wave-v1/capture-terminal.json'),'helper':descriptor(Path(__file__)),'measured_iterations':1200,'total_samples':sum(x['whole_process']['samples'] for x in children),'qualified_samples':sum(x['qualified']['samples'] for x in children),'whole_period':sum(x['whole_process']['period'] for x in children),'qualified_period':sum(x['qualified']['period'] for x in children),'children':children,'cases':cases}
out=ROOT/'fp-wave-summary-v1.json';encoded=(json.dumps(result,indent=2)+'\n').encode();assert len(encoded)<1048576
with out.open('xb') as f:f.write(encoded)
print(json.dumps({'output':descriptor(out),'qualified_samples':result['qualified_samples'],'whole_samples':result['total_samples']}))
