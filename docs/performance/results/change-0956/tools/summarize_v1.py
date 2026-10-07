"""Reduce a completed full comparison to bounded review metrics, never new statistics."""
import hashlib,json,pathlib,sys
source,dest=map(pathlib.Path,sys.argv[1:])
r=json.loads(source.read_text())
assert r['status']=='completed' and r['child_count']==144
out={'schema':'litchi.performance.0956.bounded-comparison.v1','status':'completed','full_comparison':{'path':str(source.resolve()),'bytes':source.stat().st_size,'sha256':hashlib.sha256(source.read_bytes()).hexdigest()},'plan':r['plan'],'packet':r['packet'],'metrics':r['metrics'],'adoption':r['adoption'],'groups':[]}
for g in r['groups']:
 d={k:g[k] for k in ('comparison_id','case_id','gate_summary','interpretation')}
 d['metrics']={}
 for name,m in g['metrics'].items():
  d['metrics'][name]={}
  for q,v in m['quantiles'].items():
   d['metrics'][name][q]={k:x for k,x in v.items() if k not in ('block_ratios','process_pair_ratios')}
   d['metrics'][name][q]['block_points']=[b['block_geometric_mean'] for b in v['block_ratios']]
 out['groups'].append(d)
data=(json.dumps(out,indent=2,sort_keys=True)+'\n').encode()
assert len(data)<1048576
with dest.open('xb') as f:f.write(data)
print(json.dumps({'path':str(dest),'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()}))
