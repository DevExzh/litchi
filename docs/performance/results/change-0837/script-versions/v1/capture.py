"""Serial matched runs; immutable baseline-produced source fixtures."""
import driver as d

def run(stage,case,label,samples,warmup):
 binary=d.read(d.P/f'build-{stage}.json')['binary'];assert d.desc(binary['path'])==binary
 output=d.P/(label+'.json');assert not output.exists()
 argv=['taskset','-c','12',binary['path'],'run',case,str(samples),str(warmup),str(d.P/'fixtures'),str(output)]
 assert d.run(label,argv)==0,label
 return dict(stage=stage,case=case,samples=samples,warmup=warmup,report=d.desc(output),receipt=d.desc(d.P/f'commands/{label}/receipt.json'),binary=binary)

def main(kind):
 plan=d.read(d.P/'plan.json');assert d.source()==d.read(d.P/'freeze-candidate.json')['source']
 rows=[]
 if kind=='qualification':
  for case in plan['cases']:
   for stage in ['baseline','candidate']:rows.append(run(stage,case,f'qualification-{stage}-{case}',1,0))
 else:
  assert kind=='native'
  admission=d.read(d.P/'capture-admission.json');assert admission['status']=='pass'
  for name,binding in admission['inputs'].items():assert d.desc(d.P/name)==binding,name
  for block in range(plan['blocks']):
   cases=plan['cases'] if block%2==0 else list(reversed(plan['cases']))
   stages=['baseline','candidate'] if block%2==0 else ['candidate','baseline']
   for case in cases:
    for stage in stages:
     rows.append(dict(block=block,**run(stage,case,f'native-{block:02}-{stage}-{case}',plan['samples'],plan['warmup'])))
 d.write(d.P/(kind+'.json'),dict(status='pass',rows=rows,reports=len(rows),samples=sum(r['samples'] for r in rows)))
 print(kind,'PASS',len(rows),'reports',flush=True)

if __name__=='__main__':
 import sys
 main(sys.argv[1])
