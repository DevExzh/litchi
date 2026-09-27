"""Root-only serial 0779 capture. No overwrite, no retry on variation."""
import os
from pathlib import Path
import subprocess
import sys
import time
import custody as c

lane=sys.argv[1]
assert lane in ('qualification','native','allocation','small-native','small-allocation')
small = lane.startswith('small-')
mode = lane.removeprefix('small-')
plan=c.read(c.P/('small-control-plan.json' if small else 'plan.json'));out=c.P/lane;assert not out.exists();out.mkdir()
legs=['before'] if lane=='qualification' else ['before','after']
builds={leg:c.read(c.P/f'build-{leg}/build.json') for leg in legs}
kind='allocation' if mode in ('qualification','allocation') else 'native'
for build in builds.values():
 for b in build['binaries'].values():assert c.artifact(b['path'])==b
frozen=c.source();c.write(out/'source.json',frozen)
rows=[]
blocks=1 if lane=='qualification' else plan[mode]['blocks']
for block in range(blocks):
 for corpus in plan['corpora']:
  source=c.ROOT/corpus['source'];assert c.sha(source)==corpus['sha256']
  for phase in plan['phases']:
   order=['before'] if lane=='qualification' else plan['native']['orders'][block]
   for leg in order:
    stem=f'{block}-{corpus["id"]}-{phase}-{leg}';report=out/f'{stem}.json';rss=out/f'{stem}.rss';log=out/f'{stem}.log'
    samples=1 if lane=='qualification' else plan[mode]['samples'];warmup=0 if lane=='qualification' else plan[mode]['warmup']
    binary=builds[leg]['binaries'][kind]
    command=['/usr/bin/time','-f','%M','-o',str(rss),'taskset','-c',str(plan['cpu']),binary['path'],'--source',str(source),'--sheet',corpus['sheet'],'--phase',phase,'--samples',str(samples),'--warmup',str(warmup),'--output',str(report)]
    start=time.time()
    with log.open('w') as f:r=subprocess.run(command,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
    row={'lane':lane,'block':block,'corpus':corpus['id'],'phase':phase,'leg':leg,'command':command,'exit_code':r.returncode,'started':start,'ended':time.time(),'binary':binary,'log':c.artifact(log),'rss':c.artifact(rss)}
    if report.exists():row['report']=c.artifact(report)
    rows.append(row);c.write(out/'receipts.json',rows)
    assert r.returncode==0,row
    data=c.read(report);assert len(data['samples'])==samples and data['source']['sha256']==corpus['sha256']
    assert c.sha(source)==corpus['sha256'];assert c.source()==frozen
    print(stem,'passed',flush=True)
c.write(out/'complete.json',{'children':len(rows),'receipts':c.artifact(out/'receipts.json'),'source':c.artifact(out/'source.json')})
