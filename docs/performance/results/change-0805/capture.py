"""Root-only fresh-process direct helper captures, no retries or overwrite."""
import sys,time,subprocess
import custody as c
lane=sys.argv[1];assert lane in ['native','profiles']
p=c.P;plan=c.read(p/'plan.json');build=c.read(p/'build/build.json');binary=build['binary'];assert c.artifact(binary['path'])==binary
cases=c.read(p/'cases.json');assert isinstance(cases,list) and len(cases)==plan['case_count'];ids=[r['id'] for r in cases];assert len(set(ids))==len(ids)
source=c.source();assert source==c.read(p/'source.json');inputs=c.read(build['inputs']['path'])
for n,h in inputs['probe'].items():assert c.sha(p/'probe-src'/n)==h
out=p/lane;assert not out.exists();out.mkdir();c.write(out/'source.json',source);rows=[]
settings=plan[lane];orders=settings['orders']
for block,order in enumerate(orders):
 for case in ids:
  for mode in plan['modes']:
   for leg in order:
    stem=f'{block}-{case}-{mode}-{leg}';report=out/f'{stem}.json';log=out/f'{stem}.log';raw=out/f'{stem}.callgrind';owner=plan['profiles']['owner_pattern'].format(leg=leg,mode=mode)
    args=[binary['path'],'--leg',leg,'--case',case,'--mode',mode,'--samples',str(settings['samples']),'--warmup',str(settings['warmup']),'--iterations',str(settings['iterations']),'--output',str(report)]
    command=['taskset','-c',str(plan['cpu'])]
    if lane=='profiles':command+=['valgrind','--tool=callgrind','--branch-sim=yes','--collect-atstart=no','--toggle-collect='+owner,'--zero-before='+owner,'--dump-after='+owner,'--callgrind-out-file='+str(raw)]
    command+=args;started=time.time()
    with log.open('w') as f:r=subprocess.run(command,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
    row={'block':block,'case':case,'mode':mode,'leg':leg,'command':command,'binary':binary,'exit_code':r.returncode,'started':started,'ended':time.time(),'log':c.artifact(log)}
    if report.exists():row['report']=c.artifact(report)
    if lane=='profiles':row['artifacts']={f.name:c.artifact(f) for f in out.glob(stem+'.*') if f.is_file()}
    rows.append(row);c.write(out/'receipts.json',rows)
    assert r.returncode==0 and report.exists(),log
    if lane=='profiles':assert (out/f'{stem}.callgrind.1').exists() and not (out/f'{stem}.callgrind.2').exists()
    print(stem,'PASS',flush=True)
 assert c.source()==source
assert c.artifact(binary['path'])==binary
for n,h in inputs['probe'].items():assert c.sha(p/'probe-src'/n)==h
c.write(out/'complete.json',{'children':len(rows),'source':c.artifact(out/'source.json'),'receipts':c.artifact(out/'receipts.json')})
