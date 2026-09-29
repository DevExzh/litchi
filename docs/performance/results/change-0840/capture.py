"""Serial matched capture; all files exclusive, failed children retained."""
import driver as d
import subprocess,sys,time

def one(lane,leg,case,block=0):
 plan=d.read(d.P/'plan.json');assert case in plan['cases'];d.check()
 cfg=plan['qualification' if lane=='allocation-preflight' else lane];kind='allocation' if lane in ['allocation','allocation-preflight'] else 'native'
 build=d.read(d.P/f'build-{leg}-{kind}.json');binary=build['binary'];assert d.desc(binary['path'])==binary
 folder=d.P/'runs'/lane/f"{block:02}-{case}-{'leg-a' if leg=='before' else 'leg-b'}";folder.mkdir(parents=True,exist_ok=False)
 args=['taskset','-c',str(plan['native']['cpu']),'/usr/bin/time','-f','%M %R %F','-o',str(folder/'time.txt'),binary['path'],'--case',case,'--samples',str(cfg['samples']),'--warmup',str(cfg['warmup']),'--output',str(folder/'report.json')]
 if lane=='observer':args+=['--observe']
 if lane in ['qualification','observer']:args+=['--artifact',str(folder/'output.cfb')]
 start=time.time();d.write(folder/'started.json',dict(lane=lane,leg=leg,case=case,block=block,argv=args,started_unix=start,binary=binary,build=d.desc(d.P/f'build-{leg}-{kind}.json')))
 with (folder/'stdout.log').open('xb') as out,(folder/'stderr.log').open('xb') as err:
  code=subprocess.run(args,cwd=d.ROOT,env=d.env(),stdout=out,stderr=err).returncode
 d.write(folder/'receipt.json',dict(exit_code=code,finished_unix=time.time(),files=[d.desc(p) for p in sorted(folder.iterdir()) if p.is_file()]))
 assert code==0,(folder,code)
 print(lane,leg,case,block,'PASS',flush=True)
 return folder

if __name__=='__main__':
 lane=sys.argv[1];plan=d.read(d.P/'plan.json');source=d.source()
 if lane in ['qualification','observer','allocation-preflight']:
  for case in plan['cases']:one(lane,sys.argv[2],case)
 else:
  assert lane in ['native','allocation']
  admission=d.read(d.P/'admission.json')
  for x in admission['files']:assert d.desc(x['path'])==x
  assert source==admission['source']
  for b in range(plan[lane]['blocks']):
   cases=plan['cases'] if b%2==0 else list(reversed(plan['cases']))
   for case in cases:
    for leg in plan[lane]['orders'][b]:one(lane,leg,case,b)
 assert d.source()==source
 d.write(d.P/(f'capture-{lane}-{sys.argv[2]}.json' if lane in ['qualification','observer','allocation-preflight'] else f'capture-{lane}.json'),dict(status='pass',source=source))
