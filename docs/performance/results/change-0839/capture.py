"""Serial immutable capture; native and diagnostic populations stay separate."""
import driver as d
import gzip,json,os,select,subprocess,sys,time
PLAN=d.read(d.P/'plan.json');AFF=','.join(map(str,PLAN['affinity']))
def identity(p):return d.desc(p)
def args(case,samples,warmup,path):
 result=[]
 for k,v in case.items():result+=['--'+k.replace('_','-'),str(v)]
 return result+['--samples',str(samples),'--warmup',str(warmup),'--output',str(path)]
def binary(leg,kind):
 name=leg if kind=='native' else leg+'-'+kind
 b=d.read(d.P/f'build-{name}.json')['binary'];assert identity(b['path'])==b
 return b['path']
def one(lane,leg,block,index,case,samples,warmup,kind='native',aff='all',protocol=0):
 dest=d.P/'runs'/lane/f'{block:02}-{index:02}-{leg}-{aff}-p{protocol}'
 dest.mkdir(parents=True,exist_ok=False);report=dest/'report.json';program=binary(leg,kind)
 if lane in ['memory','memory-preflight']:
  argv=['taskset','-c',AFF if aff=='all' else str(PLAN['affinity'][0]),str(d.TARGET/'launcher'),'--usage',str(dest/'usage.json'),program,*args(case,samples,warmup,report)]
 else:argv=['taskset','-c',AFF,'/usr/bin/time','-f','%M %R %F','-o',str(dest/'time.txt'),program,*args(case,samples,warmup,report)]
 e=d.env();e.pop('LITCHI_RSS_PHASES',None)
 if lane in ['memory','memory-preflight']:e['LITCHI_RSS_PHASES']='1'
 assert not e.get('LD_PRELOAD')
 launcher_identity=identity(d.TARGET/'launcher') if lane in ['memory','memory-preflight'] else None
 start=time.time();d.write(dest/'started.json',dict(argv=argv,lane=lane,leg=leg,block=block,index=index,case=case,samples=samples,warmup=warmup,kind=kind,affinity=aff,protocol=protocol,started_unix=start,binary=identity(program),launcher=launcher_identity,phase_environment=e.get('LITCHI_RSS_PHASES')))
 snapshots=[]
 with (dest/'stderr.log').open('xb') as err:
  if lane in ['memory','memory-preflight']:
   proc=subprocess.Popen(argv,cwd=d.ROOT,env=e,stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=err,bufsize=0)
   assert proc.stdout and proc.stdin
   data=b'';pid=None;starttime=None
   expected=[('startup',2**64-1),('corpus_ready',2**64-1),('warmup_done',2**64-1)]
   for sample_index in sorted({0,samples-1}):expected += [(n,sample_index) for n in ['package_ready','after_preload','after_operation','after_batch_drop','after_package_drop']]
   expected += [(n,2**64-1) for n in ['samples_done','report_written','report_dropped']]
   try:
    while True:
     ready,_,_=select.select([proc.stdout],[],[],60)
     if not ready:raise RuntimeError('phase timeout; stopping child '+str(proc.pid))
     chunk=os.read(proc.stdout.fileno(),4096)
     if not chunk:break
     data+=chunk
     while b'\n' in data:
      line,data=data.split(b'\n',1);tag,phase,sample,pid_text=line.decode().split('\t');assert tag=='RSS0788'
      current=int(pid_text);assert pid is None or pid==current;pid=current
      assert len(snapshots)<len(expected) and (phase,int(sample))==expected[len(snapshots)]
      base=d.Path('/proc')/str(pid)
      assert (d.Path('/proc')/str(proc.pid)/'exe').resolve()==(d.TARGET/'launcher').resolve()
      assert (base/'exe').resolve()==d.Path(program).resolve()
      raw={n:(base/n).read_text() for n in ['smaps','smaps_rollup','status','stat','maps']}
      status=dict(line.split(':',1) for line in raw['status'].splitlines() if ':' in line)
      assert int(status['PPid'])==proc.pid
      stat_fields=raw['stat'][raw['stat'].rfind(')')+2:].split();current_start=int(stat_fields[19])
      assert starttime is None or starttime==current_start;starttime=current_start
      snapshots.append(dict(phase=phase,sample=int(sample),pid=pid,parent_pid=proc.pid,starttime=starttime,launcher=launcher_identity,raw=raw))
      proc.stdin.write(b'+\n');proc.stdin.flush()
    assert not data
    proc.stdin.close();code=proc.wait(timeout=60)
    assert len(snapshots)==len(expected)
   except BaseException:
    proc.kill();proc.wait(timeout=10)
    raise
   with gzip.GzipFile(filename='',mode='wb',fileobj=(dest/'snapshots.json.gz').open('xb'),mtime=0) as f:f.write(json.dumps(snapshots,sort_keys=True).encode())
  else:
   with (dest/'stdout.log').open('xb') as out:code=subprocess.run(argv,cwd=d.ROOT,env=e,stdout=out,stderr=err).returncode
 d.write(dest/'receipt.json',dict(exit_code=code,finished_unix=time.time(),started_unix=start,argv=argv,artifacts=[identity(p) for p in sorted(dest.iterdir()) if p.is_file()]))
 assert code==0,(dest,code)
 print(lane,block,index,leg,aff,protocol,flush=True)
 return dest
if __name__=='__main__':
 lane=sys.argv[1];d.check();source=d.source()
 if lane=='qualification':
  leg=sys.argv[2]
  for i,case in enumerate(PLAN['cases']):one(lane,leg,0,i,case,1,0)
 else:
  assert lane in ['native','observer','memory','allocation']
  admission=d.read(d.P/'admission.json')
  for item in admission['files']:assert d.desc(item['path'])==item
  assert source==admission['source']
  for leg in ['before','after']:
   kind={'observer':'observer','allocation':'allocator'}.get(lane,'native');binary(leg,kind)
  settings=PLAN[lane];cases=PLAN[{'memory':'memory_cases','allocation':'allocation_cases'}.get(lane,'cases')]
  for block in range(settings['blocks']):
   order=PLAN['native']['orders'][block]
   indices=list(range(len(cases)))
   if block%2:indices.reverse()
   for i in indices:
    if lane=='memory':
     for aff in settings['affinities']:
      for pi,protocol in enumerate(settings['protocols']):
       for leg in order:one(lane,leg,block,i,cases[i],**protocol,aff=aff,protocol=pi)
    else:
     kind={'observer':'observer','allocation':'allocator'}.get(lane,'native')
     for leg in order:one(lane,leg,block,i,cases[i],settings['samples'],settings['warmup'],kind)
 assert d.source()==source
 suffix='-'+sys.argv[2] if lane=='qualification' else ''
 d.write(d.P/f'capture-{lane}{suffix}.json',dict(status='pass',source=source,lane=lane))
