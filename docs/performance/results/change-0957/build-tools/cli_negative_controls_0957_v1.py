import json,pathlib,hashlib,subprocess,os,time
root=pathlib.Path('/home/zhuhe/litchi-performance-0957')
binary=root/'executables/xlsx_managed_compression_alloc.v2'
out=root/'cli-negative-controls-v1';out.mkdir()
def desc(p): return {'path':str(p),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()}
def write(p,x):
 with p.open('x') as f: json.dump(x,f,indent=2);f.write('\n')
assert desc(binary)['sha256']=='9ba37a6b9796ad0e316e7571ebe117ba7a5f97bcf5d29153c7cb3d74c485e9a5'
env={'PATH':'/usr/bin:/bin','LANG':'C','LC_ALL':'C','TZ':'UTC'}
existing=out/'existing';existing.mkdir();sentinel=existing/'sentinel';sentinel.write_bytes(b'preserve existing output\n')
base=['--shape','medium','--workload','one-percent','--workers','1','--warmup','0','--samples','1']
cases=[]
for label,index,value in [('workers-zero',5,'0'),('workers-eight',5,'8'),('samples-zero',9,'0')]:
 args=base.copy();args[index]=value;cases.append((label,args+['--output',str(out/(label+'-output'))]))
cases += [('missing-flags',[]),('existing-output',base+['--output',str(existing)])]
records=[]
for name,args in cases:
 folder=out/name;folder.mkdir();argv=[str(binary)]+args
 write(folder/'start.json',{'argv':argv,'cwd':str(root/'source-v1'),'environment':env,'deadline_seconds':15,'binary':desc(binary),'runner':desc(pathlib.Path(__file__))})
 with (folder/'stdout.log').open('xb') as stdout,(folder/'stderr.log').open('xb') as stderr:
  p=subprocess.Popen(argv,cwd=root/'source-v1',env=env,stdout=stdout,stderr=stderr,start_new_session=True)
  try: rc=p.wait(timeout=15)
  except subprocess.TimeoutExpired:
   os.killpg(p.pid,9);p.wait();raise
 try: os.killpg(p.pid,0);clear=False
 except ProcessLookupError: clear=True
 terminal={'returncode':rc,'leader_reaped':p.poll() is not None,'process_group_clear':clear,'timed_out':False,'pid':p.pid,'terminal_unix_ns':time.time_ns()}
 write(folder/'terminal.json',terminal)
 assert rc!=0 and clear
 if name=='existing-output': assert list(existing.iterdir())==[sentinel] and sentinel.read_bytes()==b'preserve existing output\n'
 elif name!='missing-flags': assert not (out/(name+'-output')).exists()
 records.append({'name':name,'passed':True,'start':desc(folder/'start.json'),'terminal':desc(folder/'terminal.json'),'stderr':desc(folder/'stderr.log')})
write(out/'result.json',{'schema':'litchi.performance.0957.cli-negative-controls.v1','status':'passed','binary':desc(binary),'runner':desc(pathlib.Path(__file__)),'checks':records})
print(json.dumps({'status':'passed','checks':len(records),'result':desc(out/'result.json')}))
