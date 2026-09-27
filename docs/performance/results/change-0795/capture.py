"""Root-only serial operation-scoped guest profiles and native samples."""
import subprocess,sys,time,gzip,shutil
import custody as c
lane=sys.argv[1];assert lane in ['profiles','perf']
p=c.P;plan=c.read(p/'plan.json');source=c.source();assert source['files']==c.read(p/'baseline-source.json')['files']
builds={leg:c.read(p/f'build-{leg}/build.json') for leg in ['before','after']}
for build in builds.values():
 for b in build['binaries'].values():assert c.artifact(b['path'])==b
out=p/lane;assert not out.exists();out.mkdir();c.write(out/'source.json',source);rows=[]
def run(cmd,log):
 started=time.time()
 with log.open('w') as stream:r=subprocess.run(cmd,cwd=c.ROOT,stdout=stream,stderr=subprocess.STDOUT)
 return {'command':cmd,'exit_code':r.returncode,'started':started,'ended':time.time(),'log':c.artifact(log)}
def zip_file(raw):
 dest=raw.with_name(raw.name+'.gz');assert not dest.exists()
 with raw.open('rb') as src,dest.open('wb') as dst:
  with gzip.GzipFile(filename='',mode='wb',fileobj=dst,mtime=0) as z:shutil.copyfileobj(src,z)
 return c.artifact(dest)
if lane=='profiles':
 owner=plan['owner']
 for repeat,shapes in enumerate(plan['callgrind']['shape_orders']):
  for shape in shapes:
   for leg in plan['callgrind']['leg_orders'][repeat]:
    stem=f'{repeat}-{shape}-{leg}';raw=out/f'{stem}.callgrind';report=out/f'{stem}.json';log=out/f'{stem}.log';binary=builds[leg]['binaries']['profile']
    cmd=['taskset','-c',str(plan['cpu']),'valgrind','--tool=callgrind','--branch-sim=yes','--collect-atstart=no','--toggle-collect='+owner,'--zero-before='+owner,'--dump-after='+owner,'--callgrind-out-file='+str(raw),binary['path'],'--mode','capture','--shape',shape,'--samples','1','--warmup','0','--output',str(report)]
    row={'repeat':repeat,'shape':shape,'leg':leg,'binary':binary,**run(cmd,log)};row['artifacts']={f.name:c.artifact(f) for f in out.glob(stem+'.*') if f.is_file()};rows.append(row);c.write(out/'receipts.json',rows)
    assert row['exit_code']==0 and report.exists() and (out/f'{stem}.callgrind.1').exists()
    assert not (out/f'{stem}.callgrind.2').exists()
    assert c.source()==source
    print(stem,'PASS',flush=True)
else:
 for repeat,order in enumerate(plan['perf']['orders']):
  for leg in order:
   stem=f'{repeat}-{leg}';raw=out/f'{stem}.data';report=out/f'{stem}.json';log=out/f'{stem}.log';binary=builds[leg]['binaries']['fp']
   cmd=['taskset','-c',str(plan['cpu']),'perf','record','-e',plan['perf']['event'],'-F',str(plan['perf']['frequency']),'--call-graph',plan['perf']['call_graph'],'-o',str(raw),'--',binary['path'],'--mode','capture','--shape','large','--samples','100','--warmup','3','--output',str(report)]
   row={'repeat':repeat,'leg':leg,'binary':binary,**run(cmd,log)}
   if report.exists():row['report']=c.artifact(report)
   if raw.exists():row['raw']=c.artifact(raw)
   rows.append(row);c.write(out/'receipts.json',rows)
   assert row['exit_code']==0 and report.exists() and raw.exists()
   assert c.source()==source
   print(stem,'captured',flush=True)
 decodes=[]
 for row in rows:
  repeat,leg=row['repeat'],row['leg'];stem=f'{repeat}-{leg}';raw=out/f'{stem}.data';frames=out/f'{stem}.frames';log=out/f'{stem}.decode.log'
  cmd=['perf','script','--no-inline','--ns','-i',str(raw)];started=time.time()
  with frames.open('w') as dst,log.open('w') as err:r=subprocess.run(cmd,stdout=dst,stderr=err,cwd=c.ROOT)
  decoded={'repeat':repeat,'leg':leg,'command':cmd,'exit_code':r.returncode,'started':started,'ended':time.time(),'raw':row['raw'],'log':c.artifact(log)}
  assert r.returncode==0
  decoded['frames']=zip_file(frames);frames.unlink();row['compressed']=zip_file(raw);raw.unlink()
  decodes.append(decoded);c.write(out/'decodes.json',decodes);c.write(out/'receipts.json',rows)
  print(stem,'decoded',flush=True)
c.write(out/'complete.json',{'children':len(rows),'source':c.artifact(out/'source.json'),'receipts':c.artifact(out/'receipts.json')})
