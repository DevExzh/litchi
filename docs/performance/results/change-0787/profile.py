"""Root-only Callgrind region collection, preserving raw records."""
import subprocess,time
import custody as c
plan=c.read(c.P/'profile-plan.json');build=c.read(c.P/'profile-build/receipt.json');binary=build['binary'];assert c.artifact(binary['path'])==binary
frozen=c.source();assert frozen==c.read(c.P/'profile-build/source.json');out=c.P/'profiles';assert not out.exists();out.mkdir();rows=[];owner=build['owner']
for repeat,widths in enumerate(plan['orders']):
 for width in widths:
  stem=f'{repeat}-{width}';report=out/f'{stem}.json';raw=out/f'{stem}.callgrind';log=out/f'{stem}.log'
  cmd=['taskset','-c',','.join(map(str,plan['affinity'])),'valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect='+owner,'--zero-before='+owner,'--dump-after='+owner,'--callgrind-out-file='+str(raw),binary['path'],'--route','parts','--shape','large','--workers',str(width),'--task-floor','0','--state','primed','--samples','1','--warmup','0','--output',str(report)]
  start=time.time()
  with log.open('w') as f:r=subprocess.run(cmd,cwd=c.ROOT,stdout=f,stderr=subprocess.STDOUT)
  row={'repeat':repeat,'workers':width,'command':cmd,'exit_code':r.returncode,'started':start,'ended':time.time(),'binary':binary,'plan_sha256':c.sha(c.P/'profile-plan.json'),'driver_sha256':c.sha(c.P/'profile.py'),'artifacts':{f.name:c.artifact(f) for f in out.glob(stem+'.*') if f.is_file()}}
  rows.append(row);c.write(out/'receipts.json',rows);assert r.returncode==0,row;assert c.source()==frozen
  region=out/f'{stem}.callgrind.1';assert region.exists() and not (out/f'{stem}.callgrind.2').exists()
  lines=region.read_text().splitlines();totals=[int(x.split(':',1)[1]) for x in lines if x.startswith('summary:')];assert len(totals)==1 and totals[0]>1000
  assert any('read_parts_ordered' in x for x in lines),'missing operation owner'
  print(stem,'Callgrind region collected',totals[0],flush=True)
c.write(out/'complete.json',{'processes':len(rows),'receipts':c.artifact(out/'receipts.json')})
