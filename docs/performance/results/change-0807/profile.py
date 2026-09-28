"""Root-only serial Callgrind capture at the exact evidence wrapper."""
import subprocess,time
import custody as c
plan=c.read(c.P/'plan.json');build=c.read(c.P/'build/build.json');binary=build['binaries']['profile']
assert c.source()==c.read(c.P/'build/source.json')
assert c.sha(binary['path'])==binary['sha256']
assert not (c.P/'profiles').exists()
out=c.P/'profiles';out.mkdir();rows=[]
owner=plan['owner']
for block,shapes in enumerate(plan['profile']['orders']):
 for shape in shapes:
  stem=f'{block}-{shape}';report=out/f'{stem}.json';raw=out/f'{stem}.callgrind';log=out/f'{stem}.log'
  cmd=['taskset','-c',str(plan['cpu']),'valgrind','--tool=callgrind','--collect-atstart=no','--toggle-collect='+owner,'--zero-before='+owner,'--dump-after='+owner,'--callgrind-out-file='+str(raw),binary['path'],'--mode','capture','--shape',shape,'--samples','1','--warmup','0','--output',str(report)]
  started=time.time()
  with log.open('w') as f:r=subprocess.run(cmd,stdout=f,stderr=subprocess.STDOUT)
  row={'block':block,'shape':shape,'command':cmd,'started':started,'ended':time.time(),'exit_code':r.returncode,'binary':binary,'driver_sha256':c.sha(c.P/'profile.py'),'artifacts':{f.name:c.artifact(f) for f in out.glob(stem+'.*') if f.is_file()}}
  rows.append(row);c.write(out/'receipts.json',rows)
  assert r.returncode==0,log
  assert c.source()==c.read(c.P/'build/source.json')
  assert c.sha(binary['path'])==binary['sha256']
  assert (out/f'{stem}.callgrind.1').is_file(),'wrapper did not produce expected region dump'
  assert not (out/f'{stem}.callgrind.2').exists(),'unexpected additional wrapper call'
  lines=(out/f'{stem}.callgrind.1').read_text().splitlines()
  summaries=[int(line.split(':',1)[1].strip()) for line in lines if line.startswith('summary:')]
  assert len(summaries)==1 and summaries[0]>1000,'region contains too little work for public capture; inspect tail-call scope'
  assert any('capture_internal' in line for line in lines),'missing capture implementation in scoped profile'
  print(stem,'profile PASS',flush=True)
c.write(out/'complete.json',{'processes':len(rows),'plan_sha256':c.sha(c.P/'plan.json'),'build_sha256':c.sha(c.P/'build/build.json')})
