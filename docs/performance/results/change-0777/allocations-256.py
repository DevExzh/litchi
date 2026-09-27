"""Serial whole-process heap observations, separate from native timing."""
from pathlib import Path
import re,subprocess,time
import capture as c
P=c.PACKET;D=c.CAPTURE
if __name__=='__main__':
 complete=c.json_read(D/'complete.json');assert complete['complete'];out=P/'allocations-256';assert not out.exists();out.mkdir()
 bins=c.json_read(D/'binaries.json');refs={'before':c.BASE_REF,'after':complete['candidate']};roots={'before':c.BEFORE,'after':c.AFTER};inputs={leg:c.source_manifest(root,refs[leg]) for leg,root in roots.items()};rows=[]
 for case,n in [('opc_relationship_declarations',256)]:
  for leg in ['before','after']:
   binary=bins[leg]['mce-stream-probe' if n is None else 'attribute_checks'];assert c.sha256(Path(binary['path']))==binary['sha256']
   prefix=out/f'{case}-{n}-{leg}';stdout=prefix.with_suffix('.stdout');stderr=prefix.with_suffix('.stderr');report=prefix.with_suffix('.json');args=[binary['path'],'adversarial' if n is None else 'probe','--case',case,'--samples','1','--warmup','0','--json',str(report)]
   if n is not None:args+=['--n',str(n)]
   cmd=['taskset','-c',str(c.CPU),'heaptrack','--record-only','-o',str(prefix),*args];start=time.time()
   with stdout.open('w') as o,stderr.open('w') as e:r=subprocess.run(cmd,cwd=c.AFTER,stdout=o,stderr=e)
   row={'case':case,'n':n,'leg':leg,'binary_sha256':binary['sha256'],'command':cmd,'exit':r.returncode,'started':start,'ended':time.time()}
   for name,path in [('stdout',stdout),('stderr',stderr),('report',report)]:row[name]=path.name;row[name+'_sha256']=c.sha256(path) if path.exists() else None
   rows.append(row);c.json_write(out/'runs.json',rows);assert r.returncode==0
   captures=[f for f in out.glob(prefix.name+'.*') if f.suffix in ['.zst','.gz']];assert len(captures)==1;capture=captures[0];summary=prefix.with_suffix('.summary');hist=prefix.with_suffix('.histogram');cmd=['heaptrack_print','-f',str(capture),'-H',str(hist),'-p','0','-a','0','-T','0','-l','0']
   with summary.open('w') as f:r=subprocess.run(cmd,stdout=f,stderr=subprocess.STDOUT)
   row.update({'decode_command':cmd,'decode_exit':r.returncode})
   for name,path in [('capture',capture),('summary',summary),('histogram',hist)]:row[name]=path.name;row[name+'_sha256']=c.sha256(path) if path.exists() else None
   c.json_write(out/'runs.json',rows);assert r.returncode==0
   pairs=[list(map(int,line.split())) for line in hist.read_text().splitlines()];calls=sum(count for _,count in pairs);assert calls==int(re.search(r'calls to allocation functions: (\d+)',summary.read_text())[1]);row.update({'allocation_calls':calls,'allocated_bytes':sum(size*count for size,count in pairs)});c.json_write(out/'runs.json',rows);print(prefix.name+' captured',flush=True)
 for leg in inputs:assert c.source_manifest(roots[leg],refs[leg])==inputs[leg]
 assert c.fixture_manifest()==c.json_read(D/'fixtures-before.json')
 c.json_write(out/'complete.json',{'source_unchanged':True,'fixtures_unchanged':True,'runs':len(rows),'scope':'One whole-process instrumented sample per case/leg, including startup, input generation, parsing, observation and report serialization. Instrumented time is not native time.'})
