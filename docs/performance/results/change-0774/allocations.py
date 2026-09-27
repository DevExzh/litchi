"""Separate serial heaptrack lane: whole-process counts, never native timing."""
from pathlib import Path
import json,os,subprocess,time
import measure
P=measure.P;M=P/'measure-1';sha=measure.sha;write=measure.write
if __name__=='__main__':
 assert (M/'complete.json').exists();out=P/'allocations-0';assert not out.exists();out.mkdir()
 inputs={leg:measure.source(root,measure.REFS[leg]) for leg,root in measure.ROOTS.items()};bins=json.loads((M/'binaries.json').read_text());rows=[]
 for case in ['formula','numeric']:
  for leg in ['before','after']:
   binary=Path(bins[leg]['path']);assert sha(binary)==bins[leg]['sha256'];prefix=out/f'{case}-{leg}';stdout=prefix.with_suffix('.stdout');stderr=prefix.with_suffix('.stderr')
   cmd=['taskset','-c','12','heaptrack','--record-only','-o',str(prefix),str(binary),'--case',case,'--samples','1','--warmup','0'];start=time.time()
   with stdout.open('w') as o,stderr.open('w') as e:r=subprocess.run(cmd,cwd=P.parents[3],stdout=o,stderr=e)
   row={'case':case,'leg':leg,'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'stdout':stdout.name,'stdout_sha256':sha(stdout),'stderr':stderr.name,'stderr_sha256':sha(stderr)};rows.append(row);write(out/'runs.json',rows);assert r.returncode==0
   captures=[f for f in out.glob(prefix.name+'.*') if f.suffix in ['.gz','.zst']];assert len(captures)==1;capture=captures[0];row['capture']=capture.name;row['capture_sha256']=sha(capture)
   log=prefix.with_suffix('.summary');cmd=['heaptrack_print','-f',str(capture),'-p','0','-a','0','-T','0','-l','0']
   with log.open('w') as f:r=subprocess.run(cmd,stdout=f,stderr=subprocess.STDOUT)
   row.update({'decode_command':cmd,'decode_exit':r.returncode,'summary':log.name,'summary_sha256':sha(log)});write(out/'runs.json',rows);assert r.returncode==0
   print(prefix.name+' allocations captured',flush=True)
 for leg in inputs:assert measure.source(measure.ROOTS[leg],measure.REFS[leg])==inputs[leg]
 write(out/'complete.json',{'source_unchanged':True,'scope':'Whole process including startup, input generation, single serialization and post-timer SHA; instrumented times excluded from native statistics.'})
