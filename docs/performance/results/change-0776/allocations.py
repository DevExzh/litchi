"""Separate serial heaptrack observations; instrumented time is never native time."""
from pathlib import Path
import json,subprocess,time,re
import measure
P=measure.P;M=P/'measure-1';sha=measure.sha;write=measure.write
if __name__=='__main__':
 assert (M/'complete.json').exists();out=P/'allocations-0';assert not out.exists();out.mkdir()
 inputs={leg:measure.source(root,measure.REFS[leg]) for leg,root in measure.ROOTS.items()};bins=json.loads((M/'binaries.json').read_text());rows=[]
 for case in ['mce_benign_worksheet','mce_benign_document','mce_prefixed_32']:
  for leg in ['before','after']:
   binary=Path(bins[leg]['path']);assert sha(binary)==bins[leg]['sha256'];prefix=out/f'{case}-{leg}';stdout=prefix.with_suffix('.stdout');stderr=prefix.with_suffix('.stderr');report=prefix.with_suffix('.json')
   cmd=['taskset','-c','12','heaptrack','--record-only','-o',str(prefix),str(binary),'adversarial','--case',case,'--samples','1','--warmup','0','--json',str(report)];start=time.time()
   with stdout.open('w') as o,stderr.open('w') as e:r=subprocess.run(cmd,cwd=P.parents[3],stdout=o,stderr=e)
   row={'case':case,'leg':leg,'command':cmd,'exit':r.returncode,'started':start,'ended':time.time(),'stdout':stdout.name,'stdout_sha256':sha(stdout),'stderr':stderr.name,'stderr_sha256':sha(stderr),'report':report.name,'report_sha256':sha(report) if report.exists() else None};rows.append(row);write(out/'runs.json',rows);assert r.returncode==0
   captures=[f for f in out.glob(prefix.name+'.*') if f.suffix in ['.gz','.zst']];assert len(captures)==1;capture=captures[0];row['capture']=capture.name;row['capture_sha256']=sha(capture)
   log=prefix.with_suffix('.summary');hist=prefix.with_suffix('.histogram');cmd=['heaptrack_print','-f',str(capture),'-H',str(hist),'-p','0','-a','0','-T','0','-l','0']
   with log.open('w') as f:r=subprocess.run(cmd,stdout=f,stderr=subprocess.STDOUT)
   row.update({'decode_command':cmd,'decode_exit':r.returncode,'summary':log.name,'summary_sha256':sha(log),'histogram':hist.name,'histogram_sha256':sha(hist) if hist.exists() else None});write(out/'runs.json',rows);assert r.returncode==0
   pairs=[list(map(int,line.split())) for line in hist.read_text().splitlines()];calls=sum(count for _,count in pairs);assert calls==int(re.search(r'calls to allocation functions: (\d+)',log.read_text())[1]);row['allocation_calls']=calls;row['allocated_bytes']=sum(size*count for size,count in pairs);write(out/'runs.json',rows)
   print(prefix.name+' allocations captured',flush=True)
 for leg in inputs:assert measure.source(measure.ROOTS[leg],measure.REFS[leg])==inputs[leg]
 write(out/'complete.json',{'source_unchanged':True,'scope':'One whole-process instrumented sample per case/leg, including startup, XML input generation, parser, observer and report serialization. Instrumented timings excluded from native statistics.'})
