#!/usr/bin/env python3
"""Secondary native MCE controls for empty, one-name and 4096-name profiles."""
import hashlib,json,statistics,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
out=P/'oracle-controls';out.mkdir(exist_ok=True)
def sha(f):return hashlib.sha256(f.read_bytes()).hexdigest()
mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
a='http://schemas.openxmlformats.org/drawingml/2006/main'
head=f'<r xmlns:mc="{mc}" xmlns:a="{a}" xmlns:ext="urn:litchi:oracle:opaque">'
body='<a:n a:x="1"/>'*1000
for name,xml in [('ordinary',head+body+'</r>'),('opaque',head+'<ext:payload>'+body+'</ext:payload></r>')]:
 (out/(name+'.xml')).write_text(xml)
rows=[]
for leg,phase in [('a0','baseline'),('a1','baseline'),('a2','baseline'),('b0','candidate'),('b1','candidate'),('a3','baseline')]:
 binary=ROOT.parent/'litchi-0702-bin'/(phase+'-oracle')
 for case in ['ordinary','opaque']:
  source=out/(case+'.xml')
  for profile in ['baseline','opaque','opaque-many']:
   identity=subprocess.check_output([str(binary),'run',profile,str(source)],text=True)
   assert identity.startswith('OK\t')
   command=['taskset','-c','12',str(binary),'time',profile,str(source),'10','300']
   result=subprocess.run(command,capture_output=True,text=True,check=True)
   path=out/f'{case}-{profile}-{leg}.tsv';path.write_text(result.stdout)
   error=path.with_suffix('.stderr');error.write_text(result.stderr)
   samples=[int(line.rsplit('=',1)[1]) for line in result.stdout.splitlines() if line.startswith('SAMPLE\t')]
   assert len(samples)==300
   rows.append(dict(case=case,profile=profile,leg=leg,phase=phase,command=command,binary_sha256=sha(binary),source_sha256=sha(source),source=str(source.relative_to(P)),output=str(path.relative_to(P)),output_sha256=sha(path),stderr_sha256=sha(error),identity=identity,exit_code=result.returncode,p50_ns=statistics.median(samples),mean_ns=statistics.mean(samples),p95_ns=sorted(samples)[284],p99_ns=sorted(samples)[296]))
   (P/'oracle-controls.json').write_text(json.dumps(rows,indent=2)+'\n')
   print(case,profile,leg,flush=True)
for case in ['ordinary','opaque']:
 for profile in ['baseline','opaque','opaque-many']:
  assert len({r['identity'] for r in rows if r['case']==case and r['profile']==profile})==1
comparisons=[]
for case in ['ordinary','opaque']:
 for profile in ['baseline','opaque','opaque-many']:
  group={r['leg']:r for r in rows if r['case']==case and r['profile']==profile}
  for b,a in [('a1','a0'),('b0','a2'),('b1','a3'),('a3','a2')]:
   comparisons.append(dict(case=case,profile=profile,pair=b+'/'+a,delta_pct={k:(group[b][k]/group[a][k]-1)*100 for k in ['p50_ns','mean_ns','p95_ns','p99_ns']}))
(P/'oracle-control-comparisons.json').write_text(json.dumps(comparisons,indent=2)+'\n')
