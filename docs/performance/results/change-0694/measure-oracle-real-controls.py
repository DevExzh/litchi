#!/usr/bin/env python3
"""Supplemental native MCE profile controls on deterministic real XML parts."""
import hashlib,json,statistics,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
out=P/'oracle-real-controls';out.mkdir(exist_ok=True)
def sha(f):return hashlib.sha256(f.read_bytes()).hexdigest()
corpus=json.loads((P/'oracle-results/corpus.json').read_text())
selection=[]
for family in ['docx','xlsx','pptx']:
 cases=[r for r in corpus['cases'] if r['origin']['kind']==family and r['mutation']['kind']=='identity' and b'http://schemas.openxmlformats.org/markup-compatibility/2006' in (P/'oracle-results'/r['path']).read_bytes()]
 case=max(cases,key=lambda r:(r['input_len'],r['id']))
 (out/(family+'.xml')).write_bytes((P/'oracle-results'/case['path']).read_bytes())
 selection.append(dict(family=family,case=case))
(P/'oracle-real-cases.json').write_text(json.dumps(selection,indent=2)+'\n')
rows=[]
for leg,phase in [('a0','baseline'),('a1','baseline'),('a2','baseline'),('b0','candidate'),('b1','candidate'),('a3','baseline')]:
 binary=ROOT.parent/'litchi-0694-bin'/(phase+'-oracle')
 for case in ['docx','xlsx','pptx']:
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
   (P/'oracle-real-controls.json').write_text(json.dumps(rows,indent=2)+'\n')
   print(case,profile,leg,flush=True)
for case in ['docx','xlsx','pptx']:
 for profile in ['baseline','opaque','opaque-many']:
  assert len({r['identity'] for r in rows if r['case']==case and r['profile']==profile})==1
comparisons=[]
for case in ['docx','xlsx','pptx']:
 for profile in ['baseline','opaque','opaque-many']:
  group={r['leg']:r for r in rows if r['case']==case and r['profile']==profile}
  for b,a in [('a1','a0'),('b0','a2'),('b1','a3'),('a3','a2')]:
   comparisons.append(dict(case=case,profile=profile,pair=b+'/'+a,delta_pct={k:(group[b][k]/group[a][k]-1)*100 for k in ['p50_ns','mean_ns','p95_ns','p99_ns']}))
(P/'oracle-real-control-comparisons.json').write_text(json.dumps(comparisons,indent=2)+'\n')
