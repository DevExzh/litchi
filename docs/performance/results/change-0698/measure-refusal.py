#!/usr/bin/env python3
"""Native supplemental A/A and A/B/B/A, with every refusal checked by the probe."""
import hashlib,json,subprocess,sys,time
from pathlib import Path
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
mode=sys.argv[1]
assert mode in ['baseline','compare']
legs=[('a0','baseline'),('a1','baseline')] if mode=='baseline' else [('a2','baseline'),('b0','candidate'),('b1','candidate'),('a3','baseline')]
out=P/'refusal';out.mkdir(exist_ok=True)
records=[]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
for leg,phase in legs:
 binary=ROOT.parent/'litchi-0698-bin'/f'{phase}-refusal'
 binding=json.loads((P/f'build-refusal-{phase}.json').read_text())
 assert sha(binary)==binding['binary_sha256']
 command=['taskset','-c','12',str(binary),'matrix','100','5']
 output=out/f'{leg}.tsv';error=output.with_suffix('.stderr')
 started=time.monotonic()
 with output.open('w') as stdout,error.open('w') as stderr:
  result=subprocess.run(command,cwd=ROOT,stdout=stdout,stderr=stderr)
 records.append(dict(leg=leg,phase=phase,command=command,exit_code=result.returncode,seconds=time.monotonic()-started,output=str(output.relative_to(P)),output_sha256=sha(output),stderr_sha256=sha(error),binary_sha256=sha(binary)))
 (P/f'refusal-runs-{mode}.json').write_text(json.dumps(records,indent=2)+'\n')
 print(leg,result.returncode,flush=True)
 assert result.returncode==0,error.read_text()
