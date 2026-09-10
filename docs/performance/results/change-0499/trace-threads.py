#!/usr/bin/env python3
"""Whole-child thread-creation counts; traced timings are diagnostic only."""
import hashlib,json,subprocess,sys
from pathlib import Path
HERE=Path(__file__).resolve().parent
PHASE=sys.argv[1]
assert PHASE in ['before','after']
CACHE=Path('/home/zhuhe/.cache/litchi-goal-0499')
freeze=json.loads((HERE/f'{PHASE}-freeze.json').read_text())
binary=Path(freeze['binary'])
assert hashlib.sha256(binary.read_bytes()).hexdigest()==freeze['binary_sha256']
OUT=HERE/'thread-traces'/PHASE;OUT.mkdir(parents=True,exist_ok=True)
for corpus,api,workers in [('many-small','serial',1),('many-small','batch',4),('few-large','batch',4)]:
 name=f'{corpus}-{api}-w{workers}'
 scratch=CACHE/'trace-corpora'/f'{PHASE}-{name}'
 cmd=['taskset','-c','0-7','strace','-f','-qq','-c','-e','clone,clone3','-o',str(OUT/f'{name}.trace'),str(binary),
      '--corpus',corpus,'--source','owned','--api',api,'--capability','managed','--workers',str(workers),
      '--warmups','0','--samples','1','--repeats','1','--artifact-dir',str(scratch),'--output',str(OUT/f'{name}.csv')]
 with (OUT/f'{name}.stdout').open('wb') as out,(OUT/f'{name}.stderr').open('wb') as err:
  result=subprocess.run(cmd,stdout=out,stderr=err)
 record={'command':cmd,'exit_code':result.returncode,'binary_sha256':freeze['binary_sha256'],'scope':'whole-child including package setup; traced times excluded','cleanup_verified':not scratch.exists()}
 (OUT/f'{name}.json').write_text(json.dumps(record,indent=2)+'\n')
 assert result.returncode==0 and not scratch.exists()
 print(name,(OUT/f'{name}.trace').read_text(),flush=True)
