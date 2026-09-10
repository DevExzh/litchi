#!/usr/bin/env python3
"""Whole-child CPU/scheduling profile supplement to untraced latency captures."""
import hashlib,json,subprocess,sys
from pathlib import Path
HERE=Path(__file__).resolve().parent
PHASE=sys.argv[1];assert PHASE in ['before','after']
CACHE=Path('/home/zhuhe/.cache/litchi-goal-0499')
freeze=json.loads((HERE/f'{PHASE}-freeze.json').read_text());binary=Path(freeze['binary'])
assert hashlib.sha256(binary.read_bytes()).hexdigest()==freeze['binary_sha256']
OUT=HERE/'profiles'/PHASE;OUT.mkdir(parents=True,exist_ok=True)
for corpus in ['few-large','many-small']:
 name=f'{corpus}-owned-batch-w4';scratch=CACHE/'profile-corpora'/f'{PHASE}-{name}'
 cmd=['taskset','-c','0-7','perf','stat','-x,','-o',str(OUT/f'{name}.perf.csv'),'-e',
      'cycles,instructions,branches,branch-misses,cache-misses,context-switches,cpu-migrations,page-faults','--',str(binary),
      '--corpus',corpus,'--source','owned','--api','batch','--capability','managed','--workers','4',
      '--warmups','3','--samples','30','--repeats','2','--artifact-dir',str(scratch),'--output',str(OUT/f'{name}.samples.csv')]
 with (OUT/f'{name}.stdout').open('wb') as out,(OUT/f'{name}.stderr').open('wb') as err:result=subprocess.run(cmd,stdout=out,stderr=err)
 (OUT/f'{name}.json').write_text(json.dumps({'command':cmd,'exit_code':result.returncode,'binary_sha256':freeze['binary_sha256'],'scope':'whole-child including setup/verification','cleanup_verified':not scratch.exists()},indent=2)+'\n')
 assert result.returncode==0 and not scratch.exists()
 print(PHASE,name,'profile passed',flush=True)
