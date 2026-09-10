#!/usr/bin/env python3
"""Small whole-child perf counter supplement; not operation-local attribution."""
import hashlib
import json
from pathlib import Path
import subprocess
import time

HERE=Path(__file__).resolve().parent
CACHE=Path('/home/zhuhe/.cache/litchi-goal-0498')
BINARY=CACHE/'retained/source_backed_batch_perf-after'
OUT=HERE/'profiles'
OUT.mkdir(exist_ok=True)
freeze=json.loads((HERE/'after-freeze.json').read_text())
assert hashlib.sha256(BINARY.read_bytes()).hexdigest()==freeze['binary_sha256']
for source in ['owned','instrumented']:
    for api,workers in [('serial',1),('batch',1),('batch',4)]:
        name=f'few-large-{source}-{api}-w{workers}'
        scratch=CACHE/'profile-corpora'/name
        command=['taskset','-c','0-7','perf','stat','-x,','-o',str(OUT/(name+'.perf.csv')),
                 '-e','cycles,instructions,branches,branch-misses,cache-misses','--',str(BINARY),
                 '--corpus','few-large','--source',source,'--api',api,'--capability','managed',
                 '--workers',str(workers),'--warmups','3','--samples','30','--repeats','2',
                 '--artifact-dir',str(scratch),'--output',str(OUT/(name+'.samples.csv'))]
        started=time.time()
        with (OUT/(name+'.stdout')).open('wb') as out,(OUT/(name+'.stderr')).open('wb') as err:
            result=subprocess.run(command,stdout=out,stderr=err)
        (OUT/(name+'.json')).write_text(json.dumps(dict(command=command,exit_code=result.returncode,
           started_epoch=started,ended_epoch=time.time(),binary_sha256=freeze['binary_sha256'],
           cleanup_verified=not scratch.exists(),scope='whole-child including setup and verification'),indent=2)+'\n')
        print(name,result.returncode,flush=True)
        assert result.returncode==0
        assert not scratch.exists()
