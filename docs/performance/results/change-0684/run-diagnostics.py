#!/usr/bin/env python3
"""Native N=10/1010 query diagnostics with process peak RSS and hardware counters."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys

PACKET=Path(__file__).resolve().parent
ROOT=PACKET.parents[3]
phase,binary=sys.argv[1:]
out=PACKET/'diagnostics'/phase
out.mkdir(parents=True,exist_ok=True)
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
records=[]
cases=[c for c in json.loads((PACKET/'cases.json').read_text()) if c['case'] in ['54016-stored','WithCustomViews-stored','Simple-stored']]
for case in cases:
    for mode in ['owned','file']:
        for repeat in range(3):
            for n in [10,1010]:
                name=f'{case["case"]}-{mode}-{repeat}-{n}'
                command=['perf','stat','-x,','-e','cycles,instructions,branches,branch-misses,cache-misses,page-faults','-o',str(out/(name+'.csv')),
                    '/usr/bin/time','-v','-o',str(out/(name+'.rss.txt')),'taskset','-c','12',binary,mode,case['path'],str(case['sheet']),str(case['row']),str(case['column']),str(n)]
                with (out/(name+'.tsv')).open('w') as stdout,(out/(name+'.stderr')).open('w') as stderr:
                    result=subprocess.run(command,cwd=ROOT,stdout=stdout,stderr=stderr)
                records.append(dict(command=command,exit_code=result.returncode,case=case['case'],mode=mode,repeat=repeat,n=n))
                (out/'commands.json').write_text(json.dumps(records,indent=2)+'\n')
                if result.returncode:raise SystemExit(result.returncode)
        print(phase,case['case'],mode,flush=True)
source_root=Path('/home/zhuhe/code/litchi-0684-before') if phase=='baseline' else ROOT
manifest=dict(phase=phase,binary_sha256=sha(binary),
    source_sha256={str(p.relative_to(source_root)):sha(p) for p in sorted((source_root/'crates/litchi-xls').rglob('*.rs'))},
    probe_sha256={str(p.relative_to(PACKET)):sha(p) for p in sorted((PACKET/'repeat-probe').rglob('*')) if p.is_file()},
    corpus={c['path']:sha(ROOT/c['path']) for c in cases},
    raw_sha256={p.name:sha(p) for p in out.iterdir() if p.is_file() and p.name!='manifest.json'})
(out/'manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
