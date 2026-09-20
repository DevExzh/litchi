#!/usr/bin/env python3
"""Run the six retained repository evidence gates, with source custody."""
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
spec = importlib.util.spec_from_file_location('build_custody', P/'build.py')
B = importlib.util.module_from_spec(spec)
spec.loader.exec_module(B)

def main():
    source = B.census()
    (P/'source-final.json').write_text(json.dumps(source,indent=2)+'\n')
    prior = json.loads((P.parent/'change-0708/evidence/results.json').read_text())
    out = P/'evidence'
    out.mkdir(exist_ok=True)
    assert not (out/'results.json').exists(), 'refusing overwrite'
    rows=[]
    for item in prior:
        name, command = item['name'], item['command']
        log = out/(name+'.log')
        assert not log.exists()
        start=time.monotonic()
        with log.open('w') as stream:
            result=subprocess.run(command,cwd=ROOT,stdout=stream,stderr=subprocess.STDOUT)
        assert B.census() == source, 'source changed during evidence gate'
        rows.append(dict(name=name,command=command,exit_code=result.returncode,
                         seconds=time.monotonic()-start,log_sha256=B.sha(log),
                         source_manifest_sha256=B.sha(P/'source-final.json')))
        (out/'results.json').write_text(json.dumps(rows,indent=2)+'\n')
        print(name,result.returncode,flush=True)
        assert result.returncode == 0

if __name__ == '__main__': main()
