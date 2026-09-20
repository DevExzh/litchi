#!/usr/bin/env python3
"""Freeze matched native and allocator binaries with exact source custody."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / 'litchi-target-0708'
BIN = ROOT.parent / 'litchi-0708-bin'

def sha(p): return hashlib.sha256(p.read_bytes()).hexdigest()
def census():
    files = [ROOT/'Cargo.toml', ROOT/'Cargo.lock', *(ROOT/'.cargo').rglob('*')]
    files += [p for folder in ['crates','tools/perf-baseline'] for p in (ROOT/folder).rglob('*')
              if 'target' not in p.parts and (p.suffix == '.rs' or p.name in ['Cargo.toml','Cargo.lock'])]
    return {str(p.relative_to(ROOT)):sha(p) for p in sorted(files) if p.is_file()}

def main():
    phase = sys.argv[1]
    assert phase in ['baseline','candidate']
    for name,digest in json.loads((P/'constraints.json').read_text()).items():
        assert sha(ROOT/name) == digest
    source = census()
    if phase == 'candidate':
        baseline = json.loads((P/'source-baseline.json').read_text())
        changed = [name for name in set(source)|set(baseline) if source.get(name)!=baseline.get(name)]
        assert changed and all(name.startswith('crates/litchi-xlsx/') for name in changed), changed
    (P/f'source-{phase}.json').write_text(json.dumps(source,indent=2)+'\n')
    BIN.mkdir(exist_ok=True)
    records = []
    for label,binary,features in [('native','litchi-perf-baseline',[]),('alloc','litchi-perf-baseline-alloc',['--features','allocator-metrics'])]:
        command = ['cargo','build','--release','--locked','--manifest-path','tools/perf-baseline/Cargo.toml','--bin',binary,'--target-dir',str(TARGET),'-j','2',*features]
        start = time.monotonic()
        with (P/f'build-{phase}-{label}.log').open('w') as log:
            r = subprocess.run(command,cwd=ROOT,stdout=log,stderr=subprocess.STDOUT)
        assert census() == source, 'source changed during build'
        record = dict(command=command,exit_code=r.returncode,seconds=time.monotonic()-start,source_manifest_sha256=sha(P/f'source-{phase}.json'),environment={k:os.environ.get(k) for k in ['RUSTFLAGS','LD_PRELOAD','MALLOC_CONF','GLIBC_TUNABLES']})
        if r.returncode == 0:
            output = BIN/f'{phase}-{label}'
            shutil.copy2(TARGET/'release'/binary,output)
            record.update(binary=str(output),binary_sha256=sha(output),binary_bytes=output.stat().st_size)
        records.append(record)
        (P/f'build-{phase}.json').write_text(json.dumps(records,indent=2)+'\n')
        assert r.returncode == 0
        print(phase,label,'built',flush=True)

if __name__ == '__main__': main()
