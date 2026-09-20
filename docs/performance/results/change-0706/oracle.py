#!/usr/bin/env python3
"""Build and capture the separate public XML audit differential/guard probe."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

from build import census, sha, ROOT, TARGET, BIN

P = Path(__file__).resolve().parent

def inputs():
    return {str(p.relative_to(P)): sha(p) for p in sorted((P/'oracle-probe').rglob('*'))
            if p.is_file() and 'target' not in p.parts}

def save(path, value):
    path.write_text(json.dumps(value, indent=2)+'\n')

def main():
    action, role = sys.argv[1:3]
    assert role in ['baseline', 'candidate']
    binary = BIN / (role+'-oracle')
    source = census()
    expected = json.loads((P/f'source-{role}.json').read_text())
    delta = sorted(n for n in set(source)|set(expected) if source.get(n)!=expected.get(n))
    assert not delta or (role == 'baseline' and all(n.startswith('crates/xml-minifier/') for n in delta))
    record = dict(action=action, role=role, source_manifest_sha256=sha(P/f'source-{role}.json'),
                  current_source=source, retained_source_delta=delta, script_sha256=sha(Path(__file__)),
                  plan_sha256=sha(P/'plan.json'), environment={k:os.environ.get(k) for k in ['RUSTFLAGS','LD_PRELOAD','MALLOC_CONF','GLIBC_TUNABLES']})
    if action == 'build':
        assert not delta
        command = ['cargo','build','--release','--manifest-path',str(P/'oracle-probe/Cargo.toml'),
                   '--target-dir',str(TARGET),'-j','2']
        if (P/'oracle-probe/Cargo.lock').exists():
            command.append('--locked')
        before = inputs()
        started = time.monotonic()
        with (P/f'oracle-build-{role}.log').open('w') as log:
            result = subprocess.run(command, cwd=ROOT, stdout=log, stderr=subprocess.STDOUT)
        assert result.returncode == 0
        assert all(inputs().get(n)==d for n,d in before.items())
        shutil.copy2(TARGET/'release/xml-minifier-oracle-probe', binary)
        record.update(command=command, seconds=time.monotonic()-started, exit_code=result.returncode,
                      probe_inputs=inputs(), binary=str(binary), binary_sha256=sha(binary), binary_bytes=binary.stat().st_size)
        destination = P/f'oracle-build-{role}.json'
    else:
        assert action in ['differential','A1','B1','B2','A2']
        built = json.loads((P/f'oracle-build-{role}.json').read_text())
        assert inputs() == built['probe_inputs']
        assert sha(binary) == built['binary_sha256']
        command = ['taskset','-c','12',str(binary)]
        if action != 'differential': command.append('--bench')
        prefix = f'oracle-{role}-{action}'
        output, error = P/(prefix+'.json'), P/(prefix+'.stderr')
        assert not output.exists()
        started = time.monotonic()
        with output.open('w') as stdout, error.open('w') as stderr:
            result = subprocess.run(command, cwd=ROOT, stdout=stdout, stderr=stderr)
        record.update(command=command, seconds=time.monotonic()-started, exit_code=result.returncode,
                      build_record_sha256=sha(P/f'oracle-build-{role}.json'), binary=str(binary),
                      binary_sha256=sha(binary), artifacts={output.name:sha(output),error.name:sha(error)})
        destination = P/(prefix+'-receipt.json')
    assert census() == source
    save(destination, record)
    assert record['exit_code'] == 0
    print(action, role, 'passed', flush=True)

if __name__ == '__main__': main()
