#!/usr/bin/env python3
"""Build one frozen current-source native diagnostic harness."""
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import time

P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('custody0716build', P/'custody.py')
C = importlib.util.module_from_spec(spec)
spec.loader.exec_module(C)

def main():
    assert not (P/'build.json').exists()
    source = C.census()
    assert source == json.loads((P.parent/'change-0713/source-final.json').read_text())
    (P/'source.json').write_text(json.dumps(source, indent=2)+'\n')
    command = ['cargo', 'build', '--release', '--locked', '--manifest-path',
               'tools/perf-baseline/Cargo.toml', '--bin', 'litchi-perf-baseline',
               '--target-dir', str(C.TARGET), '-j', '2']
    start = time.monotonic()
    log = P/'build.log'
    with log.open('w') as stream:
        result = subprocess.run(command, cwd=C.ROOT, stdout=stream, stderr=subprocess.STDOUT)
    assert C.census() == source
    record = dict(command=command, exit_code=result.returncode,
                  seconds=time.monotonic()-start, source_sha256=C.sha(P/'source.json'),
                  log_sha256=C.sha(log))
    if result.returncode == 0:
        C.BIN.mkdir()
        binary = C.BIN/'current-native'
        shutil.copy2(C.TARGET/'release/litchi-perf-baseline', binary)
        record['binary'] = dict(path=str(binary), sha256=C.sha(binary), bytes=binary.stat().st_size)
    (P/'build.json').write_text(json.dumps(record, indent=2)+'\n')
    assert result.returncode == 0
    print('Current native build PASS', flush=True)

if __name__ == '__main__':
    main()
