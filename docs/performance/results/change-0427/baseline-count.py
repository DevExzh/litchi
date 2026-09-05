#!/usr/bin/env python3
"""Reproduce the selector-count failure with the exact prior library source."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
REVISION = 'd2f98b02e1d620c84359604a34c4ffadd9213cf3'


def main():
    path = REPO / 'tools/perf-baseline/src/lib.rs'
    current = path.read_bytes()
    baseline = subprocess.check_output(['git', 'show', REVISION + ':tools/perf-baseline/src/lib.rs'], cwd=REPO)
    assert current.replace(b'pub mod pptx_retention;\n', b'', 1) == baseline
    command = ['cargo', 'test', '--locked', '--manifest-path', 'tools/perf-baseline/Cargo.toml',
               '--all-features', '--lib', 'selectable_case_count_matches_current_enumeration',
               '--', '--test-threads=1']
    try:
        path.write_bytes(baseline)
        result = subprocess.run(command, cwd=REPO, capture_output=True)
    finally:
        path.write_bytes(current)
    sys.stdout.buffer.write(result.stdout)
    sys.stdout.buffer.write(result.stderr)
    output = result.stdout + result.stderr
    assert result.returncode == 101 and b'left: 429' in output and b'right: 427' in output
    record = {'status': 'reproduced', 'revision': REVISION, 'argv': command,
              'exit_code': result.returncode, 'baseline_lib_sha256': hashlib.sha256(baseline).hexdigest(),
              'restored_lib_sha256': hashlib.sha256(current).hexdigest(),
              'restored_exactly': path.read_bytes() == current,
              'scope': 'Exact prior lib.rs; new retention module is unreachable in this lib-only target.'}
    (ROOT / 'checks/baseline-count-custody.json').write_text(json.dumps(record, indent=2) + '\n')
    print(json.dumps(record))


if __name__ == '__main__':
    main()
