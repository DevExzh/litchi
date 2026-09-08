#!/usr/bin/env python3
"""One excluded warmup/sample per size and mode in each final binary."""
import subprocess
import sys
from common import ROOT, REPO, ENV, read, meta


if __name__ == '__main__':
    arm = sys.argv[1]
    assert arm in ('control', 'candidate')
    for instrumentation, binary in read(ROOT / f'{arm}-binaries.json').items():
        assert meta(binary['path']) == {key: binary[key] for key in ('bytes', 'sha256')}
        for mode in ('total', 'phases'):
            label = f'pilot-{arm}-{instrumentation}-{mode}'
            argv = [binary['path'], '--counts', '64,8192,131072', '--mode', mode,
                    '--samples', '1', '--warmups', '1', '--json',
                    str(ROOT / f'{label}.report.json')]
            subprocess.run([sys.executable, '-B', str(ROOT / 'gate.py'),
                            label, *argv], cwd=REPO, env=ENV, check=True)
