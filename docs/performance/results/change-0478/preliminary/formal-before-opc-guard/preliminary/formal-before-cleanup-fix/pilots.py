#!/usr/bin/env python3
"""Validate both public writer policies before the formal capture matrix."""
import subprocess
import sys
from common import ROOT, REPO, TEMP, ENV, read, meta

if __name__ == '__main__':
    binaries = read(ROOT / 'binaries.json')
    for instrumentation in ('normal', 'allocator'):
        spec = binaries[instrumentation]
        assert meta(spec['path']) == {key: spec[key] for key in ('bytes', 'sha256')}
        label = f'pilot-{instrumentation}'
        command = [spec['path'], '--mode', 'both', '--counts', '8',
                   '--samples', '1', '--warmups', '1', '--repeats', '1',
                   '--max-spool-bytes', '67108864', '--spool-buffer-bytes', '16384',
                   '--json', str(ROOT / f'{label}.report.json'),
                   '--spool-dir', str(TEMP / 'spools' / label)]
        subprocess.run([sys.executable, '-B', str(ROOT / 'gate.py'), label, *command],
                       cwd=REPO, env=ENV, check=True)
