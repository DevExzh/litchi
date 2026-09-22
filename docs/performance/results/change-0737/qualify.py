#!/usr/bin/env python3
"""Fresh exact-oracle qualification before the frozen comparative matrix."""
import subprocess
import sys
import time
from contract import P, ROOT, command, read, sha, validate_report, write

variant = sys.argv[1]
assert variant in ('legacy', 'controls')
out = P / f'qualification-{variant}'
assert not out.exists()
out.mkdir()
runs = []
for case in ('primary', 'secondary'):
    for lane in ('native', 'allocation'):
        lifecycles = ['legacy'] if variant == 'legacy' else ['legacy', 'strict-retained', 'strict-drained']
        choices = [(lifecycle, warmup) for lifecycle in lifecycles
                   for warmup in ([0, 3] if lane == 'allocation' and lifecycle != 'legacy'
                                  else [0 if lane == 'allocation' else 3])]
        for lifecycle, warmups in choices:
            row = dict(case=case, lane=lane, build=variant, lifecycle=lifecycle,
                       samples=1, warmups=warmups)
            cmd = command(row)
            name = f'{lane}-{case}-{lifecycle}' + ('-w3' if lane == 'allocation' and warmups == 3 else '')
            start = time.monotonic()
            with (out / f'{name}.json').open('wb') as stdout, (out / f'{name}.stderr').open('wb') as stderr:
                result = subprocess.run(cmd, cwd=ROOT, stdout=stdout, stderr=stderr)
            entry = dict(**row, command=cmd, exit_code=result.returncode,
                         seconds=time.monotonic()-start, output=f'{out.name}/{name}.json',
                         sha256=sha(out / f'{name}.json'), stderr_sha256=sha(out / f'{name}.stderr'))
            runs.append(entry)
            write(out / 'manifest.json', dict(status='running', runs=runs))
            assert result.returncode == 0
            validate_report(read(out / f'{name}.json'), row)
            print('PASS qualification', variant, lane, case, lifecycle, flush=True)
write(out / 'manifest.json', dict(status='passed', runs=runs))
