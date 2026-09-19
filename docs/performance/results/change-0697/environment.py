#!/usr/bin/env python3
"""Record this attribution run's host and tools without changing system state."""
import datetime
import json
import os
import subprocess
from pathlib import Path

P = Path(__file__).resolve().parent
rows = []
for command in [['uname', '-a'], ['lscpu'], ['rustc', '-Vv'], ['cargo', '-V'],
                ['perf', '--version'], ['python3', '--version'], ['ldd', '--version'],
                ['free', '-b'], ['df', '-T', str(P)], ['objdump', '--version']]:
    result = subprocess.run(command, capture_output=True, text=True)
    rows.append(dict(command=command, exit_code=result.returncode,
                     stdout=result.stdout, stderr=result.stderr))
    assert result.returncode == 0, command
assert 12 in os.sched_getaffinity(0)
(P / 'environment.json').write_text(json.dumps(dict(
    utc=datetime.datetime.now(datetime.timezone.utc).isoformat(), records=rows,
    allowed_cpus=sorted(os.sched_getaffinity(0)), measured_cpu=12,
    allocator='Rust default system allocator; no allocator instrumentation',
    host='shared; warm processing; no quiescence or cold-cache claim',
), indent=2) + '\n')
print('Current host/tool environment retained; CPU12 is allowed.')
