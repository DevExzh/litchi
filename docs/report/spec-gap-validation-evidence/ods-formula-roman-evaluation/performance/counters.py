#!/usr/bin/env python3
"""Capture serial pinned hardware counters after the timed corpus is complete."""
import datetime as dt
import hashlib
import json
from pathlib import Path
import subprocess

PERF = Path(__file__).resolve().parent

def sha(path): return hashlib.sha256(path.read_bytes()).hexdigest()
def now(): return dt.datetime.now(dt.timezone.utc).isoformat()

def main():
    out = PERF / 'perf-stat'; out.mkdir(exist_ok=True)
    lanes = [('baseline', 'comparable', 'text-if-concat', 100000),
             ('candidate', 'comparable', 'text-if-concat', 100000),
             ('candidate', 'roman', 'roman-3888-format-0', 10000),
             ('candidate', 'roman', 'roman-3888-format-4', 10000),
             ('candidate', 'roman', 'arabic-input-4096', 1000)]
    for variant, group, case, repeat in lanes:
        binary = PERF / variant / 'ods-formula-roman-evaluation-profile'
        stem = out / (variant + '-' + case)
        argv = ['perf', 'stat', '-x,', '-e', 'cycles,instructions,branches,branch-misses',
                '--', 'taskset', '-c', '6', str(binary), '--workload', 'roman-evaluation', '--revision', variant,
                '--group', group, '--phase', 'evaluate', '--case', case,
                '--warmups', '3', '--iterations', '15', '--repeat', str(repeat)]
        before = sha(binary); started = now()
        with stem.with_suffix('.stdout').open('w') as stdout, stem.with_suffix('.stderr').open('w') as stderr:
            result = subprocess.run(argv, stdout=stdout, stderr=stderr)
        finished = now(); assert before == sha(binary)
        stem.with_suffix('.status').write_text(str(result.returncode) + '\n')
        receipt = dict(argv=argv, status=result.returncode, started_at=started, finished_at=finished,
                       binary_sha256=before, binary_bytes=binary.stat().st_size,
                       stdout_sha256=sha(stem.with_suffix('.stdout')),
                       stderr_sha256=sha(stem.with_suffix('.stderr')))
        stem.with_suffix('.metadata.json').write_text(json.dumps(receipt, indent=2) + '\n')
        result.check_returncode()
        print(variant, case, 'captured', flush=True)

if __name__ == '__main__': main()
