#!/usr/bin/env python3
"""Export retained profiles after all measurements, without debuginfod access."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess

BUNDLE = Path(__file__).resolve().parent


def main():
    jobs = [
        ('samples/perf-script', ['perf', 'script', '--no-inline', '-i', str(BUNDLE / 'samples/perf.data'),
                                '-F', 'comm,pid,tid,time,event,period,ip,sym,dso']),
        ('samples/top-symbols', ['perf', 'report', '--stdio', '--no-children', '--no-inline',
                                '-i', str(BUNDLE / 'samples/perf.data'), '--sort', 'symbol', '--percent-limit', '0.5']),
        ('heaptrack/print', ['heaptrack_print', '-f', str(BUNDLE / 'heaptrack/heaptrack.zst'),
                             '-n', '30']),
    ]
    env = dict(os.environ, DEBUGINFOD_URLS='', LC_ALL='C')
    rows = []
    for name, argv in jobs:
        start = datetime.datetime.now(datetime.timezone.utc).isoformat()
        out, err = BUNDLE / (name + '.txt'), BUNDLE / (name + '.stderr')
        with out.open('x') as stdout, err.open('x') as stderr:
            result = subprocess.run(argv, stdout=stdout, stderr=stderr, env=env)
        rows.append(dict(argv=argv, started_utc=start, exit_code=result.returncode,
                         finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                         artifacts={str(p.relative_to(BUNDLE)): hashlib.sha256(p.read_bytes()).hexdigest()
                                    for p in (out, err)}))
        (BUNDLE / 'postprocess.json').write_text(json.dumps(dict(environment={'DEBUGINFOD_URLS': '', 'LC_ALL': 'C'},
                                                               commands=rows), indent=2) + '\n')
        print(name, result.returncode, flush=True)
        if result.returncode:
            raise SystemExit(result.returncode)


if __name__ == '__main__':
    main()
