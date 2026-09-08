#!/usr/bin/env python3
"""Export whole-process Heaptrack data only after timed captures finish."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parent


def main():
    records = []
    for lane in ('A-heap-clean', 'B-heap-clean'):
        folder = ROOT / lane
        argv = ['heaptrack_print', '-f', str(folder / 'heaptrack.zst'), '-n', '30']
        record = dict(argv=argv, lane=lane, started_utc=datetime.datetime.now(datetime.timezone.utc).isoformat())
        with (folder / 'print.txt').open('x') as out, (folder / 'print.stderr').open('x') as err:
            result = subprocess.run(argv, stdout=out, stderr=err, env=dict(os.environ, DEBUGINFOD_URLS='', LC_ALL='C'))
        record.update(exit_code=result.returncode, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(),
                      artifacts={name: hashlib.sha256((folder / name).read_bytes()).hexdigest() for name in ('print.txt', 'print.stderr')})
        records.append(record)
        (ROOT / 'postprocess.json').write_text(json.dumps(dict(environment={'DEBUGINFOD_URLS': '', 'LC_ALL': 'C'}, commands=records), indent=2)+'\n')
        print(lane, result.returncode, flush=True)
        if result.returncode:
            raise SystemExit(result.returncode)


if __name__ == '__main__':
    main()
