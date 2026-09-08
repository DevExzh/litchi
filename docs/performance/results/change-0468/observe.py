#!/usr/bin/env python3
"""Recompute descriptive timing/counter observations and capture chronology."""
import datetime
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent
LANES = ('normal-r1', 'counters', 'samples-fp', 'normal-r2')


def summarize(root=ROOT):
    inputs = {}

    def read(name):
        data = (root / name).read_bytes()
        inputs[name] = hashlib.sha256(data).hexdigest()
        return data.decode()

    records = []
    last_finish = datetime.datetime.fromisoformat(json.loads(read('build.json'))['finished_utc'])
    for lane in LANES:
        capture = json.loads(read(f'{lane}/capture.json'))
        start = datetime.datetime.fromisoformat(capture['started_utc'])
        finish = datetime.datetime.fromisoformat(capture['finished_utc'])
        if capture['exit_code'] != 0 or not last_finish <= start <= finish:
            raise ValueError(f'{lane}: unsuccessful capture or overlapping task chronology')
        last_finish = finish
        report = json.loads(read(f'{lane}/report.json'))
        row = report['results'][0]
        rss = re.search(r'Maximum resident set size \(kbytes\):\s*(\d+)', read(f'{lane}/resource.log'))
        if rss is None:
            raise ValueError(f'{lane}: missing process RSS')
        records.append(dict(lane=lane, samples=len(row['elapsed_ns']['samples']),
                            statistics_ns={k: row['elapsed_ns'][k] for k in ('mean', 'p50', 'p95', 'p99')},
                            process_max_rss_kib=int(rss.group(1)), sink=row.get('sink')))
    for name in ('perf-script', 'top-symbols'):
        export = json.loads(read(f'samples-fp/{name}.json'))
        start = datetime.datetime.fromisoformat(export['started_utc'])
        finish = datetime.datetime.fromisoformat(export['finished_utc'])
        if export['exit_code'] != 0 or not last_finish <= start <= finish:
            raise ValueError(f'{name}: unsuccessful or overlapping postprocessing')
        last_finish = finish
    counters = {}
    for line in read('counters/counters.csv').splitlines():
        if not line or line.startswith('#'):
            continue
        fields = line.split(';')
        if len(fields) < 5:
            raise ValueError('malformed perf stat row')
        value, unit, event, runtime, running = fields[:5]
        if event in counters:
            raise ValueError(f'duplicate counter {event}')
        counters[event] = dict(value=int(value) if value.isdigit() else None,
                               raw_value=value, unit=unit, runtime_ns=int(runtime),
                               running_percent=float(running))

    def ratio(numerator, denominator, scale=1):
        a, b = counters[numerator]['value'], counters[denominator]['value']
        return a / b * scale if a is not None and b else None

    a, b = records[0], records[-1]
    return dict(schema='litchi-0468-observations-v1', inputs=inputs, runs=records,
                normal_same_build_drift_percent={k: (b['statistics_ns'][k] / a['statistics_ns'][k] - 1) * 100
                                                for k in ('mean', 'p50', 'p95', 'p99')},
                normal_process_rss_drift_percent=(b['process_max_rss_kib'] / a['process_max_rss_kib'] - 1) * 100,
                whole_process_counters=counters,
                whole_process_ipc=ratio('instructions:u', 'cycles:u'),
                whole_process_branch_miss_percent=ratio('branch-misses:u', 'branches:u', 100),
                chronology_verified=True,
                scope='Descriptive same-build timing context; counters include setup, warmup, expected-output work, verification and teardown. GNU time wraps perf in instrumented lanes and its RSS includes profiler overhead. No cross-build speedup or operation-local counter claim.')


if __name__ == '__main__':
    print(json.dumps(summarize(), indent=2, sort_keys=True))
