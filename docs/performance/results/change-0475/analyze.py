#!/usr/bin/env python3
"""Portable normal-run and process-counter context for the 0475 profiles."""
import argparse
import hashlib
import json
import math
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent


def read(path):
    return json.loads(path.read_text())


def binding(path, root):
    return dict(path=path.relative_to(root).as_posix(), bytes=path.stat().st_size,
                sha256=hashlib.sha256(path.read_bytes()).hexdigest())


def counters(path, expected):
    result = {}
    for line in path.read_text().splitlines():
        if not line or line.startswith('#'):
            continue
        fields = line.split(';')
        assert len(fields) == 7 and fields[1] == fields[5] == fields[6] == '', line
        value, _, event, runtime, percent, _, _ = fields
        assert event in expected and event not in result, event
        assert runtime.isdecimal() and math.isfinite(float(percent)), line
        assert 0 <= float(percent) <= 100, line
        if value in ('<not supported>', '<not counted>'):
            status, count = value[1:-1].replace(' ', '_'), None
        else:
            assert value.isdecimal(), value
            status, count = 'reported', int(value)
        result[event] = dict(status=status, count=count, event_runtime_ns=int(runtime),
                             running_percent=float(percent))
    assert set(result) == set(expected)
    return result


def analyze(root):
    protocol = read(root / 'protocol.json')
    lanes = {}
    for item in protocol['order']:
        path = root / item['lane'] / 'report.json'
        report = read(path)
        row = report['results'][0]
        elapsed = row['elapsed_ns']
        rss = re.search(r'Maximum resident set size \(kbytes\): (\d+)',
                        (path.parent / 'resource.log').read_text())
        assert rss
        lanes[item['lane']] = dict(kind=item['kind'], samples=item['samples'],
            warmups=item['warmups'], report=binding(path, root),
            elapsed_ns=elapsed, whole_process_max_rss_kib=int(rss.group(1)),
            output_sha256=row['output_sha256'], sink=row['sink'])
    drift = {}
    for key in ('mean', 'p50', 'p95', 'p99'):
        a, b = (lanes[x]['elapsed_ns'][key] for x in ('normal-R1', 'normal-R2'))
        percent = 100 * (b / a - 1)
        drift[key] = dict(r1_ns=a, r2_ns=b, signed_percent=percent,
                         exceeds_review_threshold=abs(percent) > protocol['normal_repeat_review_threshold_percent'])
    pmu = counters(root / 'counters/counters.csv', protocol['counter_events'].split(','))
    return dict(schema='litchi-0475-normal-counter-context-v1', lanes=lanes,
        normal_repeat_drift=drift, normal_repeat_review_required=any(x['exceeds_review_threshold'] for x in drift.values()),
        whole_process_counters=pmu,
        counter_source=binding(root / 'counters/counters.csv', root),
        limitations=[
            'Normal runs bracket profilers on unchanged source and binary; this is descriptive repeatability, not a before/after or registered latency comparison.',
            'Profiler-instrumented elapsed vectors are diagnostic and remain separate from normal timing.',
            'perf stat values cover whole process including preflight and are scaled by perf for multiplexing; they are not operation-local counts.',
            'A reported zero hardware event is retained as reported and does not establish absence of cache misses; unsupported events remain unavailable.',
            'Whole-process RSS includes setup and profilers and cannot isolate writer heap or attribute its peak.',
        ])


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=ROOT)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    result = analyze(args.root)
    output = args.root / 'summary.json'
    if args.check:
        assert read(output) == result, 'summary does not replay exactly'
        print('normal/counter summary replay passed')
    else:
        output.write_text(json.dumps(result, indent=2, sort_keys=True) + '\n')
        print('normal/counter summary written')
