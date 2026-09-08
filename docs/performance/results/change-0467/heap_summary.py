#!/usr/bin/env python3
"""Read exported whole-process Heaptrack totals; do not compare instrumented time."""
import hashlib
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent


def summarize(root=ROOT):
    roles = {}
    for role, lane in [('control', 'A-heap-clean'), ('candidate', 'B-heap-clean')]:
        path = root / lane / 'print.txt'
        raw = path.read_bytes()
        text = raw.decode('utf-8')
        def unique(pattern):
            matches = re.findall(pattern, text, re.MULTILINE)
            if len(matches) != 1:
                raise ValueError(f'{path}: expected one total matching {pattern!r}')
            return matches[0]
        roles[role] = dict(path=str(path.relative_to(root)), sha256=hashlib.sha256(raw).hexdigest(),
                           allocation_calls=int(unique(r'^calls to allocation functions: (\d+) \(')),
                           temporary_allocations=int(unique(r'^temporary memory allocations: (\d+) \(')),
                           rounded_peak_heap=unique(r'^peak heap memory consumption: (\S+)\s*$'))
    deltas = {}
    for metric in ('allocation_calls', 'temporary_allocations'):
        a, b = roles['control'][metric], roles['candidate'][metric]
        deltas[metric] = dict(candidate_minus_control=b-a, candidate_over_control=b/a,
                              reduction_percent=100*(1-b/a))
    return dict(schema='litchi-0467-whole-process-heap-summary-v1',
                scope='Separate Heaptrack runs, five samples plus one warmup; includes generation, expected output, setup, timed commit/write, verification, and teardown. Not allocations per commit or normal latency/RSS.',
                roles=roles, deltas=deltas,
                peak_heap_precision='Heaptrack human-readable rounded value only; no exact peak-byte reduction inferred.')


if __name__ == '__main__':
    print(json.dumps(summarize(), indent=2, sort_keys=True))
