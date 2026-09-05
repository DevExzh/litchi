#!/usr/bin/env python3
"""Render the verified, descriptive repeat statistics without pooling them."""
from pathlib import Path
import json
import re

root = Path(__file__).resolve().parent
summary = json.loads((root / 'summary.json').read_text())
lines = ['# Normal repeat observations', '',
         'All elapsed values are milliseconds, shown as repeat 1 / repeat 2.',
         'Read `timing-boundaries.md` for each selector. These intervals differ;',
         'cross-selector latency ratios do not represent speedups. The >5% flags',
         'use absolute signed repeat drift and preserve both faster and slower tails.', '',
         '| Selector | p50 ms | p95 ms | p99 ms | >5% drift |',
         '| --- | ---: | ---: | ---: | --- |']
for row in summary['normal']:
    values = [' / '.join(f'{repeat[field] / 1e6:.6f}' for repeat in row['repeats'])
              for field in ('p50_ns', 'p95_ns', 'p99_ns')]
    lines.append(f"| `{row['selector']}` | " + ' | '.join(values)
                 + ' | ' + (', '.join(row['repeat_drift_flags']) or 'none') + ' |')
lines.extend(['', 'IID median order-statistic intervals and full vectors remain in',
              '`summary.json` and the raw reports. These intervals do not model',
              'shared-host drift or establish independence of successive samples.', '',
              '## Throughput and whole-process RSS', '',
              'Throughput is the reciprocal of each scoped p50 (operations/second).',
              'It is not physical input bandwidth. RSS is `/usr/bin/time -v` process',
              'maximum, including corpus construction, setup, warmups, verification',
              'and retained output, and is not an operation-local memory peak.', '',
              '| Selector | Operations/s at p50, repeat 1 / 2 | Peak RSS KiB, repeat 1 / 2 |',
              '| --- | ---: | ---: |'])
for row in summary['normal']:
    rates = ' / '.join(f"{1e9 / repeat['p50_ns']:.3f}" for repeat in row['repeats'])
    rss = []
    for repeat in row['repeats']:
        timing = (root / repeat['report']).with_suffix('.time.txt').read_text()
        matches = re.findall(r'Maximum resident set size \(kbytes\): (\d+)', timing)
        assert len(matches) == 1
        rss.append(matches[0])
    lines.append(f"| `{row['selector']}` | {rates} | {' / '.join(rss)} |")
lines.extend(['', '## Available allocation attribution', '',
              'Both repeats have identical p50/p95/p99 allocation calls and bytes',
              'for the two attributed selectors. The other 28 remain unavailable.', '',
              '| Selector | Allocation calls | Allocated bytes |',
              '| --- | ---: | ---: |'])
for row in summary['allocator']:
    if row['allocation_delta']['status'] != 'measured':
        continue
    metrics = row['allocation_delta']['metrics']
    values = []
    for field in ('allocation_calls', 'allocated_bytes'):
        stats = metrics[field]
        assert len({stats[r][p] for r in ('repeat_1', 'repeat_2')
                    for p in ('p50', 'p95', 'p99')}) == 1
        values.append(str(stats['repeat_1']['p50']))
    lines.append(f"| `{row['selector']}` | {' | '.join(values)} |")
(root / 'baseline-table.md').write_text('\n'.join(lines) + '\n')
