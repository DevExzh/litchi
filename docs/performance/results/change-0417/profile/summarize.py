#!/usr/bin/env python3
"""Strict period-weighted leaf/ancestor attribution for this CPU diagnostic."""
from pathlib import Path
from collections import Counter
import hashlib
import json
import re
import subprocess
import sys

root = Path(__file__).resolve().parent
bundle = root.parent
repo = bundle.parents[3]
sys.path.insert(0, str(repo))
from tools import summarize_crud_baseline as summary
from tools import validate_perf_corpus_binding as binding
from tools import perf_compare

def artifact_bytes(name):
    path = root / name
    if path.is_file():
        return path.read_bytes()
    return subprocess.check_output(['zstd', '-q', '-dc', str(path) + '.zst'])

build = json.loads((bundle / 'build-identity.json').read_text())
baseline = json.loads((bundle / 'normal/1-pptx_cross_copy_media_rich.json').read_text())['results'][0]
for mode in ('record', 'stat'):
    report = json.loads((root / f'{mode}.json').read_text())
    command = json.loads((root / f'{mode}-command.json').read_text())
    assert command['exit_code'] == 0
    assert command['revision'] == report['environment']['git_revision'] == build['revision']
    assert command['binary_sha256'] == report['binary_identity']['binary_sha256'] == build['binaries']['normal']['sha256']
    assert report['environment']['git_worktree_dirty'] is False
    assert report['environment']['cpu_affinity'] == '2'
    assert report['configuration']['samples_per_case'] == 20
    assert report['configuration']['warmup_iterations_per_case'] == 3
    assert report['tool']['instrumentation'] == 'none'
    assert len(report['results']) == 1
    row = report['results'][0]
    assert row['case'] == 'pptx_cross_copy_media_rich'
    assert row['corpus'] == baseline['corpus']
    assert row['output_sha256'] == baseline['output_sha256']
    summary._validate_elapsed(row, 20, mode)
    binding.validate_paths(root / f'{mode}.json', root / f'{mode}.catalog.json')
    perf_compare.validate_parallel_metrics(report)

header = re.compile(r'^(.+?)\s+(\d+)/(\d+)\s+([0-9]+\.[0-9]+):\s+(\d+)\s+cycles:u:\s*$')
frame = re.compile(r'^\s*[0-9a-f]+\s+(.*?)\s+\((.*)\)\s*$')
leaf = Counter()
ancestor = Counter()
blocks = []
current = None
for line in artifact_bytes('script.txt').decode().splitlines() + ['']:
    if not line.strip():
        if current is not None:
            assert current['frames'], 'empty stack'
            period = current['period']
            leaf[current['frames'][0]] += period
            for symbol in set(current['frames']):
                ancestor[symbol] += period
            blocks.append(current)
            current = None
        continue
    matched = header.fullmatch(line)
    if matched:
        assert current is None, 'unterminated stack'
        current = dict(period=int(matched[5]), frames=[])
        assert current['period'] > 0
    else:
        matched = frame.fullmatch(line)
        assert current is not None and matched is not None, f'unparsed perf line: {line!r}'
        current['frames'].append(matched[1])
total = sum(leaf.values())
assert total > 0
text = (root / 'self.txt').read_text()
assert '# Total Lost Samples: 0' in text
assert int(re.search(r'Event count \(approx\.\): (\d+)', text)[1]) == total

def rows(counter, count=20):
    return [dict(symbol=name, period=value, percent=value * 100 / total)
            for name, value in counter.most_common(count)]

counters = {}
for line in (root / 'stat.csv').read_text().splitlines():
    if not line or line.startswith('#'):
        continue
    fields = line.split(',')
    value, unit, event, runtime, running = fields[:5]
    counters[event] = dict(raw=value, unit=unit, counter_runtime_ns=int(runtime),
                           running_percent=float(running))
    if value == '<not supported>':
        counters[event]['status'] = 'unavailable'
    elif event.startswith('L1-'):
        counters[event]['status'] = 'unvalidated_alias_zero_not_a_measurement'
    else:
        counters[event].update(status='measured', value=float(value))

result = dict(
    scope='Whole profiled command and descendants, including setup, preflight, warmups, verification and reporting; periods are event weights, not elapsed time',
    samples=len(blocks), total_period=total, lost_samples=0,
    unresolved_leaf_percent=sum(value for name, value in leaf.items() if '[unknown]' in name) * 100 / total,
    leaf=rows(leaf), inclusive_ancestors=rows(ancestor, 40),
    deflate_leaf_percent=sum(value for name, value in leaf.items() if name.startswith(('zlib_rs::deflate', '<zlib_rs::deflate'))) * 100 / total,
    counters=counters,
    whole_command_ipc=counters['instructions']['value'] / counters['cycles']['value'],
    whole_command_branch_miss_percent=counters['branch-misses']['value'] * 100 / counters['branches']['value'],
    perf_version=json.loads((root / 'tool-identity.json').read_text())['version'],
    bindings={name: hashlib.sha256(artifact_bytes(name)).hexdigest()
              for name in ('script.txt', 'self.txt', 'stat.csv', 'record-command.json', 'stat-command.json')},
    performance_claim='none',
)
(root / 'summary.json').write_text(json.dumps(result, indent=2) + '\n')
print(json.dumps({key: result[key] for key in ('samples', 'total_period', 'unresolved_leaf_percent',
                 'deflate_leaf_percent', 'whole_command_ipc', 'whole_command_branch_miss_percent')}))
