#!/usr/bin/env python3
"""Recompute and write the bounded 0491 descriptive results review."""
import json
from pathlib import Path
import statistics

import cold_matrix as cold
import provider_matrix as provider
from support import ROOT, read, write


def checked(driver, attempt):
    builds = driver.load_builds()
    if driver is provider:
        protocol, digest = driver._load_protocol()
        entries = driver._collect(attempt, builds, protocol, digest)
    else:
        entries = driver._collect(attempt, builds)
    summary = driver.analyze_data(entries, builds)
    assert summary == read(ROOT / 'analysis' / f'{attempt}.json')
    return summary


def main():
    p = checked(provider, 'provider-formal1')
    c = checked(cold, 'cold-formal1')
    facts = {'schema': 'docx-0491-descriptive-results-v1', 'provider_inventory': p['inventory'],
             'cold_inventory': c['inventory'], 'performance_claim': 'none',
             'provider_rows': [], 'cold_rows': [], 'repeat_review_flags': []}
    lines = ['# 0491 DOCX provider and cache-state baseline', '',
             'This is a descriptive baseline on the retained machine and source build, with no before/after optimization claim. '
             'Each formal process has three warmups and 30 retained samples; two repeats are reported separately. '
             'Quantiles describe these small samples, and repeat variation is not a confidence interval. The host is shared; CPU affinity is not exclusive CPU reservation. '
             'The filesystem raw reports also retain Student-t mean intervals. Tail estimates require more samples before a regression threshold can be calibrated.', '',
             '## Source-provider lifecycle', '',
             'The timer covers typed package open, document preparation, extraction and package/document destruction. '
             'Source construction and file staging/open are outside. The returned 10,000-byte text is verified after timing. '
             'Counters are logical calls; the range adapter cap/pacing is simulated. '
             'All available media-range proofs report zero returned compressed media overlap.', '',
             '| Role | Repeat | Provider | p50 ms | p95 ms | p99 ms | Calls | Inner short reads | Heap peak increment KiB | Whole-child RSS MiB |',
             '|---|---:|---|---:|---:|---:|---:|---:|---:|---:|']
    for row in p['rows']:
        v, q = row['raw_vectors'], row['percentiles']
        latency = q['latency_ns']
        heap = v.get('allocation_peak_increment_bytes')
        fact = {'label': row['label'], 'role': row['role'], 'repeat': row['repeat'], 'arm': row['arm'],
                'latency_ns': latency, 'paragraphs_per_second_at_p50': 200e9 / latency['p50'],
                'text_bytes_per_second_at_p50': 10000e9 / latency['p50'],
                'calls': v['logical_read_calls'], 'inner_short_reads': v.get('range_short_reads'),
                'heap_peak_increment_bytes': heap, 'allocated_bytes': v.get('allocation_allocated_bytes'),
                'allocation_calls': v.get('allocation_allocation_calls'), 'whole_child_rss_kbytes': v['whole_child_rss_kbytes']}
        facts['provider_rows'].append(fact)
        calls = 'unavailable' if v['logical_read_calls'][0] is None else str(v['logical_read_calls'][0])
        short = '—' if v.get('range_short_reads') is None else str(v['range_short_reads'][0])
        h = '—' if heap is None else f'{statistics.median(heap)/1024:.2f}'
        lines.append(f"| {row['role']} | {row['repeat']} | {row['arm']} | {latency['p50']/1e6:.3f} | {latency['p95']/1e6:.3f} | {latency['p99']/1e6:.3f} | {calls} | {short} | {h} | {v['whole_child_rss_kbytes']/1024:.2f} |")
    lines += ['', 'Heap peak increment subtracts region-start live bytes from the callback-order region peak. '
              'Allocated bytes and allocation/reallocation counts remain in JSON/raw reports. '
              'Absolute region peaks include pre-existing live memory; whole-child RSS includes corpus generation and diagnostics. '
              'Neither is a total of operation allocations. The short-read arm reports inner-adapter short reads separately from the outer wrapper.', '',
              '## Filesystem lifecycle', '',
              'This timer includes facade path open, full-text extraction and document destruction. '
              'It differs from the typed provider setup scope and is not an equal-scope comparison. '
              'Each sample runs in a fresh child. Verified-cold means sync, accepted DONTNEED, zero fincore residency/dirty/writeback, '
              'and positive process read_bytes during the lifecycle; it does not identify physical-device cache state.', '',
              '| Role | Repeat | State | p50 ms | p95 ms | p99 ms | Median read_bytes delta | Heap peak increment KiB | Child high-water RSS MiB |',
              '|---|---:|---|---:|---:|---:|---:|---:|---:|']
    for row in c['rows']:
        v, q = row['raw_vectors'], row['percentiles']; latency = q['elapsed_ns']
        heap = v.get('allocation_peak_increment_bytes'); reads = v.get('read_bytes_delta')
        facts['cold_rows'].append({'label': row['label'], 'role': row['role'], 'repeat': row['repeat'],
            'state': row['cache_state'], 'latency_ns': latency, 'heap_peak_increment_bytes': heap,
            'allocated_bytes': v.get('allocation_allocated_bytes'), 'read_bytes_delta': reads,
            'process_peak_rss_bytes': v.get('peak_rss_bytes'), 'paragraphs_per_second_at_p50': 200e9 / latency['p50']})
        h = '—' if heap is None else f'{statistics.median(heap)/1024:.2f}'
        d = '—' if reads is None else str(statistics.median(reads))
        rss = statistics.median(v['peak_rss_bytes']) / 2**20
        lines.append(f"| {row['role']} | {row['repeat']} | {row['cache_state']} | {latency['p50']/1e6:.3f} | {latency['p95']/1e6:.3f} | {latency['p99']/1e6:.3f} | {d} | {h} | {rss:.2f} |")
    for kind, rows, key in [('provider', facts['provider_rows'], 'arm'), ('filesystem', facts['cold_rows'], 'state')]:
        groups = {}
        for row in rows: groups.setdefault((row['role'], row[key]), {})[row['repeat']] = row
        for (role, arm), repeats in groups.items():
            for metric in ['p50', 'p95', 'p99']:
                first, second = [repeats[i]['latency_ns'][metric] for i in [1, 2]]
                change = 100 * (second-first) / first
                if abs(change) > 5:
                    facts['repeat_review_flags'].append({'kind':kind,'role':role,'arm':arm,'metric':metric,'repeat2_vs_repeat1_percent':change})
    lines += ['', '## Repeat variance and limits', '',
              f"{len(facts['repeat_review_flags'])} latency quantiles vary by more than 5% between repeats. These are review flags, not optimization regressions:", '',
              '| Matrix | Role | Arm/state | Metric | Repeat 2 vs 1 |', '|---|---|---|---|---:|']
    for f in facts['repeat_review_flags']:
        lines.append(f"| {f['kind']} | {f['role']} | {f['arm']} | {f['metric']} | {f['repeat2_vs_repeat1_percent']:+.2f}% |")
    lines += ['', 'Provider analysis also retains RSS/absolute-region-peak repeat flags; filesystem analysis retains individual tail-spread diagnostics. '
              'No adverse row is averaged away. Prepared-query controls are explicitly cold-ineligible and produce no timed result.', '',
              'Aligned cold archives contain 564 zero comment bytes. Their one 65,536-byte EOCD search overlaps 58,808 compressed media bytes and '
              '4,500 other unselected payload bytes, with zero successful cache loads at open. '
              'The replay preserves these raw overlaps. Main preparation materializes one part; the prepared query performs zero I/O. '
              'The aligned source is structurally proven and text-identical, but is not byte-identical to the original warm source.', '',
              'Profiles are whole-child diagnostics, including corpus/setup/report work. They cannot attribute CPU/syscalls to the timed operation. '
              'The stable delayed source has 19 paced calls; bounded range coalescing is a measured follow-up opportunity, subject to identity, budget and preservation tests. '
              'Genuine non-static borrowed input, concurrency/scaling, native producers, publication and the remaining CRUD intersections are still open.', '',
              'Reproduce with the commands in README.md. The final seal rechecks build, report, protocol, profile and cleanup custody. '
              'Failed development attempts remain separate from the formal evidence.', '']
    write(ROOT/'results-review.json',facts)
    with (ROOT/'results-review.md').open('x') as stream: stream.write('\n'.join(lines))

if __name__ == '__main__': main()
