#!/usr/bin/env python3
"""Replay the frozen policy comparison; allocator timing is never normal timing."""
import math
import re
import statistics
from common import ROOT, read, write


def stats(values):
    assert values and all(type(value) is int and value >= 0 for value in values)
    ordered = sorted(values)
    mean = statistics.mean(values)
    half = 2.045 * statistics.stdev(values) / math.sqrt(len(values)) if len(values) > 1 else 0
    return dict(n=len(values), mean=mean, minimum=min(values), maximum=max(values),
                p50=ordered[math.ceil(len(values) * .5) - 1],
                p95=ordered[math.ceil(len(values) * .95) - 1],
                p99=ordered[math.ceil(len(values) * .99) - 1],
                mean_t95_interval=[mean - half, mean + half])


def percent(before, after):
    return 100 * (after - before) / before if before else None


def derive():
    protocol = read(ROOT / 'protocol.json')
    rows = {}
    identities = {}
    for capture in protocol['captures']:
        label = capture['label']
        report = read(ROOT / 'captures' / f'{label}.report.json')
        assert report['schema'] == 'zip-directory-spool-v1'
        assert (report['samples'], report['warmups'], report['repeats']) == (30, 3, 1)
        assert report['counts'] == [capture['count']]
        assert report['methods'] == [capture['method']]
        assert report['modes'] == [capture['policy']]
        assert report['spool_max_bytes'] == 64 * 1024 * 1024
        assert report['spool_buffer_bytes'] == 16384
        operations = report['operations']
        assert len(operations) == 30
        assert {row['sample'] for row in operations} == set(range(30))
        oracle = report['cases']
        assert len(oracle) == 2
        assert {item['mode'] for item in oracle} == {'control', 'spool'}
        projection = None
        for item in oracle:
            assert item['every_member_reopened'] is True
            assert item['byte_exact_control_match'] is True
            assert item['member_count'] == capture['count']
            assert item['method'] == capture['method']
            current = item['output_bytes'], item['output_sha256']
            assert projection is None or projection == current
            projection = current
        for row in operations:
            assert row['mode'] == capture['policy'] and row['method'] == capture['method']
            assert row['member_count'] == capture['count'] and row['repeat'] == 0
            assert row['output_matches_oracle'] is True
            assert (row['output_bytes'], row['output_sha256']) == projection
            assert row['source_bytes'] == capture['count'] * 256
            assert re.fullmatch('[0-9a-f]{64}', row['source_sha256'])
            identity = (row['source_sha256'], row['source_bytes'], *projection)
            key = str(capture['count']) + '-' + capture['method']
            assert key not in identities or identities[key] == identity
            identities[key] = identity
            if capture['policy'] == 'control':
                assert row['spool_bytes'] is None
            else:
                assert row['spool_bytes'] == capture['count'] * (46 + len('ppt/slides/slide-00000.xml'))
            if capture['instrumentation'] == 'allocator':
                allocation = row['allocation']
                assert allocation['status'] == 'measured'
                assert allocation['failed_allocation_calls'] == 0
                assert allocation['live_bytes_before'] == allocation['live_bytes_after']
                assert allocation['region_peak_live_bytes'] >= allocation['live_bytes_before']
            else:
                assert row['allocation'] is None
        row = dict(capture=capture, normal_timing=capture['instrumentation'] == 'normal',
                   elapsed_ns=stats([item['elapsed_ns'] for item in operations]),
                   output_bytes=projection[0], output_sha256=projection[1],
                   output_write_calls=stats([item['output_write_calls'] for item in operations]))
        resource = (ROOT / 'captures' / f'{label}.resource').read_text()
        match = re.search(r'Maximum resident set size \(kbytes\): (\d+)', resource)
        assert match
        row['process_max_rss_kib'] = int(match[1])
        if all(item['process'] is not None for item in operations):
            row['process_observer_deltas'] = {
                metric: stats([item['process'][metric] for item in operations])
                for metric in ('rchar', 'wchar', 'syscr', 'syscw', 'read_bytes', 'write_bytes',
                               'minor_faults', 'major_faults')}
        if capture['instrumentation'] == 'allocator':
            row['allocation_calls'] = stats([item['allocation']['allocation_calls'] for item in operations])
            row['requested_bytes'] = stats([item['allocation']['allocated_bytes'] for item in operations])
            row['incremental_peak_live_bytes'] = stats([
                item['allocation']['region_peak_live_bytes'] - item['allocation']['live_bytes_before']
                for item in operations])
        rows[label] = row
    pairs = []
    drifts = []
    for instrumentation in ('normal', 'allocator'):
        for count in (8, 256, 8192):
            for method in ('store', 'deflate'):
                for repeat in (1, 2):
                    before = rows[f'r{repeat}-{instrumentation}-{count}-{method}-control']
                    after = rows[f'r{repeat}-{instrumentation}-{count}-{method}-spool']
                    changes = dict(process_max_rss_kib=percent(before['process_max_rss_kib'], after['process_max_rss_kib']))
                    if instrumentation == 'normal':
                        changes.update({f'elapsed_{metric}': percent(before['elapsed_ns'][metric], after['elapsed_ns'][metric])
                                        for metric in ('mean', 'p50', 'p95', 'p99')})
                    else:
                        changes.update({metric: percent(before[metric]['mean'], after[metric]['mean'])
                                        for metric in ('allocation_calls', 'requested_bytes', 'incremental_peak_live_bytes')})
                    pairs.append(dict(instrumentation=instrumentation, count=count, method=method,
                                      repeat=repeat, percent_changes=changes,
                                      positive_review_flags=[key for key, value in changes.items()
                                                             if value is not None and value > 5]))
                for policy in ('control', 'spool'):
                    before = rows[f'r1-{instrumentation}-{count}-{method}-{policy}']
                    after = rows[f'r2-{instrumentation}-{count}-{method}-{policy}']
                    changes = dict(process_max_rss_kib=percent(before['process_max_rss_kib'], after['process_max_rss_kib']))
                    if instrumentation == 'normal':
                        changes.update({f'elapsed_{metric}': percent(before['elapsed_ns'][metric], after['elapsed_ns'][metric])
                                        for metric in ('mean', 'p50', 'p95', 'p99')})
                    drifts.append(dict(instrumentation=instrumentation, count=count, method=method,
                                       policy=policy, percent_changes=changes,
                                       absolute_review_flags=[key for key, value in changes.items()
                                                              if value is not None and abs(value) > 5]))
    return dict(schema='zip-directory-spool-summary-v1', captures=len(rows), samples=len(rows) * 30,
                rows=rows, identities={key: list(value) for key, value in identities.items()},
                pairs=pairs, repeat_drifts=drifts,
                uncertainty='nearest-rank percentiles; mean interval uses t(29)=2.045; '
                            'independent process repeats retained, no registered latency claim')


if __name__ == '__main__':
    write(ROOT / 'summary.json', derive())
