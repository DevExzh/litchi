#!/usr/bin/env python3
"""Validate frozen captures and compare independent process-level controls."""
import math
import random
import statistics as st
from contract import P, ROOT, command, read, sha, validate_report, write
from run import guard


def stats(xs):
    ordered = sorted(xs)
    result = dict(p50=st.median(xs), mean=st.mean(xs))
    if len(xs) > 1:
        result.update(p95=ordered[math.ceil(.95*len(xs))-1],
                      p99=ordered[math.ceil(.99*len(xs))-1], maximum=max(xs))
    return result


def pct(a, b):
    return 100*(b/a-1)


def summarize(values):
    plan = read(P / 'plan.json')
    rng = random.Random(plan['bootstrap_seed'])
    bootstrap = sorted(st.median(rng.choices(values, k=len(values)))
                       for _ in range(plan['bootstrap_resamples']))
    return dict(values=values, median=st.median(values), minimum=min(values), maximum=max(values),
                bootstrap_median_95=[bootstrap[math.ceil(.025*len(bootstrap))-1],
                                     bootstrap[math.ceil(.975*len(bootstrap))-1]],
                flagged_pairs=[i for i,x in enumerate(values) if abs(x)>plan['review_threshold_percent']])


def main():
    guard()
    plan = read(P / 'plan.json')
    manifest = read(P / 'captures/manifest.json')
    assert manifest['status'] == 'complete'
    assert manifest['freeze_sha256'] == sha(P / 'freeze.json')
    assert manifest['preflight_sha256'] == sha(P / 'preflight.json')
    assert read(P / 'preflight.json')['status'] == 'passed'
    assert read(P / 'preflight.json')['freeze_sha256'] == sha(P / 'freeze.json')
    assert len(manifest['runs']) == len(plan['schedule']) == 138
    expected_files = {'manifest.json'}
    processes = []
    lookup = {}
    previous_end = 0
    for row, expected in zip(manifest['runs'], plan['schedule']):
        assert {k:row[k] for k in expected} == expected
        assert row['command'] == command(expected) and row['exit_code'] == 0
        assert row['start_monotonic_ns'] >= previous_end
        assert row['end_monotonic_ns'] > row['start_monotonic_ns']
        previous_end = row['end_monotonic_ns']
        file = P / 'captures' / row['output']
        stderr = file.with_suffix('.stderr')
        expected_files.update([file.name, stderr.name])
        assert sha(file) == row['sha256'] and sha(stderr) == row['stderr_sha256']
        assert stderr.read_bytes() == b''
        report = validate_report(read(file), expected)
        key = (row['lane'], row['case'], row['arm'], row['repeat'])
        assert key not in lookup
        item = {k:row[k] for k in expected}
        if row['lane'] == 'native':
            ns = [s['phase_ns']['whole_ns'] for s in report['samples']]
            item['times_ns'] = ns
            item['stats'] = stats(ns)
            if len(ns) == 50:
                item['windows'] = {name:stats(ns[a:b]) for name,a,b in
                                   [('first10',0,10),('middle30',10,40),('last10',40,50)]}
                item['last10_vs_first10_percent'] = pct(st.median(ns[:10]),st.median(ns[-10:]))
        else:
            item['allocation'] = report['samples'][0]['allocations']['whole']
        processes.append(item)
        lookup[key] = item
    assert {f.name for f in (P/'captures').iterdir()} == expected_files
    comparisons = []
    for case in ('primary','secondary'):
        for left,right in plan['comparisons']:
            pairs = [(lookup['native',case,left,i],lookup['native',case,right,i]) for i in range(9)]
            metrics = {metric:summarize([pct(a['stats'][metric],b['stats'][metric]) for a,b in pairs])
                       for metric in ('p50','mean','p95','p99','maximum')}
            windows = {w:summarize([pct(a['windows'][w]['p50'],b['windows'][w]['p50']) for a,b in pairs])
                       for w in ('first10','middle30','last10')}
            comparisons.append(dict(case=case,left=left,right=right,metrics=metrics,windows=windows))
    allocation = []
    for case in ('primary','secondary'):
        for left,right in plan['allocation_comparisons']:
            pairs = [(lookup['allocation',case,left,i]['allocation'],
                      lookup['allocation',case,right,i]['allocation']) for i in range(3)]
            allocation.append(dict(case=case,left=left,right=right,
                fields={f:dict(before=[a[f] for a,b in pairs],after=[b[f] for a,b in pairs],
                               differences=[b[f]-a[f] for a,b in pairs]) for f in pairs[0][0]}))
    result = dict(status='passed',native_processes=108,allocation_processes=30,
                  native_samples=4518,processes=processes,comparisons=comparisons,
                  allocation_comparisons=allocation,
                  scope='Unchanged production owner; harness sensitivity only. Full-window gates retained; no candidate reinstatement.')
    write(P/'analysis.json',result)
    print('PASS 138 exact captures, 4518 native samples, eight native and eight allocation comparisons')


if __name__ == '__main__':
    main()
