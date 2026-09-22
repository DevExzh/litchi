#!/usr/bin/env python3
"""Post-hoc process-order sensitivity analysis; no new benchmark execution.

All windows are descriptive and retain the original full-window rejection.
Neither samples nor overlapping windows are independent process repetitions.
"""
import hashlib
import json
from pathlib import Path
import statistics as st

P = Path(__file__).resolve().parent
OLD = P.parent / 'change-0735'
SEAL = 'dbe26c50532792cba24871862c96dc0889b3fc21626cb57357fab6a08ff03955'
WINDOWS = {'all': (0, 50), 'first10': (0, 10), 'middle30': (10, 40),
           'last10': (40, 50), 'first25': (0, 25), 'last25': (25, 50)}


def sha(p):
    assert p.is_file() and not p.is_symlink(), p
    return hashlib.sha256(p.read_bytes()).hexdigest()


def median(xs):
    # Match 0735 midpoint p50, including even-sized windows.
    return st.median(xs)


def percent(before, after):
    return 100 * (after / before - 1)


def summary(xs):
    return dict(count=len(xs), minimum=min(xs), median=st.median(xs), maximum=max(xs),
                positive=sum(x > 0 for x in xs), above_five=sum(x > 5 for x in xs),
                below_minus_five=sum(x < -5 for x in xs))


def verify_packet():
    assert sha(OLD / 'artifact-manifest.json') == SEAL
    files = json.loads((OLD / 'artifact-manifest.json').read_text())['files']
    actual = {str(p.relative_to(OLD)) for p in OLD.rglob('*') if p.is_file()}
    assert actual == set(files) | {'artifact-manifest.json'}
    for name, row in files.items():
        p = OLD / name
        assert p.stat().st_size == row['bytes'] and sha(p) == row['sha256'], name


def main():
    verify_packet()
    manifest = json.loads((OLD / 'captures/manifest.json').read_text())
    native = [r for r in manifest['runs'] if r['lane'] == 'native']
    assert len(native) == 36
    data = {}
    processes = []
    for order, row in enumerate(native):
        report = json.loads((OLD / 'captures' / row['output']).read_text())
        samples = report['samples']
        assert [s['index'] for s in samples] == list(range(50))
        assert report['warmups'] == 3 and report['samples_requested'] == 50
        assert all(s['oracle']['oracle_ok'] for s in samples)
        ns = [s['phase_ns']['whole_ns'] for s in samples]
        assert all(type(n) is int and n > 0 for n in ns)
        key = (row['case'], row['cycle'], row['repeat'], row['variant'])
        assert key not in data
        data[key] = (order, ns)
        windows = {w: dict(p50_ns=median(ns[a:b]), mean_ns=st.mean(ns[a:b]))
                   for w, (a, b) in WINDOWS.items()}
        processes.append(dict(case=row['case'], cycle=row['cycle'], repeat=row['repeat'],
                              variant=row['variant'], process_order=order, windows=windows,
                              last10_vs_first10_percent=percent(median(ns[:10]), median(ns[-10:])),
                              serialized_sample_bytes=sum(len(json.dumps(s, separators=(',', ':')).encode())
                                                          for s in samples)))
    pairs = []
    for case in ('primary', 'secondary'):
        for cycle in range(3):
            for repeat in range(3):
                bo, b = data[(case, cycle, repeat, 'baseline')]
                co, c = data[(case, cycle, repeat, 'candidate')]
                windows = {w: dict(p50_percent=percent(median(b[a:z]), median(c[a:z])),
                                   mean_percent=percent(st.mean(b[a:z]), st.mean(c[a:z])))
                           for w, (a, z) in WINDOWS.items()}
                pairs.append(dict(case=case, cycle=cycle, repeat=repeat,
                                  first_variant='baseline' if bo < co else 'candidate', windows=windows))
    groups = {}
    for case in ('primary', 'secondary'):
        groups[case] = {}
        for first in ('all', 'baseline', 'candidate'):
            selected = [p for p in pairs if p['case'] == case and
                        (first == 'all' or p['first_variant'] == first)]
            groups[case][first] = {w: {metric: summary([p['windows'][w][metric] for p in selected])
                                     for metric in ('p50_percent', 'mean_percent')}
                                  for w in WINDOWS}
    drift = {case: {v: summary([p['last10_vs_first10_percent'] for p in processes
                               if p['case'] == case and p['variant'] == v])
                    for v in ('baseline', 'candidate')} for case in ('primary', 'secondary')}
    original = json.loads((OLD / 'analysis.json').read_text())
    for case in ('primary', 'secondary'):
        prior = next(r for r in original['native_case_summaries'] if r['case'] == case)
        for metric in ('p50', 'mean'):
            got = groups[case]['all']['all'][metric + '_percent']['median']
            assert abs(got - prior['metrics'][metric]['median_percent']) < 1e-10
    result = dict(status='passed', scope='Post-hoc descriptive sensitivity; original rejection stands.',
                  input_seal_sha256=SEAL, windows={k: list(v) for k, v in WINDOWS.items()},
                  quantile='midpoint p50 within each process; ordinary median across process-pair ratios',
                  serialized_bytes_scope='Compact UTF-8 JSON size only; not live heap bytes or RSS.',
                  processes=processes, pairs=pairs, groups=groups, within_process_drift=drift)
    (P / 'order-analysis.json').write_text(json.dumps(result, indent=2) + '\n')
    print(json.dumps(dict(status='passed', processes=len(processes), samples=1800, pairs=len(pairs))))


if __name__ == '__main__':
    main()
