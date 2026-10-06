#!/usr/bin/env python3
"""Descriptive reduction of all scheduled 0940 children; never an adoption gate.

Input index: {manifest: descriptor, independent_review: descriptor, children:
[{id, arm, case, route, block, report: descriptor}]}. Descriptors contain path,
bytes and sha256. The index must be generated from the frozen full schedule
after independent packet validation, not by selecting reports by outcome.
The independent review must certify all terminal, source, binary and semantic
checks; this reducer checks its exact manifest and scheduled-children index
digest. It does not replace that full packet checker.
"""
import argparse
import collections
import hashlib
import json
from pathlib import Path
import random
import statistics

ARMS = ('normal_off', 'normal_on', 'allocator_on')
ROUTES = ('owned_vec', 'owned_reader')
SEGMENTS = ('prep', 'observer_gap', 'open', 'combined_prep_gap_open')


def read_bound(d):
    data = Path(d['path']).read_bytes()
    assert len(data) == d['bytes'], d['path']
    assert hashlib.sha256(data).hexdigest() == d['sha256'], d['path']
    return json.loads(data)


def summary(values):
    if not values:
        return {'count': 0, 'median': None, 'min': None, 'max': None}
    return {'count': len(values), 'median': statistics.median(values),
            'min': min(values), 'max': max(values)}


def chronological(values, order):
    assert len(values) == len(order)
    assert sorted(order) == list(range(len(order)))
    result = [None] * len(values)
    for rank, acquisition in enumerate(order):
        result[acquisition] = values[rank]
    return result


def reduce_child(entry):
    report = read_bound(entry['report'])
    assert report['diagnostic_arm'] == entry['arm']
    assert report['route'] == entry['route'] and report['case'] == entry['case']
    assert report['sample_count'] == 50 and report['warmup_iterations'] == 5
    op = report['operation_elapsed_ns']
    wall = report['wall_elapsed_ns']
    order = report['sample_order']
    assert order == op['sample_order']
    op_chrono = chronological(op['samples'], op['sample_order'])
    wall_chrono = chronological(wall['samples'], wall['sample_order'])
    fields = ['preparation_ns', 'open_ns', 'observer_gap_ns']
    if entry['case'] == 'workload':
        fields += ['query_ns', 'edit_ns', 'commit_ns', 'publication_ns']
    phases = {key: chronological(report[key], order) for key in fields}
    for i in range(50):
        assert op_chrono[i] == sum(v[i] for k, v in phases.items() if k != 'observer_gap_ns')
        assert wall_chrono[i] >= op_chrono[i] + phases['observer_gap_ns'][i]
    timing = {key: summary(value) for key, value in phases.items()}
    timing['operation_ns'] = summary(op_chrono)
    timing['wall_ns'] = summary(wall_chrono)
    timing['wall_minus_operation_ns'] = summary([w-o for w,o in zip(wall_chrono,op_chrono)])
    out = {key: entry[key] for key in ('id','arm','case','route','block','report')}
    out['timing'] = timing
    out['exploratory_preparation_ge_100000ns_count'] = sum(x >= 100000 for x in phases['preparation_ns'])
    out['allocation'] = None
    if report['allocation'] is not None:
        out['allocation'] = {}
        for phase in SEGMENTS:
            rows = report['allocation'][phase]
            metrics = sorted({k for row in rows for k,v in row.items() if type(v) is int})
            out['allocation'][phase] = {
                'statuses': dict(collections.Counter(row['status'] for row in rows)),
                'metrics': {k: summary([row[k] for row in rows if row['status']=='measured']) for k in metrics},
            }
    out['process'] = None
    if report['process'] is not None:
        out['process'] = {}
        for phase in SEGMENTS:
            rows = chronological(report['process'][phase], order)
            deltas = [row['delta'] for row in rows if row['delta'] is not None]
            metrics = sorted({k for row in deltas for k,v in row.items() if type(v) is int})
            out['process'][phase] = {
                'before_statuses': dict(collections.Counter(row['before_status'] for row in rows)),
                'after_statuses': dict(collections.Counter(row['after_status'] for row in rows)),
                'available_delta_count': len(deltas),
                'delta_metrics': {k: summary([row[k] for row in deltas]) for k in metrics},
                'signed_rss_delta': summary([row['rss_delta_signed'] for row in rows if row['rss_delta_signed'] is not None]),
                'signed_rss_statuses': dict(collections.Counter(row['rss_delta_status'] for row in rows)),
            }
            if phase in ('prep','open'):
                elapsed = phases['preparation_ns' if phase=='prep' else 'open_ns']
                strata = {'zero_minor_faults': [], 'nonzero_minor_faults': [], 'unavailable': []}
                for t,row in zip(elapsed,rows):
                    key = 'unavailable' if row['delta'] is None else ('nonzero_minor_faults' if row['delta']['minor_faults'] else 'zero_minor_faults')
                    strata[key].append(t)
                out['process'][phase]['elapsed_by_minor_fault_status'] = {k:summary(v) for k,v in strata.items()}
    return out


def reduce(index):
    read_bound(index['manifest'])
    review = read_bound(index['independent_review'])
    assert review['status'] == 'passed' and review['child_count'] == 80
    entries = index['children']
    assert review['manifest_sha256'] == index['manifest']['sha256']
    assert review['children_index_sha256'] == hashlib.sha256(
        json.dumps(entries, sort_keys=True, separators=(',', ':')).encode()).hexdigest()
    assert len(entries) == 80 and len({e['id'] for e in entries}) == 80
    work = [e for e in entries if e['case']=='workload']
    calibration = [e for e in entries if e['case']=='calibration']
    expected = {(a,r,b) for a in ARMS for r in ROUTES for b in range(1,13)}
    assert len(work)==72 and {(e['arm'],e['route'],e['block']) for e in work}==expected
    expected_cal = {(a,'route_independent',b) for a in ARMS[1:] for b in (0,4,8,12)}
    assert len(calibration)==8 and {(e['arm'],e['route'],e['block']) for e in calibration}==expected_cal
    children = [reduce_child(e) for e in entries]
    by_key = {(e['arm'],e['route'],e['block']):e for e in children if e['case']=='workload'}
    groups = []
    for arm in ARMS:
        for route in ROUTES:
            rows = [by_key[arm,route,b] for b in range(1,13)]
            groups.append({'arm':arm,'route':route,'timing_child_medians':{
                metric: {'values':[r['timing'][metric]['median'] for r in rows],
                         **summary([r['timing'][metric]['median'] for r in rows])}
                for metric in rows[0]['timing']}})
    contrasts = []
    for route_index,route in enumerate(ROUTES):
        rng = random.Random(94010000+route_index)
        # One resampling table shared by all descriptive metrics for this route.
        draws = [[rng.randrange(12) for _ in range(12)] for _ in range(10000)]
        for metric in ('operation_ns','wall_ns','preparation_ns','open_ns'):
            ratios = [by_key['normal_on',route,b]['timing'][metric]['median'] /
                      by_key['normal_off',route,b]['timing'][metric]['median'] for b in range(1,13)]
            boot = sorted(statistics.median([ratios[i] for i in draw]) for draw in draws)
            contrasts.append({'route':route,'metric':metric,'on_over_off_ratios':ratios,
                              'median':statistics.median(ratios),'bootstrap_percentile_95':[boot[249],boot[9749]],
                              'seed':94010000+route_index,'draws':10000,'qualification':False})
    route_contrasts = []
    for arm in ARMS:
        for metric in ('operation_ns','wall_ns','preparation_ns','open_ns'):
            ratios = [by_key[arm,'owned_reader',b]['timing'][metric]['median'] /
                      by_key[arm,'owned_vec',b]['timing'][metric]['median'] for b in range(1,13)]
            route_contrasts.append({'arm':arm,'metric':metric,'reader_over_vec_ratios':ratios,**summary(ratios)})
    return {'schema':'litchi.performance.0940.descriptive-analysis.v1','qualification':False,
            'manifest':index['manifest'],'independent_review':index['independent_review'],
            'workload_children':72,'calibration_children':8,'groups':groups,
            'normal_probe_contrasts':contrasts,'route_contrasts':route_contrasts,'children':children,
            'calibration_by_arm_and_boundary': [
                {'arm':arm,'boundary':boundary,
                 **{key: next(c for c in children if c['case']=='calibration'
                              and c['arm']==arm and c['block']==boundary)[key]
                    for key in ('id','timing','allocation','process')}}
                for arm in ARMS[1:] for boundary in (0,4,8,12)],
            'limits':['New normal binary reference only; no historical layout equivalence.',
                      'Allocator timings are separate diagnostic observations.',
                      'Process intervals contain observer work; no pure operation causality.',
                      'Calibration is separate and never subtracted.',
                      'Small temporally related child sample; intervals are descriptive.']}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('index',type=Path)
    parser.add_argument('output',type=Path)
    args = parser.parse_args()
    index = json.loads(args.index.read_bytes())
    result = reduce(index)
    with args.output.open('x') as stream:
        json.dump(result,stream,indent=2,sort_keys=True)
        stream.write('\n')


if __name__=='__main__':
    main()
