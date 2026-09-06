#!/usr/bin/env python3
"""Recompute the predeclared 0438 retention gate from the verified summary."""
import argparse
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def derive(root=ROOT, summary_path=None, gate=None):
    root = Path(root)
    summary_path = Path(summary_path or root / 'summary.json')
    summary = json.loads(summary_path.read_text())
    rows = {(r['role'], r['mode'], r['shape'], r['repeat']): r for r in summary['rows']}
    conditions = []
    for shape in ['medium', 'large']:
        for repeat in ['R1', 'R2']:
            before, after = [rows[(role, 'normal', shape, repeat)]['elapsed_ns']
                             for role in ['before-streaming', 'after-streaming']]
            gain = 100 * (1 - after['p50'] / before['p50'])
            nonoverlap = after['confidence_interval_95']['upper'] < before['confidence_interval_95']['lower']
            conditions.append({'shape': shape, 'repeat': repeat, 'p50_improvement_percent': gain,
                               'p50_gate_pass': gain >= 5.0, 'favorable_mean_ci_nonoverlap': nonoverlap,
                               'before_p50_ns': before['p50'], 'after_p50_ns': after['p50'],
                               'before_mean_ci95': before['confidence_interval_95'],
                               'after_mean_ci95': after['confidence_interval_95']})
    allocation = [r for r in summary['comparisons'] if r['metric'] in [
        'allocation.allocation_calls', 'allocation.allocated_bytes', 'allocation.region_peak_above_entry']]
    assert len(allocation) == 18
    allocation_pass = all(r['after'] <= r['before'] for r in allocation)
    accepted = allocation_pass and all(r['p50_gate_pass'] and r['favorable_mean_ci_nonoverlap'] for r in conditions)
    status = 'accepted' if accepted else 'rejected'
    return {'change': 438, 'status': status,
            'hypothesis_sha256': hashlib.sha256((root / 'before-hypothesis.json').read_bytes()).hexdigest(),
            'summary_sha256': hashlib.sha256(summary_path.read_bytes()).hexdigest(),
            'minimum_p50_improvement_percent': 5.0, 'conditions': conditions,
            'operation_allocation_nonincrease': allocation_pass,
            'allocation_comparisons': allocation,
            'review_flags': summary['review_flags'],
            'action': 'retain measured candidate' if accepted else 'restore baseline production and retain candidate source, patch, tests and negative evidence',
            'scope': 'Two repeated warm single-worker deterministic plain ODP fresh-creation comparisons; no broad Office, cold/range, scaling or operation-only profile claim.'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true')
    args = parser.parse_args()
    value = derive()
    target = ROOT / 'decision.json'
    if args.check:
        assert json.loads(target.read_text()) == value
    else:
        with target.open('x') as stream:
            stream.write(json.dumps(value, indent=2) + '\n')
    print(json.dumps({'status': value['status'], 'conditions': value['conditions']}))


if __name__ == '__main__':
    main()
