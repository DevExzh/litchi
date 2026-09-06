#!/usr/bin/env python3
"""Re-derive the bounded ODT span retention decision from the verified summary."""
import argparse
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent

def derive():
    raw = (ROOT / 'summary.json').read_bytes()
    summary = json.loads(raw)
    rows = {(r['role'], r['mode'], r['shape'], r['repeat']): r for r in summary['rows']}
    comparisons = []
    def aligned(vectors):
        order = sorted(range(len(vectors['sample_indices'])), key=lambda i: vectors['sample_indices'][i])
        return {key: [values[i] for i in order] for key, values in vectors.items()}
    for shape in ('tiny', 'medium', 'large'):
        for repeat in ('R1', 'R2'):
            before = rows[('before-streaming', 'normal', shape, repeat)]
            after = rows[('after-streaming', 'normal', shape, repeat)]
            old, new = before['elapsed_ns'], after['elapsed_ns']
            assert new['p50_ci']['high'] < old['p50_ci']['low']
            a = rows[('before-streaming', 'allocator', shape, repeat)]['allocation']['raw_aligned_vectors']
            b = rows[('after-streaming', 'allocator', shape, repeat)]['allocation']['raw_aligned_vectors']
            assert aligned(a) == aligned(b)
            assert set(a['region_peak_above_entry']) == {420091}
            assert set(a['live_delta']) == set(a['live_balance']) == {0}
            comparisons.append({'shape': shape, 'repeat': repeat, 'before_p50_ns': old['p50'], 'after_p50_ns': new['p50'], 'p50_reduction_percent': (1 - new['p50'] / old['p50']) * 100, 'descriptive_bootstrap_intervals_separated': True, 'allocator_vectors_equal_by_sample_index': True})
    flags = [r for r in summary['comparisons'] if r['regression']]
    assert not flags
    return {'change': 436, 'decision': 'retain', 'summary_sha256': hashlib.sha256(raw).hexdigest(),
        'reason': 'Both repeats show practically useful lower normal p50 at all three sizes with separated descriptive within-report intervals, exact archive/sink identity, identical allocator vectors and no matched regression flag.',
        'comparisons': comparisons, 'matched_comparison_count': len(summary['comparisons']), 'matched_regression_flags': flags,
        'repeat_comparison_count': len(summary['repeat_flags']), 'repeat_flags': [r for r in summary['repeat_flags'] if r['flagged']],
        'scope': 'One named deterministic fresh-ODT scenario on one machine, one worker and warm in-memory inputs. Intervals resample 30 within-process observations; two repeats are not broad independent-run inference. No universal speedup, RSS, native, cold/range, copy-count or parallel-scaling claim.'}

def main():
    parser = argparse.ArgumentParser(); parser.add_argument('--verify', action='store_true'); args = parser.parse_args()
    value = derive(); path = ROOT / 'decision.json'
    if args.verify:
        assert json.loads(path.read_text()) == value
    else:
        with path.open('x') as output:
            output.write(json.dumps(value, indent=2) + '\n')
    print(json.dumps({'status': 'pass', 'decision': value['decision'], 'matched_flags': len(value['matched_regression_flags']), 'repeat_flags': len(value['repeat_flags'])}))

if __name__ == '__main__':
    main()
