#!/usr/bin/env python3
"""ABBA-interleaved, core-pinned timing runs for change 0748.

Two kinds of case share one driver:

* ``harness``: a registered ``litchi-perf-baseline`` selector. The timed value
  is each result's ``elapsed_ns.samples``.
* ``probe``: the change-0746 XLS edit probe (``xls_edit_probe``). The timed
  value is one phase vector (``commit_ns`` for the edit operations,
  ``open_ns`` for the open control).

Every case runs 8 processes in the order A B B A A B B A (A = before build,
B = after build), each pinned with ``taskset -c CORE``. Raw outputs are kept.
The summary reports per-process p50/p95/mean, the median of process p50s per
arm, the four paired after/before p50 ratios (one per AB or BA pair), their
median, and a percentile bootstrap 95% CI over the paired ratios.

usage: abba.py OUT_DIR CORE [case,case,...]
"""
import json
import os
import random
import statistics
import subprocess
import sys

OUT = sys.argv[1]
CORE = sys.argv[2]
SCRATCH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BIN = {
    'A': {
        'harness': os.path.join(SCRATCH, 'bin', 'before', 'litchi-perf-baseline'),
        'probe': os.path.join(SCRATCH, 'bin', 'before', 'xls_edit_probe'),
    },
    'B': {
        # ABBA_AA=1 runs an A/A control: a byte-identical copy of the before
        # binary takes the B role, measuring the case's own noise.
        'harness': os.path.join(SCRATCH, 'bin', 'aa' if os.environ.get('ABBA_AA') else 'after', 'litchi-perf-baseline'),
        'probe': os.path.join(SCRATCH, 'bin', 'after', 'xls_edit_probe'),
    },
}
ORDER = ['A', 'B', 'B', 'A', 'A', 'B', 'B', 'A']
FS_ROOT = os.path.join(SCRATCH, 'fsroot')
XLS_54016 = os.path.join(SCRATCH, 'fixtures', '54016.xls')
XLS_LARGE = os.path.join(SCRATCH, 'fixtures', 'xls-large.xls')
LARGE = ['--writer-shape', 'large']
TINY = ['--writer-shape', 'tiny']
FS = ['--filesystem-cache', 'warm', '--filesystem-root', FS_ROOT]

# (label, kind, selector-or-operation, warmup, samples, extra args, phase)
CASES = [
    # Treatments: generic copy-through renders become sealed (two per commit).
    ('xls_semantic_one_edit_save/large', 'harness', 'xls_semantic_one_edit_save', 5, 40, LARGE, None),
    ('xls_semantic_one_edit_save/tiny', 'harness', 'xls_semantic_one_edit_save', 5, 60, TINY, None),
    ('xls_visibility_eager_edit_save', 'harness', 'xls_visibility_eager_edit_save', 5, 40, [], None),
    ('xls_visibility_eager_batch_edit_save', 'harness', 'xls_visibility_eager_batch_edit_save', 5, 40, [], None),
    ('xls_numeric_eager_rk_mulrk_edit_save', 'harness', 'xls_numeric_eager_rk_mulrk_edit_save', 5, 60, [], None),
    # Treatments: a generic source-backed owner becomes sealed.
    ('xls_visibility_source_backed_edit_save', 'harness', 'xls_visibility_source_backed_edit_save', 5, 40, [], None),
    ('xls_visibility_source_backed_batch_edit_save', 'harness', 'xls_visibility_source_backed_batch_edit_save', 5, 40, [], None),
    # Treatments: already-owned plans stop re-hashing at view and emission.
    ('xls_comments_source_backed_edit_save', 'harness', 'xls_comments_source_backed_edit_save', 5, 30, [], None),
    ('xls_comments_source_backed_batch_edit_save', 'harness', 'xls_comments_source_backed_batch_edit_save', 5, 30, [], None),
    ('xls_numeric_source_backed_number_edit_save', 'harness', 'xls_numeric_source_backed_number_edit_save', 5, 30, [], None),
    ('xls_numeric_source_backed_rk_mulrk_edit_save', 'harness', 'xls_numeric_source_backed_rk_mulrk_edit_save', 5, 60, [], None),
    ('xls_numeric_plan_only_number_edit_save', 'harness', 'xls_numeric_plan_only_number_edit_save', 5, 30, [], None),
    ('xls_numeric_plan_only_rk_mulrk_edit_save', 'harness', 'xls_numeric_plan_only_rk_mulrk_edit_save', 5, 60, [], None),
    ('cfb_file_owned_same_length_overlay_atomic_save', 'harness', 'cfb_file_owned_same_length_overlay_atomic_save', 3, 20, FS, None),
    # Controls: generic sources keep every pass; DIFAT/length gates decline
    # copy-through before any overlay code; pure reads.
    ('cfb_file_same_length_overlay_atomic_save', 'harness', 'cfb_file_same_length_overlay_atomic_save', 3, 20, FS, None),
    ('cfb_file_same_length_overlay_atomic_save/long', 'harness', 'cfb_file_same_length_overlay_atomic_save', 3, 60, FS, None),
    ('ppt_source_backed_one_shape_text/large', 'harness', 'ppt_source_backed_one_shape_text', 5, 40, LARGE, None),
    ('xls_comments_eager_edit_save', 'harness', 'xls_comments_eager_edit_save', 3, 20, [], None),
    ('xls_numeric_eager_number_edit_save', 'harness', 'xls_numeric_eager_number_edit_save', 3, 20, [], None),
    ('ole_common_one_edit_save', 'harness', 'ole_common_one_edit_save', 5, 60, ['--shape', 'few-large', '--payload', 'incompressible'], None),
    ('ole_common_finish_render', 'harness', 'ole_common_finish_render', 5, 60, ['--shape', 'few-large', '--payload', 'incompressible'], None),
    ('doc_semantic_one_edit_save/large', 'harness', 'doc_semantic_one_edit_save', 5, 40, LARGE, None),
    ('xls_semantic_open/large', 'harness', 'xls_semantic_open', 5, 40, LARGE, None),
    # The 0746 probe: the XLS generic commit (two copy-through renders on
    # this base) and the owned source-backed paths, on a real producer file
    # and the harness writer corpus. `open` is the control.
    ('probe/54016/number-generic', 'probe', ('number-generic', XLS_54016), 3, 20, [], 'commit_ns'),
    ('probe/54016/string-generic', 'probe', ('string-generic', XLS_54016), 3, 20, [], 'commit_ns'),
    ('probe/54016/number-source-backed', 'probe', ('number-source-backed', XLS_54016), 3, 20, [], 'commit_ns'),
    ('probe/54016/number-plan', 'probe', ('number-plan', XLS_54016), 3, 20, [], 'commit_ns'),
    ('probe/54016/number-plan-publish', 'probe', ('number-plan', XLS_54016), 3, 20, [], 'publish_ns'),
    ('probe/54016/open', 'probe', ('open', XLS_54016), 3, 20, [], 'open_ns'),
    ('probe/xls-large/number-generic', 'probe', ('number-generic', XLS_LARGE), 5, 40, [], 'commit_ns'),
    ('probe/xls-large/number-source-backed', 'probe', ('number-source-backed', XLS_LARGE), 5, 40, [], 'commit_ns'),
]


def pct(values, q):
    values = sorted(values)
    k = (len(values) - 1) * q
    lo, hi = int(k), min(int(k) + 1, len(values) - 1)
    return values[lo] + (values[hi] - values[lo]) * (k - lo)


def bootstrap(ratios, iterations=10000, seed=748):
    rng = random.Random(seed)
    boots = sorted(
        statistics.median([rng.choice(ratios) for _ in ratios]) for _ in range(iterations)
    )
    return boots[int(0.025 * iterations)], boots[int(0.975 * iterations) - 1]


def run_process(label, kind, target, warmup, samples, extra, phase, leg, index):
    safe = label.replace('/', '__')
    raw = os.path.join(OUT, 'raw', f'{safe}-{index}-{leg}.json')
    # A raw report left by an earlier, interrupted run of the identical
    # command is reused, so an interrupted matrix resumes without re-measuring.
    if kind == 'harness':
        command = ['taskset', '-c', CORE, BIN[leg]['harness'], '--case', target,
                   '--samples', str(samples), '--warmup', str(warmup), '--json', raw, *extra]
        if not os.path.exists(raw):
            subprocess.run(command, capture_output=True, text=True, check=True)
        with open(raw) as handle:
            report = json.load(handle)
        results = [r for r in report['results'] if r['case'] == target]
        if len(results) != 1:
            raise SystemExit(f'{label}: expected one result, got {len(results)}')
        values = results[0]['elapsed_ns']['samples']
        corpus = results[0].get('corpus', {})
        identity = corpus.get('archive_sha256') or corpus.get('name')
    else:
        operation, fixture = target
        command = ['taskset', '-c', CORE, BIN[leg]['probe'], '--input', fixture,
                   '--operation', operation, '--warmups', str(warmup), '--samples', str(samples), *extra]
        if not os.path.exists(raw):
            completed = subprocess.run(command, capture_output=True, text=True, check=True)
            with open(raw, 'w') as handle:
                handle.write(completed.stdout)
        with open(raw) as handle:
            report = json.load(handle)
        values = report[phase]
        identity = f"{report['source_bytes']}:{report.get('published_bytes')}"
    return {
        'index': index, 'leg': leg, 'raw': os.path.relpath(raw, OUT),
        'p50': pct(values, 0.5), 'p95': pct(values, 0.95),
        'mean': statistics.fmean(values), 'n': len(values), 'identity': identity,
    }


def main():
    selected = sys.argv[3].split(',') if len(sys.argv) > 3 else None
    os.makedirs(os.path.join(OUT, 'raw'), exist_ok=True)
    os.makedirs(FS_ROOT, exist_ok=True)
    summary = []
    for label, kind, target, warmup, samples, extra, phase in CASES:
        if selected and label not in selected:
            continue
        processes = [
            run_process(label, kind, target, warmup, samples, extra, phase, leg, index)
            for index, leg in enumerate(ORDER)
        ]
        identities = {p['identity'] for p in processes}
        a = [p for p in processes if p['leg'] == 'A']
        b = [p for p in processes if p['leg'] == 'B']
        ratios, mean_ratios = [], []
        for i in range(0, len(ORDER), 2):
            pair = processes[i:i + 2]
            before = next(p for p in pair if p['leg'] == 'A')
            after = next(p for p in pair if p['leg'] == 'B')
            ratios.append(after['p50'] / before['p50'])
            mean_ratios.append(after['mean'] / before['mean'])
        lo, hi = bootstrap(ratios)
        record = {
            'case': label, 'kind': kind,
            'target': target if kind == 'harness' else list(target),
            'phase': phase or 'elapsed_ns', 'warmup': warmup, 'samples': samples,
            'extra': extra, 'core': CORE, 'identities': sorted(map(str, identities)),
            'before_median_p50': statistics.median(p['p50'] for p in a),
            'after_median_p50': statistics.median(p['p50'] for p in b),
            'before_median_p95': statistics.median(p['p95'] for p in a),
            'after_median_p95': statistics.median(p['p95'] for p in b),
            'before_median_mean': statistics.median(p['mean'] for p in a),
            'after_median_mean': statistics.median(p['mean'] for p in b),
            'paired_p50_ratios': ratios, 'paired_mean_ratios': mean_ratios,
            'median_paired_p50_ratio': statistics.median(ratios),
            'median_paired_mean_ratio': statistics.median(mean_ratios),
            'bootstrap_95ci_p50_ratio': [lo, hi],
            'processes': processes,
        }
        summary.append(record)
        with open(os.path.join(OUT, 'summary.json' if not selected else 'summary-partial.json'), 'w') as handle:
            json.dump(summary, handle, indent=1)
        print(f"{label:52s} before {record['before_median_p50'] / 1e6:9.4f} ms  after "
              f"{record['after_median_p50'] / 1e6:9.4f} ms  ratio {record['median_paired_p50_ratio']:.3f} "
              f"CI [{lo:.3f}, {hi:.3f}]  ids={len(identities)}", flush=True)
    name = 'summary.json' if not selected else 'summary-partial.json'
    with open(os.path.join(OUT, name), 'w') as handle:
        json.dump(summary, handle, indent=1)


if __name__ == '__main__':
    main()
