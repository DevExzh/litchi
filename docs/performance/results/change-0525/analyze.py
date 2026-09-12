"""Reuse the sealed 0521 numerical verifier for the fresh 0525 plan."""
import argparse
import hashlib
import importlib.util
import json
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
HELPER = HERE.parent / 'change-0521' / 'analyze.py'
spec = importlib.util.spec_from_file_location('xlsx_0521_numerical', HELPER)
assert spec is not None and spec.loader is not None
BASE = importlib.util.module_from_spec(spec)
spec.loader.exec_module(BASE)
BASE.HERE = HERE
BASE.PLAN = HERE / 'plan.json'
BASE.BOOTSTRAP_SEED = 5250525


def __getattr__(name):
    return getattr(BASE, name)


def _admission_thresholds(plan):
    """Read the numeric admission requirements from the frozen plan prose."""

    prose = plan.get('admission')
    BASE.require(isinstance(prose, str), 'plan admission text is missing')
    total_match = re.search(
        r'at least\s+([0-9]+(?:\.[0-9]+)?)%\s+p50 improvement in primary total',
        prose, re.IGNORECASE)
    commit_match = re.search(
        r'at least\s+([0-9]+(?:\.[0-9]+)?)%\s+commit Ir reduction',
        prose, re.IGNORECASE)
    BASE.require(total_match is not None,
                 'plan admission omits the primary-total p50 threshold')
    BASE.require(commit_match is not None,
                 'plan admission omits the commit-Ir threshold')
    return {
        'native_primary_total_p50_improvement_percent': float(total_match.group(1)),
        'native_primary_commit_p50_improvement_percent': float(total_match.group(1)),
        'profile_commit_ir_reduction_percent': float(commit_match.group(1)),
        'source': 'plan.admission',
    }


def _primary_rows(evidence, plan, label):
    native = evidence.get('native')
    BASE.require(isinstance(native, dict), f'{label}.native is not an object')
    rows = [row for row in native.get('rows', [])
            if row.get('kind') == 'primary' and row.get('guard') is None]
    expected = {(repeat, shape)
                for repeat in range(1, int(plan['primary']['repeats']) + 1)
                for shape in plan['primary']['shapes']}
    actual = {(row.get('repeat'), row.get('shape')) for row in rows}
    BASE.require(len(rows) == len(expected),
                 f'{label} primary native row count differs from plan: '
                 f'{len(rows)} != {len(expected)}')
    BASE.require(actual == expected,
                 f'{label} primary native matrix differs from plan: {sorted(actual ^ expected)}')
    return {(row['repeat'], row['shape']): row for row in rows}


def _improvement_percent(baseline, candidate, label):
    BASE.nonnegative_integer(baseline, f'{label}.baseline')
    BASE.nonnegative_integer(candidate, f'{label}.candidate')
    BASE.require(baseline > 0, f'{label}.baseline must be positive')
    return (baseline - candidate) / baseline * 100.0


def _native_admission(result, plan):
    thresholds = _admission_thresholds(plan)
    baseline = _primary_rows(result['baseline'], plan, 'baseline')
    candidate = _primary_rows(result['candidate'], plan, 'candidate')
    rows = []
    for repeat, shape in sorted(baseline):
        left = baseline[(repeat, shape)]
        right = candidate[(repeat, shape)]
        left_timing = left.get('timing')
        right_timing = right.get('timing')
        BASE.require(isinstance(left_timing, dict) and isinstance(right_timing, dict),
                     f'primary {repeat}/{shape} timing is not an object')
        left_elapsed = left_timing['elapsed_ns']['p50']
        right_elapsed = right_timing['elapsed_ns']['p50']
        left_commit = left_timing['commit_ns']['p50']
        right_commit = right_timing['commit_ns']['p50']
        elapsed_gain = _improvement_percent(
            left_elapsed, right_elapsed, f'primary {repeat}/{shape} elapsed p50')
        commit_gain = _improvement_percent(
            left_commit, right_commit, f'primary {repeat}/{shape} commit p50')
        total_passed = elapsed_gain >= thresholds[
            'native_primary_total_p50_improvement_percent']
        commit_passed = commit_gain >= thresholds[
            'native_primary_commit_p50_improvement_percent']
        rows.append({
            'repeat': repeat,
            'shape': shape,
            'native_primary_total_p50': {
                'baseline_ns': left_elapsed,
                'candidate_ns': right_elapsed,
                'improvement_percent': elapsed_gain,
                'required_improvement_percent': thresholds[
                    'native_primary_total_p50_improvement_percent'],
                'passed': total_passed,
            },
            'native_primary_commit_p50': {
                'baseline_ns': left_commit,
                'candidate_ns': right_commit,
                'improvement_percent': commit_gain,
                'required_improvement_percent': thresholds[
                    'native_primary_commit_p50_improvement_percent'],
                'passed': commit_passed,
            },
            'passed': total_passed and commit_passed,
        })
    return {
        'thresholds': thresholds,
        'rows': rows,
        'passed': all(row['passed'] for row in rows),
        'scope': (
            'Every planned primary shape/repeat must improve both the whole '
            'measured elapsed total and the commit phase at p50.'
        ),
    }


def analyze(stage=None):
    result = BASE.analyze(stage)
    result['scope'] = '0525 XLSX source-backed changed-row reconstruction reuse comparison'
    if result.get('stage') == 'compare':
        plan = BASE.read_json(BASE.PLAN)
        result['native_admission'] = _native_admission(result, plan)
    result['numerical_verifier'] = dict(path=str(HELPER.relative_to(HERE.parent)),
                                      sha256=hashlib.sha256(HELPER.read_bytes()).hexdigest())
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('output', nargs='?', type=Path)
    parser.add_argument('--output', dest='output_option', type=Path)
    parser.add_argument('--stage', choices=['baseline', 'candidate', 'compare'])
    args = parser.parse_args()
    stage = None if args.stage == 'compare' else args.stage
    output = args.output_option or args.output or HERE / (f'analysis-{stage}.json' if stage else 'comparison.json')
    result = analyze(stage)
    output.write_text(json.dumps(result, indent=2) + '\n')
    print('0525 evidence verified:', output)
