"""Independent v4 full-census and qualification-shape review; no workloads."""
from __future__ import annotations
import copy
import hashlib
import importlib.util
import json
from pathlib import Path

PACKET = Path('/home/zhuhe/litchi-performance-0948')
REDUCER_PATH = PACKET / 'tools/reduce_0948_v4.py'
OUT = PACKET / 'review/protocol/review_protocol_0948_v8-source-census.json'
PRIOR_FAILURE = PACKET / 'receipts/qualification-reduction-failed-v1.json'


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load_json(path: Path):
    return json.loads(path.read_text(encoding='utf-8'))


def reject(label, fn):
    try:
        fn()
    except Exception as exc:  # expected fail-closed control
        return {'label': label, 'passed': True, 'error_type': type(exc).__name__}
    raise AssertionError(f'negative control unexpectedly passed: {label}')


def main():
    assert not OUT.exists(), f'refusing overwrite: {OUT}'
    assert REDUCER_PATH.is_file() and not REDUCER_PATH.is_symlink()
    spec = importlib.util.spec_from_file_location('reduce_0948_v4_review', REDUCER_PATH)
    assert spec and spec.loader
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    reducer_sha = sha(REDUCER_PATH)
    assert reducer_sha == 'f960a94cad1544ac7f64c8a5006eb172bf5db849735b576fba58a6338a5b1d17'
    plan = mod.load_plan()
    verified = mod.source_check(plan)
    correction = mod.correction_metadata(plan)
    full, focused = mod.capture_source_summaries(verified)
    capture_path = PACKET / 'qualification/capture.json'
    capture = mod.read_json(capture_path)
    analysis = mod.read_json(PACKET / 'qualification-analysis.json')
    failure = mod.read_json(PRIOR_FAILURE)

    expected_count = plan['source']['file_count']
    assert expected_count == 10789
    assert verified['file_count'] == expected_count
    assert verified['hashed_entries'] == expected_count * 2
    assert capture['source_start'] == full
    assert capture['source_end'] == full
    assert analysis['status'] == 'pass' and analysis['stage'] == 'qualification'
    assert analysis['source'] == verified
    assert analysis['correction_receipt'] == correction['correction_receipt']
    assert correction['correction_receipt']['sha256'] == sha(PACKET / 'receipts/reducer-v4-source-census-correction-v1.json')
    receipt = mod.read_json(PACKET / 'receipts/reducer-v4-source-census-correction-v1.json')
    assert all(receipt[key] is True for key in ('no_capture_change', 'no_child_retry', 'no_gate_change'))
    assert failure['exit_code'] == 1
    assert failure['qualification_child_retries'] == 0
    assert failure['formal_native_or_allocator_children'] == 0
    assert 'checked_files/full' in failure['diagnosis']

    # Check every retained qualification child uses the full census at start and
    # the deliberately focused candidate delta at operation end.
    child_dirs = sorted(p for p in (PACKET / 'qualification').iterdir()
                        if p.is_dir() and not p.is_symlink())
    assert len(child_dirs) == 16
    child_shape_counts = {'full_start': 0, 'focused_end': 0}
    for child in child_dirs:
        start = mod.read_json(child / 'start.json')
        terminal = mod.read_json(child / 'terminal.json')
        command = mod.read_json(child / 'command.json')
        receipt_child = mod.read_json(child / 'receipt.json')
        mod.require_capture_source(start['source'], full, f'{child.name} start')
        mod.require_capture_source(terminal['source'], full, f'{child.name} terminal')
        mod.require_capture_source(command['source'], full, f'{child.name} command')
        mod.require_capture_source(receipt_child['source_start'], full, f'{child.name} receipt start')
        mod.require_capture_source(command['source_end'], focused, f'{child.name} command end')
        mod.require_capture_source(receipt_child['source_end'], focused, f'{child.name} receipt end')
        assert command['source_error'] is None and receipt_child['source_error'] is None
        child_shape_counts['full_start'] += 1
        child_shape_counts['focused_end'] += 1

    negative_controls = []
    bad = copy.deepcopy(full)
    bad['full'] = False
    negative_controls.append(reject('full_flag_false', lambda: mod.require_capture_source(bad, full, 'negative full flag')))
    bad = copy.deepcopy(full)
    bad['checked_files']['baseline'] -= 1
    negative_controls.append(reject('baseline_count_minus_one', lambda: mod.require_capture_source(bad, full, 'negative baseline count')))
    bad = copy.deepcopy(full)
    del bad['checked_files']['candidate']
    negative_controls.append(reject('candidate_count_missing', lambda: mod.require_capture_source(bad, full, 'negative candidate count')))
    bad = copy.deepcopy(full)
    bad['source_commit'] = '0' * 40
    negative_controls.append(reject('source_commit_mismatch', lambda: mod.require_capture_source(bad, full, 'negative source commit')))
    negative_controls.append(reject('focused_marked_full', lambda: mod.require_capture_source(focused, full, 'negative focused/full')))
    assert len(negative_controls) == 5 and all(row['passed'] for row in negative_controls)

    result = {
        'schema': 'litchi.performance.0948.review.source-census-v8',
        'status': 'pass',
        'scope': 'offline retained qualification source custody only; no workload launch',
        'reducer': {'path': 'tools/reduce_0948_v4.py', 'bytes': REDUCER_PATH.stat().st_size, 'sha256': reducer_sha},
        'plan': {'path': 'plan-v3.json', 'bytes': (PACKET/'plan-v3.json').stat().st_size, 'sha256': sha(PACKET/'plan-v3.json')},
        'correction_receipt': correction['correction_receipt'],
        'source_census': {
            'file_count_each_role': expected_count,
            'hashed_entries_both_roles': verified['hashed_entries'],
            'capture_full_shape_checked_files': full['checked_files'],
            'capture_focused_shape_checked_files': focused['checked_files'],
            'capture_start_end_equal': capture['source_start'] == capture['source_end'],
            'analysis_source_equivalent': analysis['source'] == verified,
        },
        'qualification_children': child_shape_counts,
        'negative_controls': negative_controls,
        'prior_failure': {
            'path': 'receipts/qualification-reduction-failed-v1.json',
            'bytes': PRIOR_FAILURE.stat().st_size,
            'sha256': sha(PRIOR_FAILURE),
            'exit_code': failure['exit_code'],
            'qualification_child_retries': failure['qualification_child_retries'],
            'formal_native_or_allocator_children': failure['formal_native_or_allocator_children'],
            'preserved': True,
        },
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open('x', encoding='utf-8') as handle:
        handle.write(json.dumps(result, indent=2, sort_keys=True) + '\n')
        handle.flush()
    print(json.dumps({'status': result['status'], 'output': str(OUT), 'children': 16,
                      'negative_controls': len(negative_controls)}))


if __name__ == '__main__':
    main()
