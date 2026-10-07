from __future__ import annotations
import hashlib, importlib.util, json, sys
from pathlib import Path

out = Path(sys.argv[1])
root = Path.cwd().resolve()
adapter_path = root / 'tools' / 'replay_0948_v2.py'
spec = importlib.util.spec_from_file_location('replay_adapter_leaf_review', adapter_path)
assert spec and spec.loader
adapter = importlib.util.module_from_spec(spec)
spec.loader.exec_module(adapter)
historical, authorization = adapter.bind_historical_root()
reducer = adapter.load_frozen_reducer()
plan = reducer.load_plan()
verified_source = reducer.source_check(plan)
full_source, focused_source = reducer.capture_source_summaries(verified_source)
binaries = reducer.binaries_check(plan)
validator = reducer.load_validator(plan)
job = reducer.jobs('qualification')[0]
folder = root / 'qualification' / reducer.child_name(job)
start_path = folder / 'start.json'
report_path = folder / 'report.json'
original_read_json = reducer.read_json

def leaf(label, mutation=None):
    restore = adapter.install_relocation_adapter(reducer, historical)
    try:
        if mutation is None:
            result = reducer.leaf_review(plan, job, binaries, validator, full_source, focused_source)
            return {'status': 'accepted', 'reviewed_label': result['label']}
        def read_json(path):
            value = original_read_json(path)
            if Path(path).resolve() == start_path.resolve():
                value = dict(value)
                mutation(value)
            return value
        reducer.read_json = read_json
        try:
            reducer.leaf_review(plan, job, binaries, validator, full_source, focused_source)
        finally:
            reducer.read_json = original_read_json
    except reducer.ReviewError as error:
        return {'status': 'rejected', 'error': str(error)}
    finally:
        reducer.read_json = original_read_json
        restore()
    raise AssertionError(f'{label} unexpectedly accepted')

positive = leaf('positive')
cwd_negative = leaf('tampered-cwd', lambda value: value.__setitem__('cwd', historical + '/tampered'))

def tamper_argv(value):
    argv = list(value['argv'])
    argv[-1] = argv[-1] + '.tampered'
    value['argv'] = argv
argv_negative = leaf('tampered-argv', tamper_argv)
# Remove only the shadow directory entry. The source relocated packet remains a separate tree;
# the hard-linked inode contents are untouched and its original directory entry is preserved.
report_path.unlink()
try:
    restore = adapter.install_relocation_adapter(reducer, historical)
    try:
        reducer.leaf_review(plan, job, binaries, validator, full_source, focused_source)
    finally:
        restore()
except reducer.ReviewError as error:
    missing = {'status': 'rejected', 'error': str(error), 'removed_shadow_entry': report_path.relative_to(root).as_posix()}
else:
    raise AssertionError('missing physical artifact unexpectedly accepted')
receipt = {
    'schema': 'litchi.performance.0948.portable-replay-leaf-controls.v1',
    'status': 'pass',
    'adapter': {'path': 'tools/replay_0948_v2.py', 'sha256': hashlib.sha256(adapter_path.read_bytes()).hexdigest()},
    'frozen_reducer': {'path': 'tools/reduce_0948_v4.py', 'sha256': hashlib.sha256((root/'tools/reduce_0948_v4.py').read_bytes()).hexdigest()},
    'plan': {'path': 'plan-v3.json', 'sha256': hashlib.sha256((root/'plan-v3.json').read_bytes()).hexdigest()},
    'historical_authorization': {'path': 'receipts/root-qualification-go-v1.json', 'sha256': hashlib.sha256((root/'receipts/root-qualification-go-v1.json').read_bytes()).hexdigest(), 'cwd': historical},
    'relocated_packet_root': str(root),
    'child': folder.relative_to(root).as_posix(),
    'controls': {'positive_actual_leaf_review': positive, 'tampered_historical_cwd_actual_leaf_review': cwd_negative, 'tampered_historical_argv_actual_leaf_review': argv_negative, 'missing_physical_report_actual_leaf_review': missing},
    'scope': 'offline leaf_review only; no workload launch; mutation controls used in-memory JSON and shadow-directory unlink',
}
with out.open('x', encoding='utf-8') as handle:
    json.dump(receipt, handle, indent=2, sort_keys=True)
    handle.write('\\n')
