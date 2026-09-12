"""Bounded in-memory corruption probes for the frozen confirmation validator."""
import copy
import json
from pathlib import Path
import eager_confirmation_guard as guard

plan = guard.plan_data()
child = plan['children'][0]
raw_path = guard.HERE / 'eager-confirmation/A1/A1.json'
raw = guard.read_json(raw_path)
binary = guard.read_json(guard.HERE / child['binary_identity'])
binding = {'path': child['binary_path'], 'sha256': binary['sha256']}
oracle = guard.reference_result(plan)
original_read = guard.read_json
current = raw

def read_value(path):
    return current if Path(path) == raw_path else original_read(path)

guard.read_json = read_value
checks = []
try:
    guard.validate_raw(raw_path, 'A1', plan, child, binding, oracle)
    for name in ('altered_p50', 'altered_sink_vector'):
        current = copy.deepcopy(raw)
        if name == 'altered_p50':
            current['results'][0]['elapsed_ns']['p50'] += 1
        else:
            current['results'][0]['operation_metrics']['sink']['accepted_bytes']['values'][0] += 1
        try:
            guard.validate_raw(raw_path, 'A1', plan, child, binding, oracle)
        except ValueError as error:
            checks.append({'case': name, 'rejected': True, 'reason': str(error)})
        else:
            raise AssertionError(f'corruption accepted: {name}')
finally:
    guard.read_json = original_read
print(json.dumps({'status': 'pass', 'valid_control_pass': True,
                  'wrapper_sha256': guard.sha(Path(guard.__file__)),
                  'report_sha256': guard.sha(raw_path), 'checks': checks,
                  'scope': 'In-memory mutations only; no raw artifact, source, or frozen-tool mutation.'}, indent=2))
