"""In-memory negative controls for retained scanner-profile attribution."""
import copy
import json
import analyze

assert analyze.analyze()['status'] == 'pass'
checks = []
original_edge = analyze.helper.target_edge_summary
original_annotation = analyze.helper.parse_annotation

def bad_edge(*args, **kwargs):
    value = copy.deepcopy(original_edge(*args, **kwargs))
    value['inclusive_ir'] += 1
    return value

def bad_self(text, selected, label):
    value = original_annotation(text, selected, label)
    if label.endswith('.self.txt'):
        value['selected_ir'] += 1
    return value

for name in ('raw_edge_cost_mismatch', 'owner_self_cost_mismatch'):
    if name == 'raw_edge_cost_mismatch':
        analyze.helper.target_edge_summary = bad_edge
    else:
        analyze.helper.parse_annotation = bad_self
    try:
        analyze.analyze()
    except ValueError as error:
        checks.append({'case': name, 'rejected': True, 'reason': str(error)})
    else:
        raise AssertionError('corruption was accepted: ' + name)
    finally:
        analyze.helper.target_edge_summary = original_edge
        analyze.helper.parse_annotation = original_annotation
print(json.dumps({'status': 'pass', 'valid_control_pass': True, 'checks': checks,
                  'scope': 'In-memory parser-result corruption only; no raw evidence mutation or new capture.'}, indent=2))
