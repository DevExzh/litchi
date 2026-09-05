#!/usr/bin/env python3
"""Exercise corrected report guards using in-memory copies of a real report."""
import copy
import hashlib
import importlib.util
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('summary422', ROOT / 'summarize.py')
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
folder = ROOT / 'runs/R1/pptx_cross_copy_media_rich_lifecycle'
report_bytes = (folder / 'report.json').read_bytes()
original = json.loads(report_bytes)
catalog = json.loads((folder / 'catalog.json').read_text())
journal = json.loads((folder / 'journal.json').read_text())
arguments = dict(label='guard_probe', selector='pptx_cross_copy_media_rich_lifecycle', journal=journal, report_sha=hashlib.sha256(report_bytes).hexdigest())
module.validate_report(original, catalog, **arguments)
checks = {'unmodified_real_report_accepted': True}

def reject(name, mutate):
    value = copy.deepcopy(original)
    mutate(value)
    try:
        module.validate_report(value, catalog, **arguments)
    except module.SummaryError as error:
        checks[name] = {'rejected': True, 'diagnostic': str(error)}
    else:
        raise AssertionError(f'{name} unexpectedly passed')

reject('missing_region_peak', lambda r: r['results'][0]['operation_metrics']['allocation'].pop('region_peak_live_bytes'))
reject('region_peak_below_live', lambda r: r['results'][0]['operation_metrics']['allocation']['region_peak_live_bytes']['values'].__setitem__(0, 0))
reject('region_peak_above_lifetime', lambda r: r['results'][0]['operation_metrics']['allocation']['region_peak_live_bytes']['values'].__setitem__(0, r['results'][0]['operation_metrics']['allocation']['peak_live_bytes_after']['values'][0] + 1))
reject('old_counter_revision', lambda r: r['tool'].__setitem__('allocator_counter_revision', 'post_update_peak_v2'))
reject('missing_counter_revision', lambda r: r['tool'].pop('allocator_counter_revision'))
reject('peak_below_live', lambda r: r['results'][0]['operation_metrics']['allocation']['peak_live_bytes_after']['values'].__setitem__(0, 0))

def decrease(r):
    row = r['results'][0]
    index = row['elapsed_ns']['sample_order'].index(0)
    values = row['operation_metrics']['allocation']['peak_live_bytes_after']['values']
    values[index] = max(values) + 1
reject('chronological_peak_decrease', decrease)
assert module.chronological_values([2, 0, 3, 1], [30, 10, 40, 20]) == [10, 20, 30, 40]
checks['elapsed_sorted_vectors_restore_chronology'] = True
record = {'change': 422, 'status': 'pass', 'classification': 'in-memory validator guards; no synthetic measurement', 'original_report_sha256': arguments['report_sha'], 'checks': checks}
(ROOT / 'checks/report-guards.json').write_text(json.dumps(record, indent=2, sort_keys=True) + '\n')
print(json.dumps({'status': 'pass', 'checks': len(checks)}))
