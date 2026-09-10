#!/usr/bin/env python3
"""Summarize formal scalar controls and the ideal edit-elimination ceiling."""
from pathlib import Path
import csv, json, math
HERE = Path(__file__).resolve().parent
rows_out = []
for path in sorted((HERE / 'before').glob('*.csv')):
    rows = [r for r in csv.DictReader(path.open()) if r['warmup'] == 'false']
    assert len(rows) == 60
    values = sorted(int(r['elapsed_ns']) for r in rows)
    edit = sum(int(r['edit_ns']) for r in rows)
    total = sum(int(r['elapsed_ns']) for r in rows)
    result = {'name': path.stem, **{f'p{p}_us': values[math.ceil(len(values)*p/100)-1]/1000 for p in [50, 95, 99]},
              'edit_share_of_total_timed_ns': edit/total,
              'ideal_upper_bound_if_edit_cost_zero': total/(total-edit),
              'charged_work': int(rows[0]['budget_after_work'])}
    rows_out.append(result)
(HERE / 'before-summary.json').write_text(json.dumps({
    'scope': 'formal baseline only; ideal bound assumes all edit cost disappears, not a forecast or measured speedup',
    'children': 12, 'measured_samples': 720, 'warmup_samples': 72, 'rows': rows_out}, indent=2) + '\n')
