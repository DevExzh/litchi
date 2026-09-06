#!/usr/bin/env python3
"""Check capture-size refinement against the retained first candidate."""
import hashlib
import json
from pathlib import Path
ROOT = Path(__file__).resolve().parent
old_path = ROOT.with_name('change-0431-first-attempt') / 'comparison.json'
new_path = ROOT / 'comparison.json'
old = json.loads(old_path.read_text())
new = json.loads(new_path.read_text())
rows = []
for current in new['captures']:
    if current['role'] != 'after':
        continue
    prior = next(c for c in old['captures'] if c['role'] == 'after' and c['name'] == current['name'])
    assert prior['input_identity'] == current['input_identity']
    assert prior['output_identity'] == current['output_identity']
    keys = ['phases.published.source_budget.input_bytes_used', 'phases.published.source_budget.work_used']
    for key in keys:
        assert prior['metrics'][key] == current['metrics'][key]
    calls = 'phases.published.source_reads.delta.logical_calls'
    a, b = prior['metrics'][calls]['p50'], current['metrics'][calls]['p50']
    if current['selector'] == 'media-rich':
        assert (a, b) == ((4265, 4265) if current['provider_label'] == 'short-range' else (1193, 425))
    rows.append({'name': current['name'], 'input_equal': True, 'output_equal': True, 'input_and_work_equal': True, 'first_publication_source_calls_p50': a, 'refined_publication_source_calls_p50': b})
record = {'status': 'pass', 'scope': 'Retained source-bound producer reports; output artifacts are not exported by this harness.', 'first_comparison_sha256': hashlib.sha256(old_path.read_bytes()).hexdigest(), 'refined_comparison_sha256': hashlib.sha256(new_path.read_bytes()).hexdigest(), 'rows': rows}
assert len(rows) == 16
output = ROOT / 'refinement-check.json'
assert not output.exists()
output.write_text(json.dumps(record, indent=2) + '\n')
print(json.dumps({'status': 'pass', 'captures': len(rows)}))
