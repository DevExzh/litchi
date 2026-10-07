"""Derive a small, lossless gate summary from the externally retained 0948 analysis."""
from pathlib import Path
import argparse
import hashlib
import json
import statistics

p = Path(__file__).resolve().parent.parent
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--output', type=Path, default=p/'reduced/summary-v1.json')
args = parser.parse_args()
analysis = p/'analysis.json'
d = json.loads(analysis.read_text())
rows = {}
for name, native in d['reduction']['cases'].items():
    metrics = {}
    for metric in ('publication_p50', 'whole_workflow_p50'):
        stat = native[metric]
        pairs = stat['pairs']
        metrics[metric] = {key: stat[key] for key in ('estimate', 'lower_95', 'upper_95', 'accepted')}
        for role in ('baseline', 'candidate'):
            metrics[metric][role+'_median_of_block_quantiles_ns'] = statistics.median(pair[role+'_'+metric] for pair in pairs)
    metrics['paired_blocks'] = native['publication_p50']['pairs']
    allocation = d['allocation_reduction']['cases'][name]
    rows[name] = {'native': metrics, 'allocator': allocation}
result = {
    'schema': 'litchi.performance.0948.reduced-summary.v1',
    'status': d['status'],
    'decision': 'reject-candidate',
    'reason': 'Both primary cases exceed the 1.05 allocated_bytes and allocation_calls limits in all three repeats.',
    'source_commit': d['source']['source_commit'],
    'analysis': {'path': 'analysis.json', 'bytes': analysis.stat().st_size, 'sha256': hashlib.sha256(analysis.read_bytes()).hexdigest()},
    'counts': {'qualification_children': len(d['qualification']), 'native_children': len(d['native']), 'allocator_children': len(d['allocator'])},
    'native_gate_pass': d['reduction']['all_case_latency_gate_pass'],
    'allocator_pair_gate_pass': d['allocation_reduction']['all_pair_gate_pass'],
    'allocator_aggregate_gate_pass': d['allocation_reduction']['all_row_safety_gate_pass'],
    'cases': rows,
    'native_review_flags': d['reduction']['review_flags'],
    'allocator_review_flags': d['allocation_reduction']['review_flags'],
    'reducer_helper': d['reducer_helper'],
    'correction_receipt': d['correction_receipt'],
    'scope': 'CPU 12, named synthetic XLSX warm serial workloads only; no adoption, broad Office speedup, cold-cache, concurrency, or allocator bootstrap claim.'
}
with args.output.open('x') as out:
    json.dump(result, out, indent=2, sort_keys=True, allow_nan=False)
    out.write('\n')
print(args.output, args.output.stat().st_size)
