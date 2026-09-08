#!/usr/bin/env python3
"""Derive exact compressor-stack totals without generic category ambiguity."""
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def analyze(root=ROOT):
    source = root / 'heap-attribution.json'
    data = json.loads(source.read_text())
    rows = []
    for trace in data['traces']:
        groups = {'zlib_rs_deflate_init': [], 'flate2_output_buffer': []}
        for stack in trace['stacks']:
            names = stack['frames_leaf_to_root']
            if any('7zlib_rs7deflate4init' in name for name in names):
                groups['zlib_rs_deflate_init'].append(stack)
            elif any('6flate23zio' in name and '3new' in name for name in names):
                groups['flate2_output_buffer'].append(stack)
        totals = {}
        for name, stacks in groups.items():
            count = sum(row['allocation_calls'] for row in stacks)
            size = sum(row['requested_bytes'] for row in stacks)
            assert count == 16421, (name, count)
            totals[name] = dict(allocation_calls=count, requested_bytes=size,
                requested_bytes_per_call=size // count, trace_ids=[s['trace_id'] for s in stacks])
            assert size == count * (380032 if name == 'zlib_rs_deflate_init' else 32768)
        phase = trace['phase_attribution']['writer-under-run']
        total = sum(row['requested_bytes'] for row in totals.values())
        rows.append(dict(lane=trace['lane'], owners=totals, run_context=phase,
            combined_requested_bytes=total, combined_share_of_run_requested_bytes_percent=100*total/phase['requested_bytes']))
    return dict(schema='litchi-0475-compressor-owners-v1',
        heap_attribution_sha256=hashlib.sha256(source.read_bytes()).hexdigest(), lanes=rows,
        scope='Exact run ancestry includes observers and report setup outside the timed operation. Owner rows use disjoint captured initialization/buffer stacks, not nearest generic category labels. Independent category peaks are not additive.')


if __name__ == '__main__':
    (ROOT / 'owners.json').write_text(json.dumps(analyze(), indent=2, sort_keys=True)+'\n')
