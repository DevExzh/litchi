"""Allocator retention-probe summary for change 0646."""
import json, sys
S = sys.argv[1]
D = sys.argv[2] if len(sys.argv) > 2 else 'alloc'
K = 'retention_probe_region_peak_live_bytes'
print("change 0646 - allocator retention probe (litchi-perf-baseline-alloc retention --api owned)")
print("5 samples after 1 warmup per leg; medians reported; allocator counters are deterministic.\n")
for c in ('plain', 'media-rich'):
    print(f'################ {c} ################')
    legs = {leg: json.load(open(f'{S}/{D}/retention-{leg}-{c}.json')) for leg in ('before', 'after')}
    print(f'  expected_output_bytes (= candidate archive) = {legs["before"]["expected_output_bytes"]:,}')
    print(f'  destination_archive_bytes                   = {legs["before"]["destination_archive_bytes"]:,}')
    keys = ['baseline_before_inputs', 'prepared_inputs_and_sink', 'opened_documents', 'planned',
            'published', 'drop_result', 'drop_plan', 'drop_document_handles', 'drop_sink']
    print(f'\n  {"checkpoint":30s} {"before live":>14s} {"after live":>14s} {"delta":>14s}')
    for k in keys:
        row = []
        for leg in ('before', 'after'):
            vals = sorted(r[k]['live_bytes'] for r in legs[leg]['samples_raw'])
            row.append(vals[len(vals) // 2])
        print(f'  {k:30s} {row[0]:>14,} {row[1]:>14,} {row[1] - row[0]:>+14,}')
    print()
    for label, f in (('region peak live bytes', lambda x: x['retention_probe'][K]),
                     ('allocation calls', lambda x: x['retention_probe']['allocation_calls']),
                     ('reallocation calls', lambda x: x['retention_probe']['reallocation_calls']),
                     ('allocated bytes', lambda x: x['retention_probe']['allocated_bytes'])):
        row = []
        for leg in ('before', 'after'):
            vals = sorted(f(r) for r in legs[leg]['samples_raw'])
            row.append(vals[len(vals) // 2])
        print(f'  {label:30s} {row[0]:>14,} {row[1]:>14,} {100.0 * (row[1] - row[0]) / row[0]:>+13.2f}%')
    print()
