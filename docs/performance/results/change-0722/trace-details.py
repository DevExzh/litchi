#!/usr/bin/env python3
"""Render the complete untimed trace matrix without pooling refusal counts."""
import hashlib
import importlib.util
import json
from pathlib import Path
import sys

P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('trace_details_0722', P / 'trace-analyze.py')
T = importlib.util.module_from_spec(spec)
spec.loader.exec_module(T)
baseline = T.parse_trace(P / 'trace/baseline/stderr')
candidate = T.parse_trace(P / 'trace/candidate/stderr')
rows = T.compare_trace_documents(baseline, candidate)
lines = [
    '# Untimed structural reader matrix', '',
    'Both lanes preserve the public report, ordered MCE vectors, structured chunks, and ranges.',
    'Counts include successful EOF reads. A parser error returned by the read itself is not a successful read.',
    'Baseline range reads are counted before admission, with an unknown end offset (`null`) because the event borrows the reader.',
    'Baseline range observers run after admission and classification; candidate observers run immediately before range classification, only while that state remains admitted.',
    'Refusal observer counts therefore have different boundaries and are not interchangeable reader counts.',
    'These are parser events, not physical I/O, timing, instruction, or allocation measurements.', '',
    '| Case | XML bytes | Baseline alt / range reads | Candidate fused reads | Baseline / candidate observers | Ordered MCE input lengths |',
    '| --- | ---: | ---: | ---: | ---: | --- |',
]
documents = {T.document_key(d): T.normalized_document(d, 'baseline') for d in baseline['documents']}
for row in rows:
    key = (row['case'], row['occurrence'])
    left = row['reader_counts']['baseline']
    right = row['reader_counts']['candidate']
    observer = row['observer_relation']
    calls = documents[key]['active_calls']
    lengths = ', '.join(str(len(call['input_offsets'])) for call in calls) or 'none'
    lines.append(f"| {row['case']} | {row['boundary']['xml_bytes']} | {left['alt']} / {left['range']} | {right['total']} | {observer['baseline_observer_events']} / {observer['candidate_observer_events']} | {lengths} |")
lines += ['', 'Input hashes:', '']
for name in ['trace/baseline/stderr', 'trace/candidate/stderr', 'trace-analyze.py', 'trace-details.py']:
    lines.append(f"- `{name}`: `{hashlib.sha256((P / name).read_bytes()).hexdigest()}`")
encoded = '\n'.join(lines) + '\n'
out = P / 'trace-details.md'
if sys.argv[1:] == ['--check']:
    assert out.read_text() == encoded
else:
    assert not sys.argv[1:]
    out.write_text(encoded)
print(f'PASS: {len(rows)} document trace rows rendered')
