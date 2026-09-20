#!/usr/bin/env python3
"""Post-capture exact-edge attribution; inclusive rows are not additive."""
import hashlib
import importlib.util
import json
from pathlib import Path
import sys

P = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location('mechanism0713attribution', P/'analyze_mechanism.py')
M = importlib.util.module_from_spec(spec)
spec.loader.exec_module(M)

def build():
    mechanism = json.loads((P/'mechanism-analysis.json').read_text())
    rows = []
    needles = ['litchi_docx::alt::codec::scan',
               'litchi_docx::parts::document_part::active_block_ranges',
               'litchi_ooxml_common::mce::codec::active_offsets',
               'litchi_ooxml_common::mce::codec::find_bytes']
    for row in mechanism['callgrind']['rows']:
        path = P/row['measured_file']
        edges = M.PARSER.raw_edges(path)
        selected = [edge for edge in edges if edge['callee'] in needles]
        totals = {}
        for callee in needles:
            incoming = [edge for edge in selected if edge['callee'] == callee]
            totals[callee] = dict(calls=sum(e['calls'] for e in incoming),
                                  inclusive_ir=sum(e['inclusive_ir'] for e in incoming),
                                  raw_edge_count=len(incoming))
        rows.append(dict(stage=row['stage'], corpus=row['corpus_id'],
                         file=path.name, sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
                         owner_ir=row['measured_ir'], raw_edges=selected,
                         all_incoming_by_exact_callee=totals))
    return dict(schema_version=1, selection='post-capture diagnostic; not an added acceptance gate',
                mechanism_analysis_sha256=hashlib.sha256((P/'mechanism-analysis.json').read_bytes()).hexdigest(),
                rows=rows,
                limitations=['Inclusive callee rows overlap and must not be added together.',
                             'All incoming calls are not an exclusive selected-parent partition.',
                             'Absent standalone find_bytes edges can mean inlining, not absent search.',
                             'Callgrind calls are not allocator request counts or native latency.'])

def main():
    output = P/'search-attribution.json'
    encoded = (json.dumps(build(), indent=2, sort_keys=True)+'\n').encode()
    if sys.argv[1:] == ['--check']:
        assert output.read_bytes() == encoded
    else:
        assert not sys.argv[1:] and not output.exists()
        output.write_bytes(encoded)
    print('PASS: eight raw-profile exact-edge attribution rows')

if __name__ == '__main__':
    main()
