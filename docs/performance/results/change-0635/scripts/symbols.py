"""change 0635: callgrind isolation-pair differencing (N=1 vs N=11)."""
import re, os, subprocess, sys, json
OUT = os.environ.get('COUNTS', 'counts')
SYMS = ['FactsBuilder', 'facts::', 'cell_column', 'row_number',
        'scan_with_limit', 'rewrite_value_only_with_provenance',
        'from_source_selected', 'worksheet_xml_and_parse_source',
        'worksheet_xml_with_facts', 'validate_xml', 'styles::parse',
        'capture_auxiliary_source', 'parse_source_with_observer',
        'from_rewritten_value_source', 'materialize_cell',
        'write_sheet_data_from_facts', 'is_spreadsheetml_name', 'from_package_with_styles', 'style_catalog']
CASES = ['one_edit-medium', 'one_edit-dense-sparse', 'one_edit-noncompact',
         'one_percent-medium', 'one_percent-dense-sparse',
         'managed_one_edit-medium', 'producer_medium_edit', 'producer_dense_edit']

def totals(path):
    text = subprocess.run(['callgrind_annotate', '--threshold=0', path],
                          capture_output=True, text=True).stdout
    m = re.search(r'^([\d,]+)\s+\(100\.0%\)\s+PROGRAM TOTALS', text, re.M)
    if not m:
        m = re.search(r'^\s*([\d,]+)\s+PROGRAM TOTALS', text, re.M)
    return int(m.group(1).replace(',', '')) if m else None

def inclusive(path):
    text = subprocess.run(['callgrind_annotate', '--inclusive=yes', '--threshold=99.99', path],
                          capture_output=True, text=True).stdout
    costs = {}
    for line in text.splitlines():
        m = re.match(r'^\s*([\d,]+)\s+\(\s*[\d.]+%\)\s+\S*?:?([^\[]+)\s*\[', line)
        if not m:
            continue
        name = m.group(2).strip()
        costs[name] = costs.get(name, 0) + int(m.group(1).replace(',', ''))
    return costs

legs = sys.argv[1:] or ['before', 'after']
rows = {}
print(f"{'leg':9s} {'case':26s} {'Ir/op':>16s}  symbol")
for leg in legs:
    for case in CASES:
        a = f'{OUT}/{leg}-{case}-n1.out'
        b = f'{OUT}/{leg}-{case}-n11.out'
        # A zero-byte profile is a run that has not finished; skip the pair.
        if not (os.path.exists(a) and os.path.exists(b)):
            continue
        if os.path.getsize(a) == 0 or os.path.getsize(b) == 0:
            continue
        ta, tb = totals(a), totals(b)
        if ta is not None and tb is not None:
            d = (tb - ta) / 10.0
            rows[(leg, case, 'WHOLE ITERATION')] = d
            print(f"{leg:9s} {case:26s} {d:16,.0f}  WHOLE ITERATION")
        ca, cb = inclusive(a), inclusive(b)
        for name in sorted(set(ca) | set(cb)):
            if not any(s in name for s in SYMS):
                continue
            d = (cb.get(name, 0) - ca.get(name, 0)) / 10.0
            if abs(d) < 2000:
                continue
            short = name.split('::')[-1]
            key = (leg, case, short)
            rows[key] = rows.get(key, 0) + d
            print(f"{leg:9s} {case:26s} {d:16,.0f}  {short}")
with open(f'{OUT}/summary-{"-".join(legs)}.json', 'w') as handle:
    json.dump({'|'.join(k): v for k, v in rows.items()}, handle, indent=1)
