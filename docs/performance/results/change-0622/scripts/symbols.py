import re, os, subprocess, sys
OUT='counts'
SYMS = ['rewrite_value_only_with_provenance', 'scan_with_limit',
        'write_sheet_data_with_provenance', 'write_sheet_data_from_facts',
        'FactsBuilder', 'materialize_cell', 'from_source_selected',
        'worksheet_xml_and_parse_source', 'parse_source_with_observer']
def inclusive(path):
    text = subprocess.run(['callgrind_annotate','--inclusive=yes','--threshold=99.99',path],
                          capture_output=True, text=True).stdout
    costs = {}
    for line in text.splitlines():
        m = re.match(r'^\s*([\d,]+)\s+\(\s*[\d.]+%\)\s+\S*?:?([^\[]+)\s*\[', line)
        if not m:
            continue
        name = m.group(2).strip()
        costs[name] = costs.get(name, 0) + int(m.group(1).replace(',', ''))
    return costs
legs = sys.argv[1:] if len(sys.argv)>1 else ['before','after']
print(f"{'leg':7s} {'case':22s} {'shape':13s} {'Ir/op':>14s}  symbol")
data={}
for leg in legs:
    for case in ['one_edit_save','one_percent_edit_save']:
        for shape in ['medium','dense-sparse','noncompact']:
            if case=='one_percent_edit_save' and shape=='noncompact': continue
            a=f'{OUT}/{leg}-{case}-{shape}-n1.out'; b=f'{OUT}/{leg}-{case}-{shape}-n11.out'
            if not (os.path.exists(a) and os.path.exists(b)): continue
            ca, cb = inclusive(a), inclusive(b)
            for name in sorted(set(ca)|set(cb)):
                if not any(s in name for s in SYMS): continue
                d=(cb.get(name,0)-ca.get(name,0))/10.0
                if abs(d) < 5000: continue
                short = name.split('::')[-1]
                data[(leg,case,shape,short)] = d
                print(f"{leg:7s} {case:22s} {shape:13s} {d:14,.0f}  {short}")
