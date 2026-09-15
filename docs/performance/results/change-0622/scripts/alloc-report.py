import json, glob, os, statistics, sys

def load(path):
    d = json.load(open(path))
    return d['results'][0]

regions = [
    'plan_allocation_metrics',
    'staging_allocation_metrics',
    'commit_core_allocation_metrics',
    'commit_allocation_metrics',
    'publication_allocation_metrics',
]

rows = []
for shape in ['medium', 'dense-sparse', 'noncompact']:
    for case in ['xlsx_source_backed_cell_values_one_edit_save',
                 'xlsx_source_backed_cell_values_one_percent_edit_save']:
        legs = {}
        for leg in ['before', 'after']:
            p = f'alloc-{leg}-{case}-{shape}.json'
            if not os.path.exists(p):
                continue
            r = load(p)
            cv = r['source']['xlsx_cell_values']
            legs[leg] = cv
        if len(legs) != 2:
            continue
        for region in regions:
            if region not in legs['before']:
                continue
            def stat(cv, key):
                vals = [s[key] for s in cv[region] if s.get('status') == 'measured']
                return max(vals) if vals else None
            for key in ['allocation_calls', 'allocated_bytes', 'region_peak_live_bytes']:
                b = stat(legs['before'], key)
                a = stat(legs['after'], key)
                if b is None or a is None:
                    continue
                delta = (a - b) / b * 100 if b else 0.0
                rows.append((shape, case.replace('xlsx_source_backed_cell_values_', ''),
                             region.replace('_allocation_metrics', ''), key, b, a, delta))
        # process peak RSS
        for leg in ['before', 'after']:
            pass

print(f"{'shape':14s} {'case':20s} {'region':14s} {'metric':24s} {'before':>14s} {'after':>14s} {'delta%':>9s}")
for row in rows:
    print(f"{row[0]:14s} {row[1]:20s} {row[2]:14s} {row[3]:24s} {row[4]:14d} {row[5]:14d} {row[6]:+9.2f}")

print()
print("process peak RSS (allocator-instrumented run, max over samples)")
for shape in ['medium', 'dense-sparse', 'noncompact']:
    for case in ['xlsx_source_backed_cell_values_one_edit_save',
                 'xlsx_source_backed_cell_values_one_percent_edit_save']:
        vals = {}
        for leg in ['before', 'after']:
            p = f'alloc-{leg}-{case}-{shape}.json'
            if not os.path.exists(p):
                continue
            r = load(p)
            proc = r['operation_metrics']['process']
            vals[leg] = max(proc['peak_rss_bytes'])
        if len(vals) == 2:
            d = (vals['after'] - vals['before']) / vals['before'] * 100
            print(f"{shape:14s} {case.replace('xlsx_source_backed_cell_values_',''):20s} "
                  f"{vals['before']:14d} {vals['after']:14d} {d:+9.2f}")
