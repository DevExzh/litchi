#!/usr/bin/env python3
"""Change 0754: per-iteration instructions, cycles, allocations and requested
bytes of each probe loop (60-iteration run minus 10-iteration run, / 50)."""
import json, os, re, sys
d = sys.argv[1]
def perf(p):
    out = {}
    for line in open(p):
        f = line.strip().split(',')
        if len(f) > 2 and f[0].isdigit():
            out[f[2]] = int(f[0])
    return out
def alloc(p):
    m = re.search(r'allocations (\d+) bytes (\d+)', open(p).read())
    return int(m.group(1)), int(m.group(2))
rows = []
for kind in ('scan', 'noop', 'one', 'text'):
    for corpus in ('semantic-large-document.xml', 'semantic-large.docx', 'semantic-medium-document.xml', 'semantic-medium.docx'):
        if not os.path.exists(f'{d}/A-{kind}-{corpus}-10.perf'):
            continue
        row = {'loop': kind, 'input': corpus}
        for leg in 'AB':
            p10, p60 = perf(f'{d}/{leg}-{kind}-{corpus}-10.perf'), perf(f'{d}/{leg}-{kind}-{corpus}-60.perf')
            a10, a60 = alloc(f'{d}/{leg}-{kind}-{corpus}-10.log'), alloc(f'{d}/{leg}-{kind}-{corpus}-60.log')
            row[leg] = {'instructions': (p60['instructions'] - p10['instructions']) / 50,
                        'cycles': (p60['cycles'] - p10['cycles']) / 50,
                        'allocations': (a60[0] - a10[0]) / 50, 'requested_bytes': (a60[1] - a10[1]) / 50}
        rows.append(row)
json.dump(rows, open(f'{d}/summary.json', 'w'), indent=1)
print('| loop | input | instructions before → after | cycles before → after | allocations before → after | requested bytes before → after |')
print('| --- | --- | --- | --- | --- | --- |')
for r in rows:
    a, b = r['A'], r['B']
    print(f"| `{r['loop']}` | {r['input']} | {a['instructions']/1e6:.2f} M → {b['instructions']/1e6:.2f} M ({(b['instructions']/a['instructions']-1)*100:+.1f}%) | {a['cycles']/1e6:.2f} M → {b['cycles']/1e6:.2f} M ({(b['cycles']/a['cycles']-1)*100:+.1f}%) | {a['allocations']:,.0f} → {b['allocations']:,.0f} | {a['requested_bytes']/1e6:.2f} MB → {b['requested_bytes']/1e6:.2f} MB |")
