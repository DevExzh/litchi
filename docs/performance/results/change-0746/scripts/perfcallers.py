#!/usr/bin/env python3
"""Bottom-up: for samples whose leaf frame matches LEAF (substring), print the
most common caller chains (DEPTH frames above the leaf, skipping noise)."""
import subprocess, sys, re, collections
data, leaf, depth = sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 6
noise = ('branch', 'from_residual', 'map', 'map_err', 'call_once', '{closure#0}', 'drop_glue', 'drop_in_place', 'drop')
out = subprocess.run(['perf', 'script', '-i', data, '-F', 'period,ip,sym'], capture_output=True, text=True).stdout
def clean(sym):
    sym = re.sub(r'\s*\(inlined\)\s*$', '', sym)
    sym = re.sub(r'\+0x[0-9a-f]+$', '', sym)
    d = 0; s2 = ''
    for ch in sym:
        if ch == '<':
            d += 1
            if d == 1: s2 += '<..>'
            continue
        if ch == '>':
            d -= 1; continue
        if d == 0: s2 += ch
    return s2
total = 0; hit = 0
counter = collections.Counter()
for block in out.split('\n\n'):
    lines = [l for l in block.split('\n') if l.strip()]
    if not lines: continue
    try:
        period = int(lines[0].split()[0])
    except Exception:
        continue
    total += period
    frames = []
    for l in lines[1:]:
        m = re.match(r'\s*[0-9a-f]+\s+(.*)$', l)
        if m: frames.append(clean(m.group(1)))
    if not frames or leaf not in frames[0]:
        continue
    hit += period
    chain = [f for f in frames[1:] if f.split('::')[-1].replace('<..>', '') not in noise][:depth]
    counter[' <- '.join(chain)] += period
print('leaf %s: %.2f%% of samples' % (leaf, 100.0 * hit / total))
for chain, w in counter.most_common(15):
    print('%6.2f%%  %s' % (100.0 * w / total, chain[:400]))
