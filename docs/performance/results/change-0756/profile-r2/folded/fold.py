#!/usr/bin/env python3
"""Fold a perf-script text file into weighted stacks.
Output lines: period<TAB>count<TAB>comm<TAB>frames(root->leaf, joined by ' ;; ')
Real (non-inlined) frames get suffix '|R:<dso>'; kernel frames are 'K'."""
import re, sys, os
from collections import defaultdict
src, dst = sys.argv[1], sys.argv[2]
agg = defaultdict(lambda: [0, 0])
off = re.compile(r'\+0x[0-9a-f]+$')
hdr = None; stack = []
def flush():
    if hdr is None: return
    parts = hdr.split()
    comm = parts[0]
    # period is the number before the event name
    period = 1
    for i, p in enumerate(parts):
        if p.endswith(':') and i + 1 < len(parts) and parts[i+1].isdigit():
            period = int(parts[i+1]); break
    key = (comm, ' ;; '.join(reversed(stack)))
    a = agg[key]; a[0] += period; a[1] += 1
with open(src, errors='replace') as fh:
    for line in fh:
        if line == '\n' or not line.strip():
            flush(); hdr = None; stack = []; continue
        if hdr is None:
            hdr = line.rstrip('\n'); continue
        t = line.strip()
        sp = t.find(' ')
        if sp < 0: continue
        addr = t[:sp]; rest = t[sp+1:]
        i = rest.rfind(' (')
        if i < 0: continue
        sym = off.sub('', rest[:i]); where = rest[i+2:-1]
        if addr.startswith('ffffffff'):
            stack.append('K'); continue
        if where == 'inlined':
            stack.append(sym)
        else:
            stack.append(sym + '|R:' + os.path.basename(where))
    flush()
with open(dst, 'w') as out:
    for (comm, st), (p, c) in sorted(agg.items(), key=lambda kv: -kv[1][0]):
        out.write(f'{p}\t{c}\t{comm}\t{st}\n')
print(dst, len(agg))
