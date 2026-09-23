#!/usr/bin/env python3
"""kids.py FOLDED PARENT_RX [DEPTH] [TOP]: period-weighted breakdown of the first DEPTH non-glue
frames leafward of the OUTERMOST frame matching PARENT_RX (kernel frames shown as K)."""
import re, sys
from collections import Counter
path, prx = sys.argv[1], re.compile(sys.argv[2])
depth = int(sys.argv[3]) if len(sys.argv) > 3 else 1
top = int(sys.argv[4]) if len(sys.argv) > 4 else 30
glue = re.compile(r'^(branch<|map_err<|and_then<|map<|into<|from$|from<|unwrap_or_else<|ok_or_else<|call_once|\{closure|try_fold|try_for_each|for_each<|fold<)')
tot = 0; inp = 0; c = Counter(); comms = Counter()
for line in (__import__('gzip').open(path,'rt') if path.endswith('.gz') else open(path)):
    p, n, comm, st = line.rstrip('\n').split('\t')
    p = int(p); comms[comm] += p
    if comm.startswith('perf-exec'): continue
    tot += p
    fr = st.split(' ;; ')
    idx = None
    for i, f in enumerate(fr):
        if prx.search(f.split('|R:')[0]): idx = i; break
    if idx is None: continue
    inp += p
    ch = []
    j = idx + 1
    while j < len(fr) and len(ch) < depth:
        s = fr[j].split('|R:')[0]
        if not glue.search(s):
            ch.append(re.sub(r'<.*', '<..>', s)[:60])
        j += 1
    c[' > '.join(ch) if ch else '<self>'] += p
print(f'{path.split("/")[-1]}: process={tot:.3e} in-parent={inp/tot*100:.1f}%  comms={dict(comms)}')
for k, v in c.most_common(top): print(f'{100*v/inp:6.2f}%  {k}')
