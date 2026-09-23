"""Attribute perf FP samples inside a root frame to phases and leaf symbols.

Usage: attrib.py SCRIPT ROOT_SUBSTRING [PHASE=substr ...]
Samples are separated by blank lines; frames listed leaf-first.
A sample counts if any frame contains ROOT_SUBSTRING. Its phase is the
first (outermost-closest-to-root) matching phase pattern searched from the
root towards the leaf, using the ordered phase list (first listed wins
when several match at the same depth).
"""
import sys, re, collections
path, root = sys.argv[1], sys.argv[2]
phases = []
for arg in sys.argv[3:]:
    name, pat = arg.split('=', 1)
    phases.append((name, pat))
samples = []
cur = []
for line in open(path, errors='replace'):
    line = line.rstrip('\n')
    if not line.strip():
        if cur:
            samples.append(cur)
            cur = []
        continue
    parts = line.strip().split(None, 1)
    sym = parts[1] if len(parts) > 1 else '?'
    sym = re.sub(r'\+0x[0-9a-f]+$', '', sym)
    cur.append(sym)
if cur:
    samples.append(cur)
inside = [s for s in samples if any(root in f for f in s)]
print(f"total samples {len(samples)}, inside root {len(inside)}")
phase_counts = collections.Counter()
leaf_by_phase = collections.defaultdict(collections.Counter)
for s in inside:
    # frames leaf-first; find root index
    ridx = max(i for i, f in enumerate(s) if root in f)
    chain = s[:ridx]  # frames below root, leaf-first
    ph = 'other'
    for name, pat in phases:
        if any(pat in f for f in chain):
            ph = name
            break
    phase_counts[ph] += 1
    leaf_by_phase[ph][s[0] if s else '?'] += 1
n = len(inside)
for ph, c in phase_counts.most_common():
    print(f"{ph:30s} {c:6d} {100.0*c/n:6.2f}%")
if '--leaves' in sys.argv or True:
    for ph, c in phase_counts.most_common():
        print(f"\n== {ph} top leaves")
        for leaf, lc in leaf_by_phase[ph].most_common(12):
            print(f"   {lc:6d} {100.0*lc/n:6.2f}%  {leaf[:150]}")
