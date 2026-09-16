"""Gate G1's mechanical half: for every sweep pair, the leading OK window, the
trailing OK window, the absence of any interior OK, and the multiset of error
classes are compared between the two legs."""
import sys, os, glob, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sweepcmp import load, leading_ok, trailing_ok, interior_ok

problems = []
rows = []
for before in sorted(glob.glob(os.path.join(sys.argv[1], '*.tsv'))):
    name = os.path.basename(before)
    after = os.path.join(sys.argv[2], name)
    _, rb = load(before)
    _, ra = load(after)
    lb, la = leading_ok(rb), leading_ok(ra)
    tb, ta = trailing_ok(rb), trailing_ok(ra)
    cb = collections.Counter(c for _, c in rb if c != 'OK')
    ca = collections.Counter(c for _, c in ra if c != 'OK')
    if lb != la:
        problems.append(f"{name}: leading OK window {lb} -> {la}")
    if tb != ta:
        problems.append(f"{name}: trailing OK window {tb} -> {ta}")
    if interior_ok(ra):
        problems.append(f"{name}: interior OK at {interior_ok(ra)}")
    moved = sorted(f"{k} {cb.get(k, 0)}->{ca.get(k, 0)}"
                   for k in set(cb) | set(ca) if cb.get(k, 0) != ca.get(k, 0))
    rows.append((name, lb, la, tb, ta, len(rb), len(ra), moved))

print("sweep\tleadOK\ttrailOK\tordinals\trefused\tclass_counts_changed")
for name, lb, la, tb, ta, nb, na, moved in rows:
    print(f"{name}\t{lb}->{la}\t{tb}->{ta}\t{nb}->{na}\t{nb-lb-tb}->{na-la-ta}\t"
          f"{' | '.join(moved) or 'none'}")
print()
if problems:
    print("PROBLEMS:")
    for p in problems:
        print("  " + p)
else:
    print("G1 mechanical half: PASS — every leg pair agrees on the leading OK window, "
          "the trailing OK window, and has no interior OK.")
