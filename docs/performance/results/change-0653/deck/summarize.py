"""Summarize the 0649-shaped real-deck phase measurement for change 0653.

The A/A floor is the widest p50 spread across the repeats of one leg in the
same window, exactly as change 0649 measured it.
"""
import glob, os, sys, collections
OUT = sys.argv[1]
rows = collections.defaultdict(lambda: collections.defaultdict(list))
for path in sorted(glob.glob(os.path.join(OUT, "*.tsv"))):
    deck, leg, _repeat, _ = os.path.basename(path).split(".")
    with open(path) as handle:
        for line in handle:
            parts = line.rstrip("\n").split("\t")
            if len(parts) == 5 and parts[1].isdigit():
                rows[(deck, parts[0])][leg].append(int(parts[1]))
print("%-42s %14s %14s %9s %9s" % ("deck / phase", "before p50", "after p50", "delta", "A/A"))
for (deck, phase), legs in rows.items():
    before, after = sorted(legs.get("before", [])), sorted(legs.get("after", []))
    if not before or not after:
        continue
    b = before[len(before) // 2]
    a = after[len(after) // 2]
    floor = 100.0 * (max(before) - min(before)) / min(before)
    print("%-42s %14d %14d %+8.2f%% %8.2f%%" % (f"{deck} / {phase}", b, a, 100.0 * (a - b) / b, floor))
