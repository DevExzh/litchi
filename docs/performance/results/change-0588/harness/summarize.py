"""Summarize the change-0588 harness ABBA: p50/p95/p99 per selector per leg."""
import json
import sys

OUT = sys.argv[1]


def load(*names):
    rows = {}
    for name in names:
        data = json.load(open(f"{OUT}/{name}.json"))
        for result in data["results"]:
            key = (result.get("case"), result.get("shape"), result.get("repeat"))
            rows.setdefault(key, []).append(result["elapsed_ns"])
    return rows


def combine(samples, field):
    return sum(entry[field] for entry in samples) / len(samples)


before = load("before-1", "before-2")
after = load("after-1", "after-2")
floor_a = load("floorA")
floor_b = load("floorB")
print("%-34s %12s %12s %8s %8s" % ("case", "before p50", "after p50", "delta", "floor"))
for key in sorted(before, key=lambda k: str(k)):
    if key not in after:
        continue
    b50 = combine(before[key], "p50")
    a50 = combine(after[key], "p50")
    f = 100.0 * (combine(floor_b[key], "p50") - combine(floor_a[key], "p50")) / combine(
        floor_a[key], "p50"
    )
    print("%-34s %12.0f %12.0f %+7.2f%% %+7.2f%%"
          % (key[0], b50, a50, 100.0 * (a50 - b50) / b50, f))
