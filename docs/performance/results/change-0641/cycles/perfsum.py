#!/usr/bin/env python3
"""Fold change 0641's `perf stat` isolation pairs.

Each leg is profiled at LOW and HIGH samples; the difference divided by
HIGH-LOW is one operation. Five repetitions per leg; the median of the five is
reported. `before2` is the before binary run a second time in the same window,
so its distance from `before` is the A/A floor for this metric.
"""
import statistics, sys
S = "/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0641"
LOW, HIGH, REPS = 20, 120, 5

def read(path, event):
    for line in open(path):
        parts = line.split(",")
        if len(parts) > 2 and parts[2].strip() == event:
            return float(parts[0])

def per_op(label, leg, event):
    out = []
    for rep in range(1, REPS + 1):
        hi = read(f"{S}/perf/{label}-{leg}-{HIGH}-{rep}.txt", event)
        lo = read(f"{S}/perf/{label}-{leg}-{LOW}-{rep}.txt", event)
        out.append((hi - lo) / (HIGH - LOW))
    return statistics.median(out)

for event in ("cycles", "instructions"):
    print(f"perf stat isolation pairs, {event} per operation (median of {REPS} repetitions, CPU 24)")
    print("%-24s %13s %13s %9s %9s" % ("scenario", "before", "after", "delta", "A/A floor"))
    for line in open(sys.argv[1]):
        label = line.split("|")[0].strip()
        if not label:
            continue
        b = per_op(label, "before", event)
        a = per_op(label, "after", event)
        f = per_op(label, "before2", event)
        print("%-24s %13.0f %13.0f %8.2f%% %8.2f%%" %
              (label, b, a, 100.0 * (a - b) / b, 100.0 * (f - b) / b))
    print()
