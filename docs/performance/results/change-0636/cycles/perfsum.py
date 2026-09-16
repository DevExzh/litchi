"""Difference perf stat isolation pairs for change 0636."""
import glob, os, sys
S = sys.argv[1]
def read(path):
    out = {}
    for line in open(path):
        parts = line.strip().split(',')
        if len(parts) > 2 and parts[2] in ('cycles', 'instructions'):
            out[parts[2]] = float(parts[0])
    return out
labels = sorted({os.path.basename(p).split('-', 1)[1].rsplit('-', 2)[0]
                 for p in glob.glob(S + '/cg/perf-*-before-20.txt')})
print(f"{'scenario':22} {'metric':12} {'before/op':>14} {'after/op':>14} {'delta':>9}")
for label in labels:
    for metric in ('cycles', 'instructions'):
        vals = {}
        for leg in ('before', 'after'):
            lo = read(f'{S}/cg/perf-{label}-{leg}-20.txt')[metric]
            hi = read(f'{S}/cg/perf-{label}-{leg}-120.txt')[metric]
            vals[leg] = (hi - lo) / 100.0
        delta = (vals['after'] - vals['before']) / vals['before'] * 100
        print(f"{label:22} {metric:12} {vals['before']:14,.0f} {vals['after']:14,.0f} {delta:+8.2f}%")
