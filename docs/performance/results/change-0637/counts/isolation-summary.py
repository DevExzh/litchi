"""Difference the callgrind isolation pairs into per-operation Ir.

Input: the raw `<tag>\t<Ir>` totals, tags shaped `<leg>-<deck>-<op>-<iters>`.
Each (leg, deck, op) has two iteration counts; the per-operation Ir is the
difference of the totals divided by the difference of the iteration counts, so
process start-up, package load and probe setup cancel.
"""
import sys, collections

rows = {}
for line in open(sys.argv[1]):
    line = line.strip()
    if not line:
        continue
    tag, ir = line.split("\t")
    parts = tag.rsplit("-", 1)
    key, iters = parts[0], int(parts[1])
    rows.setdefault(key, {})[iters] = int(ir)

per_op = {}
for key, samples in sorted(rows.items()):
    if len(samples) != 2:
        print(f"# incomplete pair: {key} {samples}")
        continue
    (n1, t1), (n2, t2) = sorted(samples.items())
    per_op[key] = (t2 - t1) / (n2 - n1)

legs = ["before", "after"]
if len(sys.argv) > 2:
    legs = sys.argv[2].split(",")
ops = []
for key in per_op:
    leg, deck, op = key.split("-", 2)
    if (deck, op) not in ops:
        ops.append((deck, op))

head = f"{'deck':9s} {'operation':12s}" + "".join(f"{leg + ' Ir/op':>18s}" for leg in legs)
print(head + f"{'delta':>18s}{'delta %':>10s}")
print("-" * len(head + f"{'delta':>18s}{'delta %':>10s}"))
for deck, op in ops:
    values = [per_op.get(f"{leg}-{deck}-{op}") for leg in legs]
    if any(v is None for v in values):
        continue
    base, last = values[0], values[-1]
    delta = last - base
    pct = 100.0 * delta / base if base else 0.0
    cells = "".join(f"{v:18,.0f}" for v in values)
    print(f"{deck:9s} {op:12s}{cells}{delta:18,.0f}{pct:9.2f}%")
