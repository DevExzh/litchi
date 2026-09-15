import json, sys, pathlib
S=pathlib.Path(sys.argv[1]); tag=sys.argv[2]
def rows(leg):
    return [json.loads(l) for l in (S/f"inter-{tag}-{leg}.jsonl").read_text().splitlines() if l.strip()]
def q(v,p):
    s=sorted(v); return s[round((len(s)-1)*p)]
B=rows("before"); A=rows("after")
print(f"n = {len(B)} before, {len(A)} after (interleaved sample by sample)")
for metric, get in (("elapsed_ns", lambda r: r["elapsed_ns"]),
                    ("minor_faults", lambda r: r["process_metrics"]["minor_faults"]),
                    ("syscr", lambda r: r["process_metrics"]["syscr"])):
    b=[get(r) for r in B]; a=[get(r) for r in A]
    print(f"\n-- {metric}")
    print(f"{'stat':<6}{'before':>12}{'after':>12}{'delta':>12}{'reduction %':>13}")
    for s,f in (("min",min),("p25",lambda v:q(v,0.25)),("p50",lambda v:q(v,0.5)),
                ("mean",lambda v:sum(v)/len(v)),("p75",lambda v:q(v,0.75)),("p95",lambda v:q(v,0.95))):
        bv,av=f(b),f(a)
        print(f"{s:<6}{bv:>12.1f}{av:>12.1f}{av-bv:>12.1f}{(bv-av)/bv*100:>13.2f}")
