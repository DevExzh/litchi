import sys, pathlib
S = pathlib.Path(sys.argv[1])
def load(tag, leg):
    txt = (S / f"probe-time-{tag}-{leg}.txt").read_text()
    line = [l for l in txt.splitlines() if l.startswith("timing_samples_ns=")][0]
    return [int(v) for v in line.split("=", 1)[1].split(",")]
def q(v, p):
    s = sorted(v); return s[round((len(s)-1)*p)]
def stats(v):
    return {"p50": q(v,0.5), "mean": sum(v)/len(v), "p95": q(v,0.95), "p99": q(v,0.99), "n": len(v)}
print("Paired timing through the retained scratch probe, in-process, file source,")
print("10 warmups then 40 timed samples per leg, CPU 22, order A1 B1 B2 A2.")
print("Positive percent = the after leg is faster. A/A is |A1-A2|/A1 in the same window.")
print()
print(f"{'scenario':<28}{'stat':<6}{'A1':>12}{'B1':>12}{'B2':>12}{'A2':>12}{'A1->B1 %':>11}{'A2->B2 %':>11}{'A/A %':>9}")
for tag in sys.argv[2:]:
    L = {leg: stats(load(tag, leg)) for leg in ("A1","B1","B2","A2")}
    for s in ("p50","mean","p95","p99"):
        a1,b1,b2,a2 = (L[k][s] for k in ("A1","B1","B2","A2"))
        print(f"{tag:<28}{s:<6}{a1:>12.0f}{b1:>12.0f}{b2:>12.0f}{a2:>12.0f}{(a1-b1)/a1*100:>11.2f}{(a2-b2)/a2*100:>11.2f}{abs((a1-a2)/a1*100):>9.2f}")
    print(f"{tag:<28}{'n':<6}" + "".join(f"{L[k]['n']:>12}" for k in ("A1","B1","B2","A2")))
    print()
