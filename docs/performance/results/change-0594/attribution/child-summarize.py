import json, sys, pathlib, statistics
S=pathlib.Path(sys.argv[1]); tag=sys.argv[2]
def load(leg):
    rows=[json.loads(l) for l in (S/f"child-{tag}-{leg}.jsonl").read_text().splitlines() if l.strip()]
    return rows
def q(v,p):
    s=sorted(v); return s[round((len(s)-1)*p)]
legs={}
for leg in ("A1","B1","B2","A2"):
    rows=load(leg)
    legs[leg]={
      "elapsed": [r["elapsed_ns"] for r in rows],
      "minor_faults": [r["process_metrics"]["minor_faults"] for r in rows],
      "syscr": [r["process_metrics"]["syscr"] for r in rows],
      "rchar": [r["process_metrics"]["rchar"] for r in rows],
    }
print(f"n per leg = {len(legs['A1']['elapsed'])}")
for metric in ("elapsed","minor_faults","syscr","rchar"):
    print(f"\n-- {metric}")
    print(f"{'stat':<6}{'A1':>12}{'B1':>12}{'B2':>12}{'A2':>12}{'A1->B1 %':>11}{'A2->B2 %':>11}{'A/A %':>9}")
    for s,f in (("p50",lambda v:q(v,0.5)),("mean",lambda v:sum(v)/len(v)),("p95",lambda v:q(v,0.95)),("min",min)):
        a1,b1,b2,a2=(f(legs[k][metric]) for k in ("A1","B1","B2","A2"))
        print(f"{s:<6}{a1:>12.1f}{b1:>12.1f}{b2:>12.1f}{a2:>12.1f}{(a1-b1)/a1*100:>11.2f}{(a2-b2)/a2*100:>11.2f}{abs((a1-a2)/a1*100):>9.2f}")
