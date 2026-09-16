#!/usr/bin/env python3
"""0657 timing analysis with an explicit, pre-declared per-leg cleanliness rule.

A leg is CLEAN when p95/p50 <= 1.05.  On this host an undisturbed leg sits at
about 1.01; a leg touched by a neighbour blows its tail out well past 1.05
while its p50 may or may not move.  A paired delta is used only when BOTH of
its legs are clean, so a single disturbed leg cannot create or hide an effect.
"""
import json,os,sys
SCR="/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0657/bench"
SELS=["xlsx_producer_medium_source_planning","xlsx_producer_dense_source_planning",
      "xlsx_producer_medium_source_one_edit_save","xlsx_producer_dense_source_one_edit_save",
      "xlsx_producer_medium_control_planning","xlsx_producer_dense_control_planning"]
FLOORSELS=["xlsx_producer_medium_source_planning","xlsx_producer_dense_source_one_edit_save"]
TAIL=1.05
def load(root,w,sel,leg):
    p=os.path.join(root,w,f"{sel}__{leg}.json")
    if not os.path.exists(p): return None
    for c in json.load(open(p))["results"]:
        if c["case"]==sel:
            e=c["elapsed_ns"]
            with open(os.path.join(root,w,f"{sel}__{leg}.durations.txt"),"w") as g:
                g.write("# elapsed_ns, one measured sample per line, ascending\n")
                for v in sorted(e["samples"]): g.write(f"{v}\n")
            r={"p50":e["p50"],"mean":e["mean"],"p95":e["p95"],"p99":e["p99"],"n":len(e["samples"])}
            r["tail"]=r["p95"]/r["p50"]; r["clean"]=r["tail"]<=TAIL
            return r
    return None
def run(root,label,out):
    o=[];A=o.append
    wins=sorted(d for d in os.listdir(root) if d.startswith("w"))
    A("="*104); A(f"{label}"); A("="*104)
    A(f"Cleanliness rule: a leg counts only if p95/p50 <= {TAIL}. Disturbed legs are listed and excluded.")
    A("")
    A("--- per-leg detail (us); FLAGGED = tail blown, excluded from the estimator ---")
    A(f"{'window':<4} {'selector':<44} {'leg':<3} {'p50':>10} {'mean':>10} {'p95':>10} {'p99':>10} {'p95/p50':>8}")
    dirty=[]
    for w in wins:
        for sel in SELS:
            for leg in ["A1","B1","B2","A2"]:
                r=load(root,w,sel,leg)
                if not r: continue
                fl="" if r["clean"] else "  FLAGGED"
                if not r["clean"]: dirty.append((w,sel,leg,r["tail"]))
                A(f"{w:<4} {sel:<44} {leg:<3} {r['p50']/1e3:>10.2f} {r['mean']/1e3:>10.2f} "
                  f"{r['p95']/1e3:>10.2f} {r['p99']/1e3:>10.2f} {r['tail']:>8.3f}{fl}")
    A("")
    A(f"Disturbed legs excluded: {len(dirty)} of {len(wins)*len(SELS)*4}")
    for w,sel,leg,t in dirty: A(f"   {w} {sel} {leg}  p95/p50={t:.2f}")
    A("")
    A("--- A/A floor from CLEAN same-binary legs only ---")
    fl=[]
    for w in wins:
        for sel in FLOORSELS:
            v={l:load(root,w,sel,l) for l in ["S1","S2","S3","S4"]}
            if any(v[l] is None for l in v): continue
            good=[l for l in ["S1","S2","S3","S4"] if v[l]["clean"]]
            for i in range(len(good)):
                for j in range(i+1,len(good)):
                    a,b=v[good[i]]["p50"],v[good[j]]["p50"]
                    d=abs((b-a)/a*100); fl.append(d)
                    A(f"   {w} {sel[:42]:<42} {good[i]}->{good[j]} {(b-a)/a*100:+7.2f}%")
    FLOOR=max(fl) if fl else 0.0
    A("")
    A(f"A/A FLOOR (max |drift| between clean same-binary legs) = {FLOOR:.2f}%  over {len(fl)} comparisons")
    A("")
    A("--- PAIRED p50 DELTAS, both directions, clean pairs only ---")
    A(f"{'selector':<44} {'pairs':>5} {'A->B mean':>10} {'min':>8} {'max':>8} {'B->A mean':>10}  verdict")
    worse=[]
    for sel in SELS:
        ds=[]
        for w in wins:
            for x,y in [("A1","B1"),("A2","B2")]:
                a,b=load(root,w,sel,x),load(root,w,sel,y)
                if a and b and a["clean"] and b["clean"]:
                    ds.append((b["p50"]-a["p50"])/a["p50"]*100)
        if not ds: A(f"{sel:<44}  no clean pair"); continue
        m=sum(ds)/len(ds)
        rev=sum((-d)/(1+d/100) for d in ds)/len(ds)
        v=("faster than the floor" if m<-FLOOR else
           "within the A/A floor" if m<FLOOR else "SLOWER than the floor")
        A(f"{sel:<44} {len(ds):>5} {m:>+9.2f}% {min(ds):>+7.2f}% {max(ds):>+7.2f}% {rev:>+9.2f}%  {v}")
        if m>5.0: worse.append((sel,m))
    A("")
    A("(A->B negative = AFTER faster. B->A is the same pairs measured the other way.)")
    A("")
    if worse:
        A("!!! WORSE BY MORE THAN 5% AT p50 !!!")
        for s,d in worse: A(f"    {s}: {d:+.2f}%")
    else:
        A("No scenario is worse by more than 5% at p50 on clean pairs.")
    t="\n".join(o)
    open(os.path.join(SCR,out),"w").write(t+"\n")
    return t,FLOOR
if __name__=="__main__":
    t,_=run(os.path.join(SCR,sys.argv[1]),sys.argv[2],sys.argv[3])
    print(t)
