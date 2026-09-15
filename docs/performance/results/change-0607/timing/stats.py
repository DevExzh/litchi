import statistics as st, sys, pathlib
def load(p): return sorted(float(x) for x in pathlib.Path(p).read_text().split())
def q(v, f): 
    i = min(len(v)-1, int(round(f*(len(v)-1))))
    return v[i]
def line(name, v):
    return f"{name}\tn={len(v)}\tp50={q(v,0.5)/1000:.1f}\tmean={st.mean(v)/1000:.1f}\tp95={q(v,0.95)/1000:.1f}\tp99={q(v,0.99)/1000:.1f} us"
d = sys.argv[1]
legs = {n: load(f"{d}/{n}.txt") for n in ("A1","B1","B2","A2")}
for n in ("A1","A2","B1","B2"): print(line(n, legs[n]))
A = sorted(legs["A1"]+legs["A2"]); B = sorted(legs["B1"]+legs["B2"])
print(line("A(resave, pooled)", A)); print(line("B(nopsave, pooled)", B))
print(f"delta p50 (A-B)\t{(q(A,0.5)-q(B,0.5))/1000:.1f} us\t{(q(A,0.5)-q(B,0.5))/q(A,0.5)*100:.2f}% of A")
print(f"delta mean (A-B)\t{(st.mean(A)-st.mean(B))/1000:.1f} us\t{(st.mean(A)-st.mean(B))/st.mean(A)*100:.2f}% of A")
print(f"A/A floor (A2 vs A1) p50\t{(q(legs['A2'],0.5)-q(legs['A1'],0.5))/q(legs['A1'],0.5)*100:+.2f}%\tp99 {(q(legs['A2'],0.99)-q(legs['A1'],0.99))/q(legs['A1'],0.99)*100:+.2f}%")
print(f"B/B floor (B2 vs B1) p50\t{(q(legs['B2'],0.5)-q(legs['B1'],0.5))/q(legs['B1'],0.5)*100:+.2f}%\tp99 {(q(legs['B2'],0.99)-q(legs['B1'],0.99))/q(legs['B1'],0.99)*100:+.2f}%")
