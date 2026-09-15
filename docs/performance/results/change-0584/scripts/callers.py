import re,sys,pathlib
NUM=re.compile(r"^\s*([\d,]+) \(\s*[\d.]+%\)\s+(.*)$")
def edges(p):
    """returns {(caller,callee):(calls,ir)} from a --tree=caller annotation"""
    out={}; pending=[]
    for line in pathlib.Path(p).read_text().splitlines():
        m=NUM.match(line)
        if not m: pending=[]; continue
        rest=m.group(2); ir=int(m.group(1).replace(",",""))
        if rest.startswith("*"):
            name=rest[1:].strip(); name=name.split(":",1)[1].split(" [")[0].strip() if ":" in name else name
            for c,calls,cir in pending: out[(c,name)]=(calls,cir)
            pending=[]
        elif rest.startswith("<"):
            body=rest[1:].strip()
            cm=re.search(r"\((\d[\d,]*)x\)",body)
            calls=int(cm.group(1).replace(",","")) if cm else 0
            name=body.split(":",1)[1].split(" (")[0].strip() if ":" in body else body
            pending.append((name,calls,ir))
        else: pending=[]
    return out
def diff(stem,op,s,l,targets):
    A=edges(f"{sys.argv[1]}/tree-head-{stem}-{op}-s{s}.txt"); B=edges(f"{sys.argv[1]}/tree-head-{stem}-{op}-s{l}.txt")
    d=l-s; rows=[]
    for k in set(A)|set(B):
        if not any(t in k[1] for t in targets): continue
        ca=(B.get(k,(0,0))[0]-A.get(k,(0,0))[0])/d; ir=(B.get(k,(0,0))[1]-A.get(k,(0,0))[1])/d
        if ir>10: rows.append((ir,ca,k[0],k[1]))
    return sorted(rows,reverse=True)
for stem,op,s,l in [("flagship","open",20,220),("flagship","one-cell",20,220),("54016","one-cell",10,60)]:
    print(f"\n### {stem}/{op} — who calls memcpy / memset (Ir per operation, calls per operation)")
    for ir,ca,caller,callee in diff(stem,op,s,l,["memcpy_avx","memset_avx"])[:12]:
        short="memcpy" if "memcpy" in callee else "memset"
        print(f"  {ir:11,.0f} Ir  {ca:8,.1f} calls  {short:6} <- {caller[:95]}")
