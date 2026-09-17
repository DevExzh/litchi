import collections, re, sys, os
SP=os.path.dirname(os.path.abspath(__file__))
def load(p):
    rows={}
    for line in open(p):
        f=line.rstrip("\n").split("\t")
        rows[f[0]]=f[1:]
    return rows
def summarize(name):
    b=load(f"{SP}/out/{name}-before.tsv"); a=load(f"{SP}/out/{name}-after.tsv")
    assert set(b)==set(a), name
    trans=collections.Counter(); digest_same=0; digest_diff=[]; newly=[]; kept_err=0; err_text_same=0; err_diff=[]
    for k in b:
        bs,as_=b[k][0],a[k][0]
        trans[(bs,as_)]+=1
        if bs=="ok" and as_=="ok":
            if b[k][3]==a[k][3]: digest_same+=1
            else: digest_diff.append(k)
        elif bs!="ok" and as_=="ok": newly.append((k,b[k][2]))
        elif bs!="ok" and as_!="ok":
            kept_err+=1
            if b[k][2]==a[k][2]: err_text_same+=1
            else: err_diff.append((k,b[k][2],a[k][2]))
    print(f"### {name}: {len(b)} fixtures")
    for (x,y),n in sorted(trans.items()): print(f"    {x} -> {y}: {n}")
    print(f"    published on both, identical digest: {digest_same}; differing: {len(digest_diff)}")
    for k in digest_diff[:5]: print(f"      DIGEST DIFF {os.path.basename(k)}")
    print(f"    newly published: {len(newly)}")
    nc=[x for x in newly if "NotCompact" in x[1]]
    print(f"      of which the before-leg error was NotCompact: {len(nc)}")
    other=[x for x in newly if "NotCompact" not in x[1]]
    for k,e in other[:6]: print(f"      newly published, other before-error: {os.path.basename(k)} :: {e[:90]}")
    print(f"    refused on both: {kept_err}; identical error text: {err_text_same}")
    for k,x,y in err_diff[:6]: print(f"      ERROR TEXT MOVED {os.path.basename(k)}\n        before={x[:110]}\n        after ={y[:110]}")
    # residual NotCompact on the after leg
    res=[k for k in a if a[k][0]!="ok" and "NotCompact" in a[k][2]]
    print(f"    residual NotCompact refusals on the after leg: {len(res)}")
    print()
for name in ["noop","edit","editnc","hide"]:
    summarize(name)
