import collections, re, statistics, sys
SP=sys.argv[1]
def load(p):
    d=collections.defaultdict(dict)
    for line in open(p):
        f=line.rstrip("\n").split(" ",2)
        d[f[0]][f[1]]=f[2] if len(f)>2 else ""
    return d
b=load(f"{SP}/docx-census-before.txt"); a=load(f"{SP}/docx-census-after.txt")
assert set(b)==set(a)
moved=collections.Counter(); same=collections.Counter()
for fx in b:
    for asp in b[fx]:
        (same if b[fx][asp]==a[fx].get(asp) else moved)[asp]+=1
print(f"fixtures: {len(b)}")
print("aspect                unchanged  moved")
for asp in ["open","source","gate","gate_source","noop_save","exact_noop","managed_edit","one_edit","one_edit_optin","one_edit_span","one_edit_reopen"]:
    print(f"  {asp:<20} {same.get(asp,0):>8}  {moved.get(asp,0):>5}")
g=collections.Counter((b[fx].get("gate","-").split(":")[0], a[fx].get("gate_source","-").split(":")[0]) for fx in b)
print("\ngate verdict transition (verify_authored -> verify_source):")
for k,v in sorted(g.items()): print(f"  {k[0]} -> {k[1]}: {v}")
for fx in b:
    if a[fx].get("gate_source","").startswith("refuses"):
        print(f"  still refused: {fx} :: {a[fx]['gate_source']}")
def nums(v): return {k:int(x) for k,x in re.findall(r"(\w+)=(\d+)", v)}
rows=[(fx,nums(b[fx]["one_edit_span"]),nums(a[fx]["one_edit_span"])) for fx in sorted(b)
      if "skipped" not in b[fx].get("one_edit_span","skipped") and "skipped" not in a[fx].get("one_edit_span","skipped")]
print(f"\none_edit_span, {len(rows)} fixtures that admit the edit:")
for label,i in (("before",1),("after",2)):
    ms=[r[i]["moved_source"] for r in rows]
    print(f"  {label}: source bytes rewritten total={sum(ms)} median={statistics.median(ms)} max={max(ms)}"
          f" published total={sum(r[i]['published'] for r in rows)}")
ch=[r for r in rows if b[r[0]].get("one_edit")!=a[r[0]].get("one_edit")]
print(f"  fixtures whose published bytes moved: {len(ch)}")
print(f"  source bytes rewritten on those: before {sum(r[1]['moved_source'] for r in ch)} -> after {sum(r[2]['moved_source'] for r in ch)}")
print("\n  largest reductions:")
for fx,x,y in sorted(ch,key=lambda r:r[1]['moved_source']-r[2]['moved_source'],reverse=True)[:6]:
    print(f"    {fx.split('/')[-1]:<55} source={x['source']:>7} rewritten {x['moved_source']:>7} -> {y['moved_source']:>6}")
