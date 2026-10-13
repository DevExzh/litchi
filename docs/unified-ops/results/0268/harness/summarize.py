import glob,statistics,sys,collections
R=sys.argv[1]
d=collections.defaultdict(lambda: collections.defaultdict(list))
for f in sorted(glob.glob(R+'/*.tsv')):
    tag='before' if 'before' in f else 'after'
    for line in open(f):
        k,v=line.split('\t'); d[k][tag].append(int(v))
print("| Operation | before median (µs) | after median (µs) | delta |")
print("|---|---:|---:|---:|")
for k in d:
    b=d[k].get('before'); a=d[k].get('after')
    bm=statistics.median(b)/1000 if b else None; am=statistics.median(a)/1000
    if bm: print(f"| {k} | {bm:.1f} | {am:.1f} | {100*(am-bm)/bm:+.1f}% |")
    else: print(f"| {k} | — | {am:.1f} | — |")
