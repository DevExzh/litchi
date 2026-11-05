import json, glob, statistics, collections
D='./'
rows=collections.defaultdict(lambda: {'before':[], 'after':[]})
for f in glob.glob(D+'*-*-[0-9].json'):
    kind, leg, _ = f.rsplit('/',1)[1][:-5].split('-')
    for r in json.load(open(f)):
        size=r.get('size', r.get('slides'))
        rows[(kind, r['scenario'], size)][leg].append(r)
out=[]
for key in sorted(rows, key=str):
    b=rows[key]['before']; a=rows[key]['after']
    if not a or not b: continue
    med=lambda xs,k: statistics.median(x[k] for x in xs)
    nb,na=med(b,'native_median_us'),med(a,'native_median_us')
    ub,ua=med(b,'unified_median_us'),med(a,'unified_median_us')
    pb,pa=med(b,'pair_ratio_median'),med(a,'pair_ratio_median')
    out.append(dict(kind=key[0],scenario=key[1],size=key[2],native_before_us=nb,native_after_us=na,native_ratio=na/nb,unified_before_us=ub,unified_after_us=ua,unified_ratio=ua/ub,pair_before=pb,pair_after=pa))
    print(f"{key[0]:8} {key[1]:34} {str(key[2]):>6} native {nb:10.1f}->{na:10.1f} ({na/nb:.3f})  unified {ub:10.1f}->{ua:10.1f} ({ua/ub:.3f})  pair {pb:.3f}->{pa:.3f}")
json.dump(out, open(D+'summary.json','w'), indent=1)
