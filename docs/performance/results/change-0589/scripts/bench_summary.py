import pathlib, statistics
base = pathlib.Path("/tmp/claude-1001/-home-zhuhe-code-litchi/709d31e6-bf68-4968-8b5c-9a2af1c22cc8/scratchpad/agents/0589/out")
def stats(vals):
    vals = sorted(vals)
    q = lambda p: vals[min(len(vals)-1, int(round(p*(len(vals)-1))))]
    return dict(n=len(vals), p50=statistics.median(vals), mean=statistics.mean(vals), p95=q(0.95), p99=q(0.99))
def load(stem, leg):
    return [int(x) for x in (base / f"bench-{stem}-{leg}.txt").read_text().split()]
for stem, label in (("docpic","doc-snapshot-open picture.doc (1,448,448 B)"),
                    ("docdup","doc-snapshot-open duplicate-style-names.doc (64,512 B)"),
                    ("pptbig","ppt-textedit-open cryptoapi-proc2356.ppt (1,341,952 B)"),
                    ("pptmid","ppt-textedit-open 45543.ppt (385,024 B)")):
    a = load(stem,"A1") + load(stem,"A2")
    b = load(stem,"B1") + load(stem,"B2")
    aa1, aa2 = load(stem,"A1")+load(stem,"A3"), load(stem,"A2")+load(stem,"A4")
    sa, sb = stats(a), stats(b)
    fa, fb = stats(aa1), stats(aa2)
    print(f"\n{label}")
    print(f"  {'leg':10s} {'n':>4s} {'p50 us':>10s} {'mean us':>10s} {'p95 us':>10s} {'p99 us':>10s}")
    for name, s in (("before", sa), ("after", sb)):
        print(f"  {name:10s} {s['n']:4d} {s['p50']/1000:10.1f} {s['mean']/1000:10.1f} {s['p95']/1000:10.1f} {s['p99']/1000:10.1f}")
    for metric in ("p50","mean","p95","p99"):
        d_ab = (sb[metric]-sa[metric])/sa[metric]*100
        d_ba = (sa[metric]-sb[metric])/sb[metric]*100
        print(f"  after vs before {metric:5s}: {d_ab:+7.1f}%   before vs after: {d_ba:+7.1f}%")
    for metric in ("p50","p99"):
        print(f"  A/A floor {metric:5s}: {(fb[metric]-fa[metric])/fa[metric]*100:+7.1f}%")
