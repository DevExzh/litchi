import subprocess, re, collections, sys
def funcs(path):
    out = subprocess.run(['callgrind_annotate','--inclusive=no','--threshold=100',path],capture_output=True,text=True).stdout
    d = collections.Counter(); total=None
    for line in out.splitlines():
        m = re.match(r'\s*([\d,]+) \([ \d.]+%\)\s+(.*)', line)
        if not m: continue
        n = int(m.group(1).replace(',',''))
        name = m.group(2)
        if 'PROGRAM TOTALS' in name: total=n; continue
        name = re.sub(r' \(\d[\d,]*x\)$', '', re.sub(r' \[.*\]$', '', name))
        name = name.split(':',1)[1] if ':' in name and not name.startswith('<') else name
        d[name]+=n
    return total, d
tag, n = sys.argv[1], int(sys.argv[2])
res={}
for leg in ['before','after']:
    tn,dn = funcs(f'cg-{tag}-{leg}-{n}.out'); t0,d0 = funcs(f'cg-{tag}-{leg}-0.out')
    res[leg]=((tn-t0)/n, {k:(dn[k]-d0.get(k,0))/n for k in dn})
    print(leg, 'per audit Ir', round((tn-t0)/n))
b=res['before'][1]; a=res['after'][1]
rows=sorted(set(a)|set(b), key=lambda k: -(abs(a.get(k,0)-b.get(k,0))))
for k in rows[:16]:
    print(f"{b.get(k,0):>12.0f} {a.get(k,0):>12.0f} {a.get(k,0)-b.get(k,0):>+12.0f}  {k[:100]}")
