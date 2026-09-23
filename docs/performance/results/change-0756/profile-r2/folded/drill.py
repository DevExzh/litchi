import sys, re
from collections import Counter
exec(open(__import__('os').path.join(__import__('os').path.dirname(__import__('os').path.abspath(__file__)), 'an.py')).read().split("if __name__")[0])
case=sys.argv[1]
tot,runner,rows=load(case)
T=sum(p for p,n,f in rows)
def sub(rx, depth, top=16, ctxleaf=3, ltop=14):
    sel=[]
    for p,n,f in rows:
        ii=[i for i,x in enumerate(f) if re.search(rx,nm(x))]
        if ii: sel.append((p,n,f[ii[0]:]))
    S=sum(p for p,n,f in sel)
    print(f'### {rx}: {100*S/T:.2f}% of timed')
    c=Counter()
    for p,n,fr in sel:
        ch=[]
        for f in fr[1:]:
            if GLUE.search(nm(f)): continue
            ch.append(short(f)[:44]+('*' if '|R:' in f else ''))
            if len(ch)>=depth: break
        c[' > '.join(ch)]+=p
    for k,v in c.most_common(top): print(f'{100*v/T:6.2f}%  {k}')
    print('  leaf:')
    c=Counter()
    for p,n,fr in sel:
        k=0
        while k<len(fr) and fr[-1-k]=='K': k+=1
        user=[short(x)[:38] for x in fr[:len(fr)-k]]
        c[('[K] ' if k else '')+' < '.join(reversed(user[-ctxleaf:]))]+=p
    for k,v in c.most_common(ltop): print(f'   {100*v/T:6.2f}%  {k}')
for a in sys.argv[2:]:
    rx,d=a.rsplit('@',1)
    sub(rx,int(d))
