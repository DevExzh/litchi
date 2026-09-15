import re,sys,collections
DEF=re.compile(r'^(c?fn)=\((\d+)\)(?:\s+(.*))?$')
def callers(path, target):
    table={}; cur_fn=None; cur_cfn=None; out=collections.Counter(); calls=collections.Counter()
    pending=None
    with open(path,'rb') as fh:
        for raw in fh:
            line=raw.decode('utf-8','replace').rstrip('\n')
            m=DEF.match(line)
            if m:
                kind,ident,name=m.groups()
                if name: table[ident]=name
                if kind=='fn': cur_fn=table.get(ident,''); cur_cfn=None
                else: cur_cfn=table.get(ident,'')
                continue
            if line.startswith('calls='):
                pending=int(line[6:].split()[0]); continue
            if pending is not None:
                # cost line following calls=
                parts=line.split()
                cost=int(parts[-1]) if parts and parts[-1].lstrip('+-').isdigit() else 0
                if cur_cfn and target in cur_cfn:
                    out[cur_fn]+=cost; calls[cur_fn]+=pending
                pending=None
    return out,calls
target=sys.argv[1]
res={}
for path in sys.argv[2:]:
    out,calls=callers(path,target)
    res[path]=(out,calls)
paths=sys.argv[2:]
keys=set()
for p in paths: keys|=set(res[p][0])
rows=[]
for k in keys:
    vals=[res[p][0].get(k,0) for p in paths]
    cs=[res[p][1].get(k,0) for p in paths]
    rows.append((max(vals),k,vals,cs))
rows.sort(reverse=True)
for _,k,vals,cs in rows[:14]:
    print(' '.join(f"{v:>12,}({c})" for v,c in zip(vals,cs)), k[:110])
