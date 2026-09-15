import struct, os, sys

def cfb_workbook(path):
    d=open(path,'rb').read()
    ssz=1<<struct.unpack_from('<H',d,30)[0]; mssz=1<<struct.unpack_from('<H',d,32)[0]
    nfat=struct.unpack_from('<I',d,44)[0]; dirstart=struct.unpack_from('<I',d,48)[0]
    difstart=struct.unpack_from('<I',d,68)[0]; ndif=struct.unpack_from('<I',d,72)[0]
    fatsect=[struct.unpack_from('<I',d,76+4*i)[0] for i in range(109)]
    s=difstart
    while ndif>0 and s<0xFFFFFFFA:
        off=(s+1)*ssz
        for i in range((ssz//4)-1): fatsect.append(struct.unpack_from('<I',d,off+4*i)[0])
        s=struct.unpack_from('<I',d,off+ssz-4)[0]; ndif-=1
    fat=[]
    for fs in fatsect[:nfat]:
        if fs>=0xFFFFFFFA: continue
        off=(fs+1)*ssz
        for i in range(ssz//4): fat.append(struct.unpack_from('<I',d,off+4*i)[0])
    def chain(start):
        out=[];s=start
        while s<0xFFFFFFFA and len(out)<4_000_000: out.append(s); s=fat[s] if s<len(fat) else 0xFFFFFFFE
        return out
    def sread(start): return b''.join(d[(s+1)*ssz:(s+2)*ssz] for s in chain(start))
    dd=sread(dirstart); ents=[]
    for i in range(len(dd)//128):
        e=dd[i*128:(i+1)*128]; nl=struct.unpack_from('<H',e,64)[0]
        ents.append((e[:max(nl-2,0)].decode('utf-16-le','replace'), e[66],
                     struct.unpack_from('<I',e,116)[0], struct.unpack_from('<Q',e,120)[0]))
    for name,typ,start,size in ents:
        if typ==2 and name in ('Workbook','Book'):
            return sread(start)[:size], ssz
    raise SystemExit("no Workbook stream")

def analyze(path):
    wb,ssz = cfb_workbook(path)
    # 1. index every record: (id, data_start, len)
    recs=[]; off=0
    while off+4<=len(wb):
        rid,rlen=struct.unpack_from('<HH',wb,off)
        recs.append((rid, off+4, rlen)); off+=4+rlen
    # 2. find SST + its Continues -> logical byte stream with source offsets
    segs=[]  # (source_offset, logical_offset, len)
    logical=0; i=0
    while i<len(recs) and recs[i][0]!=0x00FC: i+=1
    if i==len(recs): return None
    j=i
    while j<len(recs) and (j==i or recs[j][0]==0x003C):
        _,ds,dl = recs[j]
        segs.append((ds, logical, dl)); logical+=dl; j+=1
    # 3. decode SST entries: header 8 bytes (total, unique), then unique strings
    def rd(lo,n):
        out=bytearray()
        for so,llo,ln in segs:
            if lo < llo+ln and lo+n > llo:
                a=max(lo,llo); b=min(lo+n,llo+ln)
                out+=wb[so+(a-llo):so+(b-llo)]
        return bytes(out)
    hdr=rd(0,8); total,unique=struct.unpack('<II',hdr)
    pos=8; starts=[]
    for k in range(unique):
        starts.append(pos)
        cch=struct.unpack('<H',rd(pos,2))[0]; flags=rd(pos+2,1)[0]; p=pos+3
        rich = 0; ext = 0
        if flags & 0x08: rich=struct.unpack('<H',rd(p,2))[0]; p+=2
        if flags & 0x04: ext=struct.unpack('<I',rd(p,4))[0]; p+=4
        wide = flags & 0x01
        p += cch*(2 if wide else 1)
        p += rich*4 + ext
        pos=p
        if pos>logical: break
    # map a logical offset -> (source_offset, sector ordinal)
    def ordinal_of(lo):
        for so,llo,ln in segs:
            if llo<=lo<llo+ln:
                return (so+(lo-llo))//ssz
        return None
    # 4. walk LabelSst in stream order, collect the SST index sequence
    seq=[]
    for rid,ds,dl in recs:
        if rid==0x00FD and dl>=10:
            idx=struct.unpack_from('<I',wb,ds+6)[0]
            if idx<len(starts): seq.append(idx)
    ords=[ordinal_of(starts[s]) for s in seq]
    ords=[o for o in ords if o is not None]
    fwd=back=same=0
    cold_now=0; cold_hint=0; prev=None
    for o in ords:
        cold_now += o            # today: full walk from sector 0 every time
        if prev is None: cold_hint += o
        elif o>=prev: cold_hint += (o-prev); fwd+=1 if o>prev else 0; same+=1 if o==prev else 0
        else: cold_hint += o; back+=1
        prev=o
    return dict(file=os.path.basename(path), sector=ssz, workbook=len(wb),
                unique=unique, total=total, resolves=len(ords),
                ord_min=min(ords) if ords else 0, ord_max=max(ords) if ords else 0,
                fwd=fwd, same=same, back=back,
                back_pct=100.0*back/max(len(ords)-1,1),
                links_today=cold_now, links_hinted=cold_hint,
                saving_pct=100.0*(cold_now-cold_hint)/max(cold_now,1))

for p in sys.argv[1:]:
    try:
        r=analyze(p)
        if r: print(r)
    except Exception as e:
        print(f"{os.path.basename(p)}: SKIP {type(e).__name__}: {e}")
