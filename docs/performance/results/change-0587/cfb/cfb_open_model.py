import struct, sys, os, glob
ENDOFCHAIN=0xFFFFFFFE; MAXREG=0xFFFFFFFA
def parse(path):
    b=open(path,'rb').read()
    if b[:8]!=b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1': return None
    ss=1<<struct.unpack_from('<H',b,0x1E)[0]; ndir=struct.unpack_from('<I',b,0x28)[0]
    nfat=struct.unpack_from('<I',b,0x2C)[0]; fdir=struct.unpack_from('<I',b,0x30)[0]
    fmini=struct.unpack_from('<I',b,0x3C)[0]; nmini=struct.unpack_from('<I',b,0x40)[0]
    fdifat=struct.unpack_from('<I',b,0x44)[0]; ndifat=struct.unpack_from('<I',b,0x48)[0]
    fats=[struct.unpack_from('<I',b,0x4C+4*i)[0] for i in range(min(109,nfat))]
    d=fdifat
    for _ in range(ndifat):
        sec=b[(d+1)*ss:(d+2)*ss]; ents=struct.unpack_from('<%dI'%(ss//4),sec)
        fats+= [e for e in ents[:-1] if e<MAXREG]; d=ents[-1]
    fat=[]
    for s in fats[:nfat]:
        sec=b[(s+1)*ss:(s+2)*ss]; fat+=struct.unpack_from('<%dI'%(len(sec)//4),sec)
    def chain(start):
        out=[]; s=start; seen=set()
        while s<MAXREG and s not in seen and len(out)<len(fat):
            seen.add(s); out.append(s); s=fat[s] if s<len(fat) else ENDOFCHAIN
        return out
    dirsec=chain(fdir); dirdata=b''.join(b[(s+1)*ss:(s+2)*ss] for s in dirsec)
    ents=[]
    for i in range(len(dirdata)//128):
        e=dirdata[i*128:(i+1)*128]; t=e[66]
        if t==0: continue
        nl=struct.unpack_from('<H',e,64)[0]; name=e[:max(nl-2,0)].decode('utf-16le','replace')
        size=struct.unpack_from('<Q',e,120)[0] & (0xFFFFFFFF if ss==512 else 0xFFFFFFFFFFFFFFFF)
        ents.append((name,t,size))
    root=[e for e in ents if e[1]==5]; mini=root[0][2] if root else 0
    streams={n:sz for n,t,sz in ents if t==2}
    return dict(size=len(b),ss=ss,phys=len(b)//ss-1,nfat=nfat,ndifat=ndifat,nmini=nmini,ndirsec=len(dirsec),
                entries=len(ents),mini=mini,streams=streams,minifat_streams=sum(1 for v in streams.values() if v<4096))
def slurp(kind,st):
    if kind=='xls': keys=['Workbook','Book']
    elif kind=='doc': keys=['WordDocument','1Table','0Table','Data']
    else: keys=['PowerPoint Document','Current User']
    return sum(st.get(k,0) for k in keys), {k:st[k] for k in keys if k in st}
named=[('xls','test-data/ole/xls/ConditionalFormattingSamples.xls'),('xls','test-data/poi/test-data/spreadsheet/54016.xls')]
for line in open(sys.argv[1]):
    p=line.strip()
    if p: named.append((p.rsplit('.',1)[1],p))
print("fixture | bytes | sect | phys | FAT | DIFAT | MiniFAT | dirsec | entries | ministream | slurped@open | slurped/size | streams")
for kind,p in named:
    if not os.path.exists(p): print(p,"MISSING"); continue
    r=parse(p); s,which=slurp(kind,r['streams'])
    print(f"{os.path.basename(p)} | {r['size']} | {r['ss']} | {r['phys']} | {r['nfat']} | {r['ndifat']} | {r['nmini']} | {r['ndirsec']} | {r['entries']} | {r['mini']} | {s} | {s/r['size']:.2%} | {which}")
# corpus summary
tot=dict(n=0,difat=0,v4=0,minifat=0,maxsize=0,maxentries=0,maxfat=0)
for kind in ('xls','doc','ppt'):
    files=glob.glob(f'test-data/**/*.{kind}',recursive=True)
    n=difat=v4=maxsize=maxent=maxfat=0; ratios=[]
    for p in files:
        try: r=parse(p)
        except Exception: continue
        if not r: continue
        n+=1; difat+= r['ndifat']>0; v4+= r['ss']==4096; maxsize=max(maxsize,r['size']); maxent=max(maxent,r['entries']); maxfat=max(maxfat,r['nfat'])
        s,_=slurp(kind,r['streams']); ratios.append(s/r['size'])
    ratios.sort()
    med=ratios[len(ratios)//2] if ratios else 0
    print(f"corpus {kind}: files={n} with_DIFAT={difat} v4_sectors={v4} max_bytes={maxsize} max_entries={maxent} max_FAT_sectors={maxfat} median_slurped/size={med:.2%} min={ratios[0] if ratios else 0:.2%} max={ratios[-1] if ratios else 0:.2%}")
