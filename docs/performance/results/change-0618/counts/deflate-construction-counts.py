"""Deflate-compressor construction and reset counts from callgrind profiles.

The after leg inlines `<flate2::mem::Compress>::new` into
`ReusableDeflateState::new`, so the construction is counted through whichever
of the two names the profile carries, and `zlib_rs::deflate::reset` counts the
per-member resets that replace the constructions.
"""
import re, sys, collections
DEF = re.compile(r'^(c?fn)=\((\d+)\)(?:\s+(.*))?$')
NAMES = {
    'Compress::new':        '<flate2::mem::Compress>::new',
    'ReusableState::new':   '<soapberry_zip::writer::ReusableDeflateState>::new',
    'zlib deflate::reset':  'zlib_rs::deflate::reset',
    'mini-archive parse':   'from_slice',
}
def counts(path):
    table={}; cur=None; out=collections.Counter()
    with open(path,'rb') as fh:
        for raw in fh:
            line=raw.decode('utf-8','replace').rstrip('\n')
            m=DEF.match(line)
            if m:
                kind,ident,name=m.groups()
                if name: table[ident]=name
                if kind=='cfn': cur=table.get(ident,'')
                continue
            if line.startswith('calls=') and cur is not None:
                n=int(line[6:].split()[0])
                for key,needle in NAMES.items():
                    if needle in cur: out[key]+=n
                cur=None
    return out
print(f"{'profile':30s} " + ' '.join(f"{k:>20s}" for k in NAMES))
for path in sys.argv[1:]:
    c=counts(path)
    print(f"{path.split('/')[-1][:-4]:30s} " + ' '.join(f"{c[k]:>20d}" for k in NAMES))
