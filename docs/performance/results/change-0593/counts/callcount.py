import re, sys, collections
DEF = re.compile(r'^(c?fn)=\((\d+)\)(?:\s+(.*))?$')
def counts(path, names):
    table = {}
    out = collections.Counter()
    cur = None
    with open(path, 'rb') as fh:
        for raw in fh:
            line = raw.decode('utf-8', 'replace').rstrip('\n')
            m = DEF.match(line)
            if m:
                kind, ident, name = m.groups()
                if name:
                    table[ident] = name
                if kind == 'cfn':
                    cur = table.get(ident, '')
                continue
            if line.startswith('cfn='):
                cur = line[4:]
                continue
            if line.startswith('calls=') and cur is not None:
                n = int(line[6:].split()[0])
                for name in names:
                    if name in cur:
                        out[name] += n
                cur = None
    return out
names = ['verify_authored', 'try_to_xml_bytes', 'rels_uri', 'ContentType::new', 'memcmp', 'from_parts']
for path in sys.argv[1:]:
    c = counts(path, names)
    print(f"{path.split('/')[-1]:52s}", ' '.join(f"{k}={c[k]}" for k in names))
