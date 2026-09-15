import re, sys, collections
# Parse a callgrind output file and sum call counts per callee function name.
def parse(path):
    names = {}          # id -> name  (for cfn=/fn= compressed names)
    calls = collections.Counter()
    pending = None
    cur_cfn = None
    with open(path, 'r', errors='replace') as fh:
        for line in fh:
            line = line.rstrip('\n')
            m = re.match(r'^(c?fn)=\((\d+)\)(?:\s+(.*))?$', line)
            if m:
                kind, ident, nm = m.groups()
                if nm:
                    names[ident] = nm
                nm = names.get(ident, '?')
                if kind == 'cfn':
                    cur_cfn = nm
                continue
            m = re.match(r'^(c?fn)=(.*)$', line)
            if m:
                kind, nm = m.groups()
                if kind == 'cfn':
                    cur_cfn = nm
                continue
            m = re.match(r'^calls=(\d+)', line)
            if m and cur_cfn is not None:
                calls[cur_cfn] += int(m.group(1))
                continue
    return calls

if __name__ == '__main__':
    path = sys.argv[1]
    pats = sys.argv[2:]
    calls = parse(path)
    for p in pats:
        total = sum(v for k, v in calls.items() if p in k)
        matched = {k: v for k, v in calls.items() if p in k}
        print(f'{p}: {total}')
        for k, v in sorted(matched.items(), key=lambda kv: -kv[1])[:4]:
            print(f'    {v:>8}  {k}')
