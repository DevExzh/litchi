#!/usr/bin/env python3
"""Scan an OOXML corpus for .rels parts that hold two or more relationships
sharing the same (Type, Target, TargetMode) triple.

Such a duplicate is the input shape that makes
`Relationships::get_or_add` / `get_or_add_ext_rel` return a hash-order-dependent
rId, because both scan `self.rels.values()` and take the first match.
"""
import sys, os, zipfile, re, collections

REL_RE = re.compile(rb'<Relationship\b[^>]*/?>', re.I)
ATTR_RE = re.compile(rb'(\w+)\s*=\s*"([^"]*)"')

EXTS = ('.docx','.docm','.dotx','.dotm','.xlsx','.xlsm','.xltx','.xltm','.xlsb',
        '.pptx','.pptm','.potx','.potm','.ppsx','.ppsm','.thmx')

def scan(root):
    files = 0
    parsed = 0
    rels_parts = 0
    rels_total = 0
    dup_parts = 0
    dup_files = set()
    internal_dups = 0
    external_dups = 0
    examples = []
    for dirpath, _dirs, names in os.walk(root):
        for name in sorted(names):
            if not name.lower().endswith(EXTS):
                continue
            path = os.path.join(dirpath, name)
            files += 1
            try:
                z = zipfile.ZipFile(path)
                members = z.namelist()
            except Exception:
                continue
            parsed += 1
            for m in members:
                low = m.lower()
                if not low.endswith('.rels'):
                    continue
                d, f = low.rsplit('/', 1) if '/' in low else ('', low)
                if not (d == '_rels' or d.endswith('/_rels')):
                    continue
                try:
                    data = z.read(m)
                except Exception:
                    continue
                rels_parts += 1
                seen = collections.Counter()
                ids = collections.defaultdict(list)
                for tag in REL_RE.findall(data):
                    a = {k.lower(): v for k, v in ATTR_RE.findall(tag)}
                    if b'id' not in a:
                        continue
                    rels_total += 1
                    key = (a.get(b'type', b''), a.get(b'target', b''),
                           a.get(b'targetmode', b'Internal'))
                    seen[key] += 1
                    ids[key].append(a[b'id'])
                dups = {k: v for k, v in seen.items() if v > 1}
                if dups:
                    dup_parts += 1
                    dup_files.add(path)
                    for k, v in dups.items():
                        if k[2].lower() == b'external':
                            external_dups += 1
                        else:
                            internal_dups += 1
                        if len(examples) < 12:
                            examples.append((os.path.relpath(path, root), m, v,
                                             k[2].decode('ascii','replace'),
                                             k[0].decode('ascii','replace')[-48:],
                                             k[1].decode('ascii','replace')[:60],
                                             [i.decode('ascii','replace') for i in ids[k][:6]]))
            z.close()
    return dict(files=files, parsed=parsed, rels_parts=rels_parts,
                rels_total=rels_total, dup_parts=dup_parts,
                dup_files=len(dup_files), internal_dups=internal_dups,
                external_dups=external_dups, examples=examples,
                dup_file_list=sorted(dup_files))

if __name__ == '__main__':
    root = sys.argv[1]
    r = scan(root)
    print(f"root                     : {root}")
    print(f"OOXML files found        : {r['files']}")
    print(f"OOXML files opened as zip: {r['parsed']}")
    print(f".rels parts scanned      : {r['rels_parts']}")
    print(f"Relationship elements    : {r['rels_total']}")
    print(f".rels parts with a duplicate (Type,Target,TargetMode): {r['dup_parts']}")
    print(f"packages with a duplicate                            : {r['dup_files']}")
    print(f"duplicate groups, Internal / External                : {r['internal_dups']} / {r['external_dups']}")
    print()
    for e in r['examples']:
        print(f"  {e[0]}")
        print(f"    member={e[1]} count={e[2]} mode={e[3]}")
        print(f"    type=...{e[4]}")
        print(f"    target={e[5]}")
        print(f"    ids={e[6]}")
    print()
    print("packages with at least one duplicate group:")
    for p in r['dup_file_list']:
        print(f"  {os.path.relpath(p, root)}")
