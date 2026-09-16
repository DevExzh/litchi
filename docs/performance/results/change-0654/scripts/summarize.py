import json, sys, collections, re
rows = [json.loads(l) for l in open(sys.argv[1])]
print("members:", len(rows), " packages:", len({r['file'] for r in rows}))
# invariant 1: the authored verdict is identical on both legs
moved = [r for r in rows if r['before_authored'] != r['after_authored']]
print("authored-verdict changes across legs:", len(moved))
# invariant 2: the before leg's "source" column equals its authored column
same = [r for r in rows if r['before_authored'] != r['before_source']]
print("before-leg source column differs from authored:", len(same))

def cls(v):
    if v == 'ok': return 'ok'
    m = re.match(r'ERR noncompact XML (\w+) at byte', v)
    if m: return 'NotCompact:' + m.group(1)
    m = re.match(r'ERR malformed XML at byte \d+: (.*)$', v)
    if m: return 'Malformed:' + m.group(1)[:60]
    m = re.match(r'ERR XML (\w+) limit', v)
    if m: return 'Limit:' + m.group(1)
    if v.startswith('ERR DTD and DOCTYPE'): return 'Doctype'
    if v.startswith('ERR XML is not UTF-8'): return 'Encoding'
    return 'other:' + v[:60]

before = collections.Counter(cls(r['before_source']) for r in rows)
after = collections.Counter(cls(r['after_source']) for r in rows)
print("\n--- ORIGINAL-bytes audit verdicts, per XML member ---")
keys = sorted(set(before) | set(after))
print(f"{'class':55s} {'before':>7s} {'after':>7s}")
for k in keys:
    print(f"{k:55s} {before[k]:7d} {after[k]:7d}")

print("\n--- transitions (before -> after) ---")
trans = collections.Counter((cls(r['before_source']), cls(r['after_source'])) for r in rows)
for (b, a), n in sorted(trans.items(), key=lambda kv: -kv[1]):
    print(f"{n:6d}  {b}  ->  {a}")

# per-package: does every XML member now pass the source audit?
pkg_before = collections.defaultdict(lambda: True)
pkg_after = collections.defaultdict(lambda: True)
for r in rows:
    if r['before_source'] != 'ok': pkg_before[r['file']] = False
    if r['after_source'] != 'ok': pkg_after[r['file']] = False
files = sorted({r['file'] for r in rows})
print("\npackages with every XML member accepted by the original-bytes audit:")
print("  before:", sum(1 for f in files if pkg_before[f]), "/", len(files))
print("  after: ", sum(1 for f in files if pkg_after[f]), "/", len(files))
print("\npackages still carrying a refused XML member after the change:")
for f in files:
    if not pkg_after[f]:
        bad = [(r['member'], r['after_source']) for r in rows if r['file']==f and r['after_source']!='ok']
        print("  ", f)
        for m, v in bad[:6]:
            print("        ", m, "::", v)
