import sys, collections, zipfile, hashlib, os
S = os.path.dirname(os.path.abspath(__file__))
W = "/home/zhuhe/code/litchi-worktrees/scratch-0654"

def load(path):
    rows = {}
    for i, line in enumerate(open(path)):
        f, part, status, err, digest = line.rstrip("\n").split("\t")
        rows[f] = (i, part, status, err, digest)
    return rows

before = load(f"{S}/publish-before.tsv")
after = load(f"{S}/publish-after.tsv")
assert set(before) == set(after)

trans = collections.Counter()
for f in before:
    trans[(before[f][2], after[f][2])] += 1
print("--- publication outcome transitions (before -> after) ---")
for (b, a), n in sorted(trans.items(), key=lambda kv: -kv[1]):
    print(f"{n:5d}  {b} -> {a}")

print("\n--- packages refused on BOTH legs: identical error text? ---")
same = diff = 0
kept = collections.Counter()
for f in before:
    if before[f][2] == "error" and after[f][2] == "error":
        if before[f][3] == after[f][3]:
            same += 1
        else:
            diff += 1
            print("   DIFFERS:", f)
            print("     before:", before[f][3])
            print("     after :", after[f][3])
        kept[after[f][3][:110]] += 1
print(f"   identical: {same}   differing: {diff}")
for text, n in kept.most_common():
    print(f"   {n:3d}  {text}")

print("\n--- packages that published on BOTH legs: identical output digest? ---")
ok_same = ok_diff = 0
for f in before:
    if before[f][2] == "ok" and after[f][2] == "ok":
        if before[f][4] == after[f][4]:
            ok_same += 1
        else:
            ok_diff += 1
            print("   DIGEST MOVED:", f)
print(f"   identical: {ok_same}   differing: {ok_diff}")

print("\n--- untouched-member byte identity, every package that published on the after leg ---")
checked = members_checked = mismatch = 0
missing_member = 0
for f, (i, part, status, err, digest) in sorted(after.items()):
    if status != "ok":
        continue
    src = zipfile.ZipFile(f)
    out = zipfile.ZipFile(f"{W}/pub-after/{i:04d}.zip")
    replaced = part.lstrip("/")
    src_names = [n for n in src.namelist() if not n.endswith("/")]
    out_names = [n for n in out.namelist() if not n.endswith("/")]
    if sorted(src_names) != sorted(out_names):
        print("   MEMBER SET MOVED:", f, set(src_names) ^ set(out_names))
        missing_member += 1
    for name in src_names:
        if name == replaced:
            continue
        if name not in out_names:
            continue
        members_checked += 1
        if src.read(name) != out.read(name):
            mismatch += 1
            print("   MEMBER BYTES MOVED:", f, name)
    # the replaced member must be exactly the payload the probe wrote
    if replaced in out_names and out.read(replaced) != b"<litchi0654/>":
        print("   REPLACEMENT NOT WRITTEN:", f, replaced)
    checked += 1
print(f"   packages checked: {checked}   untouched members compared: {members_checked}")
print(f"   member-set changes: {missing_member}   byte mismatches: {mismatch}")
