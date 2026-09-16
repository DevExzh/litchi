"""Change 0654: audit every XML member of a fixture corpus under both legs."""
import json, subprocess, struct, sys, zipfile

BEFORE = "/home/zhuhe/code/litchi-worktrees/targets/0654-probe-before/release/litchi-probe-0654"
AFTER = "/home/zhuhe/code/litchi-worktrees/targets/0654-probe-after/release/litchi-probe-0654"


def is_xml_name(name):
    leaf = name.rsplit("/", 1)[-1]
    if leaf.lower() == "[content_types].xml":
        return True
    if "." not in leaf:
        return False
    return leaf.rsplit(".", 1)[1].lower() in ("xml", "rels", "rdf")


def collect(list_path):
    members = []
    payload = bytearray()
    for path in open(list_path).read().split():
        try:
            z = zipfile.ZipFile(path)
        except Exception as exc:
            print("ZIPFAIL", path, exc, file=sys.stderr)
            continue
        for name in z.namelist():
            if name.endswith("/") or not is_xml_name(name):
                continue
            data = z.read(name)
            members.append((path, name, len(data)))
            payload += struct.pack("<I", len(data)) + data
    return members, bytes(payload)


def run(binary, payload):
    out = subprocess.run([binary, "audit"], input=payload, stdout=subprocess.PIPE, check=True)
    return out.stdout.decode().splitlines()


def main(list_path, out_path):
    members, payload = collect(list_path)
    before = run(BEFORE, payload)
    after = run(AFTER, payload)
    assert len(before) == len(after) == len(members), (len(before), len(after), len(members))
    rows = []
    for (path, name, size), b, a in zip(members, before, after):
        b_auth, b_src = b.split("\t")
        a_auth, a_src = a.split("\t")
        rows.append({
            "file": path, "member": name, "bytes": size,
            "before_authored": b_auth, "before_source": b_src,
            "after_authored": a_auth, "after_source": a_src,
        })
    with open(out_path, "w") as handle:
        for row in rows:
            handle.write(json.dumps(row) + "\n")
    print(f"{len(rows)} XML members audited -> {out_path}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
