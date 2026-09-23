#!/usr/bin/env python3
"""Change 0750: summarize the census rows (census.jsonl[.gz]) and corpus list
(corpus.txt[.gz]) written by census.py into census/summary.txt."""
import collections, gzip, json, sys

POLICIES = [("source", "verify_source"), ("authored", "verify_authored"),
            ("compact", "verify"), ("reader", "verify_reader")]


def lines(path):
    opener = gzip.open if path.endswith(".gz") else open
    with opener(path, "rt") as handle:
        return handle.read().splitlines()


def main(census, corpus):
    rows = [json.loads(line) for line in lines(census)]
    sections = collections.defaultdict(list)
    section = None
    for line in lines(corpus):
        if line.startswith("# "):
            section = line[2:]
        else:
            sections[section].append(line)
    group = lambda path: "test-data" if path.startswith("test-data/") else "evidence packets (docs/performance/results)"
    print("Change 0750 census: every XML member publication would audit (is_xml_part by name or content type),")
    print("in every OPC package (a ZIP with [Content_Types].xml) under test-data/, docs/performance/ and crates/*/tests/,")
    print("plus every loose .xml/.rels under test-data/ooxml/, audited by the base (3174242282) and final auditors.")
    print()
    packages = collections.Counter(group(path) for path in sections["packages"])
    crates = sum(1 for path in sections["packages"] if path.startswith("crates/"))
    print(f"OPC packages found by content: {dict(packages)} ({crates} under crates/*/tests/)")
    print(f"loose OOXML XML parts under test-data/ooxml: {len(sections['loose XML parts'])}")
    unreadable = collections.Counter(group(line.split(chr(9))[0]) for line in sections["unreadable"])
    print(f"members Python zipfile could not read (bad CRC, mutated headers; deliberate corruption): {dict(unreadable)}")
    for name in ("test-data", "evidence packets (docs/performance/results)"):
        subset = [row for row in rows if group(row["file"]) == name]
        files = len({row["file"] for row in subset})
        print()
        print(f"== {name.split(' ')[0]}: {files} files with XML members, {len(subset)} XML members, "
              f"{sum(row['bytes'] for row in subset):,} bytes")
        for key, label in POLICIES:
            moves = collections.Counter()
            changed = 0
            for row in subset:
                before, after = row["before_" + key], row["after_" + key]
                moves[("accepted" if before == "OK" else "refused") + "->" +
                      ("accepted" if after == "OK" else "refused")] += 1
                changed += before != after
            print(f"  {label:<16} " + "  ".join(f"{move} {moves[move]:>5}" for move in
                  ("accepted->accepted", "accepted->refused", "refused->refused", "refused->accepted"))
                  + f"  verdict text changed {changed}")
    print()
    print("Newly refused by verify_source:")
    for row in rows:
        if row["before_source"] == "OK" and row["after_source"] != "OK":
            where = row["member"] or "(loose part)"
            print(f"   {row['file']} {where} | {row['after_source']}")
    print("Refused by verify_source on both legs (test-data):")
    for row in rows:
        if row["file"].startswith("test-data/") and row["before_source"] != "OK" and row["after_source"] != "OK":
            where = row["member"] or "(loose part)"
            same = "identical" if row["before_source"] == row["after_source"] else "CHANGED: was " + row["before_source"]
            print(f"   {row['file']} {where} | {row['after_source']} | {same}")
    depths = [row["namespaces_in_scope"] for row in rows
              if row["file"].startswith("test-data/") and row.get("namespaces_in_scope") is not None]
    print()
    print(f"Prefix declarations in scope at once (expat, accepted test-data members): max {max(depths)}, "
          f"members measured {len(depths)}, over 20: {sum(1 for depth in depths if depth > 20)}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
