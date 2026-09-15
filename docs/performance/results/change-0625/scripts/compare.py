#!/usr/bin/env python3
"""Compares the two corpus legs captured by capture.sh and prints corpus-summary.txt.

    python3 compare.py corpus-before.jsonl corpus-after.jsonl
"""
import json
import sys
from collections import Counter


def load(path):
    with open(path) as handle:
        return {record["file"]: record for record in map(json.loads, handle)}


def main(before_path, after_path):
    before, after = load(before_path), load(after_path)
    assert set(before) == set(after), "fixture sets differ"
    ok = [k for k in before if before[k]["status"] == "ok"]
    refused = [k for k in before if before[k]["status"] != "ok"]
    print(f"fixtures considered: {len(before)}   rebuilt: {len(ok)}   refused: {len(refused)}")
    for key in sorted(refused):
        print(f"  refused {key}: before={before[key]['status']} after={after[key]['status']} "
              f"same={before[key] == after[key]}")
    incomparable = [k for k in ok if before[k]["incomparable"]]
    print(f"with >=2 incomparable storage paths (susceptible): {len(incomparable)}")
    print(f"non-deterministic BEFORE (>1 digest in 16 rebuilds): "
          f"{len([k for k in ok if before[k]['distinct'] > 1])}")
    print(f"non-deterministic AFTER: {len([k for k in ok if after[k]['distinct'] > 1])}")
    identical = [k for k in ok if before[k]["distinct"] == 1
                 and before[k]["digests"] == after[k]["digests"]]
    changed = [k for k in ok if before[k]["distinct"] == 1
               and before[k]["digests"] != after[k]["digests"]]
    print(f"already deterministic BEFORE and byte-identical AFTER: {len(identical)}")
    print(f"already deterministic BEFORE but CHANGED AFTER: {len(changed)}")
    for key in sorted(changed):
        print(f"  CHANGED {key}: before={before[key]['digests']} after={after[key]['digests']}")
    lengths = [k for k in ok if before[k]["lengths"] != after[k]["lengths"]]
    print(f"output length differs before/after: {len(lengths)}")
    for key in sorted(lengths):
        print(f"  LENGTH {key}: before={before[key]['lengths']} after={after[key]['lengths']}")
    print()
    print("Susceptible fixtures (>=2 incomparable storage paths):")
    for key in sorted(incomparable):
        digest = after[key]["digests"][0]
        print(f"  {key}: storages={before[key]['storages']} streams={before[key]['streams']} "
              f"before_distinct={before[key]['distinct']} after_distinct={after[key]['distinct']} "
              f"after_digest={digest} after_in_before_set={digest in before[key]['digests']}")
    print()
    print("Storage-count histogram over rebuilt fixtures:")
    histogram = Counter(before[k]["storages"] for k in ok)
    for count in sorted(histogram):
        print(f"  {count} storages: {histogram[count]} fixtures")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
