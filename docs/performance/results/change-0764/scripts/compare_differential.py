#!/usr/bin/env python3
"""Compare two `xml_attribute_bounds differential` reports (before, after).

Prints the counts of compared keys, per check, and every key whose outcome
differs, then exits non-zero when any differs or a key is missing on one side.
Writes a JSON summary when --json is given.
"""

import argparse
import collections
import json
import sys


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("before")
    parser.add_argument("after")
    parser.add_argument("--json")
    args = parser.parse_args()
    before = json.load(open(args.before, encoding="utf-8"))
    after = json.load(open(args.after, encoding="utf-8"))
    left, right = before["results"], after["results"]
    checks = collections.Counter(key.rsplit("::", 1)[1] for key in left)
    outcomes = collections.Counter(
        (key.rsplit("::", 1)[1], value.split(":", 1)[0]) for key, value in left.items()
    )
    missing_after = sorted(set(left) - set(right))
    missing_before = sorted(set(right) - set(left))
    differing = sorted(key for key in set(left) & set(right) if left[key] != right[key])
    summary = {
        "packages": [before["packages"], after["packages"]],
        "members": [before["members"], after["members"]],
        "generated": [before["generated"], after["generated"]],
        "compared_keys": len(set(left) & set(right)),
        "checks": dict(sorted(checks.items())),
        "outcomes_before": {f"{check}:{kind}": count for (check, kind), count in sorted(outcomes.items())},
        "missing_after": missing_after,
        "missing_before": missing_before,
        "differing": [{"key": key, "before": left[key], "after": right[key]} for key in differing],
    }
    print(json.dumps({key: summary[key] for key in (
        "packages", "members", "generated", "compared_keys", "checks", "outcomes_before")}, indent=1))
    print(f"missing after: {len(missing_after)}, missing before: {len(missing_before)}, differing: {len(differing)}")
    for item in summary["differing"][:40]:
        print(json.dumps(item)[:400])
    if args.json:
        with open(args.json, "w", encoding="utf-8") as handle:
            json.dump(summary, handle, indent=1)
    return 1 if (differing or missing_after or missing_before) else 0


if __name__ == "__main__":
    sys.exit(main())
