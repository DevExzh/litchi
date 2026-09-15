#!/usr/bin/env python3
"""Extract audit call counts and inclusive costs from a callgrind output.

Callgrind records one `calls=` line per call site, so summing them per callee
gives the exact number of times a function ran, attributed to its caller.
Two runs of the same case that differ only in the sample count isolate one
operation: difference the totals and divide by the sample difference.

Usage: audit-call-counts.py <callgrind-out> [<symbol> ...]
"""

import collections
import re
import sys

DEFAULT_SYMBOLS = (
    "xml_minifier::audit::verify_authored",
    "xml_minifier::audit::package::is_xml_part",
)


def parse(path):
    names = {}
    calls = collections.defaultdict(collections.Counter)
    current = None
    pending = None
    for line in open(path, errors="replace"):
        line = line.rstrip("\n")
        match = re.match(r"^(c?fn)=\((\d+)\)(?: (.*))?$", line)
        if match:
            kind, number, name = match.group(1), match.group(2), match.group(3)
            if name:
                names[number] = name
            name = names.get(number, number)
            if kind == "fn":
                current, pending = name, None
            else:
                pending = name
            continue
        match = re.match(r"^calls=(\d+)", line)
        if match and pending is not None:
            calls[pending][current] += int(match.group(1))
    return calls


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 1
    calls = parse(sys.argv[1])
    symbols = sys.argv[2:] or list(DEFAULT_SYMBOLS)
    for symbol in symbols:
        total = sum(calls[symbol].values())
        print(f"{symbol}: {total}")
        for caller, count in calls[symbol].most_common():
            print(f"    from {caller}: {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
