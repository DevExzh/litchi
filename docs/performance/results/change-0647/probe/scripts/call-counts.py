#!/usr/bin/env python3
"""Change 0647: per-scenario call counts from a callgrind isolation pair.

Callgrind's raw format compresses function names into ids shared by the `fn=`
and `cfn=` lines, and records every call site as a `calls=<count> <position>`
line that follows its `cfn=`. Summing those counts over one file gives the
whole-process call count of that callee; differencing the N = 1 and N = 11
files and dividing by 10 gives the per-scenario count, with process start-up
and the fixture read cancelled out.

Usage: call-counts.py <callgrind-dir> <scenario> <leg> [substring...]
"""

import pathlib
import re
import sys

# Exact symbol names. Substring matching would be wrong here: the monomorphized
# sort routines `try_to_xml_bytes` calls carry its name inside their own, so a
# substring test counts a call to the sort as a call to the serializer.
DEFAULT_TARGETS = (
    "<litchi_opc::rel::Relationships>::try_to_xml_bytes",
    "xml_minifier::audit::verify_authored",
    "<litchi_opc::packuri::PackURI>::rels_uri",
    "<litchi_opc::rel::Relationships>::add_relationship",
    "<litchi_opc::rel::Relationships>::get_or_add",
    "<litchi_opc::rel::Relationships>::reuse_candidate",
)

NAME = re.compile(r"^(?:c?fn)=\((\d+)\)(?:\s+(.*))?$")
CALLS = re.compile(r"^calls=(\d+)")


def counts(path, targets):
    names = {}
    totals = {target: 0 for target in targets}
    pending = None
    for line in path.read_text(errors="replace").splitlines():
        match = NAME.match(line)
        if match:
            ident, name = match.group(1), match.group(2)
            if name:
                names[ident] = name
            pending = names.get(ident) if line.startswith("cfn=") else None
            continue
        call = CALLS.match(line)
        if call and pending:
            if pending in totals:
                totals[pending] += int(call.group(1))
            pending = None
    return totals


def main():
    root = pathlib.Path(sys.argv[1])
    scenario, leg = sys.argv[2], sys.argv[3]
    targets = tuple(sys.argv[4:]) or DEFAULT_TARGETS
    one = counts(root / f"callgrind.out.{scenario}-{leg}-1", targets)
    eleven = counts(root / f"callgrind.out.{scenario}-{leg}-11", targets)
    print(f"### {scenario} {leg}")
    for target in targets:
        per = (eleven[target] - one[target]) / 10.0
        short = target.rsplit("::", 1)[-1]
        print(f"  {short:<22} n=1:{one[target]:>7} n=11:{eleven[target]:>7} per-scenario:{per:>9.1f}")


if __name__ == "__main__":
    main()
