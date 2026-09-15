#!/usr/bin/env python3
"""Classify every before/after divergence produced by the change-0582 harness.

Usage:
    classify.py <report-before> <report-after> [<divergence-dump>]

Each report is a line-oriented record written by
`strict_scope_differential.rs`.  This script pairs the two on
(input, profile, API, member) and sorts every difference into one of:

  A  approved      refuse -> accept where the pre-change refusal was the
                   archive-wide overlap error, i.e. the row the project owner
                   approved on change 0575's witness.
  B  wider         refuse -> accept where the pre-change refusal was some other
                   per-record validation failure of a record the caller is not
                   reading.  Narrower than an overlap and not covered by the
                   approved table.
  C  widening      accept -> refuse.  A member that used to be readable is not.
  D  payload       accept -> accept with different bytes.
  E  error-drift   refuse -> refuse with a different typed error identity.
  P  panic         a panic on either side.
  O  oracle        an intra-build oracle (order independence, memo stability)
                   that reported false on either side.

Classes C, D, P and O are defects on their face.  Class B is reported
separately from class A because the approved delta was stated as an overlap
narrowing.
"""

import collections
import json
import re
import sys

MEMBER_APIS = (
    "R.read_to",
    "I.read_entry_to",
    "I.verified_reader",
    "I.precompressed",
    "R.read_stored_borrowed",
)
ORACLES = ("R.read_to.order_independent", "R.read_to.memo_stable",
           "I.read_entry_to.order_independent", "I.read_entry_to.memo_stable")
SINGLETONS = ("slice.open", "indexed.open", "I.preservation_index")

OVERLAP = 'strict streaming refuses overlapping ZIP local spans'
DUPLICATE = 'strict streaming refuses duplicate ZIP local spans'
INTO_DIRECTORY = 'strict streaming local span extends into the central directory'


def parse(path):
    """Yield (input_name, {key: verdict}) per input."""
    current = None
    rows = {}
    panics = []
    with open(path, "r", errors="replace") as handle:
        for line in handle:
            line = line.rstrip("\n")
            if line.startswith("=== "):
                if current is not None:
                    yield current, rows, panics
                current = line[4:]
                rows = {}
                panics = []
                continue
            if line.startswith("harness ") or line.startswith("TOTAL "):
                continue
            if line.startswith("PANIC "):
                panics.append(line)
                continue
            if line.startswith("slice.from_slice ") or line.startswith("slice.entries "):
                key, _, value = line.partition(" ")
                rows[("-", key, "")] = value
                continue
            parts = line.split(" ", 1)
            if len(parts) != 2:
                continue
            profile, rest = parts
            if profile not in ("fuzz", "wide"):
                continue
            api, _, tail = rest.partition(" ")
            if api in MEMBER_APIS:
                member, _, verdict = tail.partition(" ")
                rows[(profile, api, member)] = verdict
            else:
                rows[(profile, api, "")] = tail
    if current is not None:
        yield current, rows, panics


def kind_of(verdict):
    if verdict.startswith("ok"):
        return "ok"
    if verdict.startswith("err") or verdict.startswith("no-entry-id"):
        return "err"
    return "other"


def normalise(reason):
    """Collapse numeric payloads so the histogram is about identities."""
    return re.sub(r"\d+", "N", reason)


def error_text(verdict):
    match = re.search(r'msg: "([^"]*)"', verdict)
    if match:
        return match.group(1)
    inner = verdict
    for prefix in ("err-archive(", "err-transport(", "err-callback(", "err("):
        if inner.startswith(prefix):
            inner = inner[len(prefix):-1]
            break
    return inner


def main(before_path, after_path, dump_path=None):
    counts = collections.Counter()
    class_b_reasons = collections.Counter()
    class_a_inputs = set()
    class_b_inputs = set()
    defects = []
    family_counts = collections.Counter()
    reached_inputs = set()
    reached_members = collections.Counter()
    total_members = collections.Counter()
    panic_lines = []
    oracle_failures = []
    per_api = collections.Counter()
    class_e_pairs = collections.Counter()
    crafted_rows = collections.Counter()
    class_a_reasons = collections.Counter()

    before = parse(before_path)
    after = parse(after_path)

    for (before_name, before_rows, before_panics), (after_name, after_rows, after_panics) in zip(before, after):
        assert before_name == after_name, (before_name, after_name)
        name = before_name.split(" ")[0]
        family = name.split("/")[0]
        for line in before_panics:
            panic_lines.append(("before", name, line))
        for line in after_panics:
            panic_lines.append(("after", name, line))

        keys = set(before_rows) | set(after_rows)
        input_reached = False
        for key in sorted(keys):
            profile, api, member = key
            lhs = before_rows.get(key, "<absent>")
            rhs = after_rows.get(key, "<absent>")

            if api in ORACLES:
                if lhs != "true":
                    oracle_failures.append(("before", name, profile, api, lhs))
                if rhs != "true":
                    oracle_failures.append(("after", name, profile, api, rhs))
                continue

            if api in MEMBER_APIS and api != "R.read_stored_borrowed":
                total_members[(family, profile)] += 1
                # "Reached the strict-layout path" means the archive opened and
                # the member got as far as a per-target proof: everything that
                # is not a pre-proof refusal (limits, missing entry, unsupported
                # method) and not an open failure.
                if lhs != "<absent>" and lhs != "no-entry-id":
                    text = error_text(lhs)
                    pre_proof = (
                        "LimitExceeded" in lhs
                        or "FileNotFound" in lhs
                        or "UnsupportedCompressionMethod" in lhs
                        or "InvalidParallelReadLimits" in lhs
                    )
                    if not pre_proof:
                        reached_members[(family, profile)] += 1
                        input_reached = True
                        del text

            if lhs == rhs:
                continue
            if family == "crafted":
                crafted_rows[(name, api)] += 1

            lhs_kind, rhs_kind = kind_of(lhs), kind_of(rhs)
            if lhs_kind == "err" and rhs_kind == "ok":
                reason = error_text(lhs)
                if reason == OVERLAP or reason == 'borrowed access cannot prove non-overlapping ZIP spans':
                    counts["A"] += 1
                    class_a_inputs.add(name)
                    class_a_reasons[normalise(reason)] += 1
                    family_counts[("A", family)] += 1
                    per_api[("A", api)] += 1
                else:
                    counts["B"] += 1
                    class_b_reasons[normalise(reason)] += 1
                    class_b_inputs.add(name)
                    family_counts[("B", family)] += 1
                    per_api[("B", api)] += 1
                    if len(defects) < 4000:
                        defects.append(("B", name, profile, api, member, lhs, rhs))
            elif lhs_kind == "ok" and rhs_kind == "err":
                counts["C"] += 1
                family_counts[("C", family)] += 1
                per_api[("C", api)] += 1
                defects.append(("C", name, profile, api, member, lhs, rhs))
            elif lhs_kind == "ok" and rhs_kind == "ok":
                counts["D"] += 1
                family_counts[("D", family)] += 1
                per_api[("D", api)] += 1
                defects.append(("D", name, profile, api, member, lhs, rhs))
            else:
                counts["E"] += 1
                family_counts[("E", family)] += 1
                per_api[("E", api)] += 1
                class_e_pairs[(normalise(error_text(lhs)), normalise(error_text(rhs)))] += 1
                if len(defects) < 4000:
                    defects.append(("E", name, profile, api, member, lhs, rhs))

        if input_reached:
            reached_inputs.add(name)

    summary = {
        "divergences": dict(counts),
        "class_A_inputs": len(class_a_inputs),
        "class_B_inputs": len(class_b_inputs),
        "class_B_reasons": dict(class_b_reasons.most_common()),
        "per_family": {"%s/%s" % key: value for key, value in sorted(family_counts.items())},
        "members_examined": {"%s/%s" % key: value for key, value in sorted(total_members.items())},
        "members_reaching_proof": {"%s/%s" % key: value for key, value in sorted(reached_members.items())},
        "inputs_reaching_proof": len(reached_inputs),
        "panics": panic_lines,
        "oracle_failures": oracle_failures[:50],
        "oracle_failure_count": len(oracle_failures),
        "per_api": {"%s/%s" % key: value for key, value in sorted(per_api.items())},
        "class_A_reasons": dict(class_a_reasons.most_common()),
        "class_E_pairs": {"%s  ==>  %s" % key: value for key, value in class_e_pairs.most_common(60)},
        "class_E_distinct_pairs": len(class_e_pairs),
        "crafted_divergent": {"%s [%s]" % key: value for key, value in sorted(crafted_rows.items())},
    }
    print(json.dumps(summary, indent=1, sort_keys=True))

    if dump_path:
        with open(dump_path, "w") as handle:
            for row in defects:
                handle.write("\t".join(str(part) for part in row) + "\n")


if __name__ == "__main__":
    main(*sys.argv[1:])
