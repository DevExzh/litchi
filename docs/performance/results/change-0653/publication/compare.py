#!/usr/bin/env python3
"""Recompute change 0653's corpus publication evidence from `saveprobe` rows.

Usage:

    python3 compare.py --before rows-before.tsv.gz --after rows-after.tsv.gz \
        [--control rows-before-control.tsv.gz] \
        [--probe-before PATH --probe-after PATH --corpus-root DIR --scratch DIR] \
        > summary.txt

The row files are `saveprobe` output (see `saveprobe/src/main.rs`). Everything
this script prints is derived from those rows, except the canonical-XML section,
which needs the differing members' bytes: supply `--probe-before`,
`--probe-after`, `--corpus-root` and `--scratch` and the script re-runs the two
probe binaries with `SAVEPROBE_KEEP=1` over exactly the fixtures whose published
member hashes differ, extracts those members, canonicalizes them and removes the
scratch output again. Without those arguments the section reports that it was
not run.

Row grammar, first column:

    C  fixture  format  source-sha256  source-bytes  source-member-count  note
    X  fixture  route   admitted|refused  detail
    O  fixture  route   OK   published-bytes  published-member-count
    O  fixture  route   ERR  debug  display
    M  fixture  route   member  pub-sha-uncompressed  pub-size-uncompressed
                               src-sha-uncompressed-or-ABSENT  identical(0/1)
                               pub-sha-compressed  pub-size-compressed

`identical` compares the **uncompressed** member payload of the published
archive with the same-named member's uncompressed payload in the source
archive.
"""

from __future__ import annotations

import argparse
import fnmatch
import gzip
import hashlib
import os
import pathlib
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter, defaultdict

# The part each route's documented edit targets. `*_noop_save` makes no edit at
# all, so no member is "the edited part" for it. These sets come from each
# route's own documentation, not from the measured rows.
ROUTE_TARGETS: dict[str, list[str]] = {
    "docx_noop_save": [],
    "xlsx_noop_save": [],
    "pptx_noop_save": [],
    "docx_edit_save": ["word/document.xml"],
    "docx_source_backed_document": ["word/document.xml"],
    "docx_source_backed_section_layout": ["word/document.xml"],
    "xlsx_edit_save": ["xl/worksheets/*.xml"],
    "xlsx_source_backed_cell_values": ["xl/worksheets/*.xml"],
    "xlsx_source_backed_tab_state": ["xl/workbook.xml", "xl/worksheets/*.xml"],
    "pptx_edit_save": ["ppt/slides/slide*.xml"],
    "pptx_source_backed_slide": ["ppt/slides/slide*.xml"],
    "pptx_guides_publish": ["ppt/presentation.xml"],
}

# Members whose name carries an index; collapsed for the grouping report.
INDEXED = re.compile(r"\d+")


def unescape(field: str) -> str:
    out: list[str] = []
    index = 0
    while index < len(field):
        char = field[index]
        if char == "\\" and index + 1 < len(field):
            nxt = field[index + 1]
            out.append({"t": "\t", "n": "\n", "r": "\r", "\\": "\\"}.get(nxt, nxt))
            index += 2
        else:
            out.append(char)
            index += 1
    return "".join(out)


def load(path: str):
    opener = gzip.open if path.endswith(".gz") else open
    classes: dict[str, tuple] = {}
    edits: dict[tuple[str, str], tuple] = {}
    outcomes: dict[tuple[str, str], tuple] = {}
    members: dict[tuple[str, str, str], tuple] = {}
    with opener(path, "rt", encoding="utf-8", newline="") as handle:
        for line in handle:
            row = line.rstrip("\n").split("\t")
            kind = row[0]
            if kind == "C":
                classes[unescape(row[1])] = tuple(row[2:])
            elif kind == "X":
                edits[(unescape(row[1]), row[2])] = (row[3], unescape(row[4]) if len(row) > 4 else "")
            elif kind == "O":
                outcomes[(unescape(row[1]), row[2])] = tuple(row[3:])
            elif kind == "M":
                fixture, route, member = unescape(row[1]), row[2], unescape(row[3])
                members[(fixture, route, member)] = (
                    row[4],          # published sha256, uncompressed payload
                    int(row[5]),     # published uncompressed size
                    row[6],          # source sha256 or ABSENT
                    int(row[7]),     # identical-to-source
                    row[8],          # published sha256, stored/compressed payload
                    int(row[9]),     # published compressed size
                )
    return classes, edits, outcomes, members


def is_target(route: str, member: str) -> bool:
    return any(fnmatch.fnmatch(member, pattern) for pattern in ROUTE_TARGETS.get(route, []))


def group(member: str) -> str:
    return INDEXED.sub("N", member)


# A refusal message that names one member of an unordered set is stable in
# identity but not in text. Normalising it separates "the outcome changed" from
# "the message named a different element of the same set".
NAMED_RELATIONSHIP = re.compile(r"refuse package relationship '[^']*'")


def normalize_outcome(value: tuple) -> tuple:
    return tuple(NAMED_RELATIONSHIP.sub("refuse package relationship '<one-of-set>'", f) for f in value)


def canonical(xml_bytes: bytes):
    """Resolved (namespace, local) for every element and attribute, normalized
    attribute values, text nodes, and NO namespace declarations."""
    xmlns = "{http://www.w3.org/2000/xmlns/}"
    root = ET.fromstring(xml_bytes)
    nodes: list[tuple] = []

    def walk(element, depth):
        nodes.append((depth, "element", element.tag))
        for key in sorted(element.attrib):
            if key.startswith(xmlns) or key == "xmlns":
                continue
            nodes.append((depth, "attribute", key, " ".join(element.attrib[key].split())))
        text = " ".join((element.text or "").split())
        if text:
            nodes.append((depth, "text", text))
        for child in element:
            walk(child, depth + 1)
        tail = " ".join((element.tail or "").split())
        if tail:
            nodes.append((depth, "tail", tail))

    walk(root, 0)
    return nodes



def write_refusals(path: str, before: dict, after: dict, control: dict | None) -> None:
    """The outcome rows of every leg, side by side.

    `status` is OK or ERR. For an OK row `field1` is the published archive size
    in bytes and `field2` the published member count; for an ERR row `field1` is
    the error's `Debug` identity and `field2` its `Display` text, both with tabs
    and newlines escaped exactly as `saveprobe` wrote them.

    The third leg, when present, is a second run of the BEFORE binary over the
    same corpus: it is the determinism control that separates a real difference
    between the legs from run-to-run variation in a refusal message that names
    one member of an unordered set.
    """
    legs = [("before", before), ("after", after)]
    if control is not None:
        legs.append(("control", control))
    keys = sorted(set().union(*(set(leg) for _, leg in legs)))
    with open(path, "w", encoding="utf-8") as handle:
        header = ["fixture", "route"]
        for name, _ in legs:
            header += [f"{name}_status", f"{name}_field1", f"{name}_field2"]
        header += ["differs_raw", "differs_normalized"]
        handle.write("\t".join(header) + "\n")
        for key in keys:
            row = [key[0], key[1]]
            for _, leg in legs:
                value = leg.get(key)
                if value is None:
                    row += ["MISSING", "", ""]
                else:
                    row += [value[0], value[1] if len(value) > 1 else "", value[2] if len(value) > 2 else ""]
            raw = "1" if before.get(key) != after.get(key) else "0"
            norm = "1" if normalize_outcome(before.get(key, ())) != normalize_outcome(after.get(key, ())) else "0"
            row += [raw, norm]
            handle.write("\t".join(row) + "\n")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--before", required=True)
    parser.add_argument("--after", required=True)
    parser.add_argument("--control", help="a second run of the BEFORE binary, for the determinism control")
    parser.add_argument("--probe-before")
    parser.add_argument("--probe-after")
    parser.add_argument("--corpus-root")
    parser.add_argument("--scratch")
    parser.add_argument("--sample", type=int, default=12)
    parser.add_argument("--refusals", help="write the side-by-side outcome table here")
    parser.add_argument("--note", action="append", default=[],
                        help="a provenance line to echo at the top of the summary")
    args = parser.parse_args()

    b_class, b_edit, b_out, b_mem = load(args.before)
    a_class, a_edit, a_out, a_mem = load(args.after)
    c_class = c_edit = c_out = c_mem = None
    if args.control:
        c_class, c_edit, c_out, c_mem = load(args.control)

    if args.refusals:
        write_refusals(args.refusals, b_out, a_out, c_out)

    say = print
    say("change 0653 — OOXML corpus publication evidence")
    say("=" * 78)
    say("")
    say(f"BEFORE rows : {args.before}")
    say(f"AFTER rows  : {args.after}")
    if args.note:
        say("")
        say("PROVENANCE")
        say("-" * 78)
        for line in args.note:
            say(line)
    say("")
    say("Column meaning: `identical-to-source` compares the UNCOMPRESSED member")
    say("payload of the published archive against the same-named member's")
    say("uncompressed payload in the source archive. The compressed (stored)")
    say("payload hash is carried in a separate column and is never what")
    say("`identical-to-source` means.")
    say("")

    # ---------------------------------------------------------------- corpus
    say("CORPUS")
    say("-" * 78)
    say(f"fixtures scanned            : {len(b_class)}")
    formats = Counter(value[0] for value in b_class.values())
    for name, count in sorted(formats.items()):
        say(f"  classified {name:<14}: {count}")
    if b_class.keys() != a_class.keys():
        say("  !! the two legs scanned different fixtures")
    mismatched = [f for f in b_class if b_class[f] != a_class.get(f)]
    say(f"classification differs      : {len(mismatched)}")
    routes = sorted({route for _, route in b_out})
    say(f"routes                      : {len(routes)}")
    for route in routes:
        say(f"  {route}")
    say("")

    # -------------------------------------------------- 1. pristine members
    say("1. PRISTINE-PART PROOF")
    say("-" * 78)
    say("Definition A (declared): a member is 'the edited part' for a route when")
    say("its name matches that route's documented target set; every other member")
    say("is pristine. `*_noop_save` makes no edit, so it has no edited part.")
    say("")
    keys = set(b_mem) | set(a_mem)
    only_before = sorted(set(b_mem) - set(a_mem))
    only_after = sorted(set(a_mem) - set(b_mem))
    say(f"member rows, BEFORE         : {len(b_mem)}")
    say(f"member rows, AFTER          : {len(a_mem)}")
    say(f"rows present on one leg only: {len(only_before) + len(only_after)}")
    for key in (only_before + only_after)[:20]:
        say(f"  {key}")
    say("")

    non_target = [k for k in keys if not is_target(k[1], k[2])]
    b_pristine_identical = sum(1 for k in non_target if k in b_mem and b_mem[k][3] == 1)
    a_pristine_identical = sum(1 for k in non_target if k in a_mem and a_mem[k][3] == 1)
    say(f"non-edited-part member rows : {len(non_target)}")
    say(f"  identical-to-source, BEFORE: {b_pristine_identical}")
    say(f"  identical-to-source, AFTER : {a_pristine_identical}")

    flag_differs = [k for k in non_target if k in b_mem and k in a_mem and b_mem[k][3] != a_mem[k][3]]
    say(f"  identical flag DIFFERS     : {len(flag_differs)}")
    for key in flag_differs[:50]:
        say(f"    {key}  before={b_mem[key][3]} after={a_mem[key][3]}")
    say("")

    say("Definition B (observed): a member is pristine when the BEFORE leg")
    say("published it byte for byte as the source archive held it.")
    b_identical = {k for k in b_mem if b_mem[k][3] == 1}
    say(f"pristine member rows, BEFORE: {len(b_identical)}")
    lost = sorted(k for k in b_identical if k not in a_mem or a_mem[k][3] != 1)
    say(f"  no longer identical AFTER  : {len(lost)}")
    for key in lost[:50]:
        say(f"    {key}")
    hash_moved = sorted(k for k in b_identical if k in a_mem and a_mem[k][0] != b_mem[k][0])
    say(f"  published hash moved       : {len(hash_moved)}")
    for key in hash_moved[:50]:
        say(f"    {key}  before={b_mem[key][0]} after={a_mem[key][0]}")
    gained = sorted(k for k in a_mem if a_mem[k][3] == 1 and (k not in b_mem or b_mem[k][3] != 1))
    say(f"  newly identical on AFTER   : {len(gained)}")
    for key in gained[:50]:
        say(f"    {key}")
    say("")
    say("Per route (member rows / identical-to-source / not identical):")
    per_route = defaultdict(lambda: [0, 0, 0, 0])
    for key in sorted(keys):
        route = key[1]
        if key in b_mem:
            per_route[route][0] += 1
            per_route[route][1] += b_mem[key][3]
        if key in a_mem:
            per_route[route][2] += 1
            per_route[route][3] += a_mem[key][3]
    say(f"  {'route':<36}{'BEFORE':>18}{'AFTER':>18}")
    for route in sorted(per_route):
        total_b, ident_b, total_a, ident_a = per_route[route]
        say(f"  {route:<36}{ident_b:>8}/{total_b:<9}{ident_a:>8}/{total_a:<9}")
    say("")

    # ------------------------------------------------ 2. regenerated members
    say("2. REGENERATED-PART ENUMERATION (published hash differs between legs)")
    say("-" * 78)
    differing = sorted(
        k for k in keys
        if k in b_mem and k in a_mem and b_mem[k][0] != a_mem[k][0]
    )
    say(f"member rows whose published UNCOMPRESSED hash differs : {len(differing)}")
    comp_differing = sorted(
        k for k in keys
        if k in b_mem and k in a_mem and b_mem[k][4] != a_mem[k][4]
    )
    say(f"member rows whose published COMPRESSED hash differs   : {len(comp_differing)}")
    say("")
    say("by route:")
    for route, count in sorted(Counter(k[1] for k in differing).items()):
        say(f"  {route:<36}{count:>6}")
    say("by member-name pattern (digit runs collapsed to N):")
    for pattern, count in sorted(Counter(group(k[2]) for k in differing).items()):
        say(f"  {pattern:<36}{count:>6}")
    say("")
    say("every differing row:")
    say(f"  {'fixture':<72}{'route':<24}{'member':<26}{'before-bytes':>13}{'after-bytes':>12}")
    for key in differing:
        say(f"  {key[0]:<72}{key[1]:<24}{key[2]:<26}{b_mem[key][1]:>13}{a_mem[key][1]:>12}")
    say("")

    # canonical comparison
    say("canonical XML comparison (namespace-resolving, declarations removed):")
    sample = differing[: args.sample]
    if not sample:
        say("  no differing member to compare")
    elif not (args.probe_before and args.probe_after and args.corpus_root and args.scratch):
        say("  NOT RUN — pass --probe-before/--probe-after/--corpus-root/--scratch")
        say("  to let this script re-publish the differing fixtures with")
        say("  SAVEPROBE_KEEP=1 and canonicalize the differing members.")
    else:
        scratch = pathlib.Path(args.scratch)
        equal = unequal = 0
        for leg, probe in (("before", args.probe_before), ("after", args.probe_after)):
            directory = scratch / leg
            if directory.exists():
                for child in directory.iterdir():
                    child.unlink()
            directory.mkdir(parents=True, exist_ok=True)
            environment = dict(os.environ, SAVEPROBE_KEEP="1")
            for fixture in sorted({k[0] for k in sample}):
                subprocess.run(
                    [probe, args.corpus_root, fixture, str(directory)],
                    check=True, stdout=subprocess.DEVNULL, env=environment,
                )
        for key in sample:
            fixture, route, member = key
            name = hashlib.sha256(fixture.encode()).hexdigest()[:24] + f"-{route}.out"
            payload = {}
            for leg in ("before", "after"):
                with zipfile.ZipFile(scratch / leg / name) as archive:
                    payload[leg] = archive.read(member)
            try:
                same = canonical(payload["before"]) == canonical(payload["after"])
            except ET.ParseError as error:
                same = False
                say(f"  {fixture} :: {member} — XML parse error: {error}")
            declarations = {
                leg: len(re.findall(rb"xmlns:", payload[leg])) for leg in ("before", "after")
            }
            verdict = "EQUAL" if same else "NOT EQUAL"
            say(
                f"  {verdict:<10}{fixture}  [{member}]  "
                f"{len(payload['before'])} -> {len(payload['after'])} bytes "
                f"({len(payload['after']) / len(payload['before']):.3f}x), "
                f"xmlns: decls {declarations['before']} -> {declarations['after']}"
            )
            if same:
                equal += 1
            else:
                unequal += 1
                before_nodes = canonical(payload["before"])
                after_nodes = canonical(payload["after"])
                for index, (left, right) in enumerate(zip(before_nodes, after_nodes)):
                    if left != right:
                        say(f"      first structural difference at canonical node {index}")
                        say(f"        before: {left}")
                        say(f"        after : {right}")
                        break
                else:
                    say(f"      node counts differ: {len(before_nodes)} vs {len(after_nodes)}")
        for leg in ("before", "after"):
            for child in (scratch / leg).iterdir():
                child.unlink()
            (scratch / leg).rmdir()
        say("")
        say(f"  canonically EQUAL     : {equal}")
        say(f"  canonically NOT EQUAL : {unequal}")
        if len(differing) > args.sample:
            say(f"  (sampled {args.sample} of {len(differing)} differing rows)")
    say("")

    # ------------------------------------------------------- 3. refusals
    say("3. REFUSAL IDENTITY")
    say("-" * 78)
    outcome_keys = sorted(set(b_out) | set(a_out))
    say(f"outcome rows per leg        : {len(b_out)} / {len(a_out)}")
    say(f"  OK,  BEFORE / AFTER       : "
        f"{sum(1 for v in b_out.values() if v[0] == 'OK')} / "
        f"{sum(1 for v in a_out.values() if v[0] == 'OK')}")
    say(f"  ERR, BEFORE / AFTER       : "
        f"{sum(1 for v in b_out.values() if v[0] == 'ERR')} / "
        f"{sum(1 for v in a_out.values() if v[0] == 'ERR')}")
    raw_differ = [k for k in outcome_keys if b_out.get(k) != a_out.get(k)]
    norm_differ = [
        k for k in outcome_keys
        if normalize_outcome(b_out.get(k, ())) != normalize_outcome(a_out.get(k, ()))
    ]
    status_differ = [
        k for k in outcome_keys
        if b_out.get(k, ("",))[0] != a_out.get(k, ("",))[0]
    ]
    err_norm_differ = [
        k for k in norm_differ
        if b_out.get(k, ("",))[0] == "ERR" or a_out.get(k, ("",))[0] == "ERR"
    ]
    ok_norm_differ = [k for k in norm_differ if k not in err_norm_differ]
    say(f"outcomes differing, raw text: {len(raw_differ)}")
    say(f"outcomes differing, normalized for the unordered-set message: {len(norm_differ)}")
    say(f"  of those, an ERR outcome on either leg (a refusal difference): {len(err_norm_differ)}")
    for key in err_norm_differ[:50]:
        say(f"    {key}")
        say(f"      before={b_out.get(key)}")
        say(f"      after ={a_out.get(key)}")
    say(f"  of those, OK on both legs (only the published byte count moved): {len(ok_norm_differ)}")
    for key in ok_norm_differ[:50]:
        say(f"    {key}  published bytes {b_out.get(key)[1]} -> {a_out.get(key)[1]}")
    say(f"outcomes whose OK/ERR status differs: {len(status_differ)}")
    for key in status_differ[:50]:
        say(f"  {key}  before={b_out.get(key)}  after={a_out.get(key)}")
    if args.control:
        control_differ = [k for k in sorted(set(b_out) | set(c_out)) if b_out.get(k) != c_out.get(k)]
        control_mem = [k for k in set(b_mem) | set(c_mem) if b_mem.get(k) != c_mem.get(k)]
        say("")
        say("determinism control — the BEFORE binary run twice over the same corpus:")
        control_norm = [
            k for k in sorted(set(b_out) | set(c_out))
            if normalize_outcome(b_out.get(k, ())) != normalize_outcome(c_out.get(k, ()))
        ]
        say(f"  outcome rows differing, raw text     : {len(control_differ)}")
        say(f"  outcome rows differing, normalized   : {len(control_norm)}")
        say(f"  member rows differing                : {len(control_mem)}")
        say("  (a raw-text outcome difference that also appears here is run-to-run")
        say("   nondeterminism in the message, not a difference between the legs.)")
        say("  The control's own row file is not among this packet's retained")
        say("  deliverables; its outcome rows are retained, side by side with the")
        say("  two legs, in refusals.tsv.")
        say(f"  routes affected: {sorted({k[1] for k in control_differ})}")
    say("")
    say("ERR outcomes by route and Debug identity (variant before the first '(' or '{'):")
    identity = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*")
    table = defaultdict(lambda: [0, 0])
    for key in outcome_keys:
        for index, source in ((0, b_out), (1, a_out)):
            value = source.get(key)
            if value and value[0] == "ERR":
                match = identity.match(value[1])
                table[(key[1], match.group(0) if match else value[1])][index] += 1
    say(f"  {'route':<36}{'Debug variant':<26}{'BEFORE':>8}{'AFTER':>8}")
    for (route, variant), (count_b, count_a) in sorted(table.items()):
        flag = "" if count_b == count_a else "   <-- DIFFERS"
        say(f"  {route:<36}{variant:<26}{count_b:>8}{count_a:>8}{flag}")
    say("")

    say("edit-outcome rows (X):")
    x_keys = sorted(set(b_edit) | set(a_edit))
    x_differ = [k for k in x_keys if b_edit.get(k) != a_edit.get(k)]
    say(f"  rows per leg               : {len(b_edit)} / {len(a_edit)}")
    say(f"  admitted, BEFORE / AFTER   : "
        f"{sum(1 for v in b_edit.values() if v[0] == 'admitted')} / "
        f"{sum(1 for v in a_edit.values() if v[0] == 'admitted')}")
    say(f"  refused,  BEFORE / AFTER   : "
        f"{sum(1 for v in b_edit.values() if v[0] == 'refused')} / "
        f"{sum(1 for v in a_edit.values() if v[0] == 'refused')}")
    say(f"  differing                  : {len(x_differ)}")
    for key in x_differ[:50]:
        say(f"    {key}  before={b_edit.get(key)}  after={a_edit.get(key)}")
    say("")

    # -------------------------------------------------------- 4. byte totals
    say("4. BYTE TOTALS (sum of published archive bytes over successful saves)")
    say("-" * 78)
    say(f"  {'route':<36}{'saves':>7}{'BEFORE bytes':>16}{'AFTER bytes':>16}{'delta':>12}{'ratio':>9}")
    grand = [0, 0]
    for route in routes:
        total_b = total_a = saves = 0
        for key in outcome_keys:
            if key[1] != route:
                continue
            value_b, value_a = b_out.get(key), a_out.get(key)
            if value_b and value_b[0] == "OK" and value_a and value_a[0] == "OK":
                total_b += int(value_b[1])
                total_a += int(value_a[1])
                saves += 1
        grand[0] += total_b
        grand[1] += total_a
        ratio = f"{total_a / total_b:.6f}" if total_b else "-"
        say(f"  {route:<36}{saves:>7}{total_b:>16}{total_a:>16}{total_a - total_b:>12}{ratio:>9}")
    ratio = f"{grand[1] / grand[0]:.6f}" if grand[0] else "-"
    say(f"  {'ALL ROUTES':<36}{'':>7}{grand[0]:>16}{grand[1]:>16}{grand[1] - grand[0]:>12}{ratio:>9}")
    say("")
    say("published uncompressed member payload totals:")
    say(f"  {'route':<36}{'BEFORE bytes':>16}{'AFTER bytes':>16}{'delta':>12}")
    for route in routes:
        total_b = sum(b_mem[k][1] for k in b_mem if k[1] == route)
        total_a = sum(a_mem[k][1] for k in a_mem if k[1] == route)
        say(f"  {route:<36}{total_b:>16}{total_a:>16}{total_a - total_b:>12}")
    say("")
    say("end of summary")
    return 0


if __name__ == "__main__":
    sys.exit(main())
