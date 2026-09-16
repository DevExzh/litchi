#!/usr/bin/env python3
"""Recompute every number reported for the change-0653 semantic-read evidence.

Inputs (in the directory given as the first argument, default: this file's own
directory):

* ``digest-before.txt.gz`` / ``digest-after.txt.gz`` - one line per observation,
  ``<fixture>\\t<key>\\t<value>``. A refused read prints ``ERR\\t<Debug>`` in
  place of the value; a panicking read prints ``PANIC``; a fixture that exceeded
  the per-fixture timeout prints ``probe.outcome\\tTIMEOUT``.
* ``dump-before.txt.gz`` / ``dump-after.txt.gz`` - the same probe run in
  ``--dump`` mode, which prints only the markup accessors and prints the bytes
  themselves, base64-encoded, instead of a length and a hash. They are what
  makes the namespace-resolving canonical comparison possible.

Outputs: ``diff.tsv`` (every differing line, side by side, with its class) and
``summary.txt`` (every number reported).

Classes:

* ``a-value``   a text/value changed.
* ``b-outcome`` an outcome changed between OK and ERR, or between two different
  errors. These are the findings that matter most.
* ``c-markup-canonical-same`` a markup length or hash changed but the
  namespace-resolving canonical form of the bytes is unchanged.
* ``d-markup-canonical-differs`` a markup hash changed AND the canonical form
  changed too. Any of these is a blocking finding.
"""

import base64
import gzip
import os
import re
import sys
import xml.etree.ElementTree as ET

DECLARATION = re.compile(rb'xmlns(?::([A-Za-z_][\w.\-]*))?\s*=\s*"([^"]*)"')
ELEMENT_PREFIX = re.compile(rb"<\s*/?\s*([A-Za-z_][\w.\-]*):")
ATTRIBUTE_PREFIX = re.compile(rb"[\s<]([A-Za-z_][\w.\-]*):[\w.\-]+\s*=")
MARKUP_VALUE = re.compile(r"^len=(\d+);sha=([0-9a-f]{64})$")
OUTCOME_PREFIXES = ("ERR", "PANIC", "TIMEOUT")

UNRESOLVED = "urn:litchi-unresolved:"


def read_lines(path):
    opener = gzip.open if path.endswith(".gz") else open
    with opener(path, "rt", encoding="utf-8") as handle:
        return [line.rstrip("\n") for line in handle if line.strip()]


def split_line(line):
    parts = line.split("\t", 2)
    while len(parts) < 3:
        parts.append("")
    return parts[0], parts[1], parts[2]


def load_digest(path):
    order = []
    values = {}
    for line in read_lines(path):
        fixture, key, value = split_line(line)
        order.append((fixture, key))
        values[(fixture, key)] = value
    return order, values


def load_dump(path):
    values = {}
    for line in read_lines(path):
        fixture, key, value = split_line(line)
        if value.startswith("b64="):
            values[(fixture, key)] = base64.b64decode(value[4:])
    return values


def declarations(fragment):
    found = {}
    for match in DECLARATION.finditer(fragment):
        prefix = match.group(1).decode("utf-8") if match.group(1) else ""
        found.setdefault(prefix, match.group(2).decode("utf-8"))
    return found


def used_prefixes(fragment):
    prefixes = set()
    for match in ELEMENT_PREFIX.finditer(fragment):
        prefixes.add(match.group(1).decode("utf-8"))
    for match in ATTRIBUTE_PREFIX.finditer(fragment):
        prefixes.add(match.group(1).decode("utf-8"))
    # `xml` is bound implicitly by the XML specification and `xmlns` is
    # reserved: re-declaring either on the synthetic root is a parse error, and
    # neither needs a declaration for a prefixed name to resolve.
    return prefixes - {"xml", "xmlns"}


def serialize(element):
    parts = [element.tag]
    for name in sorted(element.attrib):
        parts.append("|@%s=%s" % (name, " ".join(element.attrib[name].split())))
    if element.text and element.text.strip():
        parts.append("|#" + " ".join(element.text.split()))
    for child in element:
        parts.append("|(" + serialize(child) + ")")
        if child.tail and child.tail.strip():
            parts.append("|#" + " ".join(child.tail.split()))
    return "".join(parts)


def canonical(fragment, scope, unresolved):
    """The fragment's namespace-resolving canonical form.

    `scope` is the union of the namespace declarations found in the two legs'
    fragments, applied on a synthetic root so that a prefix a leg inherits from
    its part - rather than re-declaring on the fragment root - still resolves.
    Declarations inside the fragment win over the synthetic root, so this never
    changes a binding a fragment makes for itself. Namespace declarations are
    excluded from the canonical form because ElementTree reports resolved
    `{uri}local` names and drops `xmlns` attributes.
    """
    bindings = {
        prefix: uri for prefix, uri in scope.items() if prefix not in ("xml", "xmlns")
    }
    for prefix in used_prefixes(fragment):
        if prefix not in bindings:
            bindings[prefix] = UNRESOLVED + prefix
            unresolved.add(prefix)
    attributes = "".join(
        ' xmlns="%s"' % uri if prefix == "" else ' xmlns:%s="%s"' % (prefix, uri)
        for prefix, uri in sorted(bindings.items())
    )
    document = b"<canonroot" + attributes.encode("utf-8") + b">" + fragment + b"</canonroot>"
    root = ET.fromstring(document)
    parts = []
    if root.text and root.text.strip():
        parts.append("#" + " ".join(root.text.split()))
    for child in root:
        parts.append("(" + serialize(child) + ")")
        if child.tail and child.tail.strip():
            parts.append("#" + " ".join(child.tail.split()))
    return "|".join(parts)


def classify(before_value, after_value, key_pair, dump_before, dump_after, state):
    before_outcome = before_value.split("\t", 1)[0] in OUTCOME_PREFIXES
    after_outcome = after_value.split("\t", 1)[0] in OUTCOME_PREFIXES
    if before_outcome or after_outcome:
        return "b-outcome"
    before_markup = MARKUP_VALUE.match(before_value)
    after_markup = MARKUP_VALUE.match(after_value)
    if not (before_markup and after_markup):
        return "a-value"
    before_bytes = dump_before.get(key_pair)
    after_bytes = dump_after.get(key_pair)
    if before_bytes is None or after_bytes is None:
        state["missing_dump"].append(key_pair)
        return "d-markup-canonical-differs"
    scope = declarations(before_bytes)
    for prefix, uri in declarations(after_bytes).items():
        scope.setdefault(prefix, uri)
    try:
        before_canonical = canonical(before_bytes, scope, state["unresolved"])
        after_canonical = canonical(after_bytes, scope, state["unresolved"])
    except ET.ParseError as error:
        state["parse_errors"].append((key_pair, str(error)))
        return "d-markup-canonical-differs"
    if before_canonical == after_canonical:
        return "c-markup-canonical-same"
    state["canonical_pairs"].append((key_pair, before_canonical, after_canonical))
    return "d-markup-canonical-differs"


def self_test(rows, dump_before, dump_after, state):
    """Negative controls for the canonicalizer.

    A canonical comparison that reported "same" for everything would hide a
    real change, so the same function is run against deliberately perturbed
    copies of a real differing fragment. Each control must report a difference,
    and the unperturbed copy must not.
    """
    pair = None
    for fixture, key, kind, _before, _after in rows:
        if kind.startswith(("c-markup", "d-markup")) and (fixture, key) in dump_before:
            pair = (fixture, key)
            break
    if pair is None:
        return ["      (no markup pair available to test)"]
    before_bytes = dump_before[pair]
    after_bytes = dump_after[pair]
    scope = declarations(before_bytes)
    for prefix, uri in declarations(after_bytes).items():
        scope.setdefault(prefix, uri)
    baseline = canonical(before_bytes, scope, state["unresolved"])
    root_name = re.match(rb"<\s*([\w.\-:]+)", after_bytes)
    root = root_name.group(1) if root_name else b"x"
    first_attribute = re.search(rb'\s([\w.\-:]+)="([^"]*)"', after_bytes)
    controls = []
    if first_attribute:
        controls.append(
            (
                "attribute value changed",
                after_bytes.replace(
                    first_attribute.group(0),
                    b' %s="%s"' % (first_attribute.group(1), b"perturbed"),
                    1,
                ),
            )
        )
    controls.append(
        ("attribute added", after_bytes.replace(b"<" + root, b"<" + root + b' zz:zz="1"', 1))
    )
    # Rebind the namespace the root element itself resolves through, so the
    # control actually changes a resolved name rather than adding an unused
    # declaration.
    if b":" in root:
        rebound = after_bytes.replace(
            b"<" + root, b'<%s xmlns:%s="urn:rebound"' % (root, root.split(b":")[0]), 1
        )
    else:
        rebound = re.sub(rb'\sxmlns="[^"]*"', b"", after_bytes, count=1).replace(
            b"<" + root, b'<%s xmlns="urn:rebound"' % root, 1
        )
    controls.append(("root namespace rebound", rebound))
    text_match = re.search(rb">([^<>]{3,})<", after_bytes)
    if text_match:
        controls.append(
            ("text changed", after_bytes.replace(text_match.group(0), b">perturbed<", 1))
        )
    controls.append(("unperturbed copy", after_bytes))

    report = ["      control pair: %s %s" % pair]
    for label, data in controls:
        expected_difference = label != "unperturbed copy"
        try:
            differs = canonical(data, scope, state["unresolved"]) != baseline
            verdict = "PASS" if differs == expected_difference else "FAIL"
            report.append(
                "      %-38s canonical differs=%-5s %s" % (label, differs, verdict)
            )
        except ET.ParseError as error:
            report.append("      %-38s parse error: %s" % (label, error))
    return report


def normalize_key(key):
    key = re.sub(r"\.cell\.[A-Z]+[0-9]+", ".cell.ADDR", key)
    return re.sub(r"\.\d+(?=\.|$)", ".N", key)


def main():
    directory = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))

    def resolve(name):
        for candidate in (name + ".gz", name):
            path = os.path.join(directory, candidate)
            if os.path.exists(path):
                return path
        raise SystemExit("missing input: " + name)

    before_order, before_values = load_digest(resolve("digest-before.txt"))
    after_order, after_values = load_digest(resolve("digest-after.txt"))
    dump_before = load_dump(resolve("dump-before.txt"))
    dump_after = load_dump(resolve("dump-after.txt"))

    before_fixtures = sorted({fixture for fixture, _key in before_order})
    after_fixtures = sorted({fixture for fixture, _key in after_order})

    state = {
        "unresolved": set(),
        "parse_errors": [],
        "canonical_pairs": [],
        "missing_dump": [],
    }

    keys = list(dict.fromkeys(before_order + after_order))
    rows = []
    for key_pair in keys:
        before_value = before_values.get(key_pair, "<ABSENT>")
        after_value = after_values.get(key_pair, "<ABSENT>")
        if before_value == after_value:
            continue
        kind = classify(before_value, after_value, key_pair, dump_before, dump_after, state)
        rows.append((key_pair[0], key_pair[1], kind, before_value, after_value))

    with open(os.path.join(directory, "diff.tsv"), "w", encoding="utf-8") as handle:
        handle.write("fixture\tkey\tclass\tbefore\tafter\n")
        for fixture, key, kind, before_value, after_value in rows:
            handle.write(
                "%s\t%s\t%s\t%s\t%s\n"
                % (
                    fixture,
                    key,
                    kind,
                    before_value.replace("\t", "\\t"),
                    after_value.replace("\t", "\\t"),
                )
            )

    markup_bytes = {}
    markup_lines = {}
    for label, values in (("before", before_values), ("after", after_values)):
        total = 0
        count = 0
        for value in values.values():
            match = MARKUP_VALUE.match(value)
            if match:
                total += int(match.group(1))
                count += 1
        markup_bytes[label] = total
        markup_lines[label] = count

    groups = {}
    for fixture, key, kind, before_value, after_value in rows:
        bucket = groups.setdefault(normalize_key(key), {"rows": [], "fixtures": set(), "classes": {}})
        bucket["rows"].append((fixture, key, kind, before_value, after_value))
        bucket["fixtures"].add(fixture)
        bucket["classes"][kind] = bucket["classes"].get(kind, 0) + 1

    class_counts = {}
    for _fixture, _key, kind, _before, _after in rows:
        class_counts[kind] = class_counts.get(kind, 0) + 1

    lines = []
    lines.append("change 0653 semantic-read comparison")
    lines.append("=" * 72)
    lines.append("")
    binaries = os.path.join(directory, "binaries.sha256")
    if os.path.exists(binaries):
        lines.append("0. Probe binaries (built from readprobe/, one leg each)")
        with open(binaries, encoding="utf-8") as handle:
            for entry in handle:
                if entry.strip():
                    lines.append("   " + entry.strip())
        lines.append("")
    lines.append("1. Corpus and observation counts")
    lines.append("   fixtures processed (before leg): %d" % len(before_fixtures))
    lines.append("   fixtures processed (after leg):  %d" % len(after_fixtures))
    lines.append("   observation lines (before leg):  %d" % len(before_order))
    lines.append("   observation lines (after leg):   %d" % len(after_order))
    lines.append(
        "   (fixture, key) sequence identical between legs: %s"
        % (before_order == after_order)
    )
    lines.append("   timeouts recorded: before=%d after=%d" % (
        sum(1 for (_f, k), v in before_values.items() if k == "probe.outcome" and v == "TIMEOUT"),
        sum(1 for (_f, k), v in after_values.items() if k == "probe.outcome" and v == "TIMEOUT"),
    ))
    lines.append("   refusals (ERR) recorded: before=%d after=%d" % (
        sum(1 for value in before_values.values() if value.startswith("ERR")),
        sum(1 for value in after_values.values() if value.startswith("ERR")),
    ))
    lines.append("   panics recorded: before=%d after=%d" % (
        sum(1 for value in before_values.values() if value.startswith("PANIC")),
        sum(1 for value in after_values.values() if value.startswith("PANIC")),
    ))
    lines.append("")
    lines.append("2. Differing lines")
    lines.append("   differing observation lines: %d" % len(rows))
    lines.append("   fixtures with at least one differing line: %d" % len({row[0] for row in rows}))
    lines.append("")
    for name in sorted(groups, key=lambda name: -len(groups[name]["rows"])):
        bucket = groups[name]
        lines.append(
            "   key group %s: %d lines across %d fixtures (%s)"
            % (
                name,
                len(bucket["rows"]),
                len(bucket["fixtures"]),
                ", ".join(
                    "%s=%d" % (kind, count) for kind, count in sorted(bucket["classes"].items())
                ),
            )
        )
        for fixture, key, kind, before_value, after_value in bucket["rows"][:3]:
            lines.append("      example fixture: %s" % fixture)
            lines.append("         key:    %s  [%s]" % (key, kind))
            lines.append("         before: %s" % before_value.replace("\t", " | "))
            lines.append("         after:  %s" % after_value.replace("\t", " | "))
        lines.append("")
    lines.append("3. Classes")
    for kind in (
        "a-value",
        "b-outcome",
        "c-markup-canonical-same",
        "d-markup-canonical-differs",
    ):
        lines.append("   %-28s %d" % (kind, class_counts.get(kind, 0)))
    lines.append("")
    lines.append("   Every class-b line (outcome changed):")
    outcome_rows = [row for row in rows if row[2] == "b-outcome"]
    if not outcome_rows:
        lines.append("      (none)")
    for fixture, key, _kind, before_value, after_value in outcome_rows:
        lines.append("      %s  %s" % (fixture, key))
        lines.append("         before: %s" % before_value.replace("\t", " | "))
        lines.append("         after:  %s" % after_value.replace("\t", " | "))
    lines.append("")
    lines.append("   Every class-d line (canonical form changed - blocking):")
    blocking_rows = [row for row in rows if row[2] == "d-markup-canonical-differs"]
    if not blocking_rows:
        lines.append("      (none)")
    for fixture, key, _kind, before_value, after_value in blocking_rows[:50]:
        lines.append("      %s  %s" % (fixture, key))
        lines.append("         before: %s" % before_value.replace("\t", " | "))
        lines.append("         after:  %s" % after_value.replace("\t", " | "))
    lines.append("")
    lines.append("   canonicalization diagnostics")
    lines.append("      prefixes bound to a synthetic URI because neither leg's")
    lines.append("      fragment declared them: %d %s" % (
        len(state["unresolved"]),
        sorted(state["unresolved"]),
    ))
    lines.append("      canonical XML parse failures: %d" % len(state["parse_errors"]))
    for key_pair, message in state["parse_errors"][:10]:
        lines.append("         %s %s: %s" % (key_pair[0], key_pair[1], message))
    lines.append("      markup lines with no dump-mode bytes: %d" % len(state["missing_dump"]))
    lines.append("")
    lines.append("4. Markup accessor bytes")
    lines.append("   markup accessor observations per leg: before=%d after=%d" % (
        markup_lines["before"],
        markup_lines["after"],
    ))
    lines.append("   total bytes returned (before): %d" % markup_bytes["before"])
    lines.append("   total bytes returned (after):  %d" % markup_bytes["after"])
    delta = markup_bytes["after"] - markup_bytes["before"]
    lines.append("   change: %+d bytes (%.2f%%)" % (
        delta,
        100.0 * delta / markup_bytes["before"] if markup_bytes["before"] else 0.0,
    ))
    if markup_bytes["after"]:
        lines.append(
            "   before/after ratio: %.2fx"
            % (markup_bytes["before"] / markup_bytes["after"])
        )
    lines.append("")
    lines.append("   markup accessor observations by key family (before leg):")
    families = {}
    for (fixture, key), value in before_values.items():
        if MARKUP_VALUE.match(value):
            family = families.setdefault(normalize_key(key), [0, 0, set()])
            family[0] += 1
            family[1] += int(MARKUP_VALUE.match(value).group(1))
            family[2].add(fixture)
    for name in sorted(families, key=lambda name: -families[name][0]):
        count, total, fixtures = families[name]
        after_total = sum(
            int(MARKUP_VALUE.match(after_values[(fixture, key)]).group(1))
            for (fixture, key), value in before_values.items()
            if MARKUP_VALUE.match(value)
            and normalize_key(key) == name
            and MARKUP_VALUE.match(after_values.get((fixture, key), ""))
        )
        lines.append(
            "      %-36s n=%-5d fixtures=%-4d before=%-9d after=%d"
            % (name, count, len(fixtures), total, after_total)
        )
    lines.append("")
    lines.append("5. Coverage of the accessors change 0653 migrated")
    coverage = [
        ("worksheet sheetView retained markup", r"\.view\.\d+\.retained_xml$"),
        ("worksheet sheetView ext markup", r"\.view\.\d+\.ext\.\d+\.markup$"),
        ("worksheet sheetViews collection ext markup", r"\.views\.ext\.\d+\.markup$"),
        ("worksheet pivotArea markup", r"\.pivot\.\d+\.area_markup$"),
        ("worksheet ignoredErrors present", r"\.ignored_errors\.entries$"),
        ("worksheet ignoredErrors ext markup", r"\.ierr\.ext\.\d+\.markup$"),
        ("named sheet views present", r"\.nsv\.views$"),
        ("named sheet views ext markup", r"\.nsv\.ext\.\d+\.markup$"),
        ("calculation chain present", r"^xlsx\.chain\.len$"),
        ("calculation chain extLst markup", r"^xlsx\.chain\.extlst$"),
        ("conditional formatting inline dxf markup", r"\.rule\.\d+\.dxf$"),
        ("pptx shape xml span", r"^pptx\.sld\.\d+\.shp\.\d+\.xml$"),
        ("pptx source-backed slide images", r"\.images\.len$"),
        ("docx comment xml_bytes", r"^docx\.cmt\.\d+\.xml$"),
        ("docx footnote xml_bytes", r"^docx\.fn\.\d+\.xml$"),
        ("docx endnote xml_bytes", r"^docx\.en\.\d+\.xml$"),
        ("docx opaque block xml_bytes", r"^docx\.block\.\d+\.xml$"),
        ("docx opaque inline xml_bytes", r"^docx\.par\.\d+\.inline\.\d+\.xml$"),
        ("docx opaque run content xml_bytes", r"^docx\.par\.\d+\.run\.\d+\.content\.\d+\.xml$"),
        ("docx paragraph extensions", r"^docx\.par\.\d+\.ext$"),
        ("docx table row extension ids", r"\.row\.\d+\.extids$"),
    ]
    for label, pattern in coverage:
        matcher = re.compile(pattern)
        hits = [
            (fixture, key)
            for (fixture, key) in before_values
            if matcher.search(key)
        ]
        errors = sum(
            1 for pair in hits if before_values[pair].startswith(OUTCOME_PREFIXES)
        )
        with_markup = [pair for pair in hits if MARKUP_VALUE.match(before_values[pair])]
        lines.append(
            "   %-42s observations=%-6d with-bytes=%-5d fixtures=%-4d refusals=%d"
            % (
                label,
                len(hits),
                len(with_markup),
                len({pair[0] for pair in hits}),
                errors,
            )
        )
    lines.append("")
    lines.append("   PowerPoint cross-check: the source-backed picture inventory is")
    lines.append("   the one migrated read whose fragment is only re-declared when the")
    lines.append("   owner was MCE-rewritten, so it is only covered on such a slide.")
    rewritten = {}
    images = {}
    for (fixture, key), value in before_values.items():
        if key.endswith(".scene.rewritten"):
            rewritten[(fixture, key.split(".")[2])] = value
        if key.endswith(".images.len"):
            images[(fixture, key.split(".")[2])] = int(value)
    lines.append(
        "      slides whose images() succeeded: %d; of those, MCE-rewritten: %d"
        % (
            len(images),
            sum(1 for pair in images if rewritten.get(pair) == "true"),
        )
    )
    lines.append(
        "      slides with at least one picture: %d; of those, MCE-rewritten: %d"
        % (
            sum(1 for value in images.values() if value > 0),
            sum(
                1
                for pair, value in images.items()
                if value > 0 and rewritten.get(pair) == "true"
            ),
        )
    )
    lines.append(
        "      slides whose package-side scene was MCE-rewritten at all: %d"
        % sum(1 for value in rewritten.values() if value == "true")
    )
    lines.append("")
    lines.append("6. Canonicalizer self-test (negative controls)")
    lines.extend(self_test(rows, dump_before, dump_after, state))
    lines.append("")

    text = "\n".join(lines) + "\n"
    with open(os.path.join(directory, "summary.txt"), "w", encoding="utf-8") as handle:
        handle.write(text)
    sys.stdout.write(text)


main()
