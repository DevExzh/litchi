#!/usr/bin/env python3
"""Derive the marker-bearing corpus shape of change 0664 from real fixtures.

Change 0649 found that every prior PPTX record measured on generated corpora
whose members never mention the markup-compatibility (MCE) namespace, so the
codec's rewriting branch -- the branch that costs 93.9% of a real deck's
opened-transaction edit -- was never priced on a harness selector.  Change 0601
built producer-shaped corpora, but its PPTX shape marks only the slide parts and
its DOCX shape only ``word/document.xml``.

This script is the reproducible derivation step behind the corpora change 0664
adds to ``tools/perf-baseline/src/marker_shape.rs``.  It reads three real
fixtures that are ordinary tracked files in this repository, censuses every
member, and
emits the shape parameters the harness generator authors:

* which part kinds mention the MCE namespace,
* the exact root namespace declaration list of each such part kind -- the
  codec re-declares every in-scope binding on every emitted start tag, so this
  list is what sets the output amplification,
* the ``mc:Ignorable`` value each part kind carries, if any,
* the ``mc:AlternateContent`` block each part kind carries, if any,
* the share of members and of uncompressed bytes that mention the namespace.

Three modes:

  census  <archive> [--json OUT]   per-member census of one archive
  derive  [--out DIR]              both fixtures, plus the derived shape
  verify  [--module PATH]          re-derive and check the Rust generator's
                                   constants still agree with the fixtures

``verify`` is the gate: it fails if the checked-in generator drifts from the
fixtures it claims to be derived from.  It compares the facts the generator
authors -- the namespace URIs it adds, the PPTX root binding count, the
``mc:Ignorable`` strings, the ``mc:AlternateContent`` fragments and the fixture
paths -- and not the fixtures' byte totals, which move with the corpus size the
generator chooses.

Run from the repository root.  No network, no clock, no randomness.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import pathlib
import re
import sys
import zipfile

MCE_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# The two fixtures.  Both are ordinary tracked files; neither is a submodule.
PPTX_FIXTURE = pathlib.Path(
    "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"
)
DOCX_FIXTURE = pathlib.Path(
    "test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/layout-in-cell-2.docx"
)
# The only fixture in the corpus whose MCE coverage reaches the notes parts.
# The primary PPTX fixture has no notes slides at all (change 0649 recorded
# that its ``ppt/notesSlides/`` is empty), so the notes root declaration set is
# derived from this one instead of assumed.
PPTX_NOTES_FIXTURE = pathlib.Path(
    "test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf89064.pptx"
)

ROOT_TAG = re.compile(rb"<([A-Za-z_][\w.\-]*(?::[A-Za-z_][\w.\-]*)?)([^>]*)>")
DECLARATION = re.compile(r'xmlns:([\w.\-]+)="([^"]*)"')
IGNORABLE = re.compile(r'mc:Ignorable="([^"]*)"')
ALTERNATE_CONTENT = re.compile(
    r"<mc:AlternateContent[^>]*>.*?</mc:AlternateContent>", re.S
)


def part_kind(name: str) -> str:
    """Classify a member by the part kind the generator authors."""
    rules = (
        ("ppt/slides/", "pptx-slide"),
        ("ppt/slideLayouts/", "pptx-slide-layout"),
        ("ppt/slideMasters/", "pptx-slide-master"),
        ("ppt/notesSlides/", "pptx-notes-slide"),
        ("ppt/notesMasters/", "pptx-notes-master"),
        ("ppt/theme/", "pptx-theme"),
        ("word/document.xml", "docx-document"),
        ("word/settings.xml", "docx-settings"),
        ("word/styles.xml", "docx-styles"),
        ("word/numbering.xml", "docx-numbering"),
        ("word/fontTable.xml", "docx-font-table"),
        ("word/webSettings.xml", "docx-web-settings"),
        ("word/footnotes.xml", "docx-footnotes"),
        ("word/endnotes.xml", "docx-endnotes"),
    )
    for prefix, kind in rules:
        if name.startswith(prefix):
            return kind
    if name == "ppt/presentation.xml":
        return "pptx-presentation"
    if name.startswith("word/header") or name.startswith("word/footer"):
        return "docx-header-footer"
    return "other"


def root_facts(payload: bytes) -> dict:
    """Return the root element's declarations and markup-compatibility facts."""
    body = payload
    declaration_end = body.find(b"?>")
    if declaration_end >= 0:
        body = body[declaration_end + 2 :]
    match = ROOT_TAG.search(body)
    if match is None:
        return {"root": None, "declarations": [], "mc_ignorable": None}
    root = match.group(0).decode("utf-8", "replace")
    return {
        "root": match.group(1).decode("ascii", "replace"),
        "root_tag_bytes": len(match.group(0)),
        "declarations": [list(pair) for pair in DECLARATION.findall(root)],
        "mc_ignorable": (IGNORABLE.search(root).group(1) if "mc:Ignorable" in root else None),
    }


def census(path: pathlib.Path) -> dict:
    payload = path.read_bytes()
    members = []
    with zipfile.ZipFile(path) as archive:
        for info in archive.infolist():
            content = archive.read(info.filename)
            mentions = MCE_NAMESPACE.encode() in content
            entry = {
                "part": info.filename,
                "kind": part_kind(info.filename),
                "uncompressed_bytes": len(content),
                "sha256": hashlib.sha256(content).hexdigest(),
                "markup_compatibility_namespace": mentions,
                "alternate_content_occurrences": content.count(b"<mc:AlternateContent"),
            }
            if info.filename.endswith(".xml") or info.filename.endswith(".rels"):
                entry.update(root_facts(content))
            members.append(entry)
    total = sum(member["uncompressed_bytes"] for member in members)
    marked = [member for member in members if member["markup_compatibility_namespace"]]
    marked_bytes = sum(member["uncompressed_bytes"] for member in marked)
    return {
        "fixture": str(path),
        "archive_bytes": len(payload),
        "archive_sha256": hashlib.sha256(payload).hexdigest(),
        "member_count": len(members),
        "uncompressed_bytes": total,
        "markup_compatibility_member_count": len(marked),
        "markup_compatibility_bytes": marked_bytes,
        "markup_compatibility_byte_share": (
            round(marked_bytes / total, 6) if total else 0.0
        ),
        "members": members,
    }


def first_of_kind(report: dict, kind: str) -> dict | None:
    for member in report["members"]:
        if member["kind"] == kind and member.get("markup_compatibility_namespace"):
            return member
    return None


def alternate_content_template(path: pathlib.Path, member: str) -> str | None:
    with zipfile.ZipFile(path) as archive:
        text = archive.read(member).decode("utf-8", "replace")
    found = ALTERNATE_CONTENT.search(text)
    return found.group(0) if found else None


def derive(root: pathlib.Path) -> dict:
    pptx = census(root / PPTX_FIXTURE)
    docx = census(root / DOCX_FIXTURE)
    notes = census(root / PPTX_NOTES_FIXTURE)

    shape: dict = {
        "schema": "litchi.perf-baseline.marker-shape-derivation.v1",
        "derived_by": "docs/performance/results/change-0664/scripts/derive_marker_shape.py",
        "generator": "tools/perf-baseline/src/marker_shape.rs",
        "markup_compatibility_namespace": MCE_NAMESPACE,
        "markup_compatibility_namespace_bytes": len(MCE_NAMESPACE),
        "pptx": {"fixture": str(PPTX_FIXTURE), "kinds": {}},
        "docx": {"fixture": str(DOCX_FIXTURE), "kinds": {}},
    }

    for kind in (
        "pptx-slide",
        "pptx-slide-layout",
        "pptx-slide-master",
        "pptx-presentation",
    ):
        member = first_of_kind(pptx, kind)
        if member is None:
            continue
        shape["pptx"]["kinds"][kind] = {
            "example_part": member["part"],
            "declarations": member["declarations"],
            "declaration_count": len(member["declarations"]),
            "mc_ignorable": member["mc_ignorable"],
            "alternate_content_occurrences": member["alternate_content_occurrences"],
        }
    slide = first_of_kind(pptx, "pptx-slide")
    if slide is not None and slide["alternate_content_occurrences"]:
        shape["pptx"]["alternate_content_template"] = alternate_content_template(
            root / PPTX_FIXTURE, slide["part"]
        )
    for kind in ("pptx-notes-slide", "pptx-notes-master"):
        member = first_of_kind(notes, kind)
        if member is None:
            continue
        shape["pptx"]["kinds"][kind] = {
            "fixture": str(PPTX_NOTES_FIXTURE),
            "example_part": member["part"],
            "declarations": member["declarations"],
            "declaration_count": len(member["declarations"]),
            "mc_ignorable": member["mc_ignorable"],
            "alternate_content_occurrences": member["alternate_content_occurrences"],
        }

    for kind in (
        "docx-document",
        "docx-settings",
        "docx-styles",
        "docx-numbering",
        "docx-font-table",
        "docx-web-settings",
        "docx-footnotes",
        "docx-endnotes",
    ):
        member = first_of_kind(docx, kind)
        if member is None:
            continue
        shape["docx"]["kinds"][kind] = {
            "example_part": member["part"],
            "declarations": member["declarations"],
            "declaration_count": len(member["declarations"]),
            "mc_ignorable": member["mc_ignorable"],
            "alternate_content_occurrences": member["alternate_content_occurrences"],
        }

    shape["pptx"]["census"] = {
        key: pptx[key]
        for key in (
            "archive_bytes",
            "archive_sha256",
            "member_count",
            "uncompressed_bytes",
            "markup_compatibility_member_count",
            "markup_compatibility_bytes",
            "markup_compatibility_byte_share",
        )
    }
    shape["docx"]["census"] = {
        key: docx[key]
        for key in (
            "archive_bytes",
            "archive_sha256",
            "member_count",
            "uncompressed_bytes",
            "markup_compatibility_member_count",
            "markup_compatibility_bytes",
            "markup_compatibility_byte_share",
        )
    }
    return {"shape": shape, "pptx_census": pptx, "docx_census": docx, "notes_census": notes}


def verify(root: pathlib.Path, module: pathlib.Path) -> int:
    """Check the generator's constants against the fixtures they came from.

    The generator does not re-declare what the production writer already emits:
    on a PPTX part it adds ``p14``, ``p15`` and ``mc`` to the three the writer
    writes, so what must agree with the fixture is (a) every URI the generator
    adds and (b) the *total* binding count, which is the number the codec
    re-declares on every emitted start tag.  On a DOCX part the writer's root is
    replaced outright, so every URI must be present.
    """
    derived = derive(root)["shape"]
    source = (root / module).read_text(encoding="utf-8")
    failures: list[str] = []

    def require(text: str, why: str) -> None:
        if text not in source:
            failures.append(f"{why}: {text!r} is absent from {module}")

    require(MCE_NAMESPACE, "the markup-compatibility namespace")

    # (a) the URIs the PPTX generator adds, and (b) the fixture's binding count.
    pptx_kinds = derived["pptx"]["kinds"]
    added = {"p14", "p15", "mc"}
    counts = set()
    for kind, facts in pptx_kinds.items():
        counts.add(facts["declaration_count"])
        for prefix, uri in facts["declarations"]:
            if prefix in added:
                require(uri, f"{kind} declaration xmlns:{prefix}")
        if facts["mc_ignorable"]:
            require(facts["mc_ignorable"], f"{kind} mc:Ignorable")
    if len(counts) != 1:
        failures.append(
            f"the fixtures disagree about the PPTX root binding count: {sorted(counts)}"
        )
    else:
        expected = counts.pop()
        needle = f"const PPTX_ROOT_DECLARATION_COUNT: usize = {expected};"
        require(needle, "the PPTX root binding count")

    # Every DOCX declaration, because the generator replaces the whole root.
    for kind, facts in derived["docx"]["kinds"].items():
        for prefix, uri in facts["declarations"]:
            require(uri, f"{kind} declaration xmlns:{prefix}")
        if facts["mc_ignorable"]:
            require(facts["mc_ignorable"], f"{kind} mc:Ignorable")

    # The alternate-content wrapper, compared on its load-bearing fragments
    # because the generator writes it as escaped Rust string literals.
    if derived["pptx"].get("alternate_content_template"):
        for fragment in (
            "mc:AlternateContent",
            "mc:Choice",
            "Requires=",
            "p14:dur",
            "mc:Fallback",
            "p:transition",
        ):
            require(fragment, "the PPTX mc:AlternateContent template")

    for fixture in (PPTX_FIXTURE, DOCX_FIXTURE, PPTX_NOTES_FIXTURE):
        require(str(fixture), "the derivation fixture path")

    if failures:
        for failure in failures:
            print(f"FAIL {failure}", file=sys.stderr)
        return 1
    print(
        "ok: {module} agrees with {pptx} and {docx}".format(
            module=module, pptx=PPTX_FIXTURE.name, docx=DOCX_FIXTURE.name
        )
    )
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", default=".", help="repository root")
    sub = parser.add_subparsers(dest="mode", required=True)

    one = sub.add_parser("census")
    one.add_argument("archive")
    one.add_argument("--json")

    both = sub.add_parser("derive")
    both.add_argument("--out", help="directory for the census and shape JSON")

    check = sub.add_parser("verify")
    check.add_argument("--module", default="tools/perf-baseline/src/marker_shape.rs")

    arguments = parser.parse_args()
    root = pathlib.Path(arguments.root)

    if arguments.mode == "census":
        report = census(pathlib.Path(arguments.archive))
        text = json.dumps(report, indent=2, sort_keys=True) + "\n"
        if arguments.json:
            pathlib.Path(arguments.json).write_text(text, encoding="utf-8")
        else:
            sys.stdout.write(text)
        return 0

    if arguments.mode == "derive":
        derived = derive(root)
        if arguments.out:
            out = pathlib.Path(arguments.out)
            out.mkdir(parents=True, exist_ok=True)
            for name, payload in (
                ("marker-shape.json", derived["shape"]),
                ("pptx-real-deck-census.json", derived["pptx_census"]),
                ("docx-real-file-census.json", derived["docx_census"]),
                ("pptx-notes-fixture-census.json", derived["notes_census"]),
            ):
                (out / name).write_text(
                    json.dumps(payload, indent=2, sort_keys=True) + "\n",
                    encoding="utf-8",
                )
            print(f"wrote 4 files to {out}")
        else:
            sys.stdout.write(json.dumps(derived["shape"], indent=2, sort_keys=True) + "\n")
        return 0

    return verify(root, pathlib.Path(arguments.module))


if __name__ == "__main__":
    raise SystemExit(main())
