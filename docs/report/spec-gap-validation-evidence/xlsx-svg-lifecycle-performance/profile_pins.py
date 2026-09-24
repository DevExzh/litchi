#!/usr/bin/env python3
"""Approved source identity for the XLSX SVG lifecycle profile.

The production owner was reviewed and committed at ``ac288a303``.  The
profile runner still requires a clean committed checkout and compares the
working bytes with this commit's Git blobs; a dirty hash-only workspace is not
an evidence source.
"""

from __future__ import annotations

import argparse


SOURCE_PIN = "ac288a303264f9ea0bb4081baa44031bee5b79a7"

PREREQUISITE_SOURCES = (
    "crates/litchi-opc/src/package/relationships.rs",
    "crates/litchi-opc/src/phys_pkg.rs",
    "crates/litchi-xlsx/src/drawing/source.rs",
    # Shared SVG blip contextual behavior; the approved blob is
    # e8a1e480d249b834f286e0deec083f8044257faf3b95905fb5b771e83f71605a.
    "crates/litchi-drawingml/src/svg_blip.rs",
)

# The complete XLSX lifecycle feature closure from SOURCE_PIN.  Keep every
# changed path in the production commit bound: a public adapter can compile
# while a changed module, codec, semantic wiring file, or focused source test
# silently comes from a different checkout.
FEATURE_SOURCES = (
    "crates/litchi-xlsx/docs/FEATURE_MATRIX.md",
    "crates/litchi-xlsx/src/drawing/source_tests.rs",
    "crates/litchi-xlsx/src/edit.rs",
    "crates/litchi-xlsx/src/workbook/edit/codec.rs",
    "crates/litchi-xlsx/src/workbook/edit/mod.rs",
    "crates/litchi-xlsx/src/workbook/edit/model.rs",
    "crates/litchi-xlsx/src/workbook/edit/package.rs",
    "crates/litchi-xlsx/src/workbook/edit/semantic/mod.rs",
    "crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs",
    "crates/litchi-xlsx/src/workbook/edit/semantic/worksheet.rs",
    "crates/litchi-xlsx/src/workbook/edit/svg.rs",
    "crates/litchi-xlsx/src/workbook/edit/svg_lifecycle.rs",
    "crates/litchi-xlsx/src/workbook/edit/svg_lifecycle/relationship_ids.rs",
    "crates/litchi-xlsx/src/workbook/edit/svg_lifecycle/topology.rs",
    "crates/litchi-xlsx/tests/drawing_svg_lifecycle.rs",
)

assert len(FEATURE_SOURCES) == 15
PINNED_SOURCES = PREREQUISITE_SOURCES + FEATURE_SOURCES


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--paths", action="store_true")
    args = parser.parse_args()
    if args.paths:
        for path in PINNED_SOURCES:
            print(path)
    else:
        print(f"source_pin={SOURCE_PIN}")


if __name__ == "__main__":
    main()
