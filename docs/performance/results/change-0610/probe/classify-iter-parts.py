#!/usr/bin/env python3
"""Classify every `iter_parts()` call site by the migration shape a lazy
first-access decode (C2') would need.

C2' keeps `Part::blob(&self) -> &[u8]` infallible, because every route that
hands out a `&dyn Part` except `iter_parts()` is already fallible and can force
the decode before returning. `iter_parts()` is therefore the whole migration
surface. This script sizes it and splits it by what each site does with the
iterator.

Two scopes are reported for every site, because neither alone is honest:

  narrow   a fixed window of N lines starting at the call. A site whose
           payload read is further away than N lines is missed, so the `bytes`
           count from this scope is a LOWER bound.
  function the whole enclosing function body, found by scanning back to the
           nearest `fn` at a lower indent and brace-matching forward. A
           function that iterates names and separately reads a payload
           elsewhere is attributed to `bytes`, so this scope's `bytes` count is
           an UPPER bound.

The truth is between them; sites where the two scopes disagree are listed so
they can be read by hand.

Categories:
  bytes      the scope reads payload bytes (`blob()`, `blob_arc()`,
             `rel_ref_count`, `set_blob`) -> needs `try_iter_parts()` and `?`
  metadata   only `partname()`, `content_type()`, `rels()` and similar
  count      only counts or collects the iterator
  unclear    neither pattern matched

Receiver: a site whose enclosing function signature or window names
`SourceBackedPackage` or `PartView` is tagged `source-backed` and is OUT OF
SCOPE — `SourceBackedPackage::iter_parts` yields `PartView`, which has no
`blob()` at all and whose `data()` is already fallible.

Inline `#[cfg(test)]` modules are excluded, and only an inline module counts:
a `#[cfg(test)] mod name;` *declaration* leaves the rest of the file
production, which a naive "first `#[cfg(test)]` line is the cut" rule gets
wrong (`crates/litchi-opc/src/package.rs:1238` is exactly that case).

Usage: classify.py <tree-with-crates> [window-lines]
"""

from __future__ import annotations

import collections
import pathlib
import re
import sys

CALL = re.compile(r"\.iter_parts\(\)")
BYTES = re.compile(r"\.blob\(\)|\.blob_arc\(\)|rel_ref_count|set_blob|\bblob\b")
METADATA = re.compile(r"partname\(\)|content_type\(\)|\.rels\(\)|\.uri\b|as_str\(\)")
COUNT = re.compile(r"\.count\(\)|\.collect\(|\.len\(\)")
SOURCE_BACKED = re.compile(r"SourceBackedPackage|PartView|\.data\(\)")
CFGTEST = re.compile(r"^\s*#\[cfg\(test\)\]\s*$")
INLINE_MOD = re.compile(r"^\s*(pub\s+)?mod\s+[A-Za-z0-9_]+\s*\{")
FN = re.compile(r"^(\s*)(pub(\([^)]*\))?\s+)?(const\s+|async\s+|unsafe\s+|extern\s+\"[^\"]*\"\s+)*fn\s")


def inline_test_spans(lines: list[str]) -> list[tuple[int, int]]:
    """Line ranges (inclusive, 0-based) covered by an inline `#[cfg(test)] mod X { .. }`.

    A `#[cfg(test)] mod X;` declaration covers nothing: the rest of the file is
    production code.
    """
    spans = []
    index = 0
    while index < len(lines):
        if CFGTEST.match(lines[index]):
            probe = index + 1
            while probe < len(lines) and lines[probe].strip().startswith("#["):
                probe += 1
            if probe < len(lines) and INLINE_MOD.match(lines[probe]):
                depth = 0
                end = probe
                for cursor in range(probe, len(lines)):
                    depth += lines[cursor].count("{") - lines[cursor].count("}")
                    if depth <= 0 and cursor > probe:
                        end = cursor
                        break
                    end = cursor
                spans.append((index, end))
                index = end + 1
                continue
        index += 1
    return spans


def enclosing_function(lines: list[str], site: int) -> str:
    """The body of the function containing `site`, brace-matched."""
    start = None
    for cursor in range(site, -1, -1):
        if FN.match(lines[cursor]):
            start = cursor
            break
    if start is None:
        return "\n".join(lines[max(0, site - 4) : site + 40])
    depth = 0
    seen = False
    for cursor in range(start, len(lines)):
        depth += lines[cursor].count("{") - lines[cursor].count("}")
        if lines[cursor].count("{"):
            seen = True
        if seen and depth <= 0:
            return "\n".join(lines[start : cursor + 1])
    return "\n".join(lines[start:])


def classify(scope: str) -> str:
    if BYTES.search(scope):
        return "bytes"
    if METADATA.search(scope):
        return "metadata"
    if COUNT.search(scope):
        return "count"
    return "unclear"


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    root = pathlib.Path(sys.argv[1])
    span = int(sys.argv[2]) if len(sys.argv) > 2 else 8

    narrow: dict[tuple[str, str], int] = collections.Counter()
    wide: dict[tuple[str, str], int] = collections.Counter()
    source_backed: dict[str, int] = collections.Counter()
    disagreements: list[str] = []
    sb_sites: list[str] = []
    scanned = skipped = 0

    for path in sorted((root / "crates").glob("*/src/**/*.rs")):
        crate = path.relative_to(root / "crates").parts[0]
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
        spans = inline_test_spans(lines)
        for index, line in enumerate(lines):
            if not CALL.search(line):
                continue
            scanned += 1
            if any(low <= index <= high for low, high in spans):
                skipped += 1
                continue
            body = enclosing_function(lines, index)
            window = "\n".join(lines[index : index + span])
            where = f"{path.relative_to(root)}:{index + 1}"
            if SOURCE_BACKED.search(body):
                source_backed[crate] += 1
                sb_sites.append(where)
                continue
            near, far = classify(window), classify(body)
            narrow[(crate, near)] += 1
            wide[(crate, far)] += 1
            if near != far:
                disagreements.append(f"{where}  narrow={near} function={far}")

    kinds = ["bytes", "metadata", "count", "unclear"]
    crates = sorted({c for c, _ in narrow} | {c for c, _ in wide} | set(source_backed))

    def table(name: str, tally: dict[tuple[str, str], int]) -> None:
        print(f"\n## {name}")
        print(f"  {'crate':<22}" + "".join(f"{k:>10}" for k in kinds) + f"{'total':>10}")
        for crate in crates:
            row = [tally[(crate, k)] for k in kinds]
            if not sum(row):
                continue
            print(f"  {crate:<22}" + "".join(f"{v:>10}" for v in row) + f"{sum(row):>10}")
        totals = [sum(tally[(c, k)] for c in crates) for k in kinds]
        print(f"  {'TOTAL':<22}" + "".join(f"{v:>10}" for v in totals) + f"{sum(totals):>10}")

    print(f"# iter_parts() sites under crates/*/src, narrow window = {span} lines")
    print(f"# scanned {scanned}; {skipped} inside an inline #[cfg(test)] mod; "
          f"{sum(source_backed.values())} on SourceBackedPackage (out of scope); "
          f"{scanned - skipped - sum(source_backed.values())} classified")
    table("narrow window (bytes = LOWER bound)", narrow)
    table("enclosing function (bytes = UPPER bound)", wide)
    print("\n## source-backed, out of scope")
    for crate in crates:
        if source_backed[crate]:
            print(f"  {crate:<22}{source_backed[crate]:>10}")
    print(f"  {'TOTAL':<22}{sum(source_backed.values()):>10}")
    for site in sb_sites:
        print(f"    {site}")
    print(f"\n## sites where the two scopes disagree ({len(disagreements)})")
    for site in disagreements:
        print(f"  {site}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
