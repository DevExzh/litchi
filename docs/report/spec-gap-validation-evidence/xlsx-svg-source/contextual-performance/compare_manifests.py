#!/usr/bin/env python3
"""Compare before/after source manifests while admitting the two frozen host files."""

from __future__ import annotations

import argparse
import hashlib
from pathlib import Path


EXPECTED_DELTA_PATHS = {
    "crates/litchi-xlsx/src/drawing/source.rs",
    "crates/litchi-xlsx/src/drawing/source_tests.rs",
}


def normalize(path: Path) -> list[str]:
    output = []
    for line in path.read_text().splitlines():
        if line.startswith("metadata_sha256="):
            continue
        fields = line.split("\t")
        if line.startswith("package=litchi-xlsx\t"):
            fields[-1] = "<expected-litchi-xlsx-tree-delta>"
            line = "\t".join(fields)
        elif line.startswith("file=litchi-xlsx\t") and len(fields) == 5:
            shown = fields[3]
            if shown in EXPECTED_DELTA_PATHS:
                fields[-1] = "<expected-frozen-host-delta>"
                line = "\t".join(fields)
        output.append(line)
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--before", type=Path, required=True)
    parser.add_argument("--after", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    before = normalize(args.before)
    after = normalize(args.after)
    if before != after:
        for index, (left, right) in enumerate(zip(before, after)):
            if left != right:
                raise SystemExit(f"normalized manifest differs at line {index + 1}:\n{left}\n{right}")
        raise SystemExit(f"normalized manifest lengths differ: {len(before)} vs {len(after)}")
    digest = hashlib.sha256("\n".join(before).encode()).hexdigest()
    args.output.write_text(
        "format=xlsx-svg-contextual-normalized-manifest-v1\n"
        f"before={args.before}\n"
        f"after={args.after}\n"
        f"normalized_sha256={digest}\n"
        "admitted_source_deltas=crates/litchi-xlsx/src/drawing/source.rs,crates/litchi-xlsx/src/drawing/source_tests.rs\n"
    )


if __name__ == "__main__":
    main()
