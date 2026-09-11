#!/usr/bin/env python3
"""Emit one compact SHA-256 identity row per deterministic corpus lane."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from summarize import LANES


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    lines = ["lane\tinput_bytes\tpictures\tnamespace_declarations\tsha256"]
    for lane in LANES:
        paths = sorted(args.results.glob(f"{lane}-p*.json"))
        if not paths:
            raise SystemExit(f"missing receipt for {lane}")
        value = json.loads(paths[0].read_text())
        lines.append(
            "\t".join(
                (
                    lane,
                    str(value["input_bytes"]),
                    str(value["picture_count"]),
                    str(value["namespace_declarations"]),
                    str(value["input_hash_sha256"]),
                )
            )
        )
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
