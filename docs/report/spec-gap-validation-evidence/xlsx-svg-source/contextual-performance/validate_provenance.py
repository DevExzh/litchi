#!/usr/bin/env python3
"""Validate the build hash and provenance while the fresh binary still exists."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--expected-commit", required=True)
    parser.add_argument("--expected-label", required=True)
    parser.add_argument("--expected-source-hash", required=True)
    parser.add_argument("--expected-test-hash", required=True)
    parser.add_argument("--expected-codec-hash", required=True)
    args = parser.parse_args()

    if not args.binary.is_file() or not args.binary.stat().st_mode & 0o111:
        raise SystemExit(f"fresh profile binary is missing or not executable: {args.binary}")
    recorded = args.results.joinpath("binary.sha256").read_text().split()
    if len(recorded) != 2 or recorded[1] != str(args.binary):
        raise SystemExit("binary.sha256 does not identify the built binary")
    actual = digest(args.binary)
    if recorded[0] != actual or len(actual) != 64:
        raise SystemExit("binary.sha256 does not match the built binary")

    lines = args.results.joinpath("build-provenance.txt").read_text().splitlines()
    values = {}
    source_hashes = {}
    in_source_hashes = False
    for line in lines:
        if line.startswith("source_sha256:"):
            in_source_hashes = True
            continue
        if line.startswith("source_api="):
            continue
        if line.startswith("snapshot_label="):
            values["snapshot_label"] = line.split("=", 1)[1]
        elif line.startswith("git_head="):
            values["git_head"] = line.split("=", 1)[1]
        elif line.startswith("binary="):
            values["binary"] = line.split("=", 1)[1]
        elif in_source_hashes and len(line.split()) == 2 and len(line.split()[0]) == 64:
            source_hashes[line.rsplit(None, 1)[1]] = line.split()[0]
    if values.get("snapshot_label") != args.expected_label:
        raise SystemExit("build provenance snapshot label mismatch")
    if values.get("git_head") != args.expected_commit:
        raise SystemExit("build provenance commit mismatch")
    if values.get("binary") != str(args.binary):
        raise SystemExit("build provenance binary path mismatch")
    expected_hashes = {
        "source.rs": args.expected_source_hash,
        "source_tests.rs": args.expected_test_hash,
        "svg_blip.rs": args.expected_codec_hash,
    }
    for suffix, expected in expected_hashes.items():
        matches = [value for path, value in source_hashes.items() if path.endswith(suffix)]
        if matches != [expected]:
            raise SystemExit(f"build provenance source hash mismatch for {suffix}")
    if len(source_hashes) != 3 or any(len(value) != 64 for value in source_hashes.values()):
        raise SystemExit("build provenance source hash set is incomplete")
    args.results.joinpath("provenance-verification.json").write_text(
        json.dumps(
            {
                "passed": True,
                "snapshot_label": args.expected_label,
                "expected_commit": args.expected_commit,
                "binary_sha256": actual,
                "source_sha256": {
                    "source.rs": args.expected_source_hash,
                    "source_tests.rs": args.expected_test_hash,
                    "svg_blip.rs": args.expected_codec_hash,
                },
                "fresh_binary_validated_before_target_cleanup": True,
            },
            indent=2,
        )
        + "\n"
    )


if __name__ == "__main__":
    main()
