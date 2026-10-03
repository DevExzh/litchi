#!/usr/bin/env python3
"""Compare the sealed f790 baseline and frozen 0890 candidate reports."""

from __future__ import annotations

import gzip
import hashlib
import json
from pathlib import Path


PACKET = Path(__file__).resolve().parent
BASELINE = PACKET / "attempts/baseline-f790c2228c-retry1"
CANDIDATE = PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1"
OUTPUT = PACKET / "cross-version-diagnostic-comparison.json"


def descriptor(path: Path) -> dict[str, object]:
    raw = path.read_bytes()
    plain = gzip.decompress(raw) if path.suffix == ".gz" else raw
    return {
        "path": str(path.relative_to(PACKET)),
        "bytes": len(raw),
        "sha256": hashlib.sha256(raw).hexdigest(),
        "decompressed_bytes": len(plain),
        "decompressed_sha256": hashlib.sha256(plain).hexdigest(),
    }


def main() -> int:
    pairs = (
        ("first", BASELINE / "differential-first.json.gz", CANDIDATE / "differential-first.json.gz"),
        ("second", BASELINE / "differential-second.json.gz", CANDIDATE / "differential-second.json.gz"),
    )
    records = {}
    for name, baseline, candidate in pairs:
        baseline_raw = gzip.decompress(baseline.read_bytes())
        candidate_raw = gzip.decompress(candidate.read_bytes())
        if baseline_raw != candidate_raw:
            raise SystemExit(f"baseline/candidate report differs: {name}")
        if json.loads(baseline_raw) != json.loads(candidate_raw):
            raise SystemExit(f"baseline/candidate parsed report differs: {name}")
        records[name] = {
            "byte_identical": True,
            "baseline": descriptor(baseline),
            "candidate": descriptor(candidate),
        }
    baseline_identity = {key: value for key, value in records["first"]["baseline"].items() if key != "path"}
    baseline_second_identity = {key: value for key, value in records["second"]["baseline"].items() if key != "path"}
    candidate_identity = {key: value for key, value in records["first"]["candidate"].items() if key != "path"}
    candidate_second_identity = {key: value for key, value in records["second"]["candidate"].items() if key != "path"}
    if baseline_identity != baseline_second_identity:
        raise SystemExit("baseline repeated report descriptors differ")
    if candidate_identity != candidate_second_identity:
        raise SystemExit("candidate repeated report descriptors differ")
    result = {
        "schema": "litchi.mce-0890-cross-version-diagnostic-comparison.v1",
        "status": "pass",
        "baseline_attempt": "attempts/baseline-f790c2228c-retry1",
        "candidate_attempt": "attempts/candidate-diagnostic-f790c2228c-v2-retry1",
        "generated": 20000,
        "aliasing": 20000,
        "packages": 3,
        "selected_members": 67,
        "repeat_byte_identical": True,
        "cross_version_byte_identical": True,
        "reports": records,
    }
    OUTPUT.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
