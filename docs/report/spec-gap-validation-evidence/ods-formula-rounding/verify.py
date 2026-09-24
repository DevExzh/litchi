#!/usr/bin/env python3
"""Fail-closed verification for the retained ODS rounding capture."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
RESULTS = HERE / "results"
CASES = {
    "scalar-control",
    "scalar-int",
    "scalar-floor",
    "scalar-round",
    "scalar-rounddown",
    "scalar-ceiling",
    "scalar-mround",
    "scalar-roundup",
    "scalar-trunc",
    "array-control-4x4",
    "array-int-4x4",
    "array-floor-4x4",
    "array-round-4x4",
    "array-rounddown-4x4",
    "array-ceiling-4x4",
    "array-mround-4x4",
    "array-roundup-4x4",
    "array-trunc-4x4",
    "array-roundup-16x16",
}
PHASES = {"evaluate", "parse-evaluate"}
SOURCE_FILES = [
    "crates/litchi-ods/src/codec/formula/evaluation.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/rounding.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value.rs",
    "crates/litchi-ods/Cargo.toml",
    "crates/litchi-core/Cargo.toml",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/harness/Cargo.toml",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/harness/Cargo.lock",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/harness/src/main.rs",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/run_profile.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/verify.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/summarize.py",
]


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    source = json.loads((RESULTS / "source-manifest.json").read_text())
    assert source["unchanged"]
    assert source["source_hashes_unchanged"]
    assert source["before"]["source_sha256"] == source["after"]["source_sha256"]
    assert source["before"]["source_sha256"] == {
        name: digest(ROOT / name) for name in SOURCE_FILES
    }
    assert source["binary_sha256"]
    cleanup = json.loads((RESULTS / "target-cleanup.json").read_text())
    assert cleanup["removed"]

    records = [
        json.loads(line)
        for line in (RESULTS / "measurements.jsonl").read_text().splitlines()
        if line.strip()
    ]
    assert len(records) == len(CASES) * len(PHASES)
    assert {(row["case"], row["phase"]) for row in records} == {
        (case, phase) for case in CASES for phase in PHASES
    }
    for row in records:
        assert row["rss_kib"] > 0
        assert row["elapsed_ns_p50"] > 0
        assert row["elapsed_ns_p95"] >= row["elapsed_ns_p50"]
        assert row["elapsed_ns_p99"] >= row["elapsed_ns_p95"]
        assert row["repeat"] > 0 and row["iterations"] == 20
        assert row["checksum_p50"] != 0
        assert row["source_git_head"] == source["before"]["git_head"]
        assert len(row["samples"]) == row["iterations"]
        assert row["samples"]
        for sample in row["samples"]:
            assert sample["elapsed_ns"] > 0
            assert sample["work"] > 0
            assert sample["alloc_calls"] >= 0
            assert sample["dealloc_calls"] >= 0
            assert sample["requested_bytes"] >= 0
            assert sample["released_bytes"] >= 0
            assert sample["peak_live_delta"] >= 0
            assert sample["memory_retained"] >= 0
        stem = f'{row["case"]}.{row["phase"]}'
        raw = json.loads((RESULTS / f"{stem}.stdout.json").read_text())
        assert raw["case"] == row["case"] and raw["phase"] == row["phase"]
        assert (RESULTS / f"{stem}.stderr.log").read_text() == ""
        time_text = (RESULTS / f"{stem}.time.txt").read_text()
        assert re.search(r"Maximum resident set size \(kbytes\):\s*\d+", time_text)
        assert raw["rss_kib"] is None

    assert {row["case"] for row in records if row["phase"] == "evaluate"} == CASES
    print(
        json.dumps(
            {
                "verified": True,
                "cases": len(CASES),
                "phases": len(PHASES),
                "rows": len(records),
                "source_git_head": source["before"]["git_head"],
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
