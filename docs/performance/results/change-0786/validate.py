"""Fail-closed offline validator for the 0786 execution-scaling packet."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

import analyze


PACKET = analyze.PACKET


def require(condition: bool, message: str) -> None:
    analyze.require(condition, message)


def validate(*, require_final_seal: bool = False) -> dict[str, Any]:
    result = analyze.analyze(check=True)
    require(result["schema"] == analyze.ANALYSIS_SCHEMA, "analysis schema changed")
    require(result["plan_schema"] == "litchi.execution-scaling.0786.v1", "plan schema changed")
    require(result["report_schema"] == analyze.REPORT_SCHEMA, "probe schema changed")
    require(result["counts"] == {
        "reports": 1080,
        "samples": 22200,
        "native_reports": 720,
        "native_samples": 21600,
        "observer_reports": 240,
        "observer_samples": 480,
        "qualification_reports": 120,
        "qualification_samples": 120,
    }, "aggregate report/sample cardinality changed")
    require(result["quality"]["gates"] == 6, "quality gate cardinality changed")
    require(result["source"]["production_file_count"] == 9196
            and result["source"]["production_byte_identical_to_base"] is True
            and result["source"]["tool_source_frozen"] is True,
            "source custody changed")
    architecture = result["architecture_inputs"]
    require(architecture["count"] == 35
            and architecture["revision"] == analyze.origin()["base"]
            and architecture["live_files_match"] is True
            and architecture["origin_blob_hashes_match"] is True,
            "architecture input custody changed")
    require(result["bootstrap"] == {
        "seed": analyze.BOOTSTRAP_SEED,
        "resamples": analyze.BOOTSTRAP_RESAMPLES,
        "confidence": analyze.BOOTSTRAP_CONFIDENCE,
        "statistic": "median",
    }, "bootstrap contract changed")
    require(len(result["scaling"]) == 120, "scaling row cardinality changed")
    raw_audit = result["raw_audit"]
    require(raw_audit["reports"] == 1080 and raw_audit["samples"] == 22200
            and raw_audit["independently_reconstructed_payloads"] == 96
            and raw_audit["rows"] == 120
            and raw_audit["paired_curves_match"] is True,
            "independent raw audit changed")
    require(result["observer"]["timings_pooled_with_native"] is False
            and result["observer"]["reports"] == 360,
            "observer separation changed")
    verification = result["verification"]
    for key in (
        "plan_and_order_checked", "receipts_checked", "report_schema_checked",
        "output_parity_checked", "resource_limits_checked",
        "final_permits_released_checked", "production_base_git_blobs_checked",
        "architecture_inputs_checked", "tool_source_checked",
        "quality_commands_checked_exactly", "observer_timing_separated",
        "amdahl_descriptive_only",
    ):
        require(verification.get(key) is True, f"verification marker missing: {key}")
    for row in result["reports"]:
        require(row["verification_ok"] is True and row["resource_ok"] is True,
                "report correctness witness changed")
        require(row["report"]["sha256"] == analyze.sha256(
            analyze.resolve_path(row["report"]["path"])),
                "retained report digest changed")
    if require_final_seal:
        seal = PACKET / "seal.json"
        final_seal = PACKET / "final-seal.json"
        require(seal.is_file() and not seal.is_symlink(), "final seal is missing")
        require(not final_seal.exists(), "ambiguous second final seal")
        value = analyze.read_json(seal)
        require(value.get("schema") == "litchi.execution-scaling-seal.v1",
                "final seal schema changed")
        files = value.get("files")
        require(isinstance(files, dict) and files, "final seal files are missing")
        actual = {}
        for path in PACKET.rglob("*"):
            if path == seal or "__pycache__" in path.parts:
                continue
            require(not path.is_symlink(), f"symlink in sealed packet: {path}")
            if path.is_file():
                actual[str(path.relative_to(PACKET))] = analyze.sha256(path)
        require(actual == files, "final seal file set or payload hash mismatch")
        require(value.get("payload_count") == len(actual), "final seal count changed")
    return {
        "reports": result["counts"]["reports"],
        "samples": result["counts"]["samples"],
        "native_reports": result["counts"]["native_reports"],
        "observer_reports": result["counts"]["observer_reports"],
        "qualification_reports": result["counts"]["qualification_reports"],
        "scaling_rows": len(result["scaling"]),
        "raw_audit_rows": raw_audit["rows"],
        "quality_gates": result["quality"]["gates"],
        "seal_checked": require_final_seal,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--require-final-seal", action="store_true")
    args = parser.parse_args(argv)
    try:
        print(json.dumps(validate(require_final_seal=args.require_final_seal),
                         indent=2, sort_keys=True))
    except analyze.ReplayError as error:
        print(f"0786 validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
