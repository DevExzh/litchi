#!/usr/bin/env python3
"""Write a compact hash receipt for the final XLSB projection evidence."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path


HERE = Path(__file__).resolve().parent
RUN = HERE / "runs/final-release"


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def main() -> None:
    raw_paths = sorted(RUN.glob("*.json"))
    raw_manifest = "".join(
        f"{path.relative_to(HERE)}  {digest(path)}\n" for path in raw_paths
    ).encode()
    raw_manifest_path = HERE / "raw-json-sha256.txt"
    raw_manifest_path.write_bytes(raw_manifest)

    before = json.loads((HERE / "source-provenance-before.json").read_text())
    after = json.loads((HERE / "source-provenance-after.json").read_text())
    summary = json.loads((HERE / "matrix-summary.json").read_text())
    receipt = {
        "schema": "litchi-xlsb-drawing-projection-performance-receipt-v1",
        "fixture": after["fixture"],
        "binary": after["binary"],
        "toolchain": {
            "rustc": after["rustc"],
            "cargo": after["cargo"],
            "rustup_toolchain": after["rustup_toolchain"],
        },
        "locked_inputs": {
            "Cargo.lock": before["files"]["Cargo.lock"],
            "tools/perf-baseline/Cargo.lock": before["files"]["tools/perf-baseline/Cargo.lock"],
        },
        "source_stability": {
            "git_head_before": before["git_head"],
            "git_head_after": after["git_head"],
            "files_equal_before_after": before["files"] == after["files"],
            "fixture_equal_before_after": before["fixture"] == after["fixture"],
            "before_record_sha256": digest(HERE / "source-provenance-before.json"),
            "after_record_sha256": digest(HERE / "source-provenance-after.json"),
        },
        "matrix": {
            "supported_backend_cases": summary["contract"]["supported_backend_cases"],
            "processes_per_backend_case": summary["contract"]["processes_per_backend_case"],
            "warmup_per_process": summary["contract"]["warmup_per_process"],
            "samples_per_process": summary["contract"]["samples_per_process"],
            "raw_json_reports": len(raw_paths),
            "raw_samples": len(raw_paths) * summary["contract"]["samples_per_process"],
            "raw_json_manifest": str(raw_manifest_path.relative_to(HERE)),
            "raw_json_manifest_sha256": digest(raw_manifest_path),
            "summary": str((HERE / "matrix-summary.json").relative_to(HERE)),
            "summary_sha256": digest(HERE / "matrix-summary.json"),
            "cross_lane_comparison": "No equivalent-work speedup inferred; backend API/validation scopes differ.",
        },
        "reproduction_files": {
            relative: digest(HERE / relative)
            for relative in (
                "README.md",
                "record-provenance.py",
                "run-release-matrix.sh",
                "verify-matrix.py",
                "write-receipt.py",
                "release-build-command.txt",
                "runs/final-release/commands.txt",
                "runs/final-release/completed.txt",
            )
        },
    }
    (HERE / "performance-receipt.json").write_text(
        json.dumps(receipt, indent=2, sort_keys=True) + "\n"
    )
    print(json.dumps(receipt, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
