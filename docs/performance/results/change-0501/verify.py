#!/usr/bin/env python3
"""Replay 0501 source, executable, report, and profile custody checks."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from typing import Any

from custody import HERE, REPO, canonical_json, sha_file, source_identity, source_snapshot


def load_module(name: str, path: Path) -> Any:
    import importlib.util

    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def check_freeze(phase: str) -> dict[str, Any]:
    path = HERE / f"{phase}-freeze.json"
    freeze = json.loads(path.read_text())
    if freeze.get("change") != 501 or freeze.get("phase") != phase:
        raise ValueError(f"{path}: identity mismatch")
    binary = Path(freeze["binary"]["path"])
    if not binary.is_file() or sha_file(binary) != freeze["binary"]["sha256"] or binary.stat().st_size != freeze["binary"]["bytes"]:
        raise ValueError(f"{path}: binary identity mismatch")
    if sha_file(HERE / "protocol.json") != freeze["protocol_sha256"]:
        raise ValueError(f"{path}: protocol changed after freeze")
    if source_identity(source_snapshot()) != freeze["source_manifest"]:
        raise ValueError(f"{path}: current source differs from freeze")
    for name, expected in freeze["bound_files"].items():
        target = HERE / name if name.endswith((".py", ".json")) else REPO / name
        if sha_file(target) != expected:
            raise ValueError(f"{path}: bound file changed: {name}")
    return freeze


def check_profiles(phase: str) -> dict[str, Any] | None:
    path = HERE / "profiles" / phase / "summary.json"
    if not path.is_file():
        return None
    summary = json.loads(path.read_text())
    if summary.get("change") != 501 or summary.get("phase") != phase:
        raise ValueError(f"{path}: profile summary identity")
    for receipt_row in summary.get("profiles", []):
        receipt_path = HERE / "profiles" / phase / f"{receipt_row['lane']['corpus']}-{receipt_row['lane']['provider_label']}.receipt.json"
        receipt = json.loads(receipt_path.read_text())
        if receipt.get("status") != "pass" or receipt.get("sha_hotness_detected") is not True:
            raise ValueError(f"{receipt_path}: SHA hotness requirement failed")
        raw = Path(receipt["raw_perf_path"])
        if not raw.is_relative_to(Path("/tmp/litchi-goal-0501")) or not raw.is_file():
            raise ValueError(f"{receipt_path}: raw perf data escaped owned tmp root")
        if receipt.get("raw_perf_bytes") != raw.stat().st_size or sha_file(raw) != receipt.get("raw_perf_sha256"):
            raise ValueError(f"{receipt_path}: raw perf identity")
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--phase", choices=("before", "after"), action="append")
    parser.add_argument("--include-range", action="store_true")
    parser.add_argument("--require-comparison", action="store_true")
    args = parser.parse_args()
    protocol = json.loads((HERE / "protocol.json").read_text())
    if protocol.get("change") != 501:
        raise ValueError("protocol identity mismatch")
    phases = args.phase or ["before", "after"]
    verifier = load_module("verify_report_bundle", HERE / "verify-report.py")
    phase_counts = {}
    for phase in phases:
        check_freeze(phase)
        folder = HERE / phase
        expected = protocol["core_order"]
        if args.include_range:
            expected = expected + protocol["optional_range"]["order"]
        expected_names = {f"{lane['corpus']}-{lane['provider_label']}-{lane['repeat'].lower()}" for lane in expected}
        receipts = {path.name.removesuffix(".receipt.json") for path in folder.glob("*.receipt.json")}
        if receipts != expected_names:
            raise ValueError(f"{phase}: receipt inventory differs from protocol")
        # load_phase rechecks each report, artifact hash, cleanup receipt, and
        # source/output identity.  It is imported only after freeze custody.
        compare = load_module("compare_bundle", HERE / "compare.py")
        records = compare.load_phase(phase, protocol)
        if set(records) != expected_names:
            raise ValueError(f"{phase}: report inventory differs from protocol")
        check_profiles(phase)
        phase_counts[phase] = len(records)
    if set(phases) == {"before", "after"}:
        comparison_path = HERE / "comparison.json"
        if args.require_comparison and not comparison_path.is_file():
            raise ValueError("comparison.json is missing")
        if comparison_path.is_file():
            comparison = json.loads(comparison_path.read_text())
            if comparison.get("change") != 501 or comparison.get("before_reports") != comparison.get("after_reports"):
                raise ValueError("comparison identity/count mismatch")
            subprocess.run([sys.executable, "-B", str(HERE / "compare.py"), *( ["--include-range"] if args.include_range else [])], cwd=REPO, check=True)
    result = {"status": "pass", "phases": phase_counts}
    (HERE / "verification.json").write_bytes(canonical_json(result))
    print(json.dumps(result, sort_keys=True))


if __name__ == "__main__":
    main()
