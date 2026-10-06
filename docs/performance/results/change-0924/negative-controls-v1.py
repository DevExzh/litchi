#!/usr/bin/env python3
"""Run fail-closed tests against copies of the retained 0923 packet.

The semantic fixtures rebind only the copied raw-report index and the copied
validator's raw-index digest.  This keeps the production validator's custody
binding intact while proving that its structural checks are reached after a
test-only evidence rebind.
"""

from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
from pathlib import Path
from typing import Any, Callable


ROOT = Path("/home/zhuhe/code/litchi")
SCRATCH = Path("/tmp/litchi-performance-0924")
FINAL_SCRIPT = SCRATCH / "phase-share-v2.py"
FIXTURE_ROOT = SCRATCH / "negative-fixtures"
CASE = "generic_read_at-pair-01-leg-2-candidate"
REPORT_REL = Path("docs/performance/results/change-0923/capture-final/packet/children") / CASE / "report.json"
RAW_INDEX_REL = Path("docs/performance/results/change-0923/root-raw-sample-summary.json")


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def copy_packet(destination: Path) -> None:
    if destination.exists():
        shutil.rmtree(destination)
    for relative in (
        Path("docs/performance/results/change-0923/capture-final/packet/comparison-manifest-v2.json"),
        Path("docs/performance/results/change-0923/comparison-result-summary-v1.json"),
        RAW_INDEX_REL,
    ):
        target = destination / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / relative, target)
    source_reports = ROOT / "docs/performance/results/change-0923/capture-final/packet/children"
    for source in source_reports.glob("*-candidate/report.json"):
        target = destination / "docs/performance/results/change-0923/capture-final/packet/children" / source.parent.name / "report.json"
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)


def report_path(destination: Path) -> Path:
    return destination / REPORT_REL


def load_report(destination: Path) -> dict[str, Any]:
    path = report_path(destination)
    return json.loads(path.read_text())


def save_report(destination: Path, report: dict[str, Any]) -> None:
    report_path(destination).write_text(json.dumps(report, indent=2, sort_keys=True) + "\n")


def rebind_test_only_raw_index(destination: Path) -> tuple[Path, str]:
    path = destination / RAW_INDEX_REL
    raw = json.loads(path.read_text())
    for entry in raw["reports"]:
        if Path(entry["path"]).parent.name == CASE:
            changed = report_path(destination)
            entry["bytes"] = changed.stat().st_size
            entry["sha256"] = digest(changed)
            break
    else:
        raise RuntimeError("test fixture report is absent from copied raw index")
    path.write_text(json.dumps(raw, indent=2, sort_keys=True) + "\n")
    return path, digest(path)


def rebound_validator(destination: Path, raw_index_sha: str, name: str) -> Path:
    text = FINAL_SCRIPT.read_text()
    old = 'RAW_SAMPLE_SUMMARY_SHA256 = "edac04901aa00de5d726aef7229daca96ce58fb9441b13f093d50f26a9e43b7b"'
    replacement = f'RAW_SAMPLE_SUMMARY_SHA256 = "{raw_index_sha}"'
    if old not in text:
        raise RuntimeError("final validator raw-index binding was not found")
    path = destination / f"{name}-test-only-rebound.py"
    path.write_text(text.replace(old, replacement, 1))
    return path


def mutate_corrupt_hash(destination: Path) -> None:
    path = report_path(destination)
    data = path.read_bytes()
    needle = b"5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
    replacement = b"5e87f9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec"
    if needle not in data:
        raise RuntimeError("expected digest was not found in report")
    path.write_bytes(data.replace(needle, replacement, 1))


def mutate_duplicate_index(destination: Path) -> None:
    report = load_report(destination)
    result = report["results"][0]
    result["elapsed_ns"]["sample_order"][0] = result["elapsed_ns"]["sample_order"][1]
    result["source"]["docx_section_layout"]["sample_order"][0] = result["source"]["docx_section_layout"]["sample_order"][1]
    save_report(destination, report)


def mutate_missing_index(destination: Path) -> None:
    report = load_report(destination)
    result = report["results"][0]
    result["elapsed_ns"]["sample_order"].pop()
    result["source"]["docx_section_layout"]["sample_order"].pop()
    save_report(destination, report)


def mutate_phase_sum(destination: Path) -> None:
    report = load_report(destination)
    detail = report["results"][0]["source"]["docx_section_layout"]
    detail["preparation_ns"][0] += 1
    save_report(destination, report)


def mutate_binary_identity(destination: Path) -> None:
    report = load_report(destination)
    report["binary_identity"]["binary_sha256"] = "0" * 64
    save_report(destination, report)


def run_validator(destination: Path, script: Path, output_name: str) -> tuple[int, str]:
    output = destination / output_name
    process = subprocess.run(
        ["python3", "-B", str(script), "--repo", str(destination), "--output", str(output)],
        capture_output=True,
        text=True,
    )
    return process.returncode, (process.stdout + process.stderr).strip()


def main() -> int:
    if FIXTURE_ROOT.exists():
        shutil.rmtree(FIXTURE_ROOT)
    FIXTURE_ROOT.mkdir(parents=True)
    moved_destination = FIXTURE_ROOT / "moved-packet"
    copy_packet(moved_destination)
    moved_return_code, moved_output = run_validator(
        moved_destination,
        FINAL_SCRIPT,
        "moved-validation-output.json",
    )
    canonical_result = json.loads((SCRATCH / "phase-share-v2.json").read_text())
    moved_result = json.loads(
        (moved_destination / "moved-validation-output.json").read_text()
    ) if moved_return_code == 0 else None
    moved_replay = {
        "packet_root": str(moved_destination),
        "return_code": moved_return_code,
        "success": moved_return_code == 0,
        "metrics_equal_to_canonical": bool(
            moved_result is not None
            and moved_result["routes"] == canonical_result["routes"]
            and moved_result["all_routes_sum_of_raw_samples"]
            == canonical_result["all_routes_sum_of_raw_samples"]
        ),
        "output_bytes": (
            (moved_destination / "moved-validation-output.json").stat().st_size
            if moved_return_code == 0
            else None
        ),
        "output_sha256": (
            digest(moved_destination / "moved-validation-output.json")
            if moved_return_code == 0
            else None
        ),
        "stdout_tail": moved_output[-500:],
    }
    if not moved_replay["success"] or not moved_replay["metrics_equal_to_canonical"]:
        raise RuntimeError(json.dumps(moved_replay, indent=2))
    tests: list[dict[str, Any]] = []
    cases: list[tuple[str, Callable[[Path], None], bool, str]] = [
        ("corrupted_report_hash", mutate_corrupt_hash, False, "report hash mismatch"),
        ("duplicate_sample_index", mutate_duplicate_index, True, "elapsed order is not a permutation"),
        ("missing_sample_index", mutate_missing_index, True, "elapsed sample order missing"),
        ("changed_phase_sum", mutate_phase_sum, True, "phase sum mismatch"),
        ("wrong_binary_identity", mutate_binary_identity, True, "binary hash mismatch"),
    ]
    for name, mutation, rebind, expected in cases:
        destination = FIXTURE_ROOT / name
        copy_packet(destination)
        mutation(destination)
        script = FINAL_SCRIPT
        rebound_sha = None
        if rebind:
            _, rebound_sha = rebind_test_only_raw_index(destination)
            script = rebound_validator(destination, rebound_sha, name)
        return_code, output = run_validator(destination, script, "validator-output.json")
        tests.append(
            {
                "name": name,
                "expected_rejection": True,
                "observed_return_code": return_code,
                "observed_failure_contains": expected in output,
                "observed_output_tail": output[-500:],
                "test_only_raw_index_rebound": rebind,
                "test_only_raw_index_sha256": rebound_sha,
                "validator_sha256": digest(script),
            }
        )
    if not all(test["observed_return_code"] != 0 and test["observed_failure_contains"] for test in tests):
        raise RuntimeError(json.dumps(tests, indent=2))
    output = {
        "schema": "litchi.performance.0924.phase-share-negative-controls.v1",
        "purpose": "fail-closed validator tests on copied retained reports; no frozen evidence was changed",
        "test_runner": {"path": str(Path(__file__)), "sha256": digest(Path(__file__))},
        "validator": {"path": str(FINAL_SCRIPT), "sha256": digest(FINAL_SCRIPT)},
        "moved_packet_replay": moved_replay,
        "tests": tests,
        "all_rejected": True,
        "notes": [
            "corrupted_report_hash uses the production raw-index binding unchanged",
            "semantic fixtures rebind only their copied raw index and a copied validator constant; this is a test-only rebind to reach structural checks",
            "all fixture packet directories and test-only scripts may be removed after this receipt is retained",
        ],
    }
    receipt = SCRATCH / "phase-share-negative-controls-v1.json"
    receipt.write_text(json.dumps(output, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"receipt": str(receipt), "bytes": receipt.stat().st_size, "sha256": digest(receipt)}, sort_keys=True))
    for test in tests:
        print(json.dumps({"name": test["name"], "return_code": test["observed_return_code"], "matched": test["observed_failure_contains"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
