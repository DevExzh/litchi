#!/usr/bin/env python3
"""Run resealed negative controls against the portable 0465 verifier.

Each probe copies the complete evidence directory, changes one authenticated
input, rewrites that copy's SHA256SUMS inventory, and requires the verifier to
reject the resulting bundle for the expected semantic reason.  The retained
bundle is never modified.
"""

from __future__ import annotations

import hashlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable


ROOT = Path(__file__).resolve().parent


class ProbeError(RuntimeError):
    pass


def load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def save(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def reseal(root: Path) -> None:
    rows: list[str] = []
    for path in sorted(root.rglob("*")):
        if path.is_symlink() or not path.is_file() or path.name == "SHA256SUMS":
            continue
        rows.append(f"{sha(path)}  {path.relative_to(root).as_posix()}")
    (root / "SHA256SUMS").write_text("\n".join(rows) + "\n", encoding="utf-8")


def report_row(report: dict[str, Any], shape: str = "tiny") -> dict[str, Any]:
    for row in report.get("results", []):
        if row.get("case") == "odp_existing_append_lifecycle" and row.get("corpus", {}).get("shape") == shape:
            return row
    raise ProbeError(f"ODP {shape} row is absent")


def refresh_report_receipt(root: Path, lane: str, *, report_rows: int | None = None) -> None:
    report_path = root / "captures" / lane / "report.json"
    receipt_path = root / "captures" / lane / "receipt.json"
    receipt = load(receipt_path)
    artifact = receipt["artifact_hashes"]["report"]
    artifact.update(bytes=report_path.stat().st_size, sha256=sha(report_path))
    if report_rows is not None:
        receipt["report_validation"]["report_rows"] = report_rows
    save(receipt_path, receipt)


def run_probe(name: str, mutate: Callable[[Path], None], needle: str) -> dict[str, str]:
    with tempfile.TemporaryDirectory(prefix=f"litchi-0465-{name}-") as temporary:
        copy_root = Path(temporary) / "bundle"
        shutil.copytree(ROOT, copy_root, symlinks=True)
        mutate(copy_root)
        reseal(copy_root)
        result = subprocess.run(
            [sys.executable, "-B", "verify.py", "--portable"],
            cwd=copy_root,
            capture_output=True,
            text=True,
            check=False,
            env={**__import__("os").environ, "PYTHONDONTWRITEBYTECODE": "1"},
        )
        try:
            output = json.loads(result.stdout)
        except json.JSONDecodeError as error:
            raise ProbeError(f"{name}: verifier did not return JSON: {error}; stderr={result.stderr[-500:]}") from error
        error = output.get("error", "")
        if result.returncode == 0 or needle not in error:
            raise ProbeError(f"{name}: expected rejection containing {needle!r}, got rc={result.returncode}, error={error!r}")
    return {"status": "pass", "rejection": needle}


def probes() -> dict[str, dict[str, str]]:
    def summary_p50(root: Path) -> None:
        summary_path = root / "summary.json"
        summary = load(summary_path)
        value = summary["odp_append"]["normal_latency"]["R1"]["tiny"]["latency_ns"]["p50"]
        summary["odp_append"]["normal_latency"]["R1"]["tiny"]["latency_ns"]["p50"] = value + 1
        save(summary_path, summary)

    def gate_false(root: Path) -> None:
        path = root / "captures" / "R1-normal" / "report.json"
        report = load(path)
        report_row(report)["source"]["odp_append"]["patch_replay_verified"] = False
        save(path, report)
        refresh_report_receipt(root, "R1-normal")

    def missing_odp(root: Path) -> None:
        path = root / "captures" / "R1-normal" / "report.json"
        report = load(path)
        target = report_row(report)
        report["results"].remove(target)
        save(path, report)
        refresh_report_receipt(root, "R1-normal")

    def checked_input(root: Path) -> None:
        path = root / "checked" / "perf-regression-default-manifest-v1.json"
        identity = load(path)
        identity["result_count"] = 200
        save(path, identity)

    def normal_allocator(root: Path) -> None:
        path = root / "captures" / "R1-normal" / "report.json"
        report = load(path)
        row = report_row(report)
        allocator_report = load(root / "captures" / "R1-allocator" / "report.json")
        allocation = report_row(allocator_report)["operation_metrics"]["allocation"]
        allocation = json.loads(json.dumps(allocation))
        allocation.pop("region_peak_live_bytes", None)
        row["operation_metrics"]["allocation"] = allocation
        save(path, report)
        refresh_report_receipt(root, "R1-normal")

    return {
        "summary_p50_plus_one": run_probe("summary-p50", summary_p50, "summary.json: deterministic recomputation differs"),
        "odp_gate_false": run_probe("odp-gate", gate_false, "required gate is not true"),
        "missing_odp_row": run_probe("missing-odp", missing_odp, "expected 201 rows"),
        "changed_checked_input": run_probe("checked-input", checked_input, "checked default identity counts differ"),
        "normal_allocator_presence": run_probe("normal-allocation", normal_allocator, "normal report exposes allocator fields"),
    }


def main() -> int:
    try:
        result = {"schema": "litchi-0465-negative-probes-v1", "change": 465, "status": "pass", "probes": probes()}
    except (OSError, ProbeError, TypeError, ValueError) as error:
        result = {"schema": "litchi-0465-negative-probes-v1", "change": 465, "status": "fail", "error": str(error)}
    print(json.dumps(result, ensure_ascii=False, sort_keys=True, indent=2))
    return 0 if result["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
