#!/usr/bin/env python3
"""Replay small semantic corruptions against the retained 0420 control row.

Each mutation is validated outside the capture directory and without passing
``--capture``.  A rejection therefore comes from the report/corpus semantics,
not from an unchanged capture-manifest hash.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
from typing import Any, Callable


ROOT = Path(__file__).resolve().parent
REPORT = ROOT / "runs" / "normal" / "A1" / "pptx_cross_copy_media_rich_lifecycle" / "report.json"
CATALOG = ROOT / "runs" / "normal" / "A1" / "pptx_cross_copy_media_rich_lifecycle" / "catalog.json"
SELECTOR = "pptx_cross_copy_media_rich_lifecycle"


def repo_root() -> Path:
    for candidate in (ROOT, *ROOT.parents):
        if (candidate / "tools" / "perf_abba_summary.py").is_file():
            return candidate
    raise RuntimeError("cannot locate repository root")


def load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def run_verifier(report: Path, catalog: Path) -> tuple[list[str], subprocess.CompletedProcess[str]]:
    command = [
        sys.executable,
        str(ROOT / "verify.py"),
        "--repo-root",
        str(repo_root()),
        "--report",
        str(report),
        "--catalog",
        str(catalog),
        "--selector",
        SELECTOR,
        "--samples",
        "100",
        "--warmups",
        "10",
    ]
    result = subprocess.run(command, text=True, capture_output=True, check=False)
    return command, result


def command_template(command: list[str], temp: Path) -> list[str]:
    catalog_placeholder = (
        "<REPO>/docs/performance/results/change-0420/"
        "runs/normal/A1/pptx_cross_copy_media_rich_lifecycle/catalog.json"
    )
    return [
        "<PYTHON>" if value == sys.executable else
        "<REPO>/docs/performance/results/change-0420/verify.py"
        if value == str(ROOT / "verify.py") else
        "<REPO>" if value == str(repo_root()) else
        "<TMP>/report.json" if value == str(temp / "report.json") else
        "<TMP>/catalog.json" if value == str(temp / "catalog.json") else
        catalog_placeholder if value == str(CATALOG) else
        value
        for value in command
    ]


def run_probe(
    name: str,
    expected_field: str,
    mutate: Callable[[dict[str, Any]], None],
    original: dict[str, Any],
    catalog: Path,
    temp: Path,
) -> dict[str, Any]:
    report = copy.deepcopy(original)
    mutate(report)
    report_path = temp / "report.json"
    report_path.write_text(json.dumps(report, sort_keys=True) + "\n", encoding="utf-8")
    command, result = run_verifier(report_path, catalog)
    stderr = result.stderr.strip()
    if result.returncode == 0:
        raise RuntimeError(f"{name} unexpectedly passed")
    if expected_field not in stderr:
        raise RuntimeError(f"{name} diagnostic omitted {expected_field!r}: {stderr}")
    if "capture" in stderr.lower():
        raise RuntimeError(f"{name} was rejected by capture metadata: {stderr}")
    return {
        "name": name,
        "mutation": name,
        "capture_manifest_used": False,
        "expected_field": expected_field,
        "status": "rejected",
        "returncode": result.returncode,
        "command": command_template(command, temp),
        "diagnostic": stderr,
    }


def corrupt_sample_order(report: dict[str, Any]) -> None:
    order = report["results"][0]["elapsed_ns"].get("sample_order")
    if not isinstance(order, list) or len(order) < 2:
        raise RuntimeError("control report sample_order needs at least two entries")
    order[0] = order[1]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "checks" / "verify-probes.json")
    args = parser.parse_args()
    if not REPORT.is_file() or not CATALOG.is_file():
        print("FAIL: retained control report/catalog are missing", file=sys.stderr)
        return 2
    original = load(REPORT)
    result_rows = original.get("results")
    if not isinstance(result_rows, list) or len(result_rows) != 1:
        print("FAIL: control report does not have exactly one result", file=sys.stderr)
        return 2
    result = result_rows[0]
    if not isinstance(result, dict) or result.get("case") != SELECTOR:
        print("FAIL: control report selector differs", file=sys.stderr)
        return 2
    configuration = original.get("configuration")
    elapsed = result.get("elapsed_ns")
    if not isinstance(configuration, dict) or not isinstance(elapsed, dict):
        print("FAIL: control report shape is incomplete", file=sys.stderr)
        return 2
    if configuration.get("samples_per_case") != 100 or configuration.get("warmup_iterations_per_case") != 10:
        print("FAIL: control report is not the retained 100/10 control", file=sys.stderr)
        return 2

    baseline_command, baseline = run_verifier(REPORT, CATALOG)
    if baseline.returncode != 0:
        print(f"FAIL: unmodified control report was rejected: {baseline.stderr.strip()}", file=sys.stderr)
        return 2

    mutations: tuple[tuple[str, str, Callable[[dict[str, Any]], None]], ...] = (
        (
            "wrong-output-vector",
            "output_sha256",
            lambda report: report["results"][0]["source"]["pptx_cross_copy"]["output_sha256"].__setitem__(0, "0" * 64),
        ),
        (
            "false-correctness-gate",
            "gates",
            lambda report: report["results"][0]["source"]["pptx_cross_copy"]["gates"].__setitem__("semantic_output_verified", False),
        ),
        (
            "invalid-sample-order",
            "sample_order",
            corrupt_sample_order,
        ),
    )
    with tempfile.TemporaryDirectory(prefix="litchi-goal-0420-probes-") as directory:
        temp = Path(directory)
        probes = [
            run_probe(name, field, mutate, original, CATALOG, temp)
            for name, field, mutate in mutations
        ]

    output = {
        "change": 420,
        "status": "pass",
        "verifier_sha256": sha256(ROOT / "verify.py"),
        "source_report": "runs/normal/A1/pptx_cross_copy_media_rich_lifecycle/report.json",
        "source_catalog": "runs/normal/A1/pptx_cross_copy_media_rich_lifecycle/catalog.json",
        "control_report_shape": {
            "top_level_keys": sorted(original),
            "result_keys": sorted(result),
            "selector": result["case"],
            "samples": configuration["samples_per_case"],
            "warmups": configuration["warmup_iterations_per_case"],
            "elapsed_sample_count": len(elapsed.get("samples", [])),
        },
        "baseline": {
            "status": "pass",
            "capture_manifest_used": False,
            "command": [
                "<PYTHON>",
                "<REPO>/docs/performance/results/change-0420/verify.py",
                "--repo-root",
                "<REPO>",
                "--report",
                "<REPO>/docs/performance/results/change-0420/runs/normal/A1/pptx_cross_copy_media_rich_lifecycle/report.json",
                "--catalog",
                "<REPO>/docs/performance/results/change-0420/runs/normal/A1/pptx_cross_copy_media_rich_lifecycle/catalog.json",
                "--selector",
                SELECTOR,
                "--samples",
                "100",
                "--warmups",
                "10",
            ],
            "returncode": baseline.returncode,
        },
        "probes": probes,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(output, sort_keys=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": output["status"], "probes": len(probes)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
