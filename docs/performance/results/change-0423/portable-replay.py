#!/usr/bin/env python3
"""Replay the complete 0423 bundle after its worktrees and binaries are gone.

The replay export contains the evidence bundle and exactly the four pinned
validator modules.  It runs the deterministic summary in that isolated
namespace; the summary replays every retained journal, report, catalog,
artifact hash, and single-report verifier result.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from typing import Any


ROOT = Path(__file__).resolve().parent
PINNED_TOOLS = ROOT / "replay-tools"
FILES = (
    "perf_abba_summary.py",
    "perf_compare.py",
    "perf_resource_profile.py",
    "validate_perf_corpus_binding.py",
)


def fail(message: str) -> None:
    raise RuntimeError(message)


def strict_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            raise ValueError(f"duplicate JSON key {key!r}")
        value[key] = item
    return value


def reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON value {value!r}")


def load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=strict_pairs,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        raise RuntimeError(f"cannot load {label}: {error}") from error
    if not isinstance(value, dict):
        raise RuntimeError(f"{label} must be an object")
    return value


def bundle_path(bundle: Path, value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value or Path(value).is_absolute() or ".." in Path(value).parts:
        raise RuntimeError(f"{label} must be a relative path")
    path = (bundle / value).resolve()
    try:
        path.relative_to(bundle.resolve())
    except ValueError as error:
        raise RuntimeError(f"{label} escapes the bundle") from error
    return path


def guard_runs(bundle: Path, export: Path, temporary: str, record: dict[str, Any]) -> None:
    capture = load_json(bundle / "capture.json", "capture")
    if capture.get("change") != 423 or capture.get("status") != "pass":
        raise RuntimeError("capture is not a passing 0423 bundle")
    runs = capture.get("runs")
    if not isinstance(runs, list):
        raise RuntimeError("capture.runs is not a list")
    selected = [item for item in runs if isinstance(item, dict) and item.get("repeat") == "R1"]
    if len(selected) != 8:
        raise RuntimeError(f"expected eight R1 runs for guard probes, found {len(selected)}")
    expected = {
        (lane, corpus, role)
        for lane in ("normal", "allocator")
        for corpus in ("plain", "media_rich")
        for role in ("owned", "source")
    }
    observed = {(item.get("lane"), item.get("corpus"), item.get("role")) for item in selected}
    if observed != expected:
        raise RuntimeError("R1 guard runs do not cover all lane/corpus/role scopes")
    guard_receipts: list[dict[str, Any]] = []
    for item in selected:
        lane = item["lane"]
        corpus = item["corpus"]
        role = item["role"]
        journal_path = bundle_path(bundle, item.get("journal"), f"{lane}/{corpus}/{role}.journal")
        journal = load_json(journal_path, str(journal_path))
        artifacts = journal.get("artifacts")
        if not isinstance(artifacts, dict):
            raise RuntimeError(f"{lane}/{corpus}/{role} journal has no artifacts")
        report_entry = artifacts.get("report")
        catalog_entry = artifacts.get("catalog")
        if not isinstance(report_entry, dict) or not isinstance(catalog_entry, dict):
            raise RuntimeError(f"{lane}/{corpus}/{role} journal report/catalog custody is malformed")
        report = bundle_path(bundle, report_entry.get("path"), f"{lane}/{corpus}/{role}.report")
        catalog = bundle_path(bundle, catalog_entry.get("path"), f"{lane}/{corpus}/{role}.catalog")
        selector = item.get("selector")
        if not isinstance(selector, str) or not selector:
            raise RuntimeError(f"{lane}/{corpus}/{role} selector is missing")
        samples, warmups = ((100, 10) if lane == "normal" else (30, 3))
        command = [
            sys.executable, str(bundle / "check-report-guards.py"),
            "--repo-root", str(export), "--report", str(report), "--catalog", str(catalog),
            "--selector", selector, "--lane", lane, "--contract", "formal",
            "--samples", str(samples), "--warmups", str(warmups),
        ]
        environment = os.environ.copy()
        environment["PYTHONDONTWRITEBYTECODE"] = "1"
        environment["PYTHONPATH"] = str(export)
        result = subprocess.run(command, cwd=export, env=environment, capture_output=True, text=True, check=False)
        if result.returncode != 0:
            detail = (result.stderr or result.stdout).replace(temporary, "<EXPORT>").strip()
            raise RuntimeError(f"{lane}/{corpus}/{role} report guards failed: {detail}")
        try:
            output = json.loads(result.stdout, object_pairs_hook=strict_pairs, parse_constant=reject_constant)
        except (UnicodeError, json.JSONDecodeError, ValueError) as error:
            raise RuntimeError(f"{lane}/{corpus}/{role} report guard output is invalid JSON: {error}") from error
        if not isinstance(output, dict) or output.get("validated_original") is not True or output.get("claim_authorized") is not False:
            raise RuntimeError(f"{lane}/{corpus}/{role} report guard output is not a no-claim proof")
        probes = output.get("checks")
        if not isinstance(probes, list) or not probes or any(
            not isinstance(probe, dict) or probe.get("rejected") is not True or probe.get("unexpected_exception") is True
            for probe in probes
        ):
            raise RuntimeError(f"{lane}/{corpus}/{role} report guard mutations were not all rejected cleanly")
        if output.get("report_sha256") is None or output.get("catalog_sha256") is None:
            raise RuntimeError(f"{lane}/{corpus}/{role} report guard output lacks artifact hashes")
        guard_receipts.append({
            "lane": lane,
            "corpus": corpus,
            "role": role,
            "repeat": "R1",
            "selector": selector,
            "samples": samples,
            "warmups": warmups,
            "argv": [
                "check-report-guards.py", "--report", str(report.relative_to(bundle)),
                "--catalog", str(catalog.relative_to(bundle)), "--selector", selector,
                "--lane", lane, "--contract", "formal", "--samples", str(samples),
                "--warmups", str(warmups),
            ],
            "exit_code": result.returncode,
            "stderr_empty": not bool(result.stderr.strip()),
            "report_sha256": output["report_sha256"],
            "catalog_sha256": output["catalog_sha256"],
            "probe_count": len(probes),
            "rejected_probe_count": sum(1 for probe in probes if probe.get("rejected") is True),
            "unexpected_exception_count": sum(1 for probe in probes if probe.get("unexpected_exception") is True),
            "status": "pass",
        })
    record["guard_checks"] = guard_receipts


def load_manifest() -> dict[str, Any]:
    try:
        value = json.loads(
            (PINNED_TOOLS / "manifest.json").read_text(encoding="utf-8"),
            object_pairs_hook=strict_pairs,
        )
        if not isinstance(value, dict):
            raise ValueError("manifest must be an object")
        entries = value.get("files")
        if not isinstance(entries, list) or len(entries) != len(FILES):
            raise ValueError("pinned validator manifest has the wrong file count")
        expected = {}
        for entry in entries:
            if not isinstance(entry, dict) or set(entry) - {"path", "sha256", "source_path"}:
                raise ValueError("pinned validator manifest entry is malformed")
            path = entry.get("path")
            digest = entry.get("sha256")
            if path not in FILES or path in expected or not isinstance(digest, str) or len(digest) != 64:
                raise ValueError("pinned validator manifest entry is invalid")
            expected[path] = digest
        if set(expected) != set(FILES):
            raise ValueError("pinned validator manifest differs from the required module set")
        for name in FILES:
            path = PINNED_TOOLS / name
            if not path.is_file():
                raise ValueError(f"pinned validator is missing: {name}")
            actual = hashlib.sha256(path.read_bytes()).hexdigest()
            if actual != expected[name]:
                raise ValueError(f"pinned validator hash mismatch: {name}")
        value["files"] = [
            {"path": name, "sha256": expected[name]}
            for name in FILES
        ]
        return value
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError, TypeError) as error:
        raise RuntimeError(str(error)) from error


def run() -> int:
    checks_dir = ROOT / "checks"
    checks_dir.mkdir(parents=True, exist_ok=True)
    receipt = checks_dir / "portable-replay.json"
    try:
        manifest = load_manifest()
    except RuntimeError as error:
        receipt.write_text(
            json.dumps(
                {
                    "change": 423,
                    "status": "failed",
                    "pinned_tools": [],
                    "checks": [],
                    "error": str(error),
                    "temporary_export_removed": True,
                },
                indent=2,
                sort_keys=True,
            )
            + "\n",
            encoding="utf-8",
        )
        print(f"0423 portable replay failed: {error}", file=sys.stderr)
        return 1

    record: dict[str, Any] = {
        "change": 423,
        "status": "running",
        "pinned_tools": manifest["files"],
        "checks": [],
        "temporary_export_removed": False,
    }
    try:
        with tempfile.TemporaryDirectory(prefix="litchi-goal-0423-portable-") as temporary:
            export = Path(temporary)
            bundle = export / "docs/performance/results/change-0423"
            shutil.copytree(ROOT, bundle, ignore=shutil.ignore_patterns("__pycache__"))
            tools = export / "tools"
            tools.mkdir()
            (tools / "__init__.py").write_text("", encoding="utf-8")
            for name in FILES:
                shutil.copy2(PINNED_TOOLS / name, tools / name)

            command = [
                sys.executable,
                str(bundle / "summarize.py"),
                "--root",
                str(bundle),
                "--repo-root",
                str(export),
                "--replay",
            ]
            environment = os.environ.copy()
            environment["PYTHONDONTWRITEBYTECODE"] = "1"
            environment["PYTHONPATH"] = str(export)
            result = subprocess.run(
                command,
                cwd=export,
                env=environment,
                capture_output=True,
                text=True,
                check=False,
            )
            record["checks"].append({
                "name": "summary-replay",
                "argv": ["summarize.py", "--root", "<EXPORT>/docs/performance/results/change-0423", "--repo-root", "<EXPORT>", "--replay"],
                "exit_code": result.returncode,
                "stdout_status": "pass" if result.returncode == 0 else "failed",
                "stderr_empty": not bool(result.stderr.strip()),
            })
            if result.returncode:
                record["status"] = "failed"
            else:
                guard_runs(bundle, export, temporary, record)
                record["status"] = "pass"
        record["temporary_export_removed"] = not export.exists()
    except (OSError, subprocess.SubprocessError, RuntimeError) as error:
        record["status"] = "failed"
        detail = str(error)
        if "temporary" in locals():
            detail = detail.replace(str(temporary), "<EXPORT>")
        record["error"] = detail

    try:
        receipt.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    except OSError as error:
        print(f"0423 portable replay could not write receipt: {error}", file=sys.stderr)
        return 1
    print(json.dumps({key: record[key] for key in ("change", "status", "temporary_export_removed")}, sort_keys=True))
    return 0 if record["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(run())
