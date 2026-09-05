#!/usr/bin/env python3
"""Replay the matched 0424 summary and both retained Heaptrack profile trees.

The export is made in a temporary directory, with the four shared validators
copied to ``<export>/tools`` and checked against ``pinned/tools-manifest.json``.
Replay needs the retained reports, catalogs, journals, traces, and analysis
manifests; it does not need either profiling worktree or binary.  The control
profile remains under ``runs/`` and the candidate profile under
``candidate-profile/runs/``.
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
PINNED_MANIFEST = ROOT / "pinned" / "tools-manifest.json"
TOOLS = (
    "pinned/tools/perf_abba_summary.py",
    "pinned/tools/perf_compare.py",
    "pinned/tools/perf_resource_profile.py",
    "pinned/tools/validate_perf_corpus_binding.py",
)
CORPORA = ("plain", "media_rich")
ROLES = ("control", "candidate")
LANES = ("normal", "allocator")
R1_GUARD_COUNT = {"normal": 12, "allocator": 14}
HISTORICAL_PARSER_PIN = "pinned/analyze-heaptrack.py"


class ReplayError(RuntimeError):
    """Replay failure carrying all completed child receipts so far."""

    def __init__(
        self,
        message: str,
        *,
        checks: list[dict[str, Any]] | None = None,
        report_guard_summary: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.checks = list(checks or [])
        self.report_guard_summary = report_guard_summary


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
        raise ReplayError(f"cannot load {label}: {error}") from error
    if not isinstance(value, dict):
        raise ReplayError(f"{label} must be an object")
    return value


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        raise ReplayError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest()


def validate_pinned_tools(root: Path) -> list[dict[str, str]]:
    manifest_path = root / "pinned" / "tools-manifest.json"
    manifest = load_json(manifest_path, "pinned tools manifest")
    entries = manifest.get("files")
    if not isinstance(entries, list) or len(entries) != len(TOOLS):
        raise ReplayError("pinned tools manifest has the wrong file count")
    expected: dict[str, str] = {}
    for entry in entries:
        if not isinstance(entry, dict):
            raise ReplayError("pinned tools manifest entry is malformed")
        path = entry.get("path")
        digest = entry.get("sha256")
        if path not in TOOLS or path in expected or not isinstance(digest, str) or len(digest) != 64:
            raise ReplayError("pinned tools manifest entry is invalid")
        expected[path] = digest
    if set(expected) != set(TOOLS):
        raise ReplayError("pinned tools manifest does not cover the required files")
    result: list[dict[str, str]] = []
    for name in TOOLS:
        path = root / name
        if not path.is_file() or path.is_symlink():
            raise ReplayError(f"pinned tool is missing: {name}")
        observed = sha256_file(path)
        if observed != expected[name]:
            raise ReplayError(f"pinned tool hash mismatch: {name}")
        result.append({"path": name, "sha256": observed})
    return result


def copy_pinned_tools(bundle: Path, export: Path) -> None:
    destination = export / "tools"
    destination.mkdir(parents=True, exist_ok=True)
    (destination / "__init__.py").write_text("", encoding="utf-8")
    for name in TOOLS:
        source = bundle / name
        target = destination / Path(name).name
        shutil.copy2(source, target)


def copy_external_parser(
    export: Path, source_root: Path
) -> tuple[list[str], dict[str, Any] | None]:
    """Retain the historical parser path and inventory its pinned bytes."""
    copied: list[str] = []
    pin_inventory: dict[str, Any] | None = None
    for corpus in CORPORA:
        for role_root in ("runs", "candidate-profile/runs"):
            manifest_path = export / "docs/performance/results/change-0424" / role_root / corpus / "analysis.json"
            if not manifest_path.is_file():
                continue
            manifest = load_json(manifest_path, str(manifest_path))
            parser = manifest.get("capture", {}).get("parser", {})
            path = parser.get("path") if isinstance(parser, dict) else None
            digest = parser.get("sha256") if isinstance(parser, dict) else None
            if not isinstance(path, str) or not isinstance(digest, str):
                continue
            pinned = source_root / HISTORICAL_PARSER_PIN
            if pinned.is_file():
                if sha256_file(pinned) != digest:
                    raise ReplayError(
                        "pinned historical parser hash differs from analysis binding"
                    )
                source = pinned
                pin_inventory = {
                    "path": HISTORICAL_PARSER_PIN,
                    "bytes": pinned.stat().st_size,
                    "sha256": digest,
                }
            else:
                source = source_root / path
                if not source.is_file() or sha256_file(source) != digest:
                    raise ReplayError(f"historical parser binding is unavailable: {path}")
            target = export / "docs/performance/results/change-0424" / path
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists() and sha256_file(target) != digest:
                raise ReplayError(f"export parser binding conflicts: {target}")
            if not target.exists():
                shutil.copy2(source, target)
            copied.append(path)
    return sorted(set(copied)), pin_inventory


def run_command(command: list[str], cwd: Path, temporary: str) -> dict[str, Any]:
    environment = os.environ.copy()
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    environment["PYTHONPATH"] = str(cwd)
    try:
        result = subprocess.run(
            command, cwd=cwd, env=environment,
            capture_output=True, text=True, check=False,
        )
    except OSError as error:
        raise ReplayError(f"cannot execute replay command: {error}") from error
    return {
        "argv": [item.replace(temporary, "<EXPORT>") for item in command],
        "exit_code": result.returncode,
        "stdout": result.stdout.replace(temporary, "<EXPORT>"),
        "stderr": result.stderr.replace(temporary, "<EXPORT>"),
        "status": "pass" if result.returncode == 0 else "failed",
    }


def summary_root(bundle: Path) -> Path:
    for candidate in (bundle, bundle / "matched"):
        if (candidate / "capture.json").is_file() and (candidate / "summary.json").is_file():
            return candidate
    raise ReplayError("matched summary root with capture.json and summary.json is missing")


def safe_bundle_path(root: Path, value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value or Path(value).is_absolute() or ".." in Path(value).parts:
        raise ReplayError(f"{label} must be a safe relative path")
    path = (root / value).resolve()
    try:
        path.relative_to(root.resolve())
    except ValueError as error:
        raise ReplayError(f"{label} escapes the matched bundle") from error
    return path


def guard_child_receipt(
    label: str,
    lane: str,
    corpus: str,
    role: str,
    selector: str,
    samples: int,
    warmups: int,
    result: dict[str, Any],
    status: str,
) -> dict[str, Any]:
    return {
        "name": f"report-guards-{label}",
        "status": status,
        "lane": lane,
        "corpus": corpus,
        "role": role,
        "repeat": "R1",
        "selector": selector,
        "samples": samples,
        "warmups": warmups,
        "command": result,
    }


def _run_report_guards(
    matched: Path, bundle: Path, export: Path, temporary: str,
    progress: dict[str, Any],
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Run compact mutation proofs for the eight retained R1 reports."""
    capture = load_json(matched / "capture.json", "matched capture")
    runs = capture.get("runs")
    if not isinstance(runs, list):
        raise ReplayError("matched capture.runs is not a list")
    selected = [
        item for item in runs
        if isinstance(item, dict) and item.get("repeat") == "R1"
    ]
    expected_keys = {
        (lane, corpus, role)
        for lane in LANES
        for corpus in CORPORA
        for role in ROLES
    }
    observed_keys = {
        (item.get("lane"), item.get("corpus"), item.get("role"))
        for item in selected
    }
    if observed_keys != expected_keys or len(selected) != len(expected_keys):
        raise ReplayError("matched capture does not contain exactly eight R1 runs")

    receipts: list[dict[str, Any]] = []
    progress["receipts"] = receipts
    for lane in LANES:
        samples, warmups = ((100, 10) if lane == "normal" else (30, 3))
        for corpus in CORPORA:
            for role in ROLES:
                label = f"{lane}/{corpus}/{role}-R1"
                row = next(
                    item for item in selected
                    if item.get("lane") == lane
                    and item.get("corpus") == corpus
                    and item.get("role") == role
                )
                journal_path = safe_bundle_path(matched, row.get("journal"), f"{label}.journal")
                journal = load_json(journal_path, f"{label} journal")
                if journal.get("status") != "pass":
                    raise ReplayError(f"{label} journal is not passing")
                if row.get("journal_sha256") != sha256_file(journal_path):
                    raise ReplayError(f"{label} journal hash differs from matched capture")
                artifacts = journal.get("artifacts")
                if not isinstance(artifacts, dict):
                    raise ReplayError(f"{label} journal artifacts are missing")
                report_entry = artifacts.get("report")
                catalog_entry = artifacts.get("catalog")
                if not isinstance(report_entry, dict) or not isinstance(catalog_entry, dict):
                    raise ReplayError(f"{label} report/catalog custody is malformed")
                report = safe_bundle_path(matched, report_entry.get("path"), f"{label}.report")
                catalog = safe_bundle_path(matched, catalog_entry.get("path"), f"{label}.catalog")
                if report_entry.get("sha256") != sha256_file(report):
                    raise ReplayError(f"{label} report hash differs from journal")
                if catalog_entry.get("sha256") != sha256_file(catalog):
                    raise ReplayError(f"{label} catalog hash differs from journal")
                selector = row.get("selector")
                if not isinstance(selector, str) or not selector:
                    raise ReplayError(f"{label} selector is missing")
                result = run_command(
                    [
                        sys.executable, str(bundle / "check-report-guards.py"),
                        "--repo-root", str(bundle / "pinned"), "--report", str(report),
                        "--catalog", str(catalog), "--selector", selector,
                        "--lane", lane, "--contract", "formal",
                        "--samples", str(samples), "--warmups", str(warmups),
                    ],
                    export,
                    temporary,
                )
                if result["exit_code"] != 0:
                    receipts.append(guard_child_receipt(
                        label, lane, corpus, role, selector, samples, warmups,
                        result, "failed",
                    ))
                    raise ReplayError(f"{label} report guards failed")
                try:
                    output = json.loads(
                        result["stdout"], object_pairs_hook=strict_pairs,
                        parse_constant=reject_constant,
                    )
                except (UnicodeError, json.JSONDecodeError, ValueError) as error:
                    receipts.append(guard_child_receipt(
                        label, lane, corpus, role, selector, samples, warmups,
                        result, "failed",
                    ))
                    raise ReplayError(f"{label} guard output is invalid JSON: {error}") from error
                if (
                    not isinstance(output, dict)
                    or output.get("validated_original") is not True
                    or output.get("claim_authorized") is not False
                    or output.get("performance_claim") is not None
                ):
                    receipts.append(guard_child_receipt(
                        label, lane, corpus, role, selector, samples, warmups,
                        result, "failed",
                    ))
                    raise ReplayError(f"{label} guard output is not a no-claim proof")
                probes = output.get("checks")
                expected_probe_count = R1_GUARD_COUNT[lane]
                if not isinstance(probes, list) or len(probes) != expected_probe_count:
                    receipts.append(guard_child_receipt(
                        label, lane, corpus, role, selector, samples, warmups,
                        result, "failed",
                    ))
                    raise ReplayError(
                        f"{label} expected {expected_probe_count} mutation probes, "
                        f"observed {len(probes) if isinstance(probes, list) else 'invalid'}"
                    )
                if any(
                    not isinstance(probe, dict)
                    or probe.get("rejected") is not True
                    or probe.get("unexpected_exception") is True
                    for probe in probes
                ):
                    receipts.append(guard_child_receipt(
                        label, lane, corpus, role, selector, samples, warmups,
                        result, "failed",
                    ))
                    raise ReplayError(f"{label} has a non-rejected or unexpected probe")
                if output.get("report_sha256") != sha256_file(report):
                    receipts.append(guard_child_receipt(
                        label, lane, corpus, role, selector, samples, warmups,
                        result, "failed",
                    ))
                    raise ReplayError(f"{label} guard report hash differs")
                if output.get("catalog_sha256") != sha256_file(catalog):
                    receipts.append(guard_child_receipt(
                        label, lane, corpus, role, selector, samples, warmups,
                        result, "failed",
                    ))
                    raise ReplayError(f"{label} guard catalog hash differs")
                receipt = guard_child_receipt(
                    label, lane, corpus, role, selector, samples, warmups,
                    result, "pass",
                )
                receipt.update({
                    "report_sha256": output["report_sha256"],
                    "catalog_sha256": output["catalog_sha256"],
                    "probe_count": len(probes),
                    "rejected_probe_count": sum(
                        1 for probe in probes if probe.get("rejected") is True
                    ),
                    "unexpected_exception_count": sum(
                        1 for probe in probes if probe.get("unexpected_exception") is True
                    ),
                })
                receipts.append(receipt)
    expected_total = sum(R1_GUARD_COUNT[lane] * 4 for lane in LANES)
    observed_total = sum(item["probe_count"] for item in receipts)
    if observed_total != expected_total:
        raise ReplayError(f"report guard probe total {observed_total} differs from {expected_total}")
    summary = guard_summary_for(receipts, "pass")
    return receipts, summary


def guard_summary_for(
    receipts: list[dict[str, Any]], status: str
) -> dict[str, Any]:
    expected_total = sum(R1_GUARD_COUNT[lane] * 4 for lane in LANES)
    observed_total = sum(
        item.get("probe_count", 0)
        for item in receipts
        if isinstance(item.get("probe_count", 0), int)
    )
    complete = (
        len(receipts) == 8
        and observed_total == expected_total
        and all(
            item.get("rejected_probe_count") == item.get("probe_count")
            and item.get("unexpected_exception_count") == 0
            for item in receipts
        )
    )
    return {
        "status": status,
        "run_count": len(receipts),
        "expected_run_count": 8,
        "probe_counts_by_lane": dict(R1_GUARD_COUNT),
        "expected_probe_total": expected_total,
        "observed_probe_total": observed_total,
        "all_originals_validated": complete,
        "all_probes_rejected": complete,
        "claim_authorized": False,
        "performance_claim": None,
    }


def run_report_guards(
    matched: Path, bundle: Path, export: Path, temporary: str
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    progress: dict[str, Any] = {}
    try:
        return _run_report_guards(matched, bundle, export, temporary, progress)
    except ReplayError as error:
        receipts = progress.get("receipts", [])
        if not isinstance(receipts, list):
            receipts = []
        error.checks = list(receipts)
        error.report_guard_summary = guard_summary_for(receipts, "failed")
        raise


def run_replay(
    bundle: Path, export: Path, temporary: str
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    checks: list[dict[str, Any]] = []
    summary = summary_root(bundle)
    summary_result = run_command(
        [
            sys.executable, str(bundle / "summarize.py"),
            "--root", str(summary), "--repo-root", str(bundle / "pinned"), "--replay",
        ],
        export,
        temporary,
    )
    summary_result["name"] = "matched-summary-replay"
    checks.append(summary_result)
    if summary_result["exit_code"] != 0:
        raise ReplayError("matched summary replay failed", checks=checks)
    try:
        guard_receipts, guard_summary = run_report_guards(
            summary_root(bundle), bundle, export, temporary
        )
    except ReplayError as error:
        raise ReplayError(
            str(error),
            checks=checks + error.checks,
            report_guard_summary=error.report_guard_summary,
        ) from error
    checks.extend(guard_receipts)
    for role in ROLES:
        result = run_command(
            [
                sys.executable, str(bundle / "analyze.py"),
                "--root", str(bundle), "--role", role,
                "--replay",
            ],
            export,
            temporary,
        )
        result["name"] = f"heaptrack-{role}-replay"
        checks.append(result)
        if result["exit_code"] != 0:
            raise ReplayError(
                f"{role} Heaptrack replay failed",
                checks=checks,
                report_guard_summary=guard_summary,
            )
    return checks, guard_summary


def modified_tool_guard(bundle: Path, tools: list[dict[str, str]]) -> dict[str, Any]:
    target = bundle / tools[0]["path"]
    original = target.read_bytes()
    try:
        with target.open("ab") as stream:
            stream.write(b"\n")
        try:
            validate_pinned_tools(bundle)
        except ReplayError as error:
            return {"name": "modified-pinned-tool-rejected", "status": "pass", "error": str(error)}
        raise ReplayError("modified pinned tool was not rejected")
    finally:
        target.write_bytes(original)


def run() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args()
    source_root = args.root.expanduser().resolve()
    receipt_path = source_root / "checks" / "portable-replay.json"
    record: dict[str, Any] = {
        "change": 424,
        "status": "running",
        "temporary_export_removed": False,
        "pinned_tools": [],
        "checks": [],
        "historical_parser_paths": [],
        "historical_parser_pin": None,
        "report_guard_summary": None,
    }
    temporary_path: Path | None = None
    try:
        record["pinned_tools"] = validate_pinned_tools(source_root)
        with tempfile.TemporaryDirectory(prefix="litchi-goal-0424-portable-") as temporary:
            temporary_path = Path(temporary)
            export = temporary_path / "docs/performance/results/change-0424"
            shutil.copytree(source_root, export, ignore=shutil.ignore_patterns("__pycache__"))
            copy_pinned_tools(export, temporary_path)
            parser_paths, parser_pin = copy_external_parser(temporary_path, source_root)
            record["historical_parser_paths"] = parser_paths
            record["historical_parser_pin"] = parser_pin
            record["checks"].append(modified_tool_guard(export, record["pinned_tools"]))
            replay_checks, guard_summary = run_replay(export, temporary_path, temporary)
            record["checks"].extend(replay_checks)
            record["report_guard_summary"] = guard_summary
            record["status"] = "pass"
        record["temporary_export_removed"] = not temporary_path.exists()
        if not record["temporary_export_removed"]:
            raise ReplayError("temporary export was not removed")
    except (OSError, ReplayError, subprocess.SubprocessError) as error:
        record["status"] = "failed"
        record["error"] = str(error).replace(str(temporary_path) if temporary_path else "", "<EXPORT>")
        if isinstance(error, ReplayError):
            record["checks"].extend(error.checks)
            if error.report_guard_summary is not None:
                record["report_guard_summary"] = error.report_guard_summary
        if temporary_path is not None:
            record["temporary_export_removed"] = not temporary_path.exists()
    try:
        receipt_path.parent.mkdir(parents=True, exist_ok=True)
        receipt_path.write_text(
            json.dumps(record, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
            encoding="utf-8",
        )
    except OSError as error:
        print(f"0424 portable replay could not write receipt: {error}", file=sys.stderr)
        return 1
    print(json.dumps({key: record[key] for key in ("change", "status", "temporary_export_removed")}, sort_keys=True))
    return 0 if record["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(run())
