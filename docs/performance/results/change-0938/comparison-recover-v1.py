#!/usr/bin/env python3
"""Recover descriptive metrics from the terminal 0938 packet.

The frozen runner marked every child failed after its 0937 validator returned a
result without the already-validated ``output_sha256`` field.  This offline
tool preserves that terminal failure, independently checks the retained raw
reports and streams, and reduces only those original children.  It never
starts a workload and never rewrites packet evidence.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import sys
from pathlib import Path
from typing import Any


REPO_ROOT = Path("/home/zhuhe/code/litchi")
RUNNER_PATH = REPO_ROOT / "docs/performance/results/change-0938/comparison-runner.py"
RUNNER_SHA256 = "1b660ffbc21b2e45e9748b076583c215ce36197c0eb5a94d8224fcbc227acc41"
FROZEN_VALIDATOR_SHA256 = "b70a2b38d098e3abd2b567d1bea70f40e9d27e62acaca8f80ff633dd2c596712"
MAX_JSON_BYTES = 1 << 20


class RecoveryError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RecoveryError(message)


def sha256_path(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, *, root: Path | None = None) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    value: dict[str, Any] = {
        "path": str(path if root is None else path.relative_to(root.resolve())),
        "bytes": path.stat().st_size,
        "sha256": sha256_path(path),
    }
    return value


def verify_descriptor(value: Any, *, label: str, root: Path | None = None) -> Path:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}: descriptor path missing")
    path = Path(raw)
    if root is not None:
        root = root.resolve()
        require(not path.is_absolute(), f"{label}: packet path is absolute")
        path = (root / path).resolve()
        require(path.is_relative_to(root), f"{label}: packet path escapes root")
    else:
        path = path.resolve()
    actual = descriptor(path)
    require(actual["bytes"] == value.get("bytes") and actual["sha256"] == value.get("sha256"),
            f"{label}: descriptor changed")
    return path


def read_json(path: Path, *, label: str) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"{label}: missing JSON")
    require(path.stat().st_size < MAX_JSON_BYTES, f"{label}: JSON is at least 1 MiB")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise RecoveryError(f"{label}: invalid JSON: {error}") from error
    require(isinstance(value, dict), f"{label}: JSON is not an object")
    return value


def load_frozen_runner() -> Any:
    require(sha256_path(RUNNER_PATH) == RUNNER_SHA256, "frozen 0938 runner hash changed")
    spec = importlib.util.spec_from_file_location("comparison_runner_0938_frozen", RUNNER_PATH)
    require(spec is not None and spec.loader is not None, "cannot load frozen runner")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def verify_authorized_manifest(runner: Any, manifest: Path, packet: Path) -> dict[str, Any]:
    previous = os.environ.get("LITCHI_PERF_0938_CAPTURE_APPROVED")
    os.environ["LITCHI_PERF_0938_CAPTURE_APPROVED"] = "1"
    try:
        value = runner.verify_manifest(manifest, authorized=True, packet=packet)
    finally:
        if previous is None:
            os.environ.pop("LITCHI_PERF_0938_CAPTURE_APPROVED", None)
        else:
            os.environ["LITCHI_PERF_0938_CAPTURE_APPROVED"] = previous
    return value


def normalized_command(command: Any, *, report_path: Path) -> list[str]:
    require(isinstance(command, list) and all(isinstance(value, str) for value in command),
            "terminal argv is not a string list")
    normalized = list(command)
    require("--json" in normalized, "terminal argv has no report argument")
    index = normalized.index("--json")
    require(index + 1 < len(normalized), "terminal argv report argument is missing")
    require(Path(normalized[index + 1]).name == report_path.name,
            "terminal argv report basename changed")
    normalized[index + 1] = "<report.json>"
    return normalized


def recover(args: argparse.Namespace) -> dict[str, Any]:
    packet = args.packet.resolve()
    manifest_path = args.manifest.resolve()
    outer_path = args.outer_terminal.resolve()
    output = args.output.resolve()
    require(not output.is_relative_to(REPO_ROOT), "recovery output must be external")
    require(not output.exists(), f"refusing to overwrite recovery output: {output}")
    runner = load_frozen_runner()
    require(sha256_path(runner.COMPARISON_PATH) == FROZEN_VALIDATOR_SHA256,
            "frozen 0937 validator dependency hash changed")
    manifest = verify_authorized_manifest(runner, manifest_path, packet)
    outer = read_json(outer_path, label="outer terminal")
    require(outer.get("status") == "completed_outer_runner_failed_children",
            "outer terminal status is not the recorded failed cohort")
    require(outer.get("outer_exit_code") == 0 and outer.get("children_run") == runner.CHILDREN,
            "outer terminal does not bind the completed 96-child invocation")
    require(outer.get("children_failed") == runner.CHILDREN
            and outer.get("retries") == 0 and outer.get("recaptures") == 0,
            "outer terminal failure/retry counts changed")
    state_path = packet / "capture-state.json"
    state = read_json(state_path, label="capture state")
    packet_manifest_descriptor = runner.descriptor(packet / "manifest.json")
    require(state.get("schema") == "litchi.performance.0938.capture-state.v1",
            "capture state schema changed")
    require(state.get("status") == "failed" and state.get("completed") == runner.CHILDREN
            and state.get("planned") == runner.CHILDREN and state.get("failed") == runner.CHILDREN,
            "capture state is not the retained failed full cohort")
    require(state.get("retries") == 0 and state.get("recaptures") == 0,
            "capture state contains retry/recapture")
    rows = runner.schedule_rows()
    expected = {row["id"]: row for row in rows}
    children = state.get("rows")
    require(isinstance(children, list) and len(children) == runner.CHILDREN,
            "capture state child rows changed")
    require({child.get("id") for child in children} == set(expected),
            "capture state child IDs changed")
    comparison = runner.load_comparison()
    require(sha256_path(runner.COMPARISON_PATH) == FROZEN_VALIDATOR_SHA256,
            "validator dependency changed during recovery")
    expected_output = comparison.OUTPUT_SHA256
    validated: list[dict[str, Any]] = []
    records: list[dict[str, Any]] = []
    failure_errors: dict[str, int] = {}
    for child in children:
        row = expected[child["id"]]
        require(child.get("status") == "failed" and child.get("exit_code") == 0,
                f"{row['id']}: retained child status changed")
        terminal_path = verify_descriptor(child.get("terminal"), label=f"{row['id']} terminal", root=packet)
        terminal = read_json(terminal_path, label=f"{row['id']} terminal")
        require(terminal.get("schedule") == row and terminal.get("label") == row["label"],
                f"{row['id']}: terminal schedule changed")
        require(terminal.get("exit_code") == 0 and terminal.get("deadline_exceeded") is False
                and terminal.get("runner_error") is None and terminal.get("process_group_closed") is True,
                f"{row['id']}: terminal process status changed")
        require(terminal.get("binary") == runner.descriptor(runner.BINARY),
                f"{row['id']}: terminal binary binding changed")
        require(terminal.get("source_revision") == runner.BASE_REVISION,
                f"{row['id']}: terminal source revision changed")
        report_path = verify_descriptor(child.get("report"), label=f"{row['id']} report", root=packet)
        require(terminal.get("report") == child.get("report"),
                f"{row['id']}: state/terminal report descriptor differs")
        require(terminal.get("manifest_sha256") == packet_manifest_descriptor["sha256"],
                f"{row['id']}: terminal packet-manifest binding changed")
        require(terminal.get("runner_sha256") == RUNNER_SHA256,
                f"{row['id']}: terminal runner binding changed")
        require(normalized_command(terminal.get("argv"), report_path=report_path)
                == normalized_command(runner.child_command(row, report_path), report_path=report_path),
                f"{row['id']}: terminal argv changed")
        require(terminal.get("environment") == runner.child_environment(row),
                f"{row['id']}: terminal environment changed")
        require(terminal.get("cwd") == str(runner.SOURCE_ROOT),
                f"{row['id']}: terminal cwd changed")
        for name in ("stdout", "stderr", "time"):
            verify_descriptor(terminal.get(name), label=f"{row['id']} {name}", root=packet)
        metrics = runner.parse_time_metrics(verify_descriptor(terminal.get("time"),
                                                               label=f"{row['id']} time", root=packet))
        require(metrics == terminal.get("time_metrics"), f"{row['id']}: time metrics changed")
        validation_path = verify_descriptor(child.get("validation"),
                                             label=f"{row['id']} validation", root=packet)
        validation = read_json(validation_path, label=f"{row['id']} validation")
        error = validation.get("error")
        require(validation.get("status") == "failed" and error == "KeyError: 'output_sha256'",
                f"{row['id']}: frozen validator failure changed")
        failure_errors[error] = failure_errors.get(error, 0) + 1
        report = runner.read_json(report_path)
        result = comparison.validate_report(report, row["route"], warmups=runner.WARMUPS,
                                            samples=runner.SAMPLES, label=str(report_path),
                                            binary=runner.descriptor(runner.BINARY))
        report_result = report["results"][0]
        section = report_result["source"]["docx_section_layout"]
        require(report_result.get("output_sha256") == expected_output,
                f"{row['id']}: report output identity changed")
        require(section.get("output_sha256") == [expected_output] * runner.SAMPLES,
                f"{row['id']}: per-sample output identity changed")
        validated.append({"row": row, "result": result, "time": metrics})
        records.append({
            "id": row["id"], "pair_id": row["pair_id"], "block": row["block"],
            "route": row["route"], "label": row["label"],
            "report": child["report"], "terminal": child["terminal"],
            "validation": child["validation"], "time_metrics": metrics,
            "raw_report_revalidated": True,
            "frozen_validation_error": error,
            "output_sha256": expected_output,
        })
    reduction = runner.reduce_validated(validated)
    result: dict[str, Any] = {
        "schema": "litchi.performance.0938.offline-recovery.v1",
        "status": "recovered_raw_reports_validator_contract_failed",
        "scope": "Descriptive reduction of the one completed 96-child packet; no adoption or causal claim.",
        "packet": str(packet),
        "manifest": runner.descriptor(manifest_path),
        "packet_manifest": runner.descriptor(packet / "manifest.json"),
        "capture_state": runner.descriptor(state_path),
        "outer_terminal": runner.descriptor(outer_path),
        "children": runner.CHILDREN,
        "raw_reports_revalidated": len(records),
        "frozen_validation_errors": failure_errors,
        "frozen_validator": {
            "path": str(runner.COMPARISON_PATH),
            "sha256": FROZEN_VALIDATOR_SHA256,
            "failure": "KeyError: 'output_sha256'",
            "correction": "The report output identity is checked independently; the corrected validator result contract is used only offline.",
        },
        "records": records,
        "reduction": reduction,
        "no_retry_or_recapture": True,
    }
    raw = (json.dumps(result, indent=2, sort_keys=True) + "\n").encode()
    require(len(raw) < MAX_JSON_BYTES, "recovery output is at least 1 MiB")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("xb") as stream:
        stream.write(raw)
        stream.flush()
        os.fsync(stream.fileno())
    result["output"] = descriptor(output)
    return result


def parser() -> argparse.ArgumentParser:
    value = argparse.ArgumentParser(description=__doc__)
    value.add_argument("--manifest", type=Path, required=True)
    value.add_argument("--packet", type=Path, required=True)
    value.add_argument("--outer-terminal", type=Path, required=True)
    value.add_argument("--output", type=Path, required=True)
    return value


def main() -> int:
    args = parser().parse_args()
    print(json.dumps(recover(args), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"0938 offline recovery failed: {type(error).__name__}: {error}", file=sys.stderr)
        raise
