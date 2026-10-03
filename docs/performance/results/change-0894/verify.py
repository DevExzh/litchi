#!/usr/bin/env python3
"""Independently replay custody and semantic checks for an 0894 capture.

This verifier launches no workload.  It reopens every retained receipt and
report, recomputes artifact identities, reconstructs the frozen command, and
calls the strict child validator from :mod:`experiment`, including every 0893
worksheet guard vector on all primary children.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
_EXPERIMENT_PATH = PACKET / "experiment.py"
_EXPERIMENT_SPEC = importlib.util.spec_from_file_location("litchi_selected_metadata_experiment_0894_verify", _EXPERIMENT_PATH)
if _EXPERIMENT_SPEC is None or _EXPERIMENT_SPEC.loader is None:
    raise ImportError(f"cannot load local experiment helper: {_EXPERIMENT_PATH}")
experiment = importlib.util.module_from_spec(_EXPERIMENT_SPEC)
sys.modules[_EXPERIMENT_SPEC.name] = experiment
_EXPERIMENT_SPEC.loader.exec_module(experiment)


SCHEMA = "litchi.xlsx-selected-metadata-verification-0894.v1"


def _read(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def _sha(path: Path) -> str:
    return experiment._sha(path)


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def _descriptor_matches(raw: Any, path: Path, label: str) -> None:
    _require(isinstance(raw, dict), f"{label} descriptor is missing")
    actual = experiment.descriptor(path)
    for key in ("path", "bytes", "sha256", "mode", "executable"):
        _require(raw.get(key) == actual[key], f"{label} {key} mismatch")


def _regular_path(path: Path, label: str, *, directory: bool = False) -> Path:
    path = Path(path)
    _require(not path.is_symlink(), f"{label} must not be a symlink: {path}")
    _require(path.is_dir() if directory else path.is_file(), f"{label} is missing: {path}")
    return path.resolve()


def _time(value: str) -> _datetime.datetime:
    return _datetime.datetime.fromisoformat(value)


def _verify_plan_receipts(plan: dict[str, Any]) -> None:
    source = plan.get("source", {})
    source_ref = source.get("receipt") if isinstance(source, dict) else None
    _require(isinstance(source_ref, dict), "plan source receipt descriptor is missing")
    source_path = _regular_path(Path(str(source_ref.get("path", ""))), "plan source receipt")
    _descriptor_matches(source_ref, source_path, "source receipt")
    builds = plan.get("builds")
    _require(isinstance(builds, dict), "plan build receipt descriptors are missing")
    for role in ("before", "after"):
        raw = builds.get(role)
        _require(isinstance(raw, dict), f"plan {role} build receipt descriptor is missing")
        path = _regular_path(Path(str(raw.get("path", ""))), f"plan {role} build receipt")
        _descriptor_matches(raw, path, f"{role} build receipt")


def _verify_child(plan_path: Path, plan: dict[str, Any], root: Path, row: dict[str, Any], expected: dict[str, Any], binaries: dict[tuple[str, str], dict[str, Any]], previous_finished: _datetime.datetime | None) -> _datetime.datetime:
    mode = "native" if row["lane"] == "native" else "allocator"
    binary = binaries[(row["role"], mode)]
    folder = root / "runs" / f"{row['ordinal']:03d}-{row['lane']}-{row['pair']}-{row['role']}"
    _require(folder.is_dir(), f"missing child directory: {row['ordinal']}")
    receipt_path = folder / "receipt.json"
    command_path = folder / "command.json"
    report_path = folder / "report.json"
    peak_path = folder / "peak-rss-kib.txt"
    stdout_path = folder / "stdout.txt"
    stderr_path = folder / "stderr.txt"
    semantic_path = folder / "semantic.json"
    for path, label in ((receipt_path, "receipt"), (command_path, "command"), (report_path, "report"), (peak_path, "peak RSS"), (stdout_path, "stdout"), (stderr_path, "stderr"), (semantic_path, "semantic")):
        _require(path.is_file(), f"child {label} missing: {row['ordinal']}")
    receipt = _read(receipt_path)
    _require(receipt.get("schema") == experiment.RECEIPT_SCHEMA, f"receipt schema mismatch: {row['ordinal']}")
    _require(receipt.get("row") == row and receipt.get("exit_code") == 0, f"receipt row/status mismatch: {row['ordinal']}")
    plan_sha = _sha(plan_path)
    _require(receipt.get("plan_sha256") == plan_sha, f"receipt plan binding mismatch: {row['ordinal']}")
    _require(receipt.get("binary") == binary, f"receipt binary mismatch: {row['ordinal']}")
    command = _read(command_path)
    _require(command.get("schema") == experiment.COMMAND_SCHEMA, f"command schema mismatch: {row['ordinal']}")
    _require(command.get("row") == row and command.get("plan_sha256") == plan_sha, f"command identity mismatch: {row['ordinal']}")
    _require(command.get("binary") == binary and command.get("environment") == experiment.CHILD_ENV, f"command environment/binary mismatch: {row['ordinal']}")
    _require(command.get("cwd") == str(experiment.PACKET), f"command cwd mismatch: {row['ordinal']}")
    expected_argv = experiment._command(binary["path"], row["case"], row["warmup"], row["samples"], report_path, folder / "producer.json", peak_path)
    _require(command.get("argv") == expected_argv and receipt.get("argv") == expected_argv, f"command argv mismatch: {row['ordinal']}")
    _require(receipt.get("cwd") == command.get("cwd") == str(experiment.PACKET), f"receipt/command cwd mismatch: {row['ordinal']}")
    _require(receipt.get("environment") == command.get("environment") == experiment.CHILD_ENV, f"receipt/command environment mismatch: {row['ordinal']}")
    _require(receipt.get("exit_code") == command.get("exit_code") == 0, f"receipt/command exit mismatch: {row['ordinal']}")
    _require(receipt.get("started_utc") == command.get("started_utc") and receipt.get("finished_utc") == command.get("finished_utc"), f"receipt/command time binding mismatch: {row['ordinal']}")
    _descriptor_matches(receipt.get("command"), command_path, f"command {row['ordinal']}")
    artifacts = receipt.get("artifacts")
    _require(isinstance(artifacts, dict), f"artifact map missing: {row['ordinal']}")
    required = {"stdout.txt": stdout_path, "stderr.txt": stderr_path, "report.json": report_path, "peak-rss-kib.txt": peak_path}
    producer_path = folder / "producer.json"
    if producer_path.is_file():
        required["producer.json"] = producer_path
    _require(set(artifacts) == set(required), f"artifact set mismatch: {row['ordinal']}")
    for name, path in required.items():
        _descriptor_matches(artifacts.get(name), path, f"{row['ordinal']} {name}")
    semantic = _read(semantic_path)
    _require(semantic.get("status") == "pass", f"semantic receipt does not pass: {row['ordinal']}")
    _require(semantic.get("report_sha256") == _sha(report_path), f"semantic report hash mismatch: {row['ordinal']}")
    producer = _read(producer_path) if producer_path.is_file() else None
    _require(semantic.get("producer_sha256") == (_sha(producer_path) if producer_path.is_file() else None), f"semantic producer hash mismatch: {row['ordinal']}")
    if row["case"] == experiment.PRIMARY_CASE:
        _require(semantic.get("guard_contract") == experiment.GUARD_CONTRACT, f"semantic guard contract missing: {row['ordinal']}")
    else:
        _require(semantic.get("guard_contract") is None, f"unexpected semantic guard contract: {row['ordinal']}")
    report = _read(report_path)
    experiment._validate_child(report, producer, row["case"], mode, row["warmup"], row["samples"], expected[row["case"]], binary)
    command_started = _time(command["started_utc"])
    command_finished = _time(command["finished_utc"])
    _require(command_finished >= command_started, f"command interval invalid: {row['ordinal']}")
    started = _time(receipt["started_utc"])
    finished = _time(receipt["finished_utc"])
    _require(finished >= started, f"child interval invalid: {row['ordinal']}")
    if previous_finished is not None:
        _require(started >= previous_finished, f"children were not sequential: {row['ordinal']}")
    _require(experiment._read_peak(peak_path) > 0, f"child peak RSS is nonpositive: {row['ordinal']}")
    return finished


def verify(plan_path: Path, experiment_root: Path, output: Path | None = None) -> dict[str, Any]:
    plan_path = _regular_path(plan_path, "experiment plan")
    root = _regular_path(experiment_root, "experiment root", directory=True)
    plan = experiment.load_plan(plan_path)
    _verify_plan_receipts(plan)
    receipt_path = root / "experiment.json"
    _require(receipt_path.is_file(), "experiment.json is missing")
    receipt = _read(receipt_path)
    _require(receipt.get("schema") == experiment.EXPERIMENT_SCHEMA and receipt.get("status") == "pass", "experiment receipt is not passing")
    plan_sha = _sha(plan_path)
    _require(receipt.get("plan_sha256") == plan_sha, "experiment plan binding mismatch")
    rows = experiment._schedule()
    _require(receipt.get("children") == len(rows) == experiment.TOTAL_CHILDREN, "experiment child count mismatch")
    expected_rows = [{"path": f"runs/{row['ordinal']:03d}-{row['lane']}-{row['pair']}-{row['role']}/receipt.json", "row": row} for row in rows]
    _require(receipt.get("rows") == expected_rows, "experiment schedule receipt mismatch")
    expected = experiment._load_expected_corpora()
    binaries = {(role, mode): experiment._binary(plan, role, mode) for role in ("before", "after") for mode in ("native", "allocator")}
    previous_finished = None
    for row in rows:
        previous_finished = _verify_child(plan_path, plan, root, row, expected, binaries, previous_finished)
    runs_root = root / "runs"
    observed = {p.name for p in runs_root.iterdir() if p.is_dir()} if runs_root.is_dir() else set()
    expected_dirs = {f"{row['ordinal']:03d}-{row['lane']}-{row['pair']}-{row['role']}" for row in rows}
    _require(observed == expected_dirs, "capture contains unexpected or missing child directories")
    result = {
        "schema": SCHEMA,
        "status": "pass",
        "plan": experiment.descriptor(plan_path),
        "experiment": experiment.descriptor(receipt_path),
        "children": len(rows),
        "native_children": experiment.TOTAL_NATIVE_CHILDREN,
        "allocator_children": experiment.TOTAL_ALLOCATOR_CHILDREN,
        "primary_guard_children": sum(row["case"] == experiment.PRIMARY_CASE for row in rows),
        "guard_contract": experiment.GUARD_CONTRACT,
        "workload_launched": False,
        "performance_claim": None,
    }
    if output is not None:
        _require(not output.exists() and not output.is_symlink(), f"verification output already exists or is symlinked: {output}")
        _write_json(output, result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--experiment-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(verify(args.plan, args.experiment_root, args.output), indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
