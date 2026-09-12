#!/usr/bin/env python3
"""Capture and compare the bounded dense-sparse eager confirmation lane.

This wrapper is intentionally separate from the frozen 0525 primary driver and
the original eager guard.  It consumes the already-built normal binaries,
records an explicit A1/B1/B2/A2 custody chain, and compares only total elapsed
time and whole-child RSS.  The raw eager schema has no source or phase vectors;
that absence is an oracle rather than a missing measurement.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "eager-confirmation-plan.json"
RUN_PATH = HERE / "run.py"
REFERENCE_FALLBACK = HERE / "baseline" / "eager-r2-dense-sparse.json"
TIMING_STATS = ("p50", "p95", "p99", "mean")
EXPECTED_ORDER = ("A1", "B1", "B2", "A2")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory confirmation artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n",
                    encoding="utf-8")


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside confirmation bundle: {path}") from error


def load_module(path: Path, name: str) -> Any:
    require(path.is_file(), f"missing helper {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None,
            f"cannot load helper {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_run() -> Any:
    require(RUN_PATH.is_file(), f"missing frozen driver {RUN_PATH}")
    spec = importlib.util.spec_from_file_location("xlsx_0525_confirmation_run", RUN_PATH)
    require(spec is not None and spec.loader is not None,
            f"cannot load frozen driver {RUN_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


RUN = load_run()
REPO = RUN.REPO
COMPANION = load_module(HERE / "analyze_eager_guard.py",
                        "xlsx_0525_eager_confirmation_companion")
BASE = COMPANION.BASE


def plan_data() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "confirmation plan is not an object")
    require(plan.get("schema") == "litchi-0525-eager-confirmation-plan-v1",
            "confirmation plan schema differs")
    require(plan.get("status") == "frozen-supplemental-before-capture",
            "confirmation plan is not frozen")
    require(plan.get("primary_plan_sha256") == sha(HERE / "plan.json"),
            "confirmation plan is bound to a different primary plan")
    require(plan.get("capture_wrapper") == Path(__file__).name,
            "confirmation wrapper name is not bound")
    helpers = plan.get("retained_helpers")
    require(isinstance(helpers, dict), "confirmation retained helper bindings are missing")
    for name, helper in helpers.items():
        require(isinstance(helper, dict) and isinstance(helper.get("path"), str),
                f"confirmation helper binding is malformed: {name}")
        helper_path = HERE / helper["path"]
        require(helper.get("sha256") == sha(helper_path),
                f"confirmation helper binding differs: {name}")
    require(plan.get("order") == list(EXPECTED_ORDER),
            "confirmation ABBA order differs")
    require(plan.get("case") == "xlsx_eager_cell_values_one_percent_edit_save",
            "confirmation case differs")
    require(plan.get("shape") == "dense-sparse", "confirmation shape differs")
    require(plan.get("cpu") == 2 and plan.get("warmup") == 20
            and plan.get("samples") == 100,
            "confirmation CPU/warmup/sample counts differ")
    children = plan.get("children")
    require(isinstance(children, list) and len(children) == len(EXPECTED_ORDER),
            "confirmation child matrix is incomplete")
    labels = [item.get("label") for item in children if isinstance(item, dict)]
    require(labels == list(EXPECTED_ORDER), "confirmation child labels differ")
    require(plan.get("oracle_reference") == relative(REFERENCE_FALLBACK),
            "confirmation oracle reference differs")
    reference_path = HERE / plan["oracle_reference"]
    require(reference_path.is_file(), "confirmation oracle reference is missing")
    require(plan.get("oracle_reference_sha256") == sha(reference_path),
            "confirmation oracle reference digest differs")
    require(plan.get("tmpdir") == str(Path(plan["owned_paths"][1]) / "test-tmp"),
            "confirmation TMPDIR is not target-owned")
    custody = plan.get("custody")
    require(isinstance(custody, dict)
            and custody.get("fresh_child_per_capture") is True
            and custody.get("fresh_child_per_sample") is False,
            "confirmation custody must be one fresh child per capture")
    return plan


def child_data(plan: dict[str, Any], label: str) -> dict[str, Any]:
    require(label in EXPECTED_ORDER, f"unknown confirmation label: {label}")
    for child in plan["children"]:
        if child["label"] == label:
            require(child["stage"] in ("baseline", "candidate"),
                    f"{label}: stage differs")
            require(child["repeat"] in (1, 2), f"{label}: repeat differs")
            require(isinstance(child.get("source_manifest"), str)
                    and isinstance(child.get("working_source_manifest"), str),
                    f"{label}: source manifest bindings are incomplete")
            return child
    raise EvidenceError(f"confirmation child is missing: {label}")


def manifest_value(path: Path, label: str) -> dict[str, str]:
    require(path.is_file() and not path.is_symlink(),
            f"{label}: source manifest is missing or non-regular")
    manifest = read_json(path)
    require(isinstance(manifest, dict) and manifest,
            f"{label}: source manifest is empty")
    require(all(isinstance(name, str) and isinstance(digest, str)
                and len(digest) == 64 for name, digest in manifest.items()),
            f"{label}: source manifest entry is malformed")
    return manifest


def source_manifest_snapshot(plan: dict[str, Any], child: dict[str, Any]) -> dict[str, str]:
    build_path = HERE / child["source_manifest"]
    working_path = HERE / child["working_source_manifest"]
    build_manifest = manifest_value(build_path, f"{child['label']} build")
    working_manifest = manifest_value(working_path, f"{child['label']} working")
    actual = {name for name in RUN.manifest_paths() if (REPO / name).is_file()}
    require(set(working_manifest) == actual,
            f"{child['label']}: working source manifest inventory differs from live source")
    for name, digest in working_manifest.items():
        source = REPO / name
        require(source.is_file() and not source.is_symlink(),
                f"{child['label']}: source is missing or non-regular: {name}")
        require(isinstance(digest, str) and sha(source) == digest,
                f"{child['label']}: working source digest differs: {name}")
    return {
        "build_source_manifest_sha256": sha(build_path),
        "working_source_manifest_sha256": sha(working_path),
    }


def binary_snapshot(plan: dict[str, Any], child: dict[str, Any],
                    manifests: dict[str, str], require_live: bool) -> dict[str, Any]:
    identity_path = HERE / child["binary_identity"]
    identity = read_json(identity_path)
    require(isinstance(identity, dict), f"{child['label']}: binary identity is not an object")
    path = Path(child["binary_path"])
    require(identity.get("path") == str(path),
            f"{child['label']}: binary path differs from frozen plan")
    digest = identity.get("sha256")
    require(isinstance(digest, str) and len(digest) == 64,
            f"{child['label']}: binary digest is malformed")
    require(identity.get("source_manifest_sha256") ==
            manifests["build_source_manifest_sha256"],
            f"{child['label']}: binary manifest binding differs")
    live = path.is_file() and not path.is_symlink()
    if require_live:
        require(live, f"{child['label']}: existing normal binary is missing or non-regular")
    elif not live:
        cleanup_path = HERE / "cleanup.json"
        cleanup = read_json(cleanup_path)
        require(cleanup.get("owned_paths_absent") is True
                and cleanup.get("accessible_process_references") == []
                and cleanup.get("removed") == plan["owned_paths"]
                and all(not Path(name).exists() for name in plan["owned_paths"]),
                f"{child['label']}: missing binary has no completed cleanup custody")
    if live:
        require(sha(path) == digest and identity.get("bytes") == path.stat().st_size,
                f"{child['label']}: live normal binary custody differs")
    return {
        "path": str(path),
        "sha256": digest,
        "bytes": identity["bytes"],
        "identity_path": relative(identity_path),
        "identity_sha256": sha(identity_path),
        "live": live,
    }


def reference_result(plan: dict[str, Any]) -> dict[str, Any]:
    raw = read_json(HERE / plan["oracle_reference"])
    require(isinstance(raw, dict), "confirmation oracle report is not an object")
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1,
            "confirmation oracle report result matrix differs")
    result = results[0]
    require(isinstance(result, dict), "confirmation oracle result is not an object")
    require("source" not in result, "confirmation oracle unexpectedly has source evidence")
    return result


def validate_raw(path: Path, label: str, plan: dict[str, Any],
                 child: dict[str, Any], binary: dict[str, Any],
                 oracle: dict[str, Any]) -> dict[str, Any]:
    raw = read_json(path)
    job = {
        "name": label,
        "kind": "primary",
        "guard": None,
        "repeat": child["repeat"],
        "case": plan["case"],
        "shape": plan["shape"],
        "warmup": plan["warmup"],
        "samples": plan["samples"],
    }
    # Reuse the companion's complete eager envelope, canonical elapsed
    # statistic recomputation, not-applicable operation sections, sink vectors,
    # corpus, and parallel-metric validators.  The extra oracle below binds
    # this confirmation to the already captured dense-sparse identity.
    COMPANION.validate_eager_report(
        raw, read_json(HERE / "plan.json"), plan, job,
        {"sha256": binary["sha256"], "path": binary["path"]},
    )
    result_list = raw.get("results")
    require(isinstance(result_list, list) and len(result_list) == 1,
            f"{label}: result matrix differs")
    result = result_list[0]
    require(isinstance(result, dict), f"{label}: result is not an object")
    require(result.get("corpus") == oracle.get("corpus"),
            f"{label}: corpus oracle differs")
    require(result.get("sink") == oracle.get("sink"),
            f"{label}: sink oracle differs")
    require(result.get("output_sha256") == oracle.get("output_sha256"),
            f"{label}: output digest oracle differs")
    return result


def command_for(plan: dict[str, Any], child: dict[str, Any], output: Path,
                rss: Path) -> list[str]:
    binary = child["binary_path"]
    return [
        "taskset", "-c", str(plan["cpu"]), "/usr/bin/time", "-f",
        '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
        "-o", str(rss), binary, "--warmup", str(plan["warmup"]),
        "--samples", str(plan["samples"]), "--case", plan["case"],
        "--xlsx-cell-crud-shape", plan["shape"], "--json", str(output),
    ]


def state_path(plan: dict[str, Any]) -> Path:
    return HERE / plan["capture_root"] / "capture-state.json"


def read_state(plan: dict[str, Any]) -> dict[str, Any] | None:
    path = state_path(plan)
    if not path.is_file():
        return None
    value = read_json(path)
    require(isinstance(value, dict), "confirmation capture state is not an object")
    require(value.get("schema") == "litchi-0525-eager-confirmation-state-v1"
            and value.get("plan_sha256") == sha(PLAN_PATH),
            "confirmation capture state binding differs")
    completed = value.get("completed")
    require(isinstance(completed, list)
            and completed == list(EXPECTED_ORDER[:len(completed)]),
            "confirmation capture state order differs")
    return value


def write_state(plan: dict[str, Any], completed: list[str],
                entries: dict[str, Any], failed: str | None = None) -> None:
    path = state_path(plan)
    value = {
        "schema": "litchi-0525-eager-confirmation-state-v1",
        "plan_sha256": sha(PLAN_PATH),
        "order": list(EXPECTED_ORDER),
        "completed": completed,
        "status": "complete" if len(completed) == len(EXPECTED_ORDER) else "in_progress",
        "entries": entries,
    }
    if failed is not None:
        value["failed_label"] = failed
        value["status"] = "failed"
    write_json(path, value)


def capture(label: str) -> None:
    plan = plan_data()
    child = child_data(plan, label)
    state = read_state(plan)
    completed = [] if state is None else list(state["completed"])
    require(len(completed) < len(EXPECTED_ORDER), "confirmation ABBA sequence is already complete")
    require(label == EXPECTED_ORDER[len(completed)],
            f"confirmation next label must be {EXPECTED_ORDER[len(completed)]}")
    root = HERE / plan["capture_root"]
    folder = root / label
    folder.mkdir(parents=True, exist_ok=False)
    output_path = folder / f"{label}.json"
    rss_path = folder / f"{label}.rss.json"
    receipt_path = folder / f"{label}.receipt.json"
    binding_path = folder / f"{label}.binding.json"
    stdout_path = folder / f"{label}.stdout"
    stderr_path = folder / f"{label}.stderr"
    manifest_before = source_manifest_snapshot(plan, child)
    binary_before = binary_snapshot(plan, child, manifest_before, True)
    command = command_for(plan, child, output_path, rss_path)
    environment = dict(os.environ)
    environment["TMPDIR"] = plan["tmpdir"]
    Path(plan["tmpdir"]).mkdir(parents=True, exist_ok=True)
    start = datetime.datetime.now(datetime.timezone.utc).isoformat()
    tick = time.monotonic()
    with stdout_path.open("w", encoding="utf-8") as stdout, \
            stderr_path.open("w", encoding="utf-8") as stderr:
        result = subprocess.run(command, cwd=REPO, stdout=stdout, stderr=stderr,
                                env=environment, check=False)
    seconds = time.monotonic() - tick
    end = datetime.datetime.now(datetime.timezone.utc).isoformat()
    manifest_after = source_manifest_snapshot(plan, child)
    binary_after = binary_snapshot(plan, child, manifest_after, True)
    require(manifest_before == manifest_after,
            f"{label}: source manifest changed during child")
    require(binary_before["sha256"] == binary_after["sha256"],
            f"{label}: binary changed during child")
    artifacts = {
        relative(path): sha(path)
        for path in (output_path, rss_path, stdout_path, stderr_path)
        if path.is_file()
    }
    receipt = {
        "schema": "litchi-0525-eager-confirmation-receipt-v1",
        "label": label,
        "stage": child["stage"],
        "repeat": child["repeat"],
        "command": command,
        "start_utc": start,
        "end_utc": end,
        "seconds": seconds,
        "exit_code": result.returncode,
        "primary_plan_sha256": sha(HERE / "plan.json"),
        "supplemental_plan_sha256": sha(PLAN_PATH),
        "run_script_sha256": sha(RUN_PATH),
        "wrapper_sha256": sha(Path(__file__)),
        "retained_helpers": plan["retained_helpers"],
        "source_manifest": child["source_manifest"],
        "working_source_manifest": child["working_source_manifest"],
        "source_manifest_sha256": manifest_after["build_source_manifest_sha256"],
        "working_source_manifest_sha256": manifest_after["working_source_manifest_sha256"],
        "binary_path": child["binary_path"],
        "binary_sha256": binary_after["sha256"],
        "environment": {key: environment.get(key)
                        for key in ("TMPDIR", "RUSTFLAGS", "LD_PRELOAD",
                                    "MALLOC_CONF", "GLIBC_TUNABLES")},
        "artifacts": artifacts,
    }
    write_json(receipt_path, receipt)
    validation_error: str | None = None
    if result.returncode == 0:
        try:
            oracle = reference_result(plan)
            validate_raw(output_path, label, plan, child, binary_after, oracle)
            BASE.validate_rss(rss_path)
        except ValueError as error:
            validation_error = str(error)
    else:
        validation_error = f"child exited with code {result.returncode}"
    receipt["validation"] = "pass" if validation_error is None else "failed"
    if validation_error is not None:
        receipt["validation_error"] = validation_error
    write_json(receipt_path, receipt)
    binding = {
        "schema": "litchi-0525-eager-confirmation-binding-v1",
        "label": label,
        "stage": child["stage"],
        "repeat": child["repeat"],
        "primary_plan_sha256": sha(HERE / "plan.json"),
        "supplemental_plan_sha256": sha(PLAN_PATH),
        "run_script_sha256": sha(RUN_PATH),
        "wrapper_sha256": sha(Path(__file__)),
        "retained_helpers": plan["retained_helpers"],
        "source_manifest_sha256": manifest_after["build_source_manifest_sha256"],
        "working_source_manifest_sha256": manifest_after["working_source_manifest_sha256"],
        "binary_sha256": binary_after["sha256"],
        "receipt_sha256": sha(receipt_path),
        "artifacts": artifacts,
        "output": relative(output_path),
        "rss": relative(rss_path),
        "validation": "pass" if validation_error is None else "failed",
    }
    if validation_error is not None:
        binding["validation_error"] = validation_error
    write_json(binding_path, binding)
    entries = {} if state is None else dict(state.get("entries", {}))
    entries[label] = {
        "receipt": relative(receipt_path),
        "receipt_sha256": sha(receipt_path),
        "binding": relative(binding_path),
        "binding_sha256": sha(binding_path),
        "stage": child["stage"],
        "repeat": child["repeat"],
    }
    if validation_error is not None:
        write_state(plan, completed, entries, label)
        raise EvidenceError(f"{label}: {validation_error}; receipt and binding retained")
    completed.append(label)
    write_state(plan, completed, entries)
    print(f"confirmation child {label} passed", flush=True)


def validate_capture(plan: dict[str, Any], label: str) -> dict[str, Any]:
    child = child_data(plan, label)
    folder = HERE / plan["capture_root"] / label
    receipt_path = folder / f"{label}.receipt.json"
    binding_path = folder / f"{label}.binding.json"
    receipt = read_json(receipt_path)
    binding = read_json(binding_path)
    require(receipt.get("schema") == "litchi-0525-eager-confirmation-receipt-v1",
            f"{label}: receipt schema differs")
    require(binding.get("schema") == "litchi-0525-eager-confirmation-binding-v1",
            f"{label}: binding schema differs")
    require(receipt.get("label") == label and binding.get("label") == label,
            f"{label}: custody label differs")
    require(receipt.get("stage") == child["stage"] and receipt.get("repeat") == child["repeat"],
            f"{label}: receipt child identity differs")
    require(receipt.get("source_manifest") == child["source_manifest"]
            and receipt.get("working_source_manifest") == child["working_source_manifest"]
            and receipt.get("binary_path") == child["binary_path"],
            f"{label}: receipt source/binary paths differ")
    require(binding.get("stage") == child["stage"]
            and binding.get("repeat") == child["repeat"],
            f"{label}: binding child identity differs")
    require(receipt.get("exit_code") == 0 and receipt.get("validation") == "pass",
            f"{label}: child receipt is not a validated pass")
    require(receipt.get("primary_plan_sha256") == sha(HERE / "plan.json")
            and receipt.get("supplemental_plan_sha256") == sha(PLAN_PATH)
            and receipt.get("run_script_sha256") == sha(RUN_PATH)
            and receipt.get("wrapper_sha256") == sha(Path(__file__))
            and receipt.get("retained_helpers") == plan["retained_helpers"],
            f"{label}: receipt code/plan custody differs")
    require(binding.get("primary_plan_sha256") == sha(HERE / "plan.json")
            and binding.get("supplemental_plan_sha256") == sha(PLAN_PATH)
            and binding.get("run_script_sha256") == sha(RUN_PATH)
            and binding.get("wrapper_sha256") == sha(Path(__file__))
            and binding.get("retained_helpers") == plan["retained_helpers"],
            f"{label}: binding code/plan custody differs")
    manifests = source_manifest_snapshot(plan, child)
    binary = binary_snapshot(plan, child, manifests, False)
    require(receipt.get("source_manifest_sha256") ==
            manifests["build_source_manifest_sha256"]
            and receipt.get("working_source_manifest_sha256") ==
            manifests["working_source_manifest_sha256"]
            and receipt.get("binary_sha256") == binary["sha256"],
            f"{label}: receipt source/binary binding differs")
    require(binding.get("source_manifest_sha256") ==
            manifests["build_source_manifest_sha256"]
            and binding.get("working_source_manifest_sha256") ==
            manifests["working_source_manifest_sha256"]
            and binding.get("binary_sha256") == binary["sha256"]
            and binding.get("receipt_sha256") == sha(receipt_path),
            f"{label}: binding source/binary/receipt differs")
    command = receipt.get("command")
    require(isinstance(command, list), f"{label}: receipt command is not a list")
    output_path = folder / f"{label}.json"
    rss_path = folder / f"{label}.rss.json"
    expected = command_for(plan, child, output_path, rss_path)
    require(command == expected, f"{label}: receipt command differs")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict)
            and set(artifacts) == {relative(output_path), relative(rss_path),
                                   relative(folder / f"{label}.stdout"),
                                   relative(folder / f"{label}.stderr")},
            f"{label}: receipt artifact inventory differs")
    for name, digest in artifacts.items():
        path = HERE / name
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{label}: artifact custody differs: {name}")
    oracle = reference_result(plan)
    result = validate_raw(output_path, label, plan, child, binary, oracle)
    rss = BASE.validate_rss(rss_path)
    return {
        "label": label,
        "stage": child["stage"],
        "repeat": child["repeat"],
        "report_sha256": sha(output_path),
        "receipt_sha256": sha(receipt_path),
        "binding_sha256": sha(binding_path),
        "binary_sha256": binary["sha256"],
        "source_manifest_sha256": manifests["build_source_manifest_sha256"],
        "working_source_manifest_sha256": manifests["working_source_manifest_sha256"],
        "elapsed_ns": result["elapsed_ns"],
        "rss": rss,
        "corpus": result["corpus"],
        "sink": result["sink"],
        "output_sha256": result["output_sha256"],
    }


def percent(left: float, right: float) -> float:
    require(left > 0, "comparison baseline metric is not positive")
    return (right / left - 1.0) * 100.0


def metric(left: Any, right: Any) -> dict[str, Any]:
    change = percent(float(left), float(right))
    return {
        "baseline": left,
        "candidate": right,
        "change_percent": change,
        "adverse_over_five_percent": change > 5.0,
    }


def compare_pair(left: dict[str, Any], right: dict[str, Any],
                 pair: str, adverse: list[dict[str, Any]]) -> dict[str, Any]:
    metrics = {"elapsed_ns": {}}
    for stat in TIMING_STATS:
        value = metric(left["elapsed_ns"][stat], right["elapsed_ns"][stat])
        metrics["elapsed_ns"][stat] = value
        if value["adverse_over_five_percent"]:
            adverse.append({"pair": pair, "metric": "elapsed_ns", "stat": stat, **value})
    rss = metric(left["rss"]["max_rss_kib"], right["rss"]["max_rss_kib"])
    metrics["max_rss_kib"] = rss
    if rss["adverse_over_five_percent"]:
        adverse.append({"pair": pair, "metric": "max_rss_kib", "stat": "peak", **rss})
    return {
        "pair": pair,
        "baseline_label": left["label"],
        "candidate_label": right["label"],
        "identity_equal": (left["corpus"] == right["corpus"]
                            and left["sink"] == right["sink"]
                            and left["output_sha256"] == right["output_sha256"]),
        "metrics": metrics,
        "baseline_raw_ranked_samples": left["elapsed_ns"]["samples"],
        "candidate_raw_ranked_samples": right["elapsed_ns"]["samples"],
    }


def compare_drift(left: dict[str, Any], right: dict[str, Any], stage: str,
                  adverse: list[dict[str, Any]]) -> dict[str, Any]:
    metrics = {"elapsed_ns": {}}
    for stat in TIMING_STATS:
        value = metric(left["elapsed_ns"][stat], right["elapsed_ns"][stat])
        metrics["elapsed_ns"][stat] = value
        if abs(value["change_percent"]) > 5.0:
            adverse.append({"stage": stage, "repeat_first": left["repeat"],
                            "repeat_second": right["repeat"],
                            "metric": "elapsed_ns", "stat": stat, **value})
    rss = metric(left["rss"]["max_rss_kib"], right["rss"]["max_rss_kib"])
    metrics["max_rss_kib"] = rss
    if abs(rss["change_percent"]) > 5.0:
        adverse.append({"stage": stage, "repeat_first": left["repeat"],
                        "repeat_second": right["repeat"],
                        "metric": "max_rss_kib", "stat": "peak", **rss})
    return {"stage": stage, "repeat_first": left["repeat"],
            "repeat_second": right["repeat"], "metrics": metrics,
            "first_raw_ranked_samples": left["elapsed_ns"]["samples"],
            "second_raw_ranked_samples": right["elapsed_ns"]["samples"]}


def analyze() -> dict[str, Any]:
    plan = plan_data()
    state = read_state(plan)
    require(state is not None and state.get("completed") == list(EXPECTED_ORDER),
            "confirmation ABBA capture is incomplete")
    rows = {label: validate_capture(plan, label) for label in EXPECTED_ORDER}
    adverse: list[dict[str, Any]] = []
    matched = [
        compare_pair(rows["A1"], rows["B1"], "A1_vs_B1", adverse),
        compare_pair(rows["A2"], rows["B2"], "A2_vs_B2", adverse),
    ]
    drift: list[dict[str, Any]] = []
    same_build = [
        compare_drift(rows["A1"], rows["A2"], "baseline", drift),
        compare_drift(rows["B1"], rows["B2"], "candidate", drift),
    ]
    pair_gate = []
    for row in matched:
        p50 = row["metrics"]["elapsed_ns"]["p50"]
        mean = row["metrics"]["elapsed_ns"]["mean"]
        pair_gate.append({
            "pair": row["pair"],
            "p50": p50,
            "mean": mean,
            "required_max_adverse_change_percent": 5.0,
            "passed": p50["change_percent"] <= 5.0 and mean["change_percent"] <= 5.0,
        })
    output = {
        "schema": "litchi-0525-eager-confirmation-comparison-v1",
        "status": "pass",
        "stage": "compare",
        "supplemental_plan_sha256": sha(PLAN_PATH),
        "primary_plan_sha256": sha(HERE / "plan.json"),
        "capture_wrapper": {"path": Path(__file__).name,
                             "sha256": sha(Path(__file__))},
        "order": list(EXPECTED_ORDER),
        "children": [rows[label] for label in EXPECTED_ORDER],
        "matched_pairs": matched,
        "same_build_drift": same_build,
        "adverse_flags_over_five_percent": adverse,
        "same_build_drift_over_five_percent": drift,
        "gate": {
            "paired_rows": pair_gate,
            "passed": all(row["passed"] for row in pair_gate),
            "scope": "Both A/B confirmation pairs require elapsed p50 and mean candidate change <=5%; no primary gain claim.",
        },
        "all_adverse_metrics_retained": True,
        "no_primary_gain_claim": True,
        "metrics_scope": {
            "elapsed_statistics": list(TIMING_STATS),
            "rss": "whole_child_process.max_rss_kib",
            "phase_metrics": "not_applicable",
        },
    }
    output_path = HERE / plan["comparison_output"]
    write_json(output_path, output)
    print(f"confirmation comparison verified: {output_path}")
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("capture", "validate", "analyze"))
    parser.add_argument("label", nargs="?", choices=EXPECTED_ORDER)
    args = parser.parse_args()
    try:
        if args.action == "capture":
            require(args.label is not None, "capture requires an ABBA label")
            capture(args.label)
        elif args.action == "validate":
            plan = plan_data()
            require(args.label is not None, "validate requires an ABBA label")
            validate_capture(plan, args.label)
            print(f"confirmation child {args.label} validated")
        else:
            require(args.label is None, "analyze takes no label")
            analyze()
    except (EvidenceError, ValueError, OSError, json.JSONDecodeError,
            subprocess.SubprocessError) as error:
        print(f"eager_confirmation_guard.py: error: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
