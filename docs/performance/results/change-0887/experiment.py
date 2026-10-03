#!/usr/bin/env python3
"""Capture the 0887 inline-4 MCE candidate experiment.

The experiment is deliberately a capture driver only.  It launches one fresh
process per case and binary, retains raw evidence, and stops at the first
failed child or semantic mismatch.  :mod:`reduce` performs paired statistics
after all 72 children have been retained.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
BASELINE_PACKET = PACKET.parent / "change-0885"
BASELINE_QUALIFICATION = PACKET / "baseline-qualification"
sys.path.insert(0, str(BASELINE_PACKET))
import capture as baseline_capture  # noqa: E402
import verify as baseline_verify  # noqa: E402


SCHEMA = "litchi.xlsx-mce-inline4-experiment-0887.v1"
COMMAND_SCHEMA = "litchi.xlsx-mce-inline4-child-command-0887.v1"
RECEIPT_SCHEMA = "litchi.xlsx-mce-inline4-child-receipt-0887.v1"
EXPERIMENT_SCHEMA = "litchi.xlsx-mce-inline4-experiment-receipt-0887.v1"
PROFILE = "release"
CPU_AFFINITY = "12"
TASKSET = "/usr/bin/taskset"
TIME = "/usr/bin/time"
CASES = (
    "xlsx_producer_medium_source_full_cell_scan",
    "xlsx_full_cell_scan",
    "docx_producer_source_selected_paragraph",
    "pptx_producer_source_selected_slide",
)
PRIMARY_CASE = CASES[0]
NATIVE_BLOCKS = 8
ALLOCATOR_BLOCKS = 4
NATIVE_WARMUP = 3
NATIVE_SAMPLES = 50
ALLOCATOR_WARMUP = 3
ALLOCATOR_SAMPLES = 15
TOTAL_NATIVE_CHILDREN = NATIVE_BLOCKS * len(CASES) * 2
TOTAL_ALLOCATOR_CHILDREN = ALLOCATOR_BLOCKS * 2
TOTAL_CHILDREN = TOTAL_NATIVE_CHILDREN + TOTAL_ALLOCATOR_CHILDREN
CHILD_ENV = dict(getattr(baseline_capture, "CHILD_ENV", {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}))
BOOTSTRAP_SEED = 887
ARCHIVE_CASES = {case: BASELINE_QUALIFICATION / case for case in CASES}


def _utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    info = path.stat()
    if not stat.S_ISREG(info.st_mode):
        raise ValueError(f"expected regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": _sha256(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def _read(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def _entry(value: Any) -> dict[str, Any]:
    if isinstance(value, str):
        return {"path": value}
    if isinstance(value, dict):
        if isinstance(value.get("identity"), dict):
            merged = dict(value["identity"])
            merged.update({key: item for key, item in value.items() if key != "identity"})
            return merged
        return value
    raise ValueError(f"invalid binary descriptor: {value!r}")


def _resolve_descriptor(packet: Path, raw: Any, name: str) -> dict[str, Any]:
    item = _entry(raw)
    _require(isinstance(item.get("path"), str), f"{name} has no path")
    path = Path(item["path"])
    if not path.is_absolute():
        path = packet / path
    resolved = path.resolve()
    _require(resolved.is_file() and os.access(resolved, os.X_OK), f"{name} is not executable: {resolved}")
    actual = descriptor(resolved)
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in item:
            _require(item[key] == actual[key], f"{name} {key} does not match the retained binary")
    _require(item.get("executable", True) is True, f"{name} is not declared executable")
    return {"path": str(resolved), "identity": actual, "declared": item}


def _binary(plan: dict[str, Any], role: str, mode: str) -> dict[str, Any]:
    binaries = plan.get("binaries")
    _require(isinstance(binaries, dict), "plan binaries must be an object")
    side = binaries.get(role)
    _require(isinstance(side, dict), f"plan has no {role} binary set")
    return _resolve_descriptor(PACKET, side.get(mode), f"{role}/{mode}")


def _load_expected_corpora() -> dict[str, dict[str, Any]]:
    expected: dict[str, dict[str, Any]] = {}
    for case, folder in ARCHIVE_CASES.items():
        report_path = folder / "report.json"
        _require(report_path.is_file(), f"sealed baseline qualification is missing {report_path}")
        report = _read(report_path)
        _require(report.get("tool", {}).get("profile") == PROFILE, f"baseline qualification {case} is not release")
        results = report.get("results")
        _require(type(results) is list and len(results) == 1 and results[0].get("case") == case, f"baseline qualification case mismatch: {case}")
        elapsed = results[0].get("elapsed_ns", {})
        _require(elapsed.get("samples") and elapsed.get("samples") == sorted(elapsed["samples"]), f"baseline qualification elapsed vector invalid: {case}")
        expected[case] = {
            "corpus": results[0]["corpus"],
            "report_sha256": _sha256(report_path),
            "producer": None,
            "producer_sha256": None,
        }
        producer_path = folder / "producer.json"
        if producer_path.is_file():
            producer = _read(producer_path)
            _require(producer.get("schema") == "litchi.perf-baseline.producer-shape-evidence.v1", f"baseline producer schema mismatch: {case}")
            expected[case]["producer"] = producer
            expected[case]["producer_sha256"] = _sha256(producer_path)
        elif case in (PRIMARY_CASE, "docx_producer_source_selected_paragraph", "pptx_producer_source_selected_slide"):
            raise ValueError(f"baseline producer sidecar missing for {case}")
        if case == PRIMARY_CASE:
            _require(expected[case]["producer"] is not None, "primary baseline producer sidecar missing")
            baseline_verify._validate_report(report, expected[case]["producer"], "native", {"warmup": 1, "samples": len(elapsed["samples"])})
    return expected


def validate_plan(plan: dict[str, Any]) -> None:
    _require(plan.get("schema") == SCHEMA, f"unexpected plan schema: {plan.get('schema')!r}")
    _require(plan.get("profile") == PROFILE, "0887 requires release binaries")
    _require(str(plan.get("cpu_affinity")) == CPU_AFFINITY, "0887 is pinned to CPU 12")
    _require(plan.get("cache") == "warm", "0887 requires warm cache")
    _require(plan.get("xlsx_shape") == "medium", "0887 fixes --xlsx-shape medium")
    _require(plan.get("cases") == list(CASES), "0887 case order is not frozen")
    _require(plan.get("native_blocks") == NATIVE_BLOCKS, "native block count mismatch")
    _require(plan.get("allocator_blocks") == ALLOCATOR_BLOCKS, "allocator block count mismatch")
    _require(plan.get("native_warmup") == NATIVE_WARMUP and plan.get("native_samples") == NATIVE_SAMPLES, "native sample protocol mismatch")
    _require(plan.get("allocator_warmup") == ALLOCATOR_WARMUP and plan.get("allocator_samples") == ALLOCATOR_SAMPLES, "allocator sample protocol mismatch")
    _require(plan.get("fail_fast") is True and plan.get("fresh_child") is True, "0887 must be fresh-child fail-fast")
    _require(isinstance(plan.get("binaries"), dict), "0887 plan must bind before/after binaries")
    for role in ("before", "after"):
        for mode in ("native", "allocator"):
            item = _entry(plan["binaries"].get(role, {}).get(mode))
            _require(isinstance(item.get("path"), str), f"missing {role}/{mode} binary")
            _require(isinstance(item.get("sha256"), str) and isinstance(item.get("bytes"), int), f"incomplete {role}/{mode} binary identity")
            _require(item.get("executable") is True, f"{role}/{mode} binary not executable")
    before_native = Path(_entry(plan["binaries"]["before"]["native"])["path"])
    before_alloc = Path(_entry(plan["binaries"]["before"]["allocator"])["path"])
    if not before_native.is_absolute():
        before_native = PACKET / before_native
    if not before_alloc.is_absolute():
        before_alloc = PACKET / before_alloc
    before_native = before_native.resolve()
    before_alloc = before_alloc.resolve()
    _require(before_native == (BASELINE_PACKET / "binaries/native").resolve(), "before native is not retained 0885 binary")
    _require(before_alloc == (BASELINE_PACKET / "binaries/allocator").resolve(), "before allocator is not retained 0885 binary")


def load_plan(path: Path) -> dict[str, Any]:
    plan = _read(path)
    _require(isinstance(plan, dict), "plan must be an object")
    validate_plan(plan)
    _load_expected_corpora()
    for role in ("before", "after"):
        for mode in ("native", "allocator"):
            _binary(plan, role, mode)
    return plan


def _command(binary: str, case: str, warmup: int, samples: int, report: Path, producer: Path, peak: Path) -> list[str]:
    return [
        TIME,
        "-f",
        "%M",
        "-o",
        str(peak),
        TASKSET,
        "-c",
        CPU_AFFINITY,
        binary,
        "--case",
        case,
        "--xlsx-shape",
        "medium",
        "--warmup",
        str(warmup),
        "--samples",
        str(samples),
        "--filesystem-cache",
        "warm",
        "--json",
        str(report),
        "--producer-evidence",
        str(producer),
    ]


def _raw_stats(samples: list[int]) -> dict[str, int]:
    _require(type(samples) is list and samples and all(type(value) is int and value > 0 for value in samples), "elapsed samples must be positive integers")
    _require(samples == sorted(samples), "elapsed samples must be sorted as emitted")
    ordered = sorted(samples)
    p50 = baseline_verify.floor_midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2])
    return {
        "min": ordered[0],
        "p50": p50,
        "p95": baseline_verify.nearest_rank(ordered, 95),
        "p99": baseline_verify.nearest_rank(ordered, 99),
        "max": ordered[-1],
    }


def _validate_binary_identity(report: dict[str, Any], binary: dict[str, Any], mode: str) -> None:
    identity = report.get("binary_identity")
    _require(isinstance(identity, dict), "report binary identity is missing")
    actual = binary["identity"]
    _require(identity.get("path") == binary["path"], "report binary path mismatch")
    _require(identity.get("binary_sha256") == actual["sha256"], "report binary hash mismatch")
    _require(identity.get("binary_bytes") == actual["bytes"], "report binary size mismatch")
    _require(identity.get("executable") is True and identity.get("profile") == PROFILE, "report binary profile/executable mismatch")
    _require(identity.get("mode_bits") == actual["mode"], "report binary mode mismatch")
    expected_name = "litchi-perf-baseline" if mode == "native" else "litchi-perf-baseline-alloc"
    _require(report.get("tool", {}).get("binary") == expected_name, "report tool binary name mismatch")


def _validate_common_report(report: dict[str, Any], case: str, mode: str, warmup: int, samples: int, expected: dict[str, Any], binary: dict[str, Any]) -> dict[str, Any]:
    _require(report.get("schema_version") == 1, "report schema version mismatch")
    tool = report.get("tool", {})
    _require(tool.get("name") == "litchi-perf-baseline" and tool.get("version") == "0.1.0", "report tool mismatch")
    _require(tool.get("profile") == PROFILE and tool.get("target_os") == "linux" and tool.get("target_arch") == "x86_64", "report release target mismatch")
    _require(tool.get("instrumentation") == ("none" if mode == "native" else "system_allocator_operation_scoped"), "report instrumentation mismatch")
    if mode == "allocator":
        _require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3", "allocator counter revision mismatch")
    config = report.get("configuration", {})
    _require(config.get("samples_per_case") == samples and config.get("warmup_iterations_per_case") == warmup, "report sample protocol mismatch")
    _require(config.get("filesystem_cache_states") == ["warm"] and config.get("filesystem_fresh_child_per_sample") is True and config.get("filesystem_process_isolated") is True, "report filesystem protocol mismatch")
    _require(config.get("xlsx_shapes") == ["medium"], "report xlsx-shape selection mismatch")
    _require(config.get("cases") == [case], "report selected case mismatch")
    result_rows = report.get("results")
    _require(type(result_rows) is list and len(result_rows) == 1 and result_rows[0].get("case") == case, "report result case mismatch")
    result = result_rows[0]
    elapsed = result.get("elapsed_ns", {})
    raw = elapsed.get("samples")
    _require(type(raw) is list and len(raw) == samples, "report elapsed vector length mismatch")
    order = elapsed.get("sample_order")
    _require(type(order) is list and sorted(order) == list(range(samples)), "report sample-order permutation mismatch")
    stats = _raw_stats(raw)
    for key, value in stats.items():
        _require(elapsed.get(key) == value, f"report raw {key} statistic mismatch")
    _require(result.get("corpus") == expected["corpus"], f"{case} corpus differs from sealed baseline qualification")
    _validate_binary_identity(report, binary, mode)
    _require(report.get("environment", {}).get("cpu_affinity") == CPU_AFFINITY, "report CPU affinity mismatch")
    return {"result": result, "stats": stats, "sample_order": order}


def _validate_child(report: dict[str, Any], producer: dict[str, Any] | None, case: str, mode: str, warmup: int, samples: int, expected: dict[str, Any], binary: dict[str, Any]) -> dict[str, Any]:
    if case == PRIMARY_CASE:
        _validate_common_report(report, case, mode, warmup, samples, expected, binary)
        # This is the 0885 semantic fence reused for the primary workload;
        # its lane parameters are supplied for 0887's sample count.
        _require(producer is not None, "primary producer sidecar is missing")
        checked = baseline_verify._validate_report(report, producer, mode, {"warmup": warmup, "samples": samples})
        _require(producer == expected["producer"], "primary producer sidecar differs from sealed baseline")
        return {"stats": checked["elapsed"], "result": checked["scan"], "allocation": checked["allocation"]}
    checked = _validate_common_report(report, case, mode, warmup, samples, expected, binary)
    _require(checked["result"].get("operation_metrics") in (None, {}), f"control {case} unexpectedly has allocator/operation metrics")
    if expected["producer"] is None:
        _require(producer is None, f"control {case} unexpectedly emitted producer evidence")
    else:
        _require(producer == expected["producer"], f"control {case} producer sidecar differs from sealed baseline")
    return {"stats": checked["stats"], "result": checked["result"], "allocation": None}


def _read_peak(path: Path) -> int:
    _require(path.is_file(), f"peak RSS output missing: {path}")
    text = path.read_text(encoding="utf-8").strip()
    _require(re.fullmatch(r"[0-9]+", text) is not None, f"peak RSS output is not one integer: {path}")
    return int(text)


def _artifacts(folder: Path) -> dict[str, dict[str, Any]]:
    result = {}
    for name in ("stdout.txt", "stderr.txt", "report.json", "peak-rss-kib.txt", "producer.json"):
        path = folder / name
        if path.is_file():
            result[name] = descriptor(path)
    return result


def _schedule() -> list[dict[str, Any]]:
    rows = []
    ordinal = 0
    for block in range(NATIVE_BLOCKS):
        for case_index, case in enumerate(CASES):
            order = ("before", "after") if (block + case_index) % 2 == 0 else ("after", "before")
            pair = f"native-b{block:02d}-c{case_index:02d}"
            for role in order:
                rows.append({"ordinal": ordinal, "lane": "native", "block": block, "case_index": case_index, "case": case, "role": role, "pair": pair, "warmup": NATIVE_WARMUP, "samples": NATIVE_SAMPLES})
                ordinal += 1
    for block in range(ALLOCATOR_BLOCKS):
        order = ("before", "after") if block % 2 == 0 else ("after", "before")
        pair = f"allocator-b{block:02d}-c00"
        for role in order:
            rows.append({"ordinal": ordinal, "lane": "allocator", "block": block, "case_index": 0, "case": PRIMARY_CASE, "role": role, "pair": pair, "warmup": ALLOCATOR_WARMUP, "samples": ALLOCATOR_SAMPLES})
            ordinal += 1
    return rows


def validate_schedule(rows: list[dict[str, Any]]) -> None:
    _require(len(rows) == TOTAL_CHILDREN, f"expected {TOTAL_CHILDREN} children")
    _require(sum(row["lane"] == "native" for row in rows) == TOTAL_NATIVE_CHILDREN, "native population mismatch")
    _require(sum(row["lane"] == "allocator" for row in rows) == TOTAL_ALLOCATOR_CHILDREN, "allocator population mismatch")
    expected = _schedule()
    _require(rows == expected, "experiment schedule differs from frozen AB/BA schedule")


def run(plan_path: Path, output_root: Path, *, popen: Any = subprocess.Popen) -> dict[str, Any]:
    plan = load_plan(plan_path)
    expected_corpora = _load_expected_corpora()
    if output_root.exists():
        raise FileExistsError(f"experiment output already exists: {output_root}")
    output_root.mkdir(parents=True)
    rows = _schedule()
    validate_schedule(rows)
    binaries = {(role, mode): _binary(plan, role, mode) for role in ("before", "after") for mode in ("native", "allocator")}
    receipts = []
    previous_finished: _datetime.datetime | None = None
    for row in rows:
        binary = binaries[(row["role"], "native" if row["lane"] == "native" else "allocator")]
        folder = output_root / "runs" / f"{row['ordinal']:03d}-{row['lane']}-{row['pair']}-{row['role']}"
        folder.mkdir(parents=True, exist_ok=False)
        report_path = folder / "report.json"
        producer_path = folder / "producer.json"
        peak_path = folder / "peak-rss-kib.txt"
        stdout_path = folder / "stdout.txt"
        stderr_path = folder / "stderr.txt"
        argv = _command(binary["path"], row["case"], row["warmup"], row["samples"], report_path, producer_path, peak_path)
        started = _utc_now()
        command = {
            "schema": COMMAND_SCHEMA,
            "row": row,
            "argv": argv,
            "cwd": str(PACKET),
            "environment": CHILD_ENV,
            "binary": binary,
            "plan_sha256": _sha256(plan_path),
            "started_utc": started,
        }
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            process = popen(argv, cwd=PACKET, env=CHILD_ENV.copy(), stdout=stdout, stderr=stderr)
            exit_code = process.wait()
        finished = _utc_now()
        command.update({"finished_utc": finished, "exit_code": exit_code})
        _write_json(folder / "command.json", command)
        receipt = {
            "schema": RECEIPT_SCHEMA,
            "row": row,
            "argv": argv,
            "cwd": str(PACKET),
            "environment": CHILD_ENV,
            "binary": binary,
            "plan_sha256": command["plan_sha256"],
            "started_utc": started,
            "finished_utc": finished,
            "exit_code": exit_code,
            "command": descriptor(folder / "command.json"),
            "artifacts": _artifacts(folder),
        }
        _write_json(folder / "receipt.json", receipt)
        receipts.append({"path": str((folder / "receipt.json").relative_to(output_root)), "row": row})
        if exit_code != 0:
            raise RuntimeError(f"child {row['ordinal']} exited {exit_code}; capture stops without retry")
        if previous_finished is not None:
            _require(_parse_time(started) >= previous_finished, "children were not launched sequentially")
        previous_finished = _parse_time(finished)
        report_path = folder / "report.json"
        _require(report_path.is_file(), f"child {row['ordinal']} report is missing")
        report = _read(report_path)
        producer = _read(producer_path) if producer_path.is_file() else None
        _validate_child(report, producer, row["case"], "native" if row["lane"] == "native" else "allocator", row["warmup"], row["samples"], expected_corpora[row["case"]], binary)
        _read_peak(peak_path)
        # Refresh the receipt artifact map only after all semantic outputs are
        # present; the receipt itself remains immutable and is rewritten only
        # by this child creation path.
        receipt["artifacts"] = _artifacts(folder)
        receipt["semantic"] = {"status": "pass", "report_sha256": _sha256(report_path), "producer_sha256": _sha256(producer_path) if producer_path.is_file() else None}
        # receipt.json was created above; preserve create-only custody by
        # recording the final semantic descriptor in a separate file.
        _write_json(folder / "semantic.json", receipt["semantic"])
    result = {
        "schema": EXPERIMENT_SCHEMA,
        "status": "pass",
        "plan": str(plan_path.resolve()),
        "plan_sha256": _sha256(plan_path),
        "schedule": "native then allocator; fixed AB/BA by (block + case_index) % 2",
        "rows": receipts,
        "children": TOTAL_CHILDREN,
        "native_children": TOTAL_NATIVE_CHILDREN,
        "allocator_children": TOTAL_ALLOCATOR_CHILDREN,
        "no_retries_or_replacements": True,
        "performance_claim": None,
    }
    _write_json(output_root / "experiment.json", result)
    return result


def _parse_time(value: str) -> _datetime.datetime:
    return _datetime.datetime.fromisoformat(value)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    result = run(args.plan.resolve(), args.output_root.resolve())
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
