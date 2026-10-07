#!/usr/bin/env python3
"""Fail-closed serial capture runner for the prospective 0959 comparison.

The runner is deliberately kept outside the checkout.  It performs no build
and cannot launch a child unless the plan is an explicit root-authorized
packet, the candidate bindings are complete, ``--execute`` is supplied, and
``LITCHI_PERF_0959_CAPTURE_APPROVED=1`` is present.  Each schedule row gets a
fresh process and a new output directory.  The process terminal is sealed
before any report, time file, or oracle parsing.  A failure stops the frozen
schedule; there is no retry or replacement path.

``--dry-run`` validates the 72-row native or 18-row allocator schedule
without starting a process.  The execution path is intentionally serial: a
native lane is completed and sealed before an allocator lane may be started by
a separate invocation.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
import time
from typing import Any, Iterable


ROOT = Path("/home/zhuhe/litchi-performance-0959").resolve()
APPROVAL_ENV = "LITCHI_PERF_0959_CAPTURE_APPROVED"
PLAN_SCHEMA = "litchi.performance.0959.inline-number-comparison-plan.v4"
RUNNER_SCHEMA = "litchi.performance.0959.capture.v4"
CPU_MASK = [12, 13, 14, 15]
DEADLINE_SECONDS = 900
REPORT_LIMIT = 8 * 1024 * 1024
MAX_SINK_WRITE_BYTES = 65_536
EXPECTED_ALLOWLIST = [
    "crates/litchi-xlsx/src/cell.rs",
    "crates/litchi-xlsx/src/raw/worksheet/semantic.rs",
    "crates/litchi-xlsx/src/raw/worksheet/selected.rs",
]

TIME_KEYS = {
    "Maximum resident set size (kbytes)": "max_rss_kib",
    "User time (seconds)": "user_seconds",
    "System time (seconds)": "system_seconds",
    "Major (requiring I/O) page faults": "major_page_faults",
    "Minor (reclaiming a frame) page faults": "minor_page_faults",
    "Voluntary context switches": "voluntary_context_switches",
    "Involuntary context switches": "involuntary_context_switches",
}

MEDIUM_LIMITS = {
    "memory": 2_956_761,
    "input_bytes": 67_108_864,
    "output_bytes": 8_519_112,
    "objects": 1_000_000,
    "depth": 256,
    "work": 1_000_000_000,
    "execution_workers": 1,
    "execution_in_flight_tasks": 1,
    "execution_in_flight_bytes": 2_956_761,
    "execution_min_parallel_bytes": 0,
    "execution_min_task_bytes": 0,
    "root_workers": 1,
    "root_io_concurrency": 1,
    "root_cpu_tasks": 5,
}
DENSE_LIMITS = {
    **MEDIUM_LIMITS,
    "memory": 3_196_436,
    "output_bytes": 8_570_408,
    "execution_in_flight_bytes": 3_196_436,
}
EDIT_LIMITS = {
    **MEDIUM_LIMITS,
    "memory": 2_751_297,
    "output_bytes": 8_518_496,
    "execution_in_flight_bytes": 2_751_297,
    "root_cpu_tasks": 2,
}

EXPECTED_CASES = (
    {
        "case_id": "medium-one-percent",
        "shape": "medium",
        "workload": "one-percent",
        "source_bytes": 4_226_429,
        "source_sha256": "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036",
        "output_bytes": 4_226_788,
        "output_sha256": "fc742a7ad139e2ce2e75fabfff8d2a33408dd5eeb10997717a2e2656d8423e6c",
        "semantic_sha256": "f705d4200e8bbe840abb16ff7d17da887cbe1d0fb53a649c69279451ff5d829c",
        "expected_cpu_tasks": 5,
        "expected_worker_permits": 0,
        "update_count": 93,
        "context_limits": MEDIUM_LIMITS,
    },
    {
        "case_id": "dense-sparse-one-percent",
        "shape": "dense-sparse",
        "workload": "one-percent",
        "source_bytes": 4_251_863,
        "source_sha256": "893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a",
        "output_bytes": 4_252_436,
        "output_sha256": "07eee3f980949f031776087ed61f06125edd82cb29394a608a3d8d5a4b28374d",
        "semantic_sha256": "d8bbafac4c31288eb48a773df88b6b2add2a51b9ad6fb652ab1aeb801c5dfa8c",
        "expected_cpu_tasks": 5,
        "expected_worker_permits": 0,
        "update_count": 178,
        "context_limits": DENSE_LIMITS,
    },
    {
        "case_id": "medium-one-edit",
        "shape": "medium",
        "workload": "one-edit",
        "source_bytes": 4_226_429,
        "source_sha256": "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036",
        "output_bytes": 4_226_480,
        "output_sha256": "9b7b66a02007eeb63498fd5de4c6b7115ace0383ce37d97e1a9560ef7bfadec1",
        "semantic_sha256": "3cd21160d4f74fa0f097ab40be08e211b3e460cea788aa2b6705a55fdece07de",
        "expected_cpu_tasks": 2,
        "expected_worker_permits": 0,
        "update_count": 1,
        "context_limits": EDIT_LIMITS,
    },
)

BLOCKS = (
    {"id": 0, "order": ["A", "B", "B", "A"]},
    {"id": 1, "order": ["B", "A", "A", "B"]},
    {"id": 2, "order": ["A", "B", "B", "A"]},
    {"id": 3, "order": ["B", "A", "A", "B"]},
    {"id": 4, "order": ["A", "B", "B", "A"]},
    {"id": 5, "order": ["B", "A", "A", "B"]},
)
ALLOCATOR_BLOCKS = (
    {"id": 0, "order": ["A", "B"]},
    {"id": 1, "order": ["B", "A"]},
    {"id": 2, "order": ["A", "B"]},
)

LANES = {
    "native": {
        "warmup_iterations": 20,
        "measured_samples": 200,
        "report_schema": "litchi.xlsx-managed-compression-benchmark.v1",
    },
    "allocator": {
        "warmup_iterations": 2,
        "measured_samples": 10,
        "report_schema": "litchi.xlsx-managed-compression-allocation.v1",
    },
}

NATIVE_SCOPES = {
    "timing_scope": "with fresh finite managed context prepared outside the clock, the timer begins at SourceBackedEditor construction and covers edit/commit, source-backed publication, returned snapshot drop, and commit/package worker-pool teardown; precomputed source/cache/context setup and bounded sink allocation are outside",
    "phase_scope": "ingress_ns is the direct SourceBackedEditor constructor; edit_commit_ns is selector staging, cell sets, and MultiCommit::commit; publication_ns is publish_multi_commit_to_stream plus returned MultiSnapshot drop; commit_drop_ns is MultiCommit drop; total_ns starts at the constructor and ends after commit teardown",
    "setup_scope": "deterministic corpus and expected output construction, frozen source/output/semantic oracles, SourceCacheLimits and output ceiling, one untimed admission preflight, inverse patch, fresh OwnedSource clone, finite managed context, and sink capacity reservation are outside measured phases",
    "verification_scope": "exact output bytes, immutable fixture source hash, output/semantic hashes, workbook verification, sink bounds, CpuTasks/OutputBytes counters, and post-drop Memory/Workers/IoConcurrency/Objects checks run outside timers; the untimed admission preflight verifies an actual managed-commit inverse (preflight_actual_commit_inverse_verified=true), while each timed sample applies the precomputed inverse patch only as a fixture to its published bytes (fixture_inverse_restored_exact_source)",
    "performance_claim": "none; raw phase timings describe this explicit in-memory harness and do not claim OS thread counts, speedup, allocation behavior, or physical I/O",
}
ALLOCATOR_SCOPES = {
    "allocator_identity": "CountingSystemAllocator(std::alloc::System)",
    "instrumentation_identity": "system_allocator_operation_scoped",
    "counter_revision": "serialized_region_peak_v3",
    "allocation_scope": "global system-allocator callbacks serialized by allocation_metrics; process-wide across threads; observer peaks are not RSS, unique heap, per-thread bytes, OS-thread counts, or timing",
    "thread_observation_scope": "calibration uses a separate measured allocator region; pre-sample /proc/self/status probes are outside operation regions; the post-core fixed-buffer drain is the fifth in-region segment; observed process thread counts do not establish admitted-worker/thread mapping or task overlap",
    "timing_scope": "none; this report contains allocator regions and diagnostic quiescence observations only",
    "performance_claim": "none; allocator counters do not claim latency, throughput, RSS, OS-thread width, worker/task overlap, speedup, or adoption; procfs Threads values are quiescence observations only",
    "peak_semantics": "observer-ordered process-live high-water evidence, not RSS, unique heap, per-thread bytes, or OS-thread width",
}
ALLOC_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)


class CaptureError(ValueError):
    """A fail-closed plan, custody, process, or report error."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CaptureError(message)


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path, *, executable: bool = False) -> dict[str, Any]:
    original = path
    require(not original.is_symlink() and original.is_file(), f"regular file required: {original}")
    path = original.resolve()
    require(not path.is_symlink() and path.is_file(), f"regular file required: {path}")
    mode = path.stat().st_mode & 0o777
    if executable:
        require(bool(mode & 0o111), f"executable bit missing: {path}")
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "mode": mode,
        "executable": bool(mode & 0o111),
        "sha256": sha256(path),
    }


def write_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def read_json(path: Path, *, limit: int = 64 * 1024 * 1024) -> Any:
    require(not path.is_symlink() and path.is_file(), f"JSON file unavailable: {path}")
    require(path.stat().st_size <= limit, f"JSON file exceeds limit: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise CaptureError(f"invalid JSON {path}: {exc}") from exc


def int_value(value: Any, location: str, *, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{location} must be integer >= {minimum}")
    return value


def external_binding(value: Any, label: str, *, executable: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} binding missing")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and Path(raw_path).is_absolute(), f"{label}.path must be absolute")
    actual = descriptor(Path(raw_path), executable=executable)
    for key in ("path", "bytes", "sha256"):
        require(value.get(key) == actual[key], f"{label}.{key} binding changed")
    if executable:
        require(value.get("executable") is True and actual["executable"], f"{label} is not executable")
    return actual


def nested_binding(value: Any, label: str, base: Path, *, executable: bool = False) -> dict[str, Any]:
    """Validate a receipt descriptor whose path is packet-relative."""
    require(isinstance(value, dict), f"{label} binding missing")
    raw_path = value.get("path")
    require(isinstance(raw_path, str), f"{label}.path missing")
    path = Path(raw_path)
    if path.is_absolute():
        actual = descriptor(path, executable=executable)
        roots = [actual]
    else:
        candidates: list[dict[str, Any]] = []
        for ancestor in (base.parent, *base.parents):
            try:
                actual = descriptor(ancestor / path, executable=executable)
            except CaptureError:
                continue
            if all(value.get(key) == actual[key] for key in ("bytes", "sha256")):
                candidates.append(actual)
        require(len(candidates) == 1, f"{label} relative path is missing or ambiguous")
        roots = candidates
    actual = roots[0]
    for key in ("bytes", "sha256"):
        require(value.get(key) == actual[key], f"{label}.{key} binding changed")
    return actual


def pending(value: Any) -> bool:
    return isinstance(value, dict) and value.get("status") == "pending_root_binding"


def manifest_entries(path: Path, label: str) -> dict[str, tuple[Any, Any, Any]]:
    value = read_json(path, limit=64 * 1024 * 1024)
    files = value.get("files") if isinstance(value, dict) else None
    require(isinstance(files, list), f"{label} files census missing")
    entries: dict[str, tuple[Any, Any, Any]] = {}
    for item in files:
        require(isinstance(item, dict), f"{label} census entry is not an object")
        path_value = item.get("path")
        require(isinstance(path_value, str) and path_value not in entries, f"{label} census path invalid or duplicated")
        entries[path_value] = (item.get("bytes"), item.get("mode"), item.get("sha256"))
    require(value.get("file_count") == len(entries), f"{label} file count does not match census")
    return entries


def validate_candidate_manifest_diff(baseline: Path, candidate: Path) -> list[str]:
    before = manifest_entries(baseline, "baseline source")
    after = manifest_entries(candidate, "candidate source")
    changed = sorted(
        path for path in set(before) | set(after)
        if before.get(path) != after.get(path)
    )
    require(changed == sorted(EXPECTED_ALLOWLIST), f"candidate source diff changed paths: {changed}")
    return changed


def validate_build_receipt(value: Any, label: str, expected_binary: dict[str, Any], candidate_manifest_sha: str) -> None:
    receipt_descriptor = external_binding(value, label)
    result = read_json(Path(receipt_descriptor["path"]), limit=16 * 1024 * 1024)
    require(isinstance(result, dict) and result.get("status") == "passed", f"{label} result is not passed")
    binary = result.get("binary")
    require(isinstance(binary, dict), f"{label} binary binding missing")
    for key in ("bytes", "sha256"):
        require(binary.get(key) == expected_binary[key], f"{label} binary {key} differs from lane binding")
    terminal = result.get("terminal")
    terminal_descriptor = nested_binding(terminal, f"{label}.terminal", Path(receipt_descriptor["path"]))
    terminal_value = read_json(Path(terminal_descriptor["path"]), limit=16 * 1024 * 1024)
    require(isinstance(terminal_value, dict), f"{label} terminal is not an object")
    require(terminal_value.get("returncode", terminal_value.get("return_code")) == 0, f"{label} terminal return code is not zero")
    require(terminal_value.get("timed_out") is False, f"{label} terminal timed out")
    if "process_group_clear" in terminal_value:
        require(terminal_value["process_group_clear"] is True, f"{label} build process group did not clear")
    if "reaped" in terminal_value:
        require(terminal_value["reaped"] is True, f"{label} build process was not reaped")
    source = terminal_value.get("source")
    require(isinstance(source, dict), f"{label} terminal source binding missing")
    manifest = source.get("manifest")
    require(isinstance(manifest, dict), f"{label} terminal source manifest missing")
    require(manifest.get("sha256") == candidate_manifest_sha, f"{label} source manifest differs from candidate binding")
    changes = source.get("changes")
    if isinstance(changes, list):
        paths = sorted(item.get("path") for item in changes if isinstance(item, dict))
        require(paths == sorted(EXPECTED_ALLOWLIST), f"{label} build source changes differ: {paths}")


def expected_children(lane: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    spec = LANES[lane]
    blocks = BLOCKS if lane == "native" else ALLOCATOR_BLOCKS
    for case in EXPECTED_CASES:
        for block in blocks:
            for position, role in enumerate(block["order"]):
                rows.append(
                    {
                        "child_id": f"{lane}__{case['case_id']}__b{block['id']:02d}__p{position}",
                        "lane": lane,
                        "case_id": case["case_id"],
                        "shape": case["shape"],
                        "workload": case["workload"],
                        "comparison_id": "baseline-vs-candidate",
                        "block": block["id"],
                        "position": position,
                        "role": role,
                        "binary_role": "baseline" if role == "A" else "candidate",
                        "workers": 1,
                        "cpu_mask": CPU_MASK,
                        "warmup_iterations": spec["warmup_iterations"],
                        "measured_samples": spec["measured_samples"],
                    }
                )
    return rows


def validate_case_matrix(plan: dict[str, Any]) -> None:
    cases = plan.get("cases")
    require(isinstance(cases, list) and len(cases) == len(EXPECTED_CASES), "three-case matrix is missing")
    for actual, expected in zip(cases, EXPECTED_CASES):
        require(isinstance(actual, dict), "case entry is not an object")
        for key, value in expected.items():
            require(actual.get(key) == value, f"case {expected['case_id']} {key} oracle changed")
    comparison = plan.get("comparison")
    require(isinstance(comparison, dict), "comparison binding missing")
    require(comparison.get("id") == "baseline-vs-candidate", "comparison id changed")
    require(comparison.get("workers") == 1, "comparison workers changed")
    require(comparison.get("block_count") == 6, "comparison block count changed")
    require(comparison.get("blocks") == list(BLOCKS), "ABBA/BAAB block order changed")
    require(comparison.get("pairs_per_case_native") == 12, "native paired block count changed")
    require(comparison.get("pairs_per_case_allocator") == 3, "allocator paired repeat count changed")
    require(comparison.get("allocator_blocks") == list(ALLOCATOR_BLOCKS), "allocator pair order changed")
    require(comparison.get("allocator_repeats_per_role") == 3, "allocator repeat count changed")


def validate_execution(plan: dict[str, Any], lane: str) -> None:
    execution = plan.get("lanes", {}).get(lane, {}).get("execution")
    require(isinstance(execution, dict), f"{lane} execution binding missing")
    require(execution.get("workload_cwd") == str(ROOT), "workload cwd is not the frozen external packet root")
    cwd = Path(execution["workload_cwd"])
    require(cwd.is_dir() and not cwd.is_symlink(), f"workload cwd is unavailable: {cwd}")
    require(execution.get("taskset_path") == "/usr/bin/taskset", "taskset path changed")
    require(execution.get("time_path") == "/usr/bin/time", "time path changed")
    require(execution.get("cpu_mask") == CPU_MASK, "CPU mask changed")
    require(execution.get("max_parallel_children") == 1, "children are not serial")
    require(execution.get("deadline_seconds") == DEADLINE_SECONDS, "child deadline changed")
    require(execution.get("terminal_before_parse") is True, "terminal-first custody was disabled")
    require(execution.get("retry_or_replacement") is False, "retry/replacement policy changed")
    expected_argv = [
        "/usr/bin/taskset", "--cpu-list", "12-15", "<BINARY>",
        "--shape", "<SHAPE>", "--workload", "<WORKLOAD>",
        "--workers", "1", "--warmup",
        str(LANES[lane]["warmup_iterations"]), "--samples",
        str(LANES[lane]["measured_samples"]), "--output", "<OUTPUT>",
    ]
    require(execution.get("argv_template") == expected_argv, "child argv template changed")
    require(execution.get("time_argv_prefix") == ["/usr/bin/time", "-v", "-o", "time-v.txt", "--"], "time argv changed")
    require(
        execution.get("environment") == {
            "LANG": "C",
            "LC_ALL": "C",
            "PATH": "/usr/bin:/bin",
            "RUST_BACKTRACE": "0",
            "TZ": "UTC",
            "approval_variable": APPROVAL_ENV,
        },
        "explicit child environment changed",
    )


def validate_plan(plan: dict[str, Any], lane: str, *, for_execution: bool) -> list[dict[str, Any]]:
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema mismatch")
    status = plan.get("status")
    if for_execution:
        require(status == "authorized", "execution requires status=authorized")
    else:
        require(status in ("prospective_candidate_pending", "prospective", "authorized"), "invalid prospective plan status")
    validate_case_matrix(plan)
    candidate_change = plan.get("candidate_change")
    require(isinstance(candidate_change, dict), "candidate change binding missing")
    require(candidate_change.get("expected_allowlist") == EXPECTED_ALLOWLIST, "candidate changed-path allowlist changed")
    lane_spec = plan.get("lanes", {}).get(lane)
    require(isinstance(lane_spec, dict), f"{lane} lane missing")
    lane_expected = LANES[lane]
    for key, expected in lane_expected.items():
        require(lane_spec.get(key) == expected, f"{lane} lane {key} changed")
    children = lane_spec.get("children")
    require(children == expected_children(lane), f"{lane} schedule differs from frozen paired matrix")
    expected_child_count = 72 if lane == "native" else 18
    require(lane_spec.get("child_count") == expected_child_count, f"{lane} child count changed")
    validate_execution(plan, lane)

    source = plan.get("source")
    require(isinstance(source, dict), "source binding missing")
    baseline_manifest = external_binding(source.get("baseline_manifest"), "source.baseline_manifest")
    require(source.get("baseline_manifest_sha256") == baseline_manifest["sha256"], "baseline manifest SHA binding changed")
    baseline_receipt = external_binding(source.get("baseline_receipt"), "source.baseline_receipt")
    baseline_root = Path(str(source.get("baseline_source_root", "")))
    require(baseline_root.is_dir() and not baseline_root.is_symlink(), "baseline source root unavailable")
    require(baseline_root == ROOT / "source-baseline", "baseline source root binding changed")
    runtime_root = Path(str(source.get("baseline_runtime_root", "")))
    require(runtime_root.is_dir() and not runtime_root.is_symlink(), "retained baseline runtime root unavailable")
    require(runtime_root == Path("/home/zhuhe/litchi-performance-0957/source-v1"), "retained baseline runtime root binding changed")
    runtime_manifest = external_binding(source.get("baseline_runtime_manifest"), "source.baseline_runtime_manifest")
    require(runtime_manifest["path"] == str((ROOT / "source-manifest-v2.json").resolve()), "retained baseline runtime manifest binding changed")
    probe = source.get("test_only_probe")
    require(isinstance(probe, dict), "test-only probe binding missing")
    require(probe.get("status") == "explicit_test_only_overlay" and probe.get("runtime") is False, "test-only probe is not excluded from runtime")
    require(probe.get("paths") == ["crates/litchi-xlsx/src/cell.rs"], "test-only probe paths changed")
    require(probe.get("manifest", {}).get("sha256") == baseline_manifest["sha256"], "probe manifest differs from baseline freeze")
    require(probe.get("receipt", {}).get("sha256") == baseline_receipt["sha256"], "probe receipt differs from baseline freeze")

    binaries: dict[str, dict[str, Any]] = {}
    for role in ("baseline_binary", "candidate_binary"):
        value = lane_spec.get(role)
        if role == "candidate_binary" and pending(value):
            require(not for_execution, "candidate binary remains pending")
        else:
            binaries[role] = external_binding(value, f"lanes.{lane}.{role}", executable=True)
    if for_execution:
        for label in ("candidate_source_manifest", "candidate_source_freeze"):
            external_binding(candidate_change.get(label), f"candidate_change.{label}")
        for label in ("candidate_manifest", "candidate_freeze"):
            external_binding(source.get(label), f"source.{label}")
        require(
            candidate_change["candidate_source_manifest"]["sha256"] == source["candidate_manifest"]["sha256"],
            "candidate source manifest bindings disagree",
        )
        changed_paths = validate_candidate_manifest_diff(
            Path(baseline_manifest["path"]),
            Path(source["candidate_manifest"]["path"]),
        )
        helpers = plan.get("helpers")
        require(isinstance(helpers, dict), "helper bindings missing for execution")
        for label in ("capture", "reducer", "controls"):
            external_binding(helpers.get(label), f"helpers.{label}")
        running = descriptor(Path(__file__), executable=True)
        external_binding(helpers.get("capture"), "helpers.capture", executable=True)
        require(helpers["capture"]["sha256"] == running["sha256"], "executing capture helper differs from bound helper")
        go = plan.get("root_go")
        require(isinstance(go, dict) and go.get("status") == "GO", "root GO binding missing")
        if "receipt" in go:
            go_receipt = external_binding(go.get("receipt"), "root_go.receipt")
            go_value = read_json(Path(go_receipt["path"]), limit=16 * 1024 * 1024)
            require(isinstance(go_value, dict) and go_value.get("status") in ("GO", "passed", "authorized"), "root GO receipt is not passed")
        host = external_binding(plan.get("host"), "host")
        host_value = read_json(Path(host["path"]), limit=16 * 1024 * 1024)
        require(isinstance(host_value, dict), "host receipt is not an object")
        candidate_build = plan.get("candidate_build")
        require(isinstance(candidate_build, dict) and candidate_build.get("status") == "bound", "candidate build receipt binding missing")
        require(candidate_build.get("toolchain") == "+1.95.0", "candidate build toolchain changed")
        require(candidate_build.get("profile") == "release, opt-level=3, fat LTO, codegen-units=1, debug=1, incremental=0, jobs=2", "candidate build profile changed")
        candidate_manifest_sha = source["candidate_manifest"]["sha256"]
        for role in ("native", "allocator"):
            validate_build_receipt(
                candidate_build.get(f"{role}_receipt"),
                f"candidate_build.{role}_receipt",
                external_binding(plan["lanes"][role]["candidate_binary"], f"lanes.{role}.candidate_binary", executable=True),
                candidate_manifest_sha,
            )
    return children


def case_for(plan: dict[str, Any], case_id: str) -> dict[str, Any]:
    for case in plan["cases"]:
        if case.get("case_id") == case_id:
            return case
    raise CaptureError(f"unknown case {case_id}")


def check_common_report(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any], lane: str, binary: Path) -> dict[str, Any]:
    spec = LANES[lane]
    require(report.get("schema") == spec["report_schema"], f"{lane} report schema changed")
    require(report.get("status") == "completed", f"{lane} report is not completed")
    # Native reports expose schedule identity at the report root.  The
    # allocator report contract keeps these fields under ``case``; do not
    # impose the native root shape on that distinct schema.
    if lane == "native":
        require(report.get("shape") == row["shape"], "report shape differs from schedule")
        require(report.get("workload") == row["workload"], "report workload differs from schedule")
        require(report.get("workers") == 1, "report workers differs from schedule")
        require(report.get("warmup_iterations") == spec["warmup_iterations"], "report warmup differs from schedule")
        require(report.get("measured_samples") == spec["measured_samples"], "report sample count differs from schedule")
    case_cfg = case_for(plan, row["case_id"])
    case = report.get("case")
    require(isinstance(case, dict), "report.case missing")
    expected_case = {
        "id": f"{row['shape']}-{row['workload']}-workers1",
        "shape": row["shape"],
        "workload": row["workload"],
        "workers": 1,
        "update_count": case_cfg["update_count"],
        "source_bytes": case_cfg["source_bytes"],
        "source_sha256": case_cfg["source_sha256"],
        "output_bytes": case_cfg["output_bytes"],
        "output_sha256": case_cfg["output_sha256"],
        "semantic_sha256": case_cfg["semantic_sha256"],
        "expected_cpu_tasks": case_cfg["expected_cpu_tasks"],
        "admission_expected_worker_permits": 0,
        "admission_observed_worker_permits": 0,
        "admission_cpu_tasks_used": case_cfg["expected_cpu_tasks"],
        "preflight_actual_commit_inverse_verified": True,
        "context_limits": case_cfg["context_limits"],
        "sink_capacity_bytes": case_cfg["context_limits"]["output_bytes"],
        "sink_max_write_bytes": MAX_SINK_WRITE_BYTES,
        "warmup_iterations": spec["warmup_iterations"],
        "measured_samples": spec["measured_samples"],
    }
    for key, expected in expected_case.items():
        require(case.get(key) == expected, f"report.case.{key} oracle/admission changed")
    actual_binary = descriptor(binary, executable=True)
    if lane == "allocator":
        identity = report.get("binary_identity")
        require(isinstance(identity, dict), "allocator binary identity missing")
        for key in ("path", "bytes", "sha256"):
            require(identity.get(key) == actual_binary[key], f"allocator binary identity {key} changed")
    return case


def validate_native(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any], binary: Path) -> None:
    check_common_report(report, row, plan, "native", binary)
    for key, expected in NATIVE_SCOPES.items():
        require(report.get(key) == expected, f"native report {key} scope changed")
    samples = report["case"].get("samples")
    require(isinstance(samples, list) and len(samples) == 200, "native sample rows missing")
    case_cfg = case_for(plan, row["case_id"])
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict), f"native sample {index} is not an object")
        require(sample.get("sample") == index, f"native sample index changed at {index}")
        values = [int_value(sample.get(name), f"native sample[{index}].{name}") for name in (
            "ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns")]
        require(sum(values[:4]) <= values[4], f"native phase sum exceeds total at {index}")
        require(sample.get("output_bytes") == case_cfg["output_bytes"], f"native output bytes changed at {index}")
        require(sample.get("output_bytes_used") == case_cfg["output_bytes"], f"native OutputBytes changed at {index}")
        require(sample.get("cpu_tasks_used") == case_cfg["expected_cpu_tasks"], f"native CpuTasks changed at {index}")
        require(type(sample.get("sink_write_calls")) is int and sample["sink_write_calls"] > 0, f"native sink calls missing at {index}")
        require(type(sample.get("sink_largest_write")) is int and 0 < sample["sink_largest_write"] <= MAX_SINK_WRITE_BYTES, f"native sink bound changed at {index}")
        require(sample.get("output_sha256") == case_cfg["output_sha256"], f"native output hash changed at {index}")
        require(sample.get("semantic_sha256") == case_cfg["semantic_sha256"], f"native semantic hash changed at {index}")
        require(sample.get("fixture_inverse_restored_exact_source") is True, f"native inverse oracle failed at {index}")
        resources = sample.get("resources_after_drop")
        require(isinstance(resources, dict), f"native resources missing at {index}")
        require(all(resources.get(name) == 0 for name in ("memory", "workers", "io_concurrency", "objects")), f"native resources not zero at {index}")


def allocation_sample(sample: Any, label: str) -> dict[str, Any]:
    require(isinstance(sample, dict), f"{label} is not an allocation sample")
    require(sample.get("status") == "measured", f"{label} status is not measured")
    require(sample.get("scope") == "operation_global_system_allocator", f"{label} scope changed")
    for field in ALLOC_FIELDS:
        int_value(sample.get(field), f"{label}.{field}")
    require(sample["failed_allocation_calls"] == 0, f"{label} recorded a failed allocation")
    require(sample["region_peak_live_bytes"] >= max(sample["live_bytes_before"], sample["live_bytes_after"]), f"{label} peak is below endpoint")
    return sample


def combine_samples(samples: Iterable[dict[str, Any]], label: str) -> dict[str, Any]:
    items = list(samples)
    require(items, f"{label} is empty")
    for left, right in zip(items, items[1:]):
        require(left["live_bytes_after"] == right["live_bytes_before"], f"{label} live boundaries are not adjacent")
    result = {
        "status": "measured",
        "scope": "operation_global_system_allocator",
        "live_bytes_before": items[0]["live_bytes_before"],
        "live_bytes_after": items[-1]["live_bytes_after"],
        "peak_live_bytes_before": items[0]["peak_live_bytes_before"],
        "peak_live_bytes_after": items[-1]["peak_live_bytes_after"],
        "region_peak_live_bytes": max(item["region_peak_live_bytes"] for item in items),
    }
    for field in ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes"):
        result[field] = sum(item[field] for item in items)
    return result


def validate_threads(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label} thread observation missing")
    require(type(value.get("first_thread_count")) is int and value["first_thread_count"] >= 1, f"{label} first thread count invalid")
    require(value.get("final_thread_count") == 1, f"{label} did not finish at one thread")
    require(type(value.get("polls")) is int and value["polls"] > 0, f"{label} poll count invalid")
    require(value.get("drained") is True, f"{label} drain failed")
    require(value.get("deadline_seconds") == 2, f"{label} deadline changed")
    require(value.get("poll_limit") == 10_000, f"{label} poll limit changed")
    require(value.get("poll_sleep_milliseconds") == 1, f"{label} poll sleep changed")


def validate_allocator(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any], binary: Path) -> None:
    check_common_report(report, row, plan, "allocator", binary)
    for key, expected in ALLOCATOR_SCOPES.items():
        require(report.get(key) == expected, f"allocator report {key} scope changed")
    calibration = report.get("calibration")
    require(isinstance(calibration, dict), "allocator calibration missing")
    for key, expected in {
        "status": "pass", "polls": 100, "thread_count_first": 1,
        "thread_count_last": 1, "allocation_calls": 0, "deallocation_calls": 0,
        "reallocation_calls": 0, "failed_allocation_calls": 0,
        "no_allocator_callbacks": True, "poll_sleep_milliseconds": 1,
        "thread_scope": ALLOCATOR_SCOPES["thread_observation_scope"],
    }.items():
        require(calibration.get(key) == expected, f"allocator calibration {key} changed")
    rows = report["case"].get("rows")
    require(isinstance(rows, list) and len(rows) == 10, "allocator measured rows missing")
    case_cfg = case_for(plan, row["case_id"])
    for index, item in enumerate(rows):
        require(isinstance(item, dict), f"allocator row {index} is not an object")
        require(item.get("sample") == index, f"allocator row index changed at {index}")
        for key, expected in {
            "context_limits": case_cfg["context_limits"],
            "workers_requested": 1,
            "expected_worker_permits": 0,
            "admitted_worker_permits": 0,
            "expected_cpu_tasks": case_cfg["expected_cpu_tasks"],
            "cpu_tasks_used": case_cfg["expected_cpu_tasks"],
            "output_bytes": case_cfg["output_bytes"],
            "output_bytes_used": case_cfg["output_bytes"],
            "source_sha256": case_cfg["source_sha256"],
            "output_sha256": case_cfg["output_sha256"],
            "semantic_sha256": case_cfg["semantic_sha256"],
            "fixture_inverse_restored_exact_source": True,
        }.items():
            require(item.get(key) == expected, f"allocator row {index} {key} oracle changed")
        require(type(item.get("sink_write_calls")) is int and item["sink_write_calls"] > 0, f"allocator row {index} sink calls missing")
        require(type(item.get("sink_largest_write")) is int and 0 < item["sink_largest_write"] <= MAX_SINK_WRITE_BYTES, f"allocator row {index} sink bound changed")
        resources = item.get("resources_after_core")
        require(isinstance(resources, dict) and all(resources.get(name) == 0 for name in ("memory", "workers", "io_concurrency", "objects")), f"allocator row {index} resources not zero")
        phases = [allocation_sample(item.get(name), f"allocator row {index}.{name}") for name in ("setup", "ingress", "edit_commit", "publication", "commit_drop", "worker_drain", "core_combined", "core_combined_with_drain", "postdrop")]
        setup, ingress, edit_commit, publication, commit_drop, drain, core, core_drain, postdrop = phases
        expected_core = combine_samples([ingress, edit_commit, publication, commit_drop], f"allocator row {index} core")
        expected_drain = combine_samples([ingress, edit_commit, publication, commit_drop, drain], f"allocator row {index} core+drain")
        require(core == expected_core, f"allocator row {index} core conservation failed")
        require(core_drain == expected_drain, f"allocator row {index} core+drain conservation failed")
        require(type(item.get("max_setup_core_peak_above_setup_live_before")) is int, f"allocator row {index} core peak diagnostic missing")
        require(type(item.get("max_setup_drained_core_peak_above_setup_live_before")) is int, f"allocator row {index} drained peak diagnostic missing")
        validate_threads(item.get("pre_sample_threads"), f"allocator row {index} pre-sample")
        validate_threads(item.get("post_core_threads"), f"allocator row {index} post-core")


def parse_time(path: Path) -> dict[str, float | int]:
    require(not path.is_symlink() and path.is_file(), f"/usr/bin/time output missing: {path}")
    values: dict[str, float | int] = {}
    for line in path.read_text(errors="replace").splitlines():
        if ":" not in line:
            continue
        key, raw = line.split(":", 1)
        field = TIME_KEYS.get(key.strip())
        if field is None:
            continue
        token = raw.strip().split()[0] if raw.strip() else ""
        try:
            values[field] = float(token) if field in ("user_seconds", "system_seconds") else int(token)
        except (ValueError, IndexError) as exc:
            raise CaptureError(f"invalid /usr/bin/time field {key!r}: {raw!r}") from exc
    require(set(TIME_KEYS.values()) <= values.keys(), f"missing /usr/bin/time fields: {path}")
    return values


def child_environment(plan: dict[str, Any]) -> dict[str, str]:
    configured = plan["lanes"]["native"]["execution"]["environment"]
    result = {key: value for key, value in configured.items() if key != "approval_variable"}
    result[APPROVAL_ENV] = "1"
    return result


def binary_for(plan: dict[str, Any], lane: str, role: str) -> Path:
    value = plan["lanes"][lane]["baseline_binary" if role == "A" else "candidate_binary"]
    return Path(value["path"]).resolve()


def command_for(plan: dict[str, Any], lane: str, row: dict[str, Any], child: Path) -> tuple[list[str], list[str], Path]:
    binary = binary_for(plan, lane, row["role"])
    output = child / "output"
    time_file = child / "time-v.txt"
    direct = [
        "/usr/bin/taskset", "--cpu-list", "12-15", str(binary),
        "--shape", row["shape"], "--workload", row["workload"],
        "--workers", "1", "--warmup", str(row["warmup_iterations"]),
        "--samples", str(row["measured_samples"]), "--output", str(output),
    ]
    wrapped = [
        "/usr/bin/taskset", "--cpu-list", "12-15", "/usr/bin/time", "-v",
        "-o", str(time_file), "--", *direct[3:],
    ]
    return wrapped, direct, output


def process_group_clear(pid: int, timeout: float = 0.0) -> bool:
    """Return true only after the whole session process group has gone."""
    deadline = time.monotonic() + timeout
    while True:
        try:
            os.killpg(pid, 0)
        except ProcessLookupError:
            return True
        except PermissionError:
            return False
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.05)


def close_process_group(process: subprocess.Popen[bytes]) -> bool:
    """Boundedly terminate, reap, and verify a child process group."""
    pid = process.pid
    if process_group_clear(pid):
        return True
    try:
        os.killpg(pid, signal.SIGTERM)
    except ProcessLookupError:
        pass
    try:
        process.wait(timeout=2)
    except subprocess.TimeoutExpired:
        pass
    if not process_group_clear(pid):
        try:
            os.killpg(pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        try:
            process.wait(timeout=2)
        except subprocess.TimeoutExpired:
            pass
    return process_group_clear(pid, timeout=5)


def run_child(index: int, row: dict[str, Any], plan: dict[str, Any], lane: str, packet: Path) -> None:
    child = packet / "children" / row["child_id"]
    child.mkdir(parents=True, exist_ok=False)
    wrapped, direct, output = command_for(plan, lane, row, child)
    binary = binary_for(plan, lane, row["role"])
    cwd = Path(plan["lanes"][lane]["execution"]["workload_cwd"])
    env = child_environment(plan)
    start_record = {
        "schema": RUNNER_SCHEMA + ".child-start",
        "index": index,
        "child": row,
        "binary": descriptor(binary, executable=True),
        "argv": wrapped,
        "workload_argv": direct,
        "cwd": str(cwd),
        "environment": env,
        "deadline_seconds": DEADLINE_SECONDS,
        "started_utc": utc(),
    }
    write_exclusive(child / "start.json", start_record)
    stdout_path = child / "stdout.log"
    stderr_path = child / "stderr.log"
    process: subprocess.Popen[bytes] | None = None
    terminal: dict[str, Any]
    started = time.monotonic()
    try:
        require(not output.exists(), "benchmark output must not exist before child launch")
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            def set_affinity() -> None:
                os.sched_setaffinity(0, set(CPU_MASK))

            process = subprocess.Popen(
                wrapped,
                cwd=str(cwd),
                env=env,
                stdin=subprocess.DEVNULL,
                stdout=stdout,
                stderr=stderr,
                start_new_session=True,
                preexec_fn=set_affinity,
            )
            observed_affinity = sorted(os.sched_getaffinity(process.pid))
            require(observed_affinity == CPU_MASK, f"child affinity differs from frozen mask: {observed_affinity}")
            write_exclusive(
                child / "launch.json",
                {
                    "schema": RUNNER_SCHEMA + ".child-launch",
                    "pid": process.pid,
                    "observed_affinity": observed_affinity,
                    "launched_utc": utc(),
                },
            )
            try:
                return_code = process.wait(timeout=DEADLINE_SECONDS)
                timed_out = False
            except subprocess.TimeoutExpired:
                group_clear_after_cleanup = close_process_group(process)
                return_code = process.returncode
                timed_out = True
            else:
                # A zero return code is insufficient: taskset/time may have
                # exited while a descendant remains in the session.
                group_clear_after_cleanup = process_group_clear(process.pid)
                if not group_clear_after_cleanup:
                    close_process_group(process)
    except BaseException as exc:
        # A parent process may have exited while a descendant remains in the
        # session.  Always use the bounded terminate/reap path for any
        # launched process; a one-shot group probe is insufficient.
        if process is not None:
            group_clear_after_cleanup = close_process_group(process)
        else:
            group_clear_after_cleanup = False
        terminal = {
            "schema": RUNNER_SCHEMA + ".child-terminal",
            "index": index,
            "child_id": row["child_id"],
            "status": "launch_or_observation_failed",
            "return_code": None if process is None else process.returncode,
            "timed_out": False,
            "process_group_clear": group_clear_after_cleanup,
            "error": str(exc),
            "ended_utc": utc(),
        }
        write_exclusive(child / "terminal.json", terminal)
        write_exclusive(child / "validation.json", {"schema": RUNNER_SCHEMA + ".child-validation", "status": "failed", "error": str(exc), "terminal": descriptor(child / "terminal.json")})
        raise
    terminal = {
        "schema": RUNNER_SCHEMA + ".child-terminal",
        "index": index,
        "child_id": row["child_id"],
        "status": "timed_out" if timed_out else ("process_group_not_clear" if not group_clear_after_cleanup else ("completed" if return_code == 0 else "process_failed")),
        "return_code": return_code,
        "timed_out": timed_out,
        "process_group_clear": group_clear_after_cleanup,
        "pid": None if process is None else process.pid,
        "elapsed_seconds": time.monotonic() - started,
        "ended_utc": utc(),
    }
    write_exclusive(child / "terminal.json", terminal)
    # Terminal custody is immutable and is intentionally written before this point's parsing.
    try:
        require(return_code == 0 and not timed_out, f"child terminal failed: rc={return_code} timeout={timed_out}")
        require(group_clear_after_cleanup is True, "child process group did not clear")
        time_metrics = parse_time(child / "time-v.txt")
        report = output / "report.json"
        require(report.stat().st_size <= REPORT_LIMIT, "report exceeds capture size limit")
        value = read_json(report, limit=REPORT_LIMIT)
        require(isinstance(value, dict), "report root is not an object")
        if lane == "native":
            validate_native(value, row, plan, binary)
        else:
            validate_allocator(value, row, plan, binary)
        validation = {
            "schema": RUNNER_SCHEMA + ".child-validation",
            "status": "passed",
            "index": index,
            "child": row,
            "terminal": descriptor(child / "terminal.json"),
            "report": descriptor(report),
            "time": descriptor(child / "time-v.txt"),
            "time_metrics": time_metrics,
        }
        write_exclusive(child / "validation.json", validation)
        write_exclusive(
            child / "post-terminal-artifacts.json",
            {
                "schema": RUNNER_SCHEMA + ".post-terminal-artifacts",
                "terminal": descriptor(child / "terminal.json"),
                "stdout": descriptor(stdout_path),
                "stderr": descriptor(stderr_path),
                "time": descriptor(child / "time-v.txt"),
                "report": descriptor(report),
                "validation": descriptor(child / "validation.json"),
            },
        )
    except BaseException as exc:
        write_exclusive(
            child / "validation.json",
            {
                "schema": RUNNER_SCHEMA + ".child-validation",
                "status": "failed",
                "index": index,
                "child": row,
                "terminal": descriptor(child / "terminal.json"),
                "error": str(exc),
            },
        )
        raise


def run_lane(plan_path: Path, plan: dict[str, Any], lane: str, rows: list[dict[str, Any]], packet: Path) -> None:
    packet.mkdir(parents=True, exist_ok=False)
    write_exclusive(
        packet / "capture-start.json",
        {
            "schema": RUNNER_SCHEMA + ".capture-start",
            "plan": descriptor(plan_path),
            "runner": descriptor(Path(__file__), executable=True),
            "lane": lane,
            "children_expected": len(rows),
            "serial": True,
            "started_utc": utc(),
        },
    )
    status = "completed"
    error = None
    started = time.monotonic()
    for index, row in enumerate(rows):
        try:
            run_child(index, row, plan, lane, packet)
        except BaseException as exc:
            status = "failed"
            error = str(exc)
            break
    terminal = {
        "schema": RUNNER_SCHEMA + ".capture-terminal",
        "status": status,
        "lane": lane,
        "children_started": index + 1 if rows else 0,
        "children_expected": len(rows),
        "terminal_before_parse": True,
        "retry_or_replacement": False,
        "elapsed_seconds": time.monotonic() - started,
        "capture_start": descriptor(packet / "capture-start.json"),
        "ended_utc": utc(),
    }
    if error is not None:
        terminal["error"] = error
    write_exclusive(packet / "capture-terminal.json", terminal)
    if status != "completed":
        raise CaptureError(error or "capture failed")
    write_exclusive(
        packet / "completed-packet.json",
        {
            "schema": RUNNER_SCHEMA + ".completed-packet",
            "status": "completed",
            "lane": lane,
            "children": [
                {"child_id": row["child_id"], "validation": descriptor(packet / "children" / row["child_id"] / "validation.json")}
                for row in rows
            ],
            "capture_terminal": descriptor(packet / "capture-terminal.json"),
        },
    )


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--plan", type=Path, default=ROOT / "plan-template-0959-v3.json")
    parser.add_argument("--lane", choices=tuple(LANES), required=True)
    parser.add_argument("--packet", type=Path)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--execute", action="store_true")
    args = parser.parse_args(argv)
    require(not (args.dry_run and args.execute), "--dry-run and --execute are mutually exclusive")
    plan_path = args.plan.resolve()
    value = read_json(plan_path, limit=16 * 1024 * 1024)
    require(isinstance(value, dict), "plan root must be an object")
    rows = validate_plan(value, args.lane, for_execution=args.execute)
    if not args.execute:
        print(json.dumps({"status": "dry-run-pass", "lane": args.lane, "children": len(rows), "cases": len(EXPECTED_CASES), "pairs_per_case": 12 if args.lane == "native" else 3}, sort_keys=True))
        return 0
    require(os.environ.get(APPROVAL_ENV) == "1", f"{APPROVAL_ENV}=1 is required for execution")
    require(args.packet is not None, "--packet is required with --execute")
    packet = args.packet.resolve()
    require(packet.parent.is_dir(), f"packet parent is unavailable: {packet.parent}")
    run_lane(plan_path, value, args.lane, rows, packet)
    print(packet)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except CaptureError as exc:
        print(f"capture_0959_v6: {exc}", file=sys.stderr)
        raise SystemExit(2)
