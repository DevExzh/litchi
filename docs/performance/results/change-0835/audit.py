#!/usr/bin/env python3
"""Independent custody and statistics audit for the 0835 baseline packet.

The root-owned driver records builds and workloads.  This reader has a separate
implementation of the packet checks: it never imports ``analyze.py``, launches
anything, runs Git, or consults a live target directory except to verify a
descriptor.  The resulting audit is intentionally descriptive.  It binds the
measurement to the committed 0834-final source and does not authorize an
optimization or before/after claim.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import random
import statistics
import sys
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
TARGET = ROOT.parent / "litchi-target-0835"
SCRATCH = ROOT.parent / "litchi-fs-0835"
BASE = "77adc4f1e24bdf76f5e34ed4dfc0113e25120c59"
TOOL = "tools/perf-baseline/Cargo.toml"
ALLOWED_SOURCES = (
    "tools/perf-baseline/src/filesystem.rs",
    "tools/perf-baseline/src/filesystem/aligned_zip.rs",
    "tools/perf-baseline/README.md",
)
CASES = (
    "opc_file_eager_open",
    "opc_file_source_open",
    "opc_file_eager_one_part_atomic_save",
    "opc_file_source_one_part_atomic_save",
    "pptx_file_eager_open_selected_slide_lifecycle",
    "pptx_file_source_open_selected_slide_lifecycle",
)
PLAN_SCHEMA = "litchi.0835.filesystem-plan.v1"
AUDIT_SCHEMA = "litchi.0835.independent-custody-audit.v1"
EXPECTED_SCOPE = (
    "Fresh descriptive filesystem route/cache baseline on committed repaired "
    "harness; iWork excluded"
)
EXPECTED_LIMITS = [
    "No before/after optimization claim",
    "Cold proves page cache and procfs only",
    "Warm and cold ZIP alignment differ; no cross-cache byte identity assumption",
    "PPTX logical ReadAt evidence is an untimed source replay",
    "Eager route logical read counters are unavailable, not zero I/O",
    "No allocation observer or hardware counter claim",
    "Default atomic save durability retained",
    "One host and deterministic synthetic corpora",
]


class AuditError(RuntimeError):
    """The retained packet is missing, stale, or internally inconsistent."""


def fail(message: str) -> None:
    raise AuditError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1 << 20), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def valid_sha(value: Any) -> bool:
    return (isinstance(value, str) and len(value) == 64
            and all(char in "0123456789abcdef" for char in value))


def path_inside(path: Path, root: Path, label: str) -> Path:
    resolved = path.resolve(strict=False)
    require(resolved.is_relative_to(root), f"{label} escaped {root}: {path}")
    return resolved


def packet_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label} path is missing")
    path = Path(raw)
    return path_inside(path if path.is_absolute() else PACKET / path,
                       PACKET, label)


def root_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw and not Path(raw).is_absolute(),
            f"{label} path is invalid")
    return path_inside(ROOT / raw, ROOT, label)


def contains_descriptor(value: Any, expected: dict[str, Any]) -> bool:
    if isinstance(value, dict):
        if all(value.get(key) == expected.get(key)
               for key in ("path", "bytes", "sha256")):
            return True
        return any(contains_descriptor(item, expected) for item in value.values())
    if isinstance(value, list):
        return any(contains_descriptor(item, expected) for item in value)
    return False


def cleanup_witness() -> dict[str, Any]:
    path = PACKET / "cleanup.json"
    require(path.is_file() and not path.is_symlink(),
            "missing cleanup witness for removed evidence")
    value = read_json(path)
    require(value.get("status") == "pass", "cleanup witness is not terminal pass")
    return value


def descriptor_value(value: Any, label: str, *, allow_cleanup: bool = False,
                     require_nonempty: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} descriptor is malformed")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and raw_path,
            f"{label} descriptor path is missing")
    path = Path(raw_path)
    if not path.is_absolute():
        path = PACKET / path
    path = path.resolve(strict=False)
    require(not path.is_symlink(), f"{label} descriptor is a symlink")
    require(type(value.get("bytes")) is int
            and value["bytes"] >= (1 if require_nonempty else 0),
            f"{label} descriptor byte count is invalid")
    require(valid_sha(value.get("sha256")), f"{label} descriptor hash is invalid")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == value["bytes"],
                f"{label} descriptor bytes changed")
        require(sha256(path) == value["sha256"],
                f"{label} descriptor hash changed")
    else:
        require(allow_cleanup, f"missing {label} without cleanup allowance")
        expected = {"path": raw_path, "bytes": value["bytes"],
                    "sha256": value["sha256"]}
        require(contains_descriptor(cleanup_witness(), expected),
                f"cleanup witness does not retain {label}")
    return {"path": raw_path, "bytes": value["bytes"],
            "sha256": value["sha256"]}


def descriptor(path: Path, label: str, *, allow_cleanup: bool = False,
               require_nonempty: bool = False) -> dict[str, Any]:
    regular = path.is_file() and not path.is_symlink()
    value = {"path": str(path),
             "bytes": path.stat().st_size if regular else (1 if require_nonempty else 0),
             "sha256": sha256(path) if regular else "0" * 64}
    return descriptor_value(value, label, allow_cleanup=allow_cleanup,
                           require_nonempty=require_nonempty)


def source_descriptor(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    return {"path": str(path), "bytes": path.stat().st_size,
            "sha256": sha256(path)}


def same_descriptor(left: Any, right: Any, label: str) -> None:
    require(isinstance(left, dict) and isinstance(right, dict),
            f"{label} descriptor is missing")
    for key in ("path", "bytes", "sha256"):
        require(left.get(key) == right.get(key), f"{label} {key} differs")


def load_origin() -> dict[str, Any]:
    origin_path = PACKET / "origin.json"
    value = read_json(origin_path)
    require(value.get("base") == BASE, "origin base changed")
    require(value.get("scope") == EXPECTED_SCOPE, "origin scope changed")
    normative = value.get("normative")
    unrelated = value.get("unrelated")
    require(isinstance(normative, dict) and normative,
            "normative custody map is malformed")
    require(isinstance(unrelated, dict) and unrelated,
            "unrelated custody map is malformed")
    for raw, expected in (*normative.items(), *unrelated.items()):
        require(valid_sha(expected), f"origin hash is malformed: {raw}")
        input_path = root_path(raw, "origin input")
        require(input_path.is_file() and sha256(input_path) == expected,
                f"origin input changed: {raw}")
    return {"path": str(origin_path.relative_to(PACKET)),
            "bytes": origin_path.stat().st_size,
            "sha256": sha256(origin_path), "base": BASE,
            "normative_count": len(normative), "unrelated_count": len(unrelated)}


def validate_freeze() -> dict[str, Any]:
    path = PACKET / "freeze-baseline.json"
    value = read_json(path)
    require(value.get("stage") == "baseline", "baseline freeze stage changed")
    source = value.get("source")
    require(isinstance(source, dict) and len(source) >= 9000,
            "baseline source inventory is missing or unexpectedly small")
    for raw, expected in source.items():
        require(isinstance(raw, str) and not Path(raw).is_absolute()
                and valid_sha(expected), f"baseline source entry is malformed: {raw}")
        live = root_path(raw, "baseline source")
        require(live.is_file() and sha256(live) == expected,
                f"baseline source changed: {raw}")
    source_root = PACKET / "sources" / "baseline"
    require(source_root.is_dir() and not source_root.is_symlink(),
            "baseline source archive is missing")
    archived: dict[str, Path] = {}
    for item in source_root.rglob("*"):
        require(not item.is_symlink(), "baseline source archive contains a symlink")
        if item.is_file():
            archived[str(item.relative_to(source_root))] = item
    require(set(archived) == set(ALLOWED_SOURCES),
            "baseline source archive boundary changed")
    for raw, item in archived.items():
        require(sha256(item) == source[raw], f"baseline source archive changed: {raw}")
    same_descriptor(value.get("driver"), source_descriptor(PACKET / "driver.py", "driver.py"),
                    "baseline freeze driver")
    same_descriptor(value.get("origin"), source_descriptor(PACKET / "origin.json", "origin.json"),
                    "baseline freeze origin")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "stage": "baseline",
            "source_count": len(source),
            "driver": value["driver"], "origin": value["origin"]}


def validate_host() -> dict[str, Any]:
    path = PACKET / "host.json"
    value = read_json(path)
    require(value.get("base") == BASE, "host base changed")
    affinity = value.get("affinity")
    require(isinstance(affinity, list) and 12 in affinity,
            "host CPU affinity changed")
    require(isinstance(value.get("rustc"), str)
            and isinstance(value.get("cargo"), str)
            and isinstance(value.get("cpu"), str)
            and isinstance(value.get("memory"), str),
            "host toolchain or hardware fields are missing")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "base": BASE,
            "affinity": affinity}


def validate_quality_reuse() -> dict[str, Any]:
    path = PACKET / "quality-reuse.json"
    value = read_json(path)
    require(value.get("status") == "pass" and value.get("reused") is True
            and value.get("source_matches_exactly") is True
            and value.get("source_count") == 9389,
            "quality reuse receipt changed")
    prior = value.get("prior_packet")
    require(isinstance(prior, str), "quality reuse prior packet is missing")
    prior_path = Path(prior).resolve(strict=False)
    require(prior_path == (PACKET.parent / "change-0834").resolve(),
            "quality reuse prior packet changed")
    inputs = value.get("inputs")
    require(isinstance(inputs, dict) and inputs, "quality reuse inputs are missing")
    for raw, binding in inputs.items():
        require(isinstance(raw, str) and not Path(raw).is_absolute(),
                f"quality reuse input path is invalid: {raw}")
        actual = prior_path / raw
        same_descriptor(binding, descriptor(actual, f"quality reuse input {raw}"),
                        f"quality reuse input {raw}")
    scope = value.get("scope")
    require(isinstance(scope, str) and "No new full-suite run claimed" in scope,
            "quality reuse scope does not disclose the reused suite")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "status": "pass", "source_count": 9389,
            "input_count": len(inputs), "prior_packet": prior}


def build_argv() -> list[str]:
    return ["cargo", "build", "--offline", "--locked", "--release",
            "--manifest-path", TOOL, "--bin", "litchi-perf-baseline"]


def workload_argv(binary: str, case: str, state: str, samples: int,
                  warmup: int, report: Path) -> list[str]:
    return ["taskset", "-c", "12", binary, "--case", case,
            "--samples", str(samples), "--warmup", str(warmup),
            "--filesystem-cache", state, "--filesystem-root", str(SCRATCH),
            "--json", str(report)]


def read_command(label: str, expected_argv: list[str], stage: str = "baseline",
                 expected_exit: int = 0) -> dict[str, Any]:
    root = PACKET / "commands" / label
    require(root.is_dir() and not root.is_symlink(),
            f"missing command directory: {label}")
    started_path, receipt_path, log_path = (root / name for name in
                                             ("started.json", "receipt.json", "output.log"))
    started, receipt = read_json(started_path), read_json(receipt_path)
    require(started.get("argv") == expected_argv
            and receipt.get("argv") == expected_argv,
            f"{label} argv changed")
    require(started.get("cwd") == str(ROOT), f"{label} cwd changed")
    started_unix, finished_unix = started.get("started_unix"), receipt.get("finished_unix")
    require(receipt.get("started_unix") == started_unix
            and type(started_unix) in (int, float)
            and not isinstance(started_unix, bool)
            and math.isfinite(float(started_unix))
            and type(finished_unix) in (int, float)
            and not isinstance(finished_unix, bool)
            and math.isfinite(float(finished_unix))
            and finished_unix >= started_unix,
            f"{label} chronology changed")
    require(receipt.get("exit_code") == expected_exit
            and receipt.get("error") is None, f"{label} exit/error changed")
    expected_freeze = descriptor(PACKET / "freeze-baseline.json", "baseline freeze")
    same_descriptor(started.get("freeze"), expected_freeze,
                    f"{label} start freeze")
    same_descriptor(receipt.get("freeze"), expected_freeze,
                    f"{label} receipt freeze")
    same_descriptor(receipt.get("log"), descriptor(log_path, f"{label} log"),
                    f"{label} log")
    return {"name": label, "started": descriptor(started_path, f"{label} started"),
            "receipt": descriptor(receipt_path, f"{label} receipt"),
            "log": descriptor(log_path, f"{label} log"), "argv": expected_argv,
            "exit_code": expected_exit, "started_unix": started_unix,
            "finished_unix": finished_unix, "stage": stage}


def validate_build() -> tuple[dict[str, Any], dict[str, Any]]:
    path = PACKET / "build-baseline.json"
    value = read_json(path)
    binary = descriptor_value(value.get("binary"), "baseline binary",
                              allow_cleanup=True, require_nonempty=True)
    binary_path = Path(binary["path"]).resolve(strict=False)
    require(binary_path.name == "litchi-perf-baseline"
            and binary_path.parent.name == "baseline"
            and binary_path.is_relative_to(TARGET),
            "baseline binary path changed")
    command = read_command("build-baseline", build_argv())
    same_descriptor(value.get("receipt"), command["receipt"],
                    "baseline build receipt")
    return ({"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
             "sha256": sha256(path), "binary": binary,
             "receipt": value["receipt"]}, command)


def validate_report_descriptor(binding: Any, path: Path, label: str) -> dict[str, Any]:
    actual = descriptor(path, label)
    same_descriptor(binding, actual, label)
    return actual


def count_report_samples(path: Path, label: str, expected_states: int) -> int:
    value = read_json(path)
    results, evidence = value.get("results"), value.get("filesystem_evidence")
    require(isinstance(results, list) and len(results) == expected_states
            and isinstance(evidence, list) and len(evidence) * 2 == expected_states,
            f"{label} result/evidence count changed")
    count = 0
    for item in evidence:
        samples = item.get("samples") if isinstance(item, dict) else None
        require(isinstance(samples, list) and len(samples) == 2,
                f"{label} warm/cold sample count changed")
        states = sorted(sample.get("cache_state") for sample in samples)
        require(states == ["cold-verified", "warm"],
                f"{label} cache states changed")
        count += len(samples)
    require(count == expected_states, f"{label} sample count changed")
    return count


def validate_workload(row: Any, label: str, binary: dict[str, Any], case: str,
                      state: str, samples: int, warmup: int,
                      stage: str = "baseline") -> dict[str, Any]:
    require(isinstance(row, dict) and row.get("case") == case
            and row.get("state") == state and row.get("exit_code") == 0,
            f"{label} workload row changed")
    report = PACKET / f"{label}.json"
    command = read_command(label,
                           workload_argv(binary["path"], case, state, samples,
                                         warmup, report), stage)
    require(report.is_file() and not report.is_symlink(),
            f"{label} report is missing")
    validate_report_descriptor(row.get("report"), report, f"{label} report")
    same_descriptor(row.get("receipt"), command["receipt"], f"{label} receipt")
    return command


def expected_plan_rows() -> list[dict[str, Any]]:
    rows = []
    for block in range(6):
        order = CASES if block % 2 == 0 else tuple(reversed(CASES))
        states = ("warm", "cold-verified") if block % 2 == 0 else (
            "cold-verified", "warm")
        for state in states:
            rows.extend({"block": block, "case": case, "cache_state": state,
                         "samples": 30, "warmup": 3} for case in order)
    return rows


def validate_plan() -> dict[str, Any]:
    path = PACKET / "measurement-plan.json"
    value = read_json(path)
    require(value.get("schema") == PLAN_SCHEMA and value.get("base") == BASE
            and value.get("cpu") == 12 and value.get("expected_reports") == 72
            and value.get("expected_samples") == 2160
            and value.get("rows") == expected_plan_rows(),
            "measurement plan changed")
    require(value.get("purpose") == "Descriptive current-source route/cache baseline; no production optimization",
            "measurement purpose changed")
    require(value.get("limits") == EXPECTED_LIMITS,
            "measurement claim limits changed")
    stats = value.get("statistics")
    require(isinstance(stats, dict)
            and stats.get("process_quantiles")
            == "p50 integer midpoint; p95/p99 nearest rank"
            and stats.get("summary") == "midpoint median of six process p50 values"
            and stats.get("paired_comparisons")
            == "eager/source per operation and cache state; block-paired p50 ratios"
            and stats.get("bootstrap_resamples") == 10000
            and stats.get("seed") == 835083
            and stats.get("sorted_endpoints") == [249, 9749]
            and stats.get("spread_flag")
            == "max/min block p50, p95, p99 or peak RSS exceeds 1.2"
            and stats.get("tail_flag")
            == "none; p99 equals block maximum with 30 samples",
            "measurement statistics plan changed")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "reports": 72, "samples": 2160}


def validate_qualification(binary: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    path = PACKET / "qualification.json"
    value = read_json(path)
    require(value.get("status") == "commands_pass"
            and isinstance(value.get("rows"), list)
            and len(value["rows"]) == 7,
            "qualification manifest changed")
    rows = value["rows"]
    commands = []
    reports = 0
    samples_count = 0
    labels = [f"qualification-{index:02}" for index in range(6)]
    labels.append("qualification-opc-pair")
    for index, label in enumerate(labels):
        case = CASES[index] if index < 6 else ",".join(CASES[2:4])
        command = validate_workload(rows[index], label, binary, case,
                                    "warm,cold-verified", 1, 0)
        commands.append(command)
        report_path = PACKET / f"{label}.json"
        samples_count += count_report_samples(report_path, label,
                                               4 if index == 6 else 2)
        reports += 1
    require(reports == 7 and samples_count == 16,
            "qualification retained counts changed")
    return ({"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
             "sha256": sha256(path), "status": "commands_pass",
             "reports": reports, "samples": samples_count}, commands)


def validate_admission() -> dict[str, Any]:
    path = PACKET / "admission.json"
    value = read_json(path)
    require(value.get("status") == "pass"
            and isinstance(value.get("inputs"), dict)
            and value["inputs"]
            and value.get("validated_sample_count") == 16
            and value.get("performance_claim")
            == "descriptive route/cache baseline only",
            "admission receipt changed")
    for raw, binding in value["inputs"].items():
        require(isinstance(raw, str) and raw and not Path(raw).is_absolute(),
                "admission input path is invalid")
        actual = packet_path(raw, "admission input")
        same_descriptor(binding, descriptor(actual, f"admission input {raw}"),
                        f"admission input {raw}")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "status": "pass",
            "input_count": len(value["inputs"])}


def validate_capture(binary: dict[str, Any], plan: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    path = PACKET / "capture.json"
    value = read_json(path)
    require(value.get("status") == "commands_pass"
            and value.get("report_count") == 72
            and value.get("sample_count") == 2160
            and isinstance(value.get("rows"), list)
            and len(value["rows"]) == 72, "formal capture receipt changed")
    started_path = PACKET / "capture-started.json"
    started = read_json(started_path)
    same_descriptor(started.get("plan"), descriptor(PACKET / "measurement-plan.json", "capture plan"),
                    "capture plan")
    same_descriptor(started.get("admission"), descriptor(PACKET / "admission.json", "capture admission"),
                    "capture admission")
    same_descriptor(started.get("driver"), descriptor(PACKET / "capture.py", "capture driver"),
                    "capture driver")
    same_descriptor(started.get("freeze"), descriptor(PACKET / "freeze-baseline.json", "capture freeze"),
                    "capture freeze")
    require(not (PACKET / "capture-failed.json").exists(),
            "formal capture has a retained failed attempt")
    commands = []
    rows = value["rows"]
    expected = expected_plan_rows()
    for index, row in enumerate(rows):
        require(row.get("plan") == expected[index],
                f"formal capture plan row changed: {index}")
        commands.append(validate_workload(row.get("result"), f"native-{index:03}",
                                          binary, expected[index]["case"],
                                          expected[index]["cache_state"], 30, 3))
    require(len(commands) == 72, "formal capture command count changed")
    return ({"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
             "sha256": sha256(path), "reports": 72, "samples": 2160},
            commands, value)


def midpoint(left: int, right: int) -> int:
    """Match the native harness' overflow-safe integer midpoint."""
    return left // 2 + right // 2 + ((left % 2 + right % 2) // 2)


def nearest_rank(values: list[int], quantile: float) -> int:
    require(values, "empty metric vector")
    ordered = sorted(values)
    return ordered[max(1, math.ceil(len(ordered) * quantile)) - 1]


def sample_summary(values: list[int | float]) -> dict[str, Any]:
    require(values and all(type(value) in (int, float)
                           and not isinstance(value, bool) for value in values),
            "invalid sample vector")
    ordered = sorted(values)
    if len(ordered) % 2:
        p50: int | float = ordered[len(ordered) // 2]
    else:
        p50 = midpoint(int(ordered[len(ordered) // 2 - 1]),
                       int(ordered[len(ordered) // 2]))
    integer_values = [int(value) for value in ordered]
    return {"min": ordered[0], "p50": p50,
            "p95": nearest_rank(integer_values, .95),
            "p99": nearest_rank(integer_values, .99),
            "max": ordered[-1], "mean": statistics.mean(ordered)}


def spread_summary(values: list[int | float]) -> dict[str, Any]:
    require(values, "empty block vector")
    minimum = min(values)
    return {"min": minimum, "median": statistics.median(values), "max": max(values),
            "max_over_min": max(values) / minimum if minimum > 0 else None}


def report_samples(report: dict[str, Any], plan_row: dict[str, Any], label: str,
                   seen_pids: set[int]) -> tuple[list[dict[str, Any]], str, dict[str, Any]]:
    """Validate one formal report and return samples in evidence order."""
    case, state = plan_row["case"], plan_row["cache_state"]
    require(isinstance(report, dict), f"{label}: report is not an object")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label}: configuration missing")
    require(configuration.get("samples_per_case") == 30
            and configuration.get("warmup_iterations_per_case") == 3
            and configuration.get("filesystem_cache_states") == [state],
            f"{label}: configuration changed")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label}: expected one timed result")
    result = results[0]
    require(isinstance(result, dict) and result.get("case") == case
            and result.get("cache_state") == state,
            f"{label}: timed result changed")
    evidence_list = report.get("filesystem_evidence")
    require(isinstance(evidence_list, list) and len(evidence_list) == 1,
            f"{label}: evidence missing")
    evidence = evidence_list[0]
    require(isinstance(evidence, dict) and evidence.get("case") == case
            and evidence.get("cache_states") == [state]
            and evidence.get("sample_count") == 30,
            f"{label}: evidence identity changed")
    raw_samples = evidence.get("samples")
    require(isinstance(raw_samples, list) and len(raw_samples) == 30,
            f"{label}: samples missing")
    by_index: dict[int, dict[str, Any]] = {}
    for position, sample in enumerate(raw_samples):
        require(isinstance(sample, dict), f"{label}.samples[{position}] malformed")
        index = sample.get("sample_index")
        require(type(index) is int and 0 <= index < 30
                and index not in by_index, f"{label}: sample index changed")
        require(sample.get("cache_state") == state,
                f"{label}.samples[{position}] state changed")
        elapsed = sample.get("elapsed_ns")
        pid = sample.get("child_process_id")
        require(type(elapsed) is int and elapsed > 0,
                f"{label}.samples[{position}] elapsed changed")
        require(type(pid) is int and pid > 0 and pid not in seen_pids,
                f"{label}.samples[{position}] process identity changed")
        seen_pids.add(pid)
        metrics = sample.get("process_metrics")
        require(isinstance(metrics, dict) and "clock_ticks_per_second" in metrics,
                f"{label}.samples[{position}] process metrics missing")
        require(all(type(value) is int and value >= 0 for value in metrics.values()),
                f"{label}.samples[{position}] process metric changed")
        if state == "cold-verified":
            proof = sample.get("cold_verified")
            require(isinstance(proof, dict) and proof.get("status") == "eligible",
                    f"{label}.samples[{position}] cold proof missing")
        by_index[index] = sample
    require(set(by_index) == set(range(30)), f"{label}: sample index set changed")
    elapsed_stats = result.get("elapsed_ns")
    require(isinstance(elapsed_stats, dict), f"{label}: elapsed statistics missing")
    sorted_values = elapsed_stats.get("samples")
    sample_order = elapsed_stats.get("sample_order")
    require(isinstance(sorted_values, list) and len(sorted_values) == 30
            and isinstance(sample_order, list) and len(sample_order) == 30
            and sorted(sample_order) == list(range(30)),
            f"{label}: elapsed ordering changed")
    reconstructed = [by_index[index]["elapsed_ns"] for index in sample_order]
    require(reconstructed == sorted_values,
            f"{label}: elapsed sample order does not match evidence")
    expected_elapsed = sample_summary(reconstructed)
    for field in ("min", "p50", "p95", "p99", "max"):
        require(elapsed_stats.get(field) == expected_elapsed[field],
                f"{label}.elapsed_ns.{field} differs from raw samples")
    require(math.isclose(float(elapsed_stats.get("mean")),
                         float(expected_elapsed["mean"]), rel_tol=0.0, abs_tol=1e-6),
            f"{label}.elapsed_ns.mean differs from raw samples")
    binary = report.get("binary_identity")
    environment = report.get("environment")
    require(isinstance(binary, dict) and valid_sha(binary.get("binary_sha256"))
            and isinstance(environment, dict), f"{label}: identity missing")
    if environment.get("git_revision") is not None:
        require(environment.get("git_revision") == BASE,
                f"{label}: report source revision changed")
    return list(by_index.values()), binary["binary_sha256"], {
        "environment": environment,
        "process_keys": sorted(key for key in by_index[0]["process_metrics"]
                                if key != "clock_ticks_per_second"),
        "elapsed": [by_index[index]["elapsed_ns"] for index in range(30)],
    }


def admission_inputs(admission: dict[str, Any]) -> dict[str, str]:
    require(admission.get("status") == "pass", "capture admission is not pass")
    inputs = admission.get("inputs")
    require(isinstance(inputs, dict) and inputs, "capture admission has no inputs")
    result: dict[str, str] = {}
    for name, binding in sorted(inputs.items()):
        if isinstance(binding, dict) and "descriptor" in binding:
            binding = binding["descriptor"]
        actual = packet_path(name, f"capture admission input {name}")
        same_descriptor(binding, descriptor(actual, f"capture admission input {name}"),
                        f"capture admission input {name}")
        result[name] = binding["sha256"]
    return result


def recompute_analysis(capture: dict[str, Any], plan_value: dict[str, Any],
                       expected_binary_sha: str) -> dict[str, Any]:
    """Recompute analyze.py's full JSON result without importing it."""
    admission_path = PACKET / "capture-admission.json"
    if not admission_path.is_file():
        admission_path = PACKET / "admission.json"
    admission = read_json(admission_path)
    input_digests = admission_inputs(admission)
    plan_path = PACKET / "measurement-plan.json"
    plan_rows = expected_plan_rows()
    require(plan_value.get("rows") == plan_rows, "analysis plan rows changed")
    seen_pids: set[int] = set()
    groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
    environments: list[dict[str, Any]] = []
    binaries: set[str] = set()
    process_keys: set[str] | None = None
    report_descriptors: list[dict[str, Any]] = []
    receipt_descriptors: list[dict[str, Any]] = []
    for index, row in enumerate(capture["rows"]):
        label = f"capture.rows[{index}]"
        plan_row = row["plan"]
        require(plan_row == plan_rows[index], f"{label}.plan changed")
        result = row["result"]
        binding = result["report"]
        report_path = packet_path(binding["path"], f"{label}.report")
        receipt_path = packet_path(result["receipt"]["path"], f"{label}.receipt")
        validate_report_descriptor(binding, report_path, f"{label}.report")
        validate_report_descriptor(result["receipt"], receipt_path, f"{label}.receipt")
        receipt = read_json(receipt_path)
        require(receipt.get("exit_code") == 0 and receipt.get("error") is None,
                f"{label}: receipt changed")
        report = read_json(report_path)
        samples, binary_digest, details = report_samples(report, plan_row, label, seen_pids)
        require(binary_digest == expected_binary_sha,
                f"{label}: report binary differs from retained build")
        binaries.add(binary_digest)
        environments.append(details["environment"])
        keys = set(details["process_keys"])
        process_keys = keys if process_keys is None else process_keys & keys
        relative_report = str(report_path.relative_to(PACKET.resolve()))
        relative_receipt = str(receipt_path.relative_to(PACKET.resolve()))
        block_record = {
            "block": plan_row["block"], "report": relative_report,
            "report_sha256": binding["sha256"], "receipt": relative_receipt,
            "receipt_sha256": result["receipt"]["sha256"], "sample_count": len(samples),
            "latency_ns": sample_summary(details["elapsed"]),
            "process_metrics": {key: sample_summary(
                [sample["process_metrics"][key] for sample in samples])
                                for key in sorted(keys)},
        }
        groups.setdefault((plan_row["case"], plan_row["cache_state"]), []).append(block_record)
        report_descriptors.append({"path": relative_report, "bytes": report_path.stat().st_size,
                                   "sha256": binding["sha256"]})
        receipt_descriptors.append({"path": relative_receipt, "bytes": receipt_path.stat().st_size,
                                    "sha256": result["receipt"]["sha256"]})
    require(len(groups) == len(CASES) * 2 and len(binaries) == 1,
            "capture group or binary coverage changed")
    require(len(set(json.dumps(item, sort_keys=True) for item in environments)) == 1,
            "formal reports use more than one environment")
    require(process_keys is not None, "formal reports expose no process metrics")
    distributions, lookup, spread_flags = [], {}, []
    for case in CASES:
        for state in ("warm", "cold-verified"):
            rows = sorted(groups[(case, state)], key=lambda value: value["block"])
            require([row["block"] for row in rows] == list(range(6)),
                    f"{case}/{state}: block coverage changed")
            latency_spread = {metric: spread_summary(
                [row["latency_ns"][metric] for row in rows])
                              for metric in ("min", "p50", "p95", "p99", "max", "mean")}
            process_spread = {key: {metric: spread_summary(
                [row["process_metrics"][key][metric] for row in rows])
                                    for metric in ("min", "p50", "p95", "p99", "max", "mean")}
                             for key in sorted(process_keys)}
            for metric in ("p50", "p95", "p99"):
                value = latency_spread[metric]["max_over_min"]
                if value is not None and value > 1.2:
                    spread_flags.append({"case": case, "cache_state": state,
                                         "metric": metric,
                                         "reason": "six-block spread exceeds 20%",
                                         "max_over_min": value})
            peak_rss_p50 = process_spread.get("peak_rss_bytes", {}).get("p50", {})
            rss_spread = peak_rss_p50.get("max_over_min")
            if rss_spread is not None and rss_spread > 1.2:
                spread_flags.append({"case": case, "cache_state": state,
                                     "metric": "peak_rss_bytes.p50",
                                     "reason": "six-block spread exceeds 20%",
                                     "max_over_min": rss_spread})
            distribution = {"case": case, "cache_state": state, "blocks": rows,
                            "six_block_spread": {"latency_ns": latency_spread,
                                                  "process_metrics": process_spread}}
            distributions.append(distribution)
            lookup[(case, state)] = distribution
    route_pairs = (("opc_open", CASES[0], CASES[1]),
                   ("opc_one_part_atomic_save", CASES[2], CASES[3]),
                   ("pptx_selected_slide_lifecycle", CASES[4], CASES[5]))
    route_ratios = []
    # The plan is the authoritative place for the bootstrap seed.  This keeps
    # the audit tied to the frozen statistics contract if the packet changes it.
    seed = plan_value["statistics"]["seed"]
    for pair_name, eager, source in route_pairs:
        for state in ("warm", "cold-verified"):
            eager_blocks = lookup[(eager, state)]["blocks"]
            source_blocks = lookup[(source, state)]["blocks"]
            eager_p50 = [row["latency_ns"]["p50"] for row in eager_blocks]
            source_p50 = [row["latency_ns"]["p50"] for row in source_blocks]
            ratios = [eager_value / source_value
                      for eager_value, source_value in zip(eager_p50, source_p50)]
            require(all(value > 0 for value in ratios), "formal paired ratio invalid")
            rng = random.Random(seed)
            bootstrap = sorted(statistics.median(rng.choices(ratios, k=6))
                               for _ in range(10_000))
            route_ratios.append({"pair": pair_name, "eager_case": eager,
                                 "source_case": source, "cache_state": state,
                                 "eager_p50_ns_by_block": eager_p50,
                                 "source_p50_ns_by_block": source_p50,
                                 "eager_over_source_p50_by_block": ratios,
                                 "median": statistics.median(ratios),
                                 "bootstrap_95": [bootstrap[249], bootstrap[9749]],
                                 "bootstrap": {"method": "median of six block ratios, six draws with replacement",
                                               "seed": seed, "resamples": 10_000,
                                               "lower_index": 249, "upper_index": 9749},
                                 "scope": "descriptive route ratio; not an optimization effect"})
    base = plan_value.get("base")
    require(isinstance(base, str) and len(base) == 40, "analysis base is not a commit hash")
    return {"schema": "litchi.0835.descriptive-analysis.v1", "base": base,
            "measurement_plan": {"path": str(plan_path.relative_to(PACKET)),
                                  "sha256": sha256(plan_path), "blocks": 6,
                                  "rows": 72, "samples_per_row": 30, "warmups_per_row": 3,
                                  "samples": 2160, "cpu": 12,
                                  "cache_states": ["warm", "cold-verified"],
                                  "cases": list(CASES)},
            "capture": {"path": "capture.json", "sha256": sha256(PACKET / "capture.json"),
                        "status": capture["status"], "reports": 72, "samples": 2160},
            "admission": {"path": str(admission_path.relative_to(PACKET)),
                          "sha256": sha256(admission_path), "status": admission["status"],
                          "inputs": input_digests},
            "binary_sha256": next(iter(binaries)), "environment": environments[0],
            "reports": 72, "samples": 2160, "distributions": distributions,
            "route_ratios": route_ratios, "spread_flags": spread_flags,
            "report_descriptors": report_descriptors,
            "receipt_descriptors": receipt_descriptors,
            "bootstrap_seed": seed, "bootstrap_resamples": 10_000,
            "limitations": [
                "Synthetic fixed corpora and one host are represented by this capture.",
                "The OPC eager-open timer drops its package, while source-open retains the package through post-timer diagnostics; their ratio is descriptive and has asymmetric lifetime scope.",
                "PPTX logical source counters come from an untimed source replay and are not latency attribution.",
                "Verified-cold evidence observes page-cache residency and process read_bytes; it is not a physical-device I/O measurement.",
                "Each row uses three independent warmups and fresh process children for measured samples; counterbalancing reduces order effects but does not remove host noise.",
                "All route ratios and spread flags describe this current-source baseline; no before/after production optimization claim is authorized.",
                "iWork is outside the scope of this packet."]}


def validate_analysis(capture: dict[str, Any], plan_value: dict[str, Any],
                      expected_binary_sha: str) -> dict[str, Any]:
    path = PACKET / "analysis.json"
    require(path.is_file() and not path.is_symlink(),
            "formal capture exists without analysis.json")
    expected = recompute_analysis(capture, plan_value, expected_binary_sha)
    require(read_json(path) == expected, "formal analysis does not independently recompute")
    return {"analysis": source_descriptor(path, "formal analysis"),
            "analyzer": source_descriptor(PACKET / "analyze.py", "formal analyzer")}


def validate_serial(rows: list[dict[str, Any]]) -> None:
    ordered = sorted(rows, key=lambda row: (row["started_unix"], row["name"]))
    for previous, current in zip(ordered, ordered[1:]):
        require(previous["finished_unix"] <= current["started_unix"],
                f"command intervals overlap: {previous['name']} and {current['name']}")


def validate_command_inventory(rows: list[dict[str, Any]]) -> None:
    root = PACKET / "commands"
    require(root.is_dir() and not root.is_symlink(), "command receipt root is missing")
    actual = {item.name for item in root.iterdir()
              if item.is_dir() and not item.is_symlink()}
    expected = {row["name"] for row in rows}
    require(actual == expected and len(rows) == 80,
            f"command inventory changed: expected 80, found {len(rows)}")


def encoded(value: dict[str, Any]) -> bytes:
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    require(len(data) < 10 * 1024 * 1024, "audit output is unexpectedly large")
    return data


def build_audit() -> dict[str, Any]:
    origin = load_origin()
    freeze = validate_freeze()
    host = validate_host()
    quality = validate_quality_reuse()
    build, build_command = validate_build()
    qualification, qualification_commands = validate_qualification(build["binary"])
    admission = validate_admission()
    plan = validate_plan()
    plan_value = read_json(PACKET / "measurement-plan.json")
    capture, capture_commands, capture_value = validate_capture(build["binary"], plan)
    analysis = validate_analysis(capture_value, plan_value, build["binary"]["sha256"])
    commands = [build_command, *qualification_commands, *capture_commands]
    validate_command_inventory(commands)
    validate_serial(commands)
    return {
        "schema": AUDIT_SCHEMA, "status": "pass", "base": BASE,
        "custody": {"origin": origin, "freeze": freeze, "quality_reuse": quality},
        "host": host, "build": build,
        "qualification": qualification, "admission": admission, "plan": plan,
        "capture": capture, "analysis": analysis,
        "commands": {"count": len(commands), "serial": True, "rows": commands},
        "claims": {"performance_claim": "none",
                    "scope": "Descriptive current-source route/cache baseline; no production optimization",
                    "iwork": "excluded"},
    }


def write_or_check(value: dict[str, Any], check: bool, preview: bool) -> None:
    path = PACKET / "audit.json"
    data = encoded(value)
    if preview:
        return
    if check or path.exists():
        require(path.is_file() and not path.is_symlink(), "audit.json is missing")
        require(path.read_bytes() == data, "audit.json does not replay deterministically")
        return
    try:
        with path.open("xb") as stream:
            stream.write(data)
    except FileExistsError:
        fail("refusing to overwrite retained audit.json")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--check", action="store_true",
                       help="replay and compare retained audit.json")
    modes.add_argument("--preview", action="store_true",
                       help="validate without writing audit.json")
    args = parser.parse_args(argv)
    try:
        write_or_check(build_audit(), args.check, args.preview)
    except (AuditError, AssertionError, OSError, UnicodeError, ValueError,
            KeyError, TypeError, IndexError) as error:
        print(f"0835 independent audit failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps({"status": "pass", "commands": 80,
                      "reports": 72, "samples": 2160}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
