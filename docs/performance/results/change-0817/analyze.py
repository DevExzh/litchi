"""Fail-closed offline reader for the 0817 ordinary-save packet.

The reader consumes only retained packet receipts and reports.  It never runs
Cargo, the exporter, or a benchmark child.  ``--write`` creates the derived
outputs once; ``--check`` recomputes them and refuses any non-deterministic or
changed primary output.
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import math
import random
import re
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable

import custody


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN_PATH = PACKET / "plan.json"
REPORT_SCHEMA_VERSION = 1
ANALYSIS_SCHEMA = "litchi.performance.0817.ordinary-save-analysis.v1"
ADMISSION_SUMMARY_SCHEMA = "litchi.performance.0817.admission-summary.v1"
ADMISSION_FAILURE_REASON = (
    "independent artifact audit failed before qualification or timing admission; "
    "no timing lane was run"
)
BOOTSTRAP_SEED = 817817
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_CONFIDENCE = 0.95
BOOTSTRAP_LOW_RANK = 250
BOOTSTRAP_HIGH_RANK = 9749
QUALITY_GATES = 6
PRODUCTION_FILES = 9_196
ARCHITECTURE_FILES = 35
OBSERVER_CONTROL_COUNT = 32
HEX = frozenset("0123456789abcdefABCDEF")
ALLOCATION_VECTOR_NAMES = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
PROCESS_VECTOR_NAMES = (
    "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
    "syscr", "syscw", "user_cpu_ticks", "system_cpu_ticks",
    "clock_ticks_per_second", "minor_faults", "major_faults",
    "voluntary_context_switches", "nonvoluntary_context_switches",
    "rss_delta_bytes", "peak_rss_bytes",
)
PROCESS_DELTA_NAMES = (
    "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
    "syscr", "syscw", "minor_faults", "major_faults", "user_cpu_ticks",
    "system_cpu_ticks", "clock_ticks_per_second", "voluntary_context_switches",
    "nonvoluntary_context_switches", "rss_bytes", "peak_rss_bytes",
)
ALLOWLISTED_TEST_SOURCE = "tools/perf-baseline/tests/xlsx_planning_allocations.rs"
TOOL_ALLOWLIST = (ALLOWLISTED_TEST_SOURCE,)
QUALITY_RECOVERY = (
    "Reuse 640 passed/1 ignored unchanged suite results from quality-1; run corrected "
    "integration under all features and allocator-only, plus all-feature doctests; "
    "fresh fmt/check/Clippy/rustdoc/boundary. Exact source and raw log checks required."
)


class ReplayError(RuntimeError):
    """Raised for missing, stale, malformed, or contradictory evidence."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def is_revision(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 40 and all(c in HEX for c in value)


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def path_candidates(raw: Path) -> list[Path]:
    candidates: list[Path] = []
    if raw.is_absolute():
        text = raw.as_posix()
        marker = "/docs/performance/results/change-0817/"
        if marker in text:
            candidates.append(PACKET / text.split(marker, 1)[1])
        candidates.append(raw)
    else:
        text = raw.as_posix()
        prefix = "docs/performance/results/change-0817/"
        if text.startswith(prefix):
            candidates.append(PACKET / text[len(prefix):])
        candidates.extend((PACKET / raw, ROOT / raw))
    result: list[Path] = []
    for candidate in candidates:
        candidate = candidate.resolve(strict=False)
        if candidate not in result:
            result.append(candidate)
    return result


def resolve_path(value: Any, *, packet_bound: bool = False) -> Path:
    require(isinstance(value, str) and value, f"invalid artifact path: {value!r}")
    candidates = path_candidates(Path(value))
    for candidate in candidates:
        if candidate.is_file() and not candidate.is_symlink():
            if packet_bound:
                try:
                    candidate.relative_to(PACKET.resolve())
                except ValueError:
                    continue
            return candidate
    fallback = candidates[0] if candidates else Path(value).resolve(strict=False)
    if packet_bound:
        try:
            fallback.relative_to(PACKET.resolve())
        except ValueError:
            fail(f"artifact path escaped packet: {value}")
    return fallback


def cleanup_value() -> dict[str, Any] | None:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return None
    value = read_json(path)
    require(isinstance(value, dict)
            and set(value) == {
                "schema", "target_removed", "scratch_removed",
                "binaries_verified_before_removal", "removed_binaries", "source",
                "removed", "started", "ended",
            }
            and value.get("schema") == "litchi.performance.0817.cleanup.v1"
            and value.get("target_removed") is True
            and value.get("scratch_removed") is True
            and value.get("binaries_verified_before_removal") is True,
            "cleanup.json removal witness is malformed")
    removed_binaries = value.get("removed_binaries")
    require(isinstance(removed_binaries, list) and len(removed_binaries) == 3
            and all(isinstance(item, dict)
                    and set(item) == {"path", "bytes", "sha256"}
                    and nonnegative_int(item.get("bytes"), "cleanup binary bytes") is None
                    and is_sha(item.get("sha256"))
                    for item in removed_binaries),
            "cleanup.json binary witness is malformed")
    removed = value.get("removed")
    require(isinstance(removed, list) and len(removed) == 2
            and all(isinstance(item, dict)
                    and set(item) == {"path", "files", "logical_bytes"}
                    and isinstance(item.get("path"), str) and item["path"]
                    and nonnegative_int(item.get("files"), "cleanup removed file count") is None
                    and nonnegative_int(item.get("logical_bytes"),
                                       "cleanup removed logical bytes") is None
                    for item in removed),
            "cleanup.json directory witness is malformed")
    require({item["path"] for item in removed} ==
            {str(custody.TARGET), str(custody.SCRATCH)},
            "cleanup.json directory paths changed")
    finite_number(value.get("started"), "cleanup start time")
    finite_number(value.get("ended"), "cleanup end time")
    require(value["started"] <= value["ended"], "cleanup timestamps are reversed")
    return value


def cleanup_contains_binary(cleanup: dict[str, Any] | None,
                            expected: dict[str, Any]) -> bool:
    if not isinstance(cleanup, dict):
        return False
    values = cleanup.get("removed_binaries")
    if not isinstance(values, list):
        return False
    # The cleanup contract is a flat artifact witness.  Build receipts may
    # carry additional nested metadata, so compare only the exact projected
    # path/size/digest descriptor required by cleanup.json.
    projected = {key: expected.get(key) for key in ("path", "bytes", "sha256")}
    return projected in values


def artifact(value: Any, label: str, *, packet_bound: bool = False,
             allow_missing: bool = False) -> Path | None:
    require(isinstance(value, dict), f"{label} is not an artifact descriptor")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    size = value.get("bytes")
    nonnegative_int(size, f"{label}.bytes")
    digest = value.get("sha256")
    require(is_sha(digest), f"{label}.sha256 is invalid")
    path = resolve_path(raw, packet_bound=packet_bound)
    if not path.is_file():
        if allow_missing:
            return None
        fail(f"missing {label}: {raw}")
    require(not path.is_symlink(), f"{label} is a symlink: {raw}")
    require(path.stat().st_size == size, f"{label}.bytes changed")
    require(sha256(path) == digest, f"{label}.sha256 changed")
    return path


def artifact_in_directory(value: Any, directory: Path, label: str,
                          *, allow_missing: bool = False) -> Path | None:
    """Resolve an exporter descriptor relative to its export directory."""
    require(isinstance(value, dict), f"{label} is not an artifact descriptor")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    raw_path = Path(raw)
    candidate = raw_path if raw_path.is_absolute() else directory / raw_path
    candidate = candidate.resolve(strict=False)
    require(candidate.is_relative_to(directory.resolve()),
            f"{label} escaped its export directory")
    projected = {"path": str(candidate), "bytes": value.get("bytes"),
                 "sha256": value.get("sha256")}
    return artifact(projected, label, allow_missing=allow_missing)


def artifact_descriptor_matches(value: Any, expected: dict[str, Any], label: str) -> Path:
    """Check a receipt descriptor against the expected packet artifact.

    Drivers sometimes retain absolute packet paths while the derived reader
    uses packet-relative paths.  Compare the resolved artifact identity and
    the recorded bytes/digest; do not compare a second, normalized path
    spelling.
    """
    path = artifact(value, label, packet_bound=True)
    expected_path = resolve_path(expected.get("path"), packet_bound=True)
    require(path is not None and path.resolve() == expected_path.resolve(),
            f"{label} path changed")
    require({key: value.get(key) for key in ("bytes", "sha256")}
            == {key: expected.get(key) for key in ("bytes", "sha256")},
            f"{label} identity changed")
    return path


def file_identity(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def origin() -> dict[str, Any]:
    value = read_json(PACKET / "origin.json")
    require(isinstance(value, dict)
            and value.get("schema") == "litchi.performance.0817.origin.v1",
            "origin.json is malformed")
    require(is_revision(value.get("base")), "origin base revision is invalid")
    require(value.get("production_changed") is False
            and value.get("tool_changed") is True
            and value.get("runtime_harness_changed") is False,
            "source-change witness changed")
    require(value.get("tool_allowlist") == list(TOOL_ALLOWLIST),
            "tool source allowlist changed")
    require(value.get("unrelated") == custody.UNRELATED,
            "origin unrelated-file witness changed")
    return value


def plan() -> dict[str, Any]:
    value = read_json(PLAN_PATH)
    require(isinstance(value, dict)
            and value.get("schema") == "litchi.performance.0817.plan.v1",
            "plan schema changed")
    base = origin()["base"]
    require(value.get("base") == base, "plan base differs from origin")
    require(value.get("affinity") == list(range(12, 20)), "capture affinity changed")
    cases = custody.plan_cases(value)
    require(value.get("expected") == {
        "artifact_cases": 6,
        "artifact_policy_outputs_per_case": 5,
        "native_reports": 72,
        "native_samples": 2160,
        "observer_reports": 24,
        "observer_samples": 72,
        "qualification_reports": 12,
        "qualification_samples": 12,
        "total_reports": 108,
        "total_samples": 2244,
    }, "plan cardinality contract changed")
    require(value.get("bootstrap") == {
        "confidence": BOOTSTRAP_CONFIDENCE,
        "high_rank": BOOTSTRAP_HIGH_RANK,
        "low_rank": BOOTSTRAP_LOW_RANK,
        "resamples": BOOTSTRAP_RESAMPLES,
        "seed": BOOTSTRAP_SEED,
        "statistic": "median of six process p50 values",
    }, "bootstrap contract changed")
    require(value.get("quantile") == "nearest rank within each process; median over six process blocks",
            "quantile contract changed")
    require(value.get("spread_flag") == "max/min block metric > 1.05"
            and value.get("tail_flag") == "median process p99 / median process p50 > 1.05",
            "diagnostic flag contract changed")
    require(value.get("regression_policy") == "baseline only; no candidate adoption or historical timing comparisons",
            "regression policy changed")
    require(value.get("quality_recovery") == QUALITY_RECOVERY,
            "quality recovery contract changed")
    require(value.get("source") == {
        "harness_changed": True,
        "lock_graph": "existing tracked perf-baseline Cargo.lock; exact differences from workspace lock retained",
        "production_changed": False,
        "runtime_harness_changed": False,
        "test_only_change": ALLOWLISTED_TEST_SOURCE,
    }, "source-change plan changed")
    lanes = value.get("lanes")
    require(lanes == {
        "native": {"binary": "native", "blocks": 6,
                    "orders": ["forward", "reverse", "forward", "reverse", "reverse", "forward"],
                    "samples": 30, "warmup": 3},
        "observer": {"binary": "observer", "blocks": 2,
                      "orders": ["forward", "reverse"], "samples": 3, "warmup": 0},
        "qualification": {"binary": "observer", "blocks": 1,
                           "orders": ["forward"], "samples": 1, "warmup": 0},
    }, "lane contract changed")
    require(value.get("build", {}).get("opt_level") == 3
            and value["build"].get("debug") == 1
            and value["build"].get("lto") == "thin"
            and value["build"].get("codegen_units") == 1
            and value["build"].get("incremental") is False
            and value["build"].get("panic") == "unwind"
            and value["build"].get("jobs") == 2
            and value["build"].get("rustflags") is None,
            "build profile contract changed")
    binaries = value.get("binaries")
    require(binaries == {
        "artifacts": {"cargo_bin": "ordinary_save_artifacts", "features": [], "latency": "untimed"},
        "native": {"cargo_bin": "litchi-perf-baseline", "features": [], "latency": "native"},
        "observer": {"cargo_bin": "litchi-perf-baseline-alloc",
                      "features": ["allocator-metrics", "ordinary-save-process-metrics"],
                      "latency": "diagnostic_only"},
    }, "binary contract changed")
    require(len(cases) == 12, "case cardinality changed")
    return value


def prior_quality_source() -> dict[str, Any]:
    """Load the pre-recovery source witness retained by quality attempt 1."""
    failure_path = PACKET / "quality-1" / "failure.json"
    failure = read_json(failure_path)
    require(isinstance(failure, dict)
            and failure.get("schema") == "litchi.performance.0817.quality-failure.v1",
            "quality-1 failure witness is missing or changed")
    source_path = artifact(failure.get("source"), "quality-1 source", packet_bound=True)
    require(source_path is not None
            and source_path.resolve() == (PACKET / "quality-1" / "source.json").resolve(),
            "quality-1 source witness path changed")
    value = read_json(source_path)
    require(isinstance(value, dict)
            and set(value) == {"production", "tool"},
            "quality-1 source witness is malformed")
    production = value.get("production")
    tool = value.get("tool")
    require(isinstance(production, dict)
            and production.get("revision") == origin()["base"]
            and isinstance(production.get("files"), dict)
            and len(production["files"]) == PRODUCTION_FILES,
            "quality-1 production source witness changed")
    require(isinstance(tool, dict) and tool
            and all(isinstance(name, str) and is_sha(digest)
                    for name, digest in tool.items()),
            "quality-1 tool source witness changed")
    # The quality driver archived the one changed test as a basename beside
    # the relative packet tree; retain that exact recovery layout.
    snapshot = PACKET / "quality-1" / "input-snapshot" / Path(ALLOWLISTED_TEST_SOURCE).name
    require(snapshot.is_file() and not snapshot.is_symlink()
            and sha256(snapshot) == tool[ALLOWLISTED_TEST_SOURCE],
            "quality-1 allowlisted test snapshot changed")
    return {"production": production, "tool": tool,
            "path": rel(source_path), "sha256": sha256(source_path)}


def base_source_matches(source_value: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(source_value, dict), f"{label} source is malformed")
    production = source_value.get("production")
    tool = source_value.get("tool")
    require(isinstance(production, dict) and is_revision(production.get("revision"))
            and production["revision"] == origin()["base"],
            f"{label} production source revision changed")
    require(isinstance(production.get("files"), dict)
            and len(production["files"]) == PRODUCTION_FILES,
            f"{label} production source census changed")
    require(isinstance(tool, dict) and tool, f"{label} tool source is missing")
    require(all(isinstance(name, str) and isinstance(digest, str) and is_sha(digest)
                for name, digest in tool.items()), f"{label} tool source is malformed")
    current_production = custody.source()
    require(current_production.get("files") == production["files"],
            f"{label} production files differ from current base files")
    current_tool = custody.tool_source()
    require(current_tool == tool, f"{label} tracked perf-baseline source changed")
    previous = prior_quality_source()
    previous_production = previous["production"]
    require(previous_production == production,
            f"{label} pre-recovery production witness differs")
    previous_tool = previous["tool"]
    require(set(previous_tool) == set(tool),
            f"{label} pre-recovery tool census changed")
    changed = [name for name in tool if tool[name] != previous_tool[name]]
    require(changed == list(TOOL_ALLOWLIST),
            f"{label} tool changes exceed the declared test allowlist")
    require(tool[ALLOWLISTED_TEST_SOURCE] != previous_tool[ALLOWLISTED_TEST_SOURCE],
            f"{label} allowlisted test source did not change")
    return {"revision": production["revision"], "files": dict(production["files"])}, \
        dict(tool)


def load_packet_custody() -> dict[str, Any]:
    o = origin()
    p = plan()
    root_inputs = custody.assert_root_inputs()
    locks = custody.lock_identity()
    corpus = custody.assert_corpus_inputs()
    provenance = custody.assert_provenance(corpus)
    architecture = custody.architecture_hashes()
    host = custody.assert_host()
    unrelated = custody.assert_unrelated()
    require(len(architecture) == ARCHITECTURE_FILES, "architecture input cardinality changed")
    return {"origin": o, "plan": p, "root_inputs": root_inputs, "locks": locks,
            "corpus": corpus, "provenance": provenance, "architecture": architecture,
            "host_sha256": host, "unrelated": unrelated}


def load_build_driver_recovery(custody_value: dict[str, Any],
                               source_value: dict[str, Any]) -> dict[str, Any]:
    """Acknowledge the pre-Cargo build-driver correction without widening it."""
    path = PACKET / "build-0" / "failure.json"
    value = read_json(path)
    require(isinstance(value, dict)
            and set(value) == {
                "cargo_started", "corrected_script", "exit_code", "failed_attempt",
                "failure", "original_script", "quality_sha256", "recorded_at", "schema",
                "scope",
            }, "build-driver recovery witness changed")
    require(value.get("schema") == "litchi.performance.0817.build-driver-recovery.v1"
            and value.get("cargo_started") is False
            and value.get("exit_code") == 1
            and value.get("failed_attempt") == 0
            and value.get("failure") ==
            "AssertionError at build.py:118: reader omitted plan binaries[*].cargo_bin"
            and value.get("scope") ==
            "Use exact existing cargo_bin plan key. Plan, Rust source, commands, features, profiles, dependencies and quality results unchanged.",
            "build-driver recovery boundary changed")
    finite_number(value.get("recorded_at"), "build-driver recovery timestamp")
    original = value.get("original_script")
    corrected = value.get("corrected_script")
    require(isinstance(original, dict) and original.get("path") == "build-0/build.py"
            and is_sha(original.get("sha256"))
            and isinstance(corrected, dict) and corrected.get("path") == "build.py"
            and is_sha(corrected.get("sha256")),
            "build-driver script recovery descriptors changed")
    original_path = PACKET / "build-0" / "build.py"
    require(original_path.is_file() and not original_path.is_symlink()
            and sha256(original_path) == original["sha256"],
            "original build driver archive changed")
    require(corrected["sha256"] == sha256(PACKET / "build.py"),
            "corrected build driver digest changed")
    require(value.get("quality_sha256") == sha256(PACKET / "quality.json"),
            "build-driver recovery quality witness changed")

    old_source_path = PACKET / "build-0" / "source.json"
    old_frozen_path = PACKET / "build-0" / "frozen-inputs.json"
    old_source = read_json(old_source_path)
    require(old_source == source_value,
            "pre-Cargo build source differs from final source")
    old_frozen = read_json(old_frozen_path)
    require(isinstance(old_frozen, dict)
            and old_frozen.get("schema") == "litchi.performance.0817.frozen-inputs.v1"
            and old_frozen.get("root_inputs") == custody_value["root_inputs"]
            and old_frozen.get("locks") == custody_value["locks"]
            and old_frozen.get("architecture") == custody_value["architecture"]
            and old_frozen.get("corpus") == custody_value["corpus"]
            and old_frozen.get("provenance") == custody_value["provenance"]
            and old_frozen.get("host") == custody_value["host_sha256"]
            and old_frozen.get("unrelated") == custody_value["unrelated"],
            "pre-Cargo build frozen inputs changed")
    current_packet = custody.packet_hashes()
    old_packet = old_frozen.get("packet")
    require(isinstance(old_packet, dict) and set(old_packet) == set(current_packet),
            "pre-Cargo build packet census changed")
    packet_changes = [name for name in current_packet if current_packet[name] != old_packet[name]]
    require(packet_changes == ["build.py"],
            "build-driver recovery changed more than its driver")
    current_drivers = custody.driver_hashes()
    old_drivers = old_frozen.get("drivers")
    require(isinstance(old_drivers, dict) and set(old_drivers) == set(current_drivers),
            "pre-Cargo build driver census changed")
    driver_changes = [name for name in current_drivers if current_drivers[name] != old_drivers[name]]
    require(driver_changes == ["build.py"],
            "build-driver recovery changed another driver")
    return {
        "schema": value["schema"], "failed_attempt": value["failed_attempt"],
        "cargo_started": value["cargo_started"], "exit_code": value["exit_code"],
        "original_script": {"path": rel(original_path), "sha256": original["sha256"]},
        "corrected_script": {"path": rel(PACKET / "build.py"),
                              "sha256": corrected["sha256"]},
        "quality_sha256": value["quality_sha256"],
        "packet_changes": packet_changes, "driver_changes": driver_changes,
    }


def load_build(custody_value: dict[str, Any], p: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "build.json"
    build = read_json(path)
    require(isinstance(build, dict)
            and build.get("schema") == "litchi.performance.0817.build.v1",
            "build schema changed")
    source_desc = build.get("source")
    source_path = artifact(source_desc, "build source", packet_bound=True)
    assert source_path is not None
    positive_int(build.get("attempt"), "build attempt")
    build_directory = PACKET / f"build-{build['attempt']}"
    require(build_directory.is_dir() and not build_directory.is_symlink(),
            "build attempt directory is missing")
    require(source_path.resolve() == (build_directory / "source.json").resolve(),
            "build source witness path changed")
    source_value = read_json(source_path)
    production, tool = base_source_matches(source_value, "build")
    driver_recovery = load_build_driver_recovery(custody_value, source_value)
    require(build.get("root_inputs") == custody_value["root_inputs"]
            and build.get("locks") == custody_value["locks"]
            and build.get("architecture") == custody_value["architecture"]
            and build.get("corpus") == custody_value["corpus"]
            and build.get("provenance") == custody_value["provenance"]
            and build.get("host") == custody_value["host_sha256"]
            and build.get("unrelated") == custody_value["unrelated"],
            "build custody witness changed")
    require(build.get("packet") == custody.packet_hashes(), "build packet witness changed")
    require(build.get("drivers") == custody.driver_hashes(), "build driver witness changed")
    frozen_desc = build.get("frozen_inputs")
    frozen_path = artifact(frozen_desc, "build frozen inputs", packet_bound=True)
    assert frozen_path is not None
    require(frozen_path.resolve() == (build_directory / "frozen-inputs.json").resolve(),
            "build frozen-input witness path changed")
    frozen = read_json(frozen_path)
    require(isinstance(frozen, dict)
            and frozen.get("schema") == "litchi.performance.0817.frozen-inputs.v1",
            "build frozen-inputs schema changed")
    require(frozen.get("packet") == build["packet"]
            and frozen.get("drivers") == build["drivers"]
            and frozen.get("root_inputs") == build["root_inputs"]
            and frozen.get("locks") == build["locks"]
            and frozen.get("architecture") == build["architecture"]
            and frozen.get("corpus") == build["corpus"]
            and frozen.get("provenance") == build["provenance"]
            and frozen.get("host") == build["host"]
            and frozen.get("unrelated") == build["unrelated"],
            "build frozen-input witness changed")
    rows = build.get("rows")
    require(isinstance(rows, list) and len(rows) == 3, "build row cardinality changed")
    expected_rows = {
        "native": ("litchi-perf-baseline", []),
        "artifacts": ("ordinary_save_artifacts", []),
        "observer": ("litchi-perf-baseline-alloc",
                      ["allocator-metrics", "ordinary-save-process-metrics"]),
    }
    seen: set[str] = set()
    for row in rows:
        require(isinstance(row, dict), "build row is malformed")
        name = row.get("binary")
        require(name in expected_rows and name not in seen, "build binary row identity changed")
        seen.add(name)
        cargo_name, features = expected_rows[name]
        require(row.get("cargo_binary") == cargo_name and row.get("features") == features
                and row.get("exit_code") == 0,
                f"build row changed: {name}")
        require(isinstance(row.get("command"), list), f"build command missing: {name}")
        command = row["command"]
        expected_command = ["cargo", "build", "--offline", "--locked", "--release",
                            "--manifest-path", str(custody.TOOL / "Cargo.toml"),
                            "--bin", cargo_name]
        if features:
            expected_command.extend(["--features", ",".join(features)])
        require(command == expected_command, f"build command changed: {name}")
        finite_number(row.get("started"), f"build {name} start")
        finite_number(row.get("ended"), f"build {name} end")
        require(row["started"] <= row["ended"], f"build {name} time order changed")
        require(row.get("environment") == {
                    "CARGO_TARGET_DIR": str(custody.TARGET),
                    "CARGO_BUILD_JOBS": str(p["build"]["jobs"]),
                    "CARGO_INCREMENTAL": "0",
                    "CARGO_PROFILE_RELEASE_OPT_LEVEL": str(p["build"]["opt_level"]),
                    "CARGO_PROFILE_RELEASE_DEBUG": str(p["build"]["debug"]),
                    "CARGO_PROFILE_RELEASE_LTO": str(p["build"]["lto"]),
                    "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": str(p["build"]["codegen_units"]),
                    "CARGO_PROFILE_RELEASE_INCREMENTAL": str(p["build"]["incremental"]).lower(),
                    "CARGO_PROFILE_RELEASE_PANIC": str(p["build"]["panic"]),
                }, f"build environment changed: {name}")
        log = artifact(row.get("log"), f"build {name} log", packet_bound=True)
        require(log is not None, f"build log missing: {name}")
        require(log.resolve() == (build_directory / f"{name}.log").resolve(),
                f"build log path changed: {name}")
    require(seen == set(expected_rows), "build binary rows incomplete")
    commands_path = build_directory / "commands.json"
    require(commands_path.is_file() and not commands_path.is_symlink()
            and read_json(commands_path) == rows,
            "build command receipt changed")
    require(build.get("profile") == {key: p["build"][key]
                                      for key in ("opt_level", "debug", "lto", "codegen_units",
                                                  "incremental", "panic")},
            "build profile witness changed")
    require(build.get("environment") == {
                "CARGO_TARGET_DIR": str(custody.TARGET),
                "CARGO_BUILD_JOBS": str(p["build"]["jobs"]),
                "CARGO_INCREMENTAL": "0",
                "CARGO_PROFILE_RELEASE_OPT_LEVEL": str(p["build"]["opt_level"]),
                "CARGO_PROFILE_RELEASE_DEBUG": str(p["build"]["debug"]),
                "CARGO_PROFILE_RELEASE_LTO": str(p["build"]["lto"]),
                "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": str(p["build"]["codegen_units"]),
                "CARGO_PROFILE_RELEASE_INCREMENTAL": str(p["build"]["incremental"]).lower(),
                "CARGO_PROFILE_RELEASE_PANIC": str(p["build"]["panic"]),
            }, "build environment witness changed")
    binaries = build.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == set(expected_rows),
            "build binary map changed")
    cleanup = cleanup_value()
    checked_binaries: dict[str, Any] = {}
    for name in expected_rows:
        entry = binaries[name]
        require(isinstance(entry, dict) and entry.get("cargo_name") == expected_rows[name][0]
                and entry.get("features") == expected_rows[name][1],
                f"build binary metadata changed: {name}")
        descriptor = entry.get("artifact")
        require(isinstance(descriptor, dict), f"build binary artifact missing: {name}")
        path_value = artifact(descriptor, f"build {name} binary", allow_missing=True)
        if path_value is None:
            require(cleanup_contains_binary(cleanup,
                                            {key: descriptor.get(key)
                                             for key in ("path", "bytes", "sha256")}),
                    f"missing {name} binary lacks exact cleanup witness")
        checked_binaries[name] = {"cargo_name": entry["cargo_name"],
                                  "features": list(entry["features"]),
                                  "artifact": {key: descriptor[key]
                                                for key in ("path", "bytes", "sha256")}}
    return {
        "receipt": file_identity(path),
        "source": {"path": rel(source_path), "production": production, "tool": tool},
        "frozen_inputs": {"path": rel(frozen_path), "sha256": sha256(frozen_path)},
        "binaries": checked_binaries,
        "rows": len(rows),
        "driver_recovery": driver_recovery,
        "profile": dict(build["profile"]),
        "provenance": build["provenance"],
        "root_inputs": build["root_inputs"], "locks": build["locks"],
        "architecture": build["architecture"], "corpus": build["corpus"],
        "host": build["host"], "packet": build["packet"], "drivers": build["drivers"],
        "unrelated": build["unrelated"],
    }


def quality_suite_records(text: str, label: str) -> list[dict[str, Any]]:
    """Parse Cargo's retained suite summaries without treating the log as proof alone."""
    summaries = re.findall(
        r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; "
        r"(\d+) ignored; (\d+) measured; (\d+) filtered out", text,
    )
    require(len(summaries) == 27, f"{label} suite-summary cardinality changed")
    require(sum(int(row[1]) for row in summaries) == 640
            and sum(int(row[2]) for row in summaries) == 1
            and sum(int(row[3]) for row in summaries) == 1
            and all(row[0] == "ok" and row[2] == "0" for row in summaries[:-1])
            and summaries[-1] == ("FAILED", "0", "1", "0", "0", "0"),
            f"{label} suite-summary totals changed")
    records: list[dict[str, Any]] = []
    active_suite: str | None = None
    for line in text.splitlines():
        match = re.match(r"\s*Running (.+) \((.+)\)$", line)
        if match:
            active_suite = match.group(1)
        match = re.match(
            r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; "
            r"(\d+) ignored; (\d+) measured; (\d+) filtered out", line,
        )
        if match:
            require(active_suite is not None, f"{label} suite name is missing")
            records.append({
                "suite": active_suite, "status": match.group(1),
                "passed": int(match.group(2)), "failed": int(match.group(3)),
                "ignored": int(match.group(4)),
            })
    require(len(records) == 27
            and records[-1]["suite"] == "tests/xlsx_planning_allocations.rs"
            and all(row["status"] == "ok" for row in records[:-1]),
            f"{label} suite records changed")
    return records


def load_quality_recovery(directory: Path, quality: dict[str, Any],
                          build: dict[str, Any],
                          custody_value: dict[str, Any]) -> dict[str, Any]:
    recovery_path = directory / "test-recovery.json"
    recovery = read_json(recovery_path)
    require(isinstance(recovery, dict)
            and set(recovery) == {
                "schema", "inherited_passed", "inherited_ignored",
                "inherited_successful_suites", "carried_suites",
                "old_plan_sha256", "amended_plan_sha256", "prior_log",
                "prior_source", "prior_frozen_inputs", "changed_test",
                "old_test_sha256", "new_test_sha256",
                "runtime_and_production_unchanged", "rows",
            }, "quality recovery receipt changed")
    require(recovery.get("schema") == "litchi.performance.0817.test-recovery.v1"
            and recovery.get("inherited_passed") == 640
            and recovery.get("inherited_ignored") == 1
            and recovery.get("inherited_successful_suites") == 26
            and recovery.get("changed_test") == ALLOWLISTED_TEST_SOURCE
            and recovery.get("runtime_and_production_unchanged") is True,
            "quality recovery summary changed")
    previous = prior_quality_source()
    old_tool = previous["tool"]
    current_tool = build["source"]["tool"]
    require(recovery.get("old_test_sha256") == old_tool[ALLOWLISTED_TEST_SOURCE]
            and recovery.get("new_test_sha256") == current_tool[ALLOWLISTED_TEST_SOURCE],
            "quality recovery source pair changed")
    require(recovery.get("old_plan_sha256") ==
            sha256(PACKET / "quality-1" / "input-snapshot" / "plan.json")
            and recovery.get("amended_plan_sha256") == sha256(PLAN_PATH)
            and recovery["old_plan_sha256"] != recovery["amended_plan_sha256"],
            "quality recovery plan relation changed")

    old_failure = read_json(PACKET / "quality-1" / "failure.json")
    require(isinstance(old_failure, dict)
            and old_failure.get("schema") == "litchi.performance.0817.quality-failure.v1"
            and old_failure.get("attempt") == 1
            and old_failure.get("failed_gate") == 3,
            "quality-1 failure boundary changed")
    failed_row = old_failure.get("row")
    require(isinstance(failed_row, dict), "quality-1 failed row is missing")
    old_checks = read_json(PACKET / "quality-1" / "checks.json")
    require(isinstance(old_checks, list) and len(old_checks) == 3
            and old_checks[2].get("exit_code") != 0,
            "quality-1 failed test gate changed")
    old_log = artifact(failed_row.get("log"),
                       "quality-1 failed test log", packet_bound=True)
    require(old_log is not None
            and old_log.resolve() == (PACKET / "quality-1" / "02.log").resolve()
            and old_checks[2].get("log") == failed_row.get("log"),
            "quality-1 failed log witness changed")
    records = quality_suite_records(old_log.read_text(encoding="utf-8"), "quality-1")
    require("to rerun pass `--test xlsx_planning_allocations`" in
            old_log.read_text(encoding="utf-8")
            and 'left: String("ordinary_save_procfs_operation_scoped")' in
            old_log.read_text(encoding="utf-8")
            and 'right: "none"' in old_log.read_text(encoding="utf-8"),
            "quality-1 failure text changed")
    old_frozen = PACKET / "quality-1" / "frozen-inputs.json"
    old_frozen_value = read_json(old_frozen)
    require(isinstance(old_frozen_value, dict)
            and old_frozen_value.get("schema") ==
            "litchi.performance.0817.frozen-inputs.v1"
            and old_frozen_value.get("root_inputs") == custody_value["root_inputs"]
            and old_frozen_value.get("locks") == custody_value["locks"]
            and old_frozen_value.get("architecture") == custody_value["architecture"]
            and old_frozen_value.get("corpus") == custody_value["corpus"]
            and old_frozen_value.get("host") == custody_value["host_sha256"]
            and old_frozen_value.get("unrelated") == custody_value["unrelated"],
            "quality-1 frozen-input witness changed")
    old_packet = old_frozen_value.get("packet")
    require(isinstance(old_packet, dict), "quality-1 packet witness is missing")
    for name, digest in old_packet.items():
        snapshot = PACKET / "quality-1" / "input-snapshot" / name
        require(snapshot.is_file() and not snapshot.is_symlink()
                and sha256(snapshot) == digest,
                f"quality-1 frozen packet snapshot changed: {name}")
    artifact_descriptor_matches(recovery.get("prior_log"), file_identity(old_log),
                                "quality recovery prior log")
    old_source_path = PACKET / "quality-1" / "source.json"
    artifact_descriptor_matches(recovery.get("prior_source"),
                                file_identity(old_source_path),
                                "quality recovery prior source")
    artifact_descriptor_matches(recovery.get("prior_frozen_inputs"),
                                file_identity(old_frozen),
                                "quality recovery prior frozen inputs")

    commands_path = directory / "test-recovery-commands.json"
    commands = read_json(commands_path)
    rows = recovery.get("rows")
    require(isinstance(rows, list) and len(rows) == 3 and commands == rows,
            "quality recovery command receipt changed")
    manifest = str(custody.TOOL / "Cargo.toml")
    expected_commands = [
        ["cargo", "test", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--test", "xlsx_planning_allocations", "--",
         "--test-threads=2"],
        ["cargo", "test", "--offline", "--locked", "--manifest-path", manifest,
         "--features", "allocator-metrics", "--test", "xlsx_planning_allocations",
         "--", "--test-threads=2"],
        ["cargo", "test", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--doc", "--", "--test-threads=2"],
    ]
    expected_environment = {
        "CARGO_TARGET_DIR": str(custody.TARGET / "quality-1"),
        "CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_DEV_DEBUG": "0", "RUSTDOCFLAGS": "-D warnings",
    }
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("command") == expected_commands[index]
                and row.get("exit_code") == 0
                and row.get("environment") == expected_environment,
                f"quality recovery command {index + 1} changed")
        started, ended = row.get("started"), row.get("ended")
        finite_number(started, f"quality recovery {index + 1} start")
        finite_number(ended, f"quality recovery {index + 1} end")
        require(started <= ended, f"quality recovery {index + 1} time order changed")
        log = artifact(row.get("log"), f"quality recovery {index + 1} log",
                       packet_bound=True)
        require(log is not None
                and log.resolve() == (directory / f"test-recovery-{index}.log").resolve(),
                f"quality recovery {index + 1} log identity changed")
    require(recovery.get("carried_suites") == records[:-1],
            "quality recovery carried-suite manifest changed")
    return {"path": rel(recovery_path), "commands": rel(commands_path),
            "inherited_passed": recovery["inherited_passed"],
            "inherited_ignored": recovery["inherited_ignored"],
            "inherited_successful_suites": recovery["inherited_successful_suites"],
            "carried_suites": recovery["carried_suites"],
            "old_plan_sha256": recovery["old_plan_sha256"],
            "amended_plan_sha256": recovery["amended_plan_sha256"],
            "changed_test": recovery["changed_test"],
            "old_test_sha256": recovery["old_test_sha256"],
            "new_test_sha256": recovery["new_test_sha256"],
            "rows": recovery["rows"]}


def load_quality(custody_value: dict[str, Any], build: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "quality.json"
    value = read_json(path)
    require(isinstance(value, dict)
            and value.get("schema") == "litchi.performance.0817.quality.v1"
            and value.get("status") == "pass" and value.get("gate_count") == QUALITY_GATES,
            "quality receipt changed")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == QUALITY_GATES,
            "quality gate cardinality changed")
    source_desc = value.get("source")
    source_path = artifact(source_desc, "quality source", packet_bound=True)
    assert source_path is not None
    source_value = read_json(source_path)
    require(source_value.get("production") == build["source"]["production"]
            and source_value.get("tool") == build["source"]["tool"],
            "quality source differs from build source")
    manifest = str(custody.TOOL / "Cargo.toml")
    attempt = value.get("attempt")
    positive_int(attempt, "quality attempt")
    quality_directory = PACKET / f"quality-{attempt}"
    require(quality_directory.is_dir() and not quality_directory.is_symlink(),
            "final quality attempt directory is missing")
    require(source_path.resolve() == (quality_directory / "source.json").resolve(),
            "quality source witness path changed")
    expected_prefixes = [
        ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        ["cargo", "check", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--all-targets"],
        ["python3", "-B", str(PACKET / "quality_tests.py"),
         str(PACKET / f"quality-{attempt}")],
        ["cargo", "clippy", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--no-deps"],
        ["python3", "-B", "tools/check_crate_boundaries.py"],
    ]
    logs: list[str] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("gate") == index + 1
                and row.get("exit_code") == 0,
                f"quality gate {index + 1} failed")
        require(row.get("command") == expected_prefixes[index],
                f"quality gate {index + 1} command changed")
        finite_number(row.get("started"), f"quality gate {index + 1} start")
        finite_number(row.get("ended"), f"quality gate {index + 1} end")
        require(row["started"] <= row["ended"],
                f"quality gate {index + 1} time order changed")
        log = artifact(row.get("log"), f"quality gate {index + 1} log", packet_bound=True)
        assert log is not None
        require(log.resolve() == (quality_directory / f"{index:02}.log").resolve(),
                f"quality gate {index + 1} log path changed")
        logs.append(rel(log))
    require(value.get("root_inputs") == custody_value["root_inputs"]
            and value.get("locks") == custody_value["locks"]
            and value.get("architecture") == custody_value["architecture"]
            and value.get("corpus") == custody_value["corpus"]
            and value.get("host") == custody_value["host_sha256"]
            and value.get("unrelated") == custody_value["unrelated"],
            "quality custody witness changed")
    historical_packet = dict(custody.packet_hashes())
    historical_drivers = dict(custody.driver_hashes())
    driver_recovery = build.get("driver_recovery")
    require(isinstance(driver_recovery, dict)
            and isinstance(driver_recovery.get("original_script"), dict)
            and is_sha(driver_recovery["original_script"].get("sha256")),
            "build-driver recovery is missing for quality provenance")
    old_build_sha = driver_recovery["original_script"]["sha256"]
    historical_packet["build.py"] = old_build_sha
    historical_drivers["build.py"] = old_build_sha
    require(value.get("packet") == historical_packet
            and value.get("drivers") == historical_drivers,
            "quality packet/driver witness changed")
    build0_frozen = read_json(PACKET / "build-0" / "frozen-inputs.json")
    require(value["packet"] == build0_frozen.get("packet")
            and value["drivers"] == build0_frozen.get("drivers"),
            "quality and pre-Cargo build frozen driver witnesses differ")
    positive_int(value.get("attempt"), "quality attempt")
    require(value["attempt"] >= 2,
            "quality receipt does not acknowledge the failed recovery attempt")
    require(value.get("provenance") == custody_value["provenance"],
            "quality provenance witness changed")
    require(value.get("environment") == {
                "CARGO_TARGET_DIR": str(custody.TARGET / "quality-1"),
                "CARGO_BUILD_JOBS": str(2),
                "CARGO_INCREMENTAL": "0",
                "CARGO_PROFILE_DEV_DEBUG": "0",
                "RUSTDOCFLAGS": "-D warnings",
            }, "quality environment witness changed")
    checks = artifact(value.get("checks"), "quality checks", packet_bound=True)
    require(checks is not None, "quality checks missing")
    require(checks.resolve() == (quality_directory / "checks.json").resolve(),
            "quality checks path changed")
    require(read_json(checks) == rows, "quality command receipt changed")
    require(resolve_path(source_desc.get("path"), packet_bound=True).is_file(),
            "quality source is missing")
    frozen_path = quality_directory / "frozen-inputs.json"
    frozen = read_json(frozen_path)
    require(isinstance(frozen, dict)
            and frozen.get("schema") == "litchi.performance.0817.frozen-inputs.v1"
            and frozen.get("packet") == value["packet"]
            and frozen.get("drivers") == value["drivers"]
            and frozen.get("root_inputs") == value["root_inputs"]
            and frozen.get("locks") == value["locks"]
            and frozen.get("architecture") == value["architecture"]
            and frozen.get("corpus") == value["corpus"]
            and frozen.get("host") == value["host"]
            and frozen.get("unrelated") == value["unrelated"],
            "quality frozen-input witness changed")
    recovery = load_quality_recovery(quality_directory, value, build, custody_value)
    return {"receipt": file_identity(path), "gates": len(rows), "logs": logs,
            "source": rel(source_path), "checks": rel(checks),
            "environment": value.get("environment"), "attempt": value.get("attempt"),
            "provenance": value["provenance"],
            "packet": value["packet"], "drivers": value["drivers"],
            "frozen_inputs": file_identity(frozen_path), "recovery": recovery,
            "packet_driver_recovery": {
                "historical_build_sha256": old_build_sha,
                "current_build_sha256": custody.driver_hashes()["build.py"],
            }}


def load_artifacts(custody_value: dict[str, Any], p: dict[str, Any], build: dict[str, Any],
                   quality: dict[str, Any]) -> dict[str, Any]:
    complete_path = PACKET / "artifacts.complete.json"
    complete = read_json(complete_path)
    require(isinstance(complete, dict)
            and complete.get("schema") == "litchi.performance.0817.artifacts.complete.v1"
            and complete.get("children") == 1 and complete.get("reports") == 0
            and complete.get("samples") == 0 and complete.get("cases") == 6,
            "artifact completion receipt changed")
    require(complete.get("plan_sha256") == sha256(PLAN_PATH)
            and complete.get("build_sha256") == sha256(PACKET / "build.json")
            and complete.get("quality_sha256") == sha256(PACKET / "quality.json"),
            "artifact completion provenance changed")
    receipt_path = artifact(complete.get("receipt"), "artifact completion receipt", packet_bound=True)
    manifest_path = artifact(complete.get("manifest"), "artifact manifest", packet_bound=True)
    source_path = artifact(complete.get("source"), "artifact completion source", packet_bound=True)
    require(receipt_path is not None and manifest_path is not None and source_path is not None,
            "artifact completion paths missing")
    require(receipt_path.resolve() == (PACKET / "artifacts-receipt.json").resolve()
            and manifest_path.resolve() == (PACKET / "artifacts" / "manifest.json").resolve(),
            "artifact completion receipt path changed")
    require(source_path.resolve() == (PACKET / "artifacts-source.json").resolve()
            and read_json(source_path) == {
                "production": build["source"]["production"],
                "tool": build["source"]["tool"],
            }, "artifact completion source identity changed")
    receipt = read_json(PACKET / "artifacts-receipt.json")
    require(isinstance(receipt, dict)
            and receipt.get("schema") == "litchi.performance.0817.artifact-receipt.v1"
            and receipt.get("lane") == "artifacts" and receipt.get("exit_code") == 0,
            "artifact export receipt changed")
    source_receipt = artifact(receipt.get("source"), "artifact export source", packet_bound=True)
    require(source_receipt is not None
            and source_receipt.resolve() == (PACKET / "artifacts-source.json").resolve(),
            "artifact export source identity changed")
    require(read_json(source_receipt) == {
                "production": build["source"]["production"],
                "tool": build["source"]["tool"],
            }, "artifact export source contents changed")
    require(receipt.get("root_inputs") == custody_value["root_inputs"]
            and receipt.get("locks") == custody_value["locks"]
            and receipt.get("architecture") == custody_value["architecture"]
            and receipt.get("corpus") == custody_value["corpus"]
            and receipt.get("host") == custody_value["host_sha256"]
            and receipt.get("packet") == custody.packet_hashes()
            and receipt.get("drivers") == custody.driver_hashes()
            and receipt.get("unrelated") == custody_value["unrelated"],
            "artifact export custody changed")
    finite_number(receipt.get("started"), "artifact export start")
    finite_number(receipt.get("ended"), "artifact export end")
    require(receipt["started"] <= receipt["ended"],
            "artifact export time order changed")
    inputs = receipt.get("inputs")
    expected_inputs = [
        {"path": name, "absolute": str(ROOT / name), **custody_value["corpus"][name]}
        for name in (
            "test-data/ooxml/docx/documentProperties.docx",
            "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx",
            "test-data/ooxml/pptx/shapes.pptx",
        )
    ]
    require(inputs == expected_inputs, "artifact export input inventory changed")
    artifact_binary = build["binaries"]["artifacts"]["artifact"]
    require(receipt.get("binary") == build["binaries"]["artifacts"],
            "artifact exporter binary identity changed")
    artifact(artifact_binary, "artifact exporter binary", allow_missing=True)
    expected_artifact_command = [
        "/usr/bin/time", "-f", "%M", "-o", str(PACKET / "artifacts.rss"),
        "taskset", "-c", ",".join(str(cpu) for cpu in p["affinity"]),
        artifact_binary["path"], "--output", str(PACKET / "artifacts"),
        "--filesystem-root", str(custody.SCRATCH),
    ]
    for item in (
        "test-data/ooxml/docx/documentProperties.docx",
        "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx",
        "test-data/ooxml/pptx/shapes.pptx",
    ):
        expected_artifact_command.extend(["--ooxml-file", str(ROOT / item)])
    require(receipt.get("command") == expected_artifact_command,
            "artifact export command changed")
    scratch_marker = receipt.get("scratch_marker")
    require(isinstance(scratch_marker, dict)
            and resolve_path(scratch_marker.get("path"), packet_bound=False).resolve() ==
            (custody.SCRATCH / ".litchi-performance-0817-owned").resolve(),
            "artifact scratch marker identity changed")
    artifact(scratch_marker, "artifact scratch marker", allow_missing=True)
    export_log = artifact(receipt.get("log"), "artifact export log", packet_bound=True)
    export_rss = artifact(receipt.get("rss"), "artifact export RSS", packet_bound=True)
    require(export_log is not None and export_log.resolve() ==
            (PACKET / "artifacts.log").resolve()
            and export_rss is not None and export_rss.resolve() ==
            (PACKET / "artifacts.rss").resolve(),
            "artifact export log/RSS identity changed")
    parse_rss(receipt, "artifact export")
    manifest = read_json(manifest_path)
    require(isinstance(manifest, dict) and manifest.get("schema_version") == 1
            and manifest.get("kind") == "ordinary-save-artifact-export"
            and manifest.get("generator") == "litchi-perf-ordinary-save-artifacts-v1",
            "artifact manifest identity changed")
    artifact_descriptor_matches(receipt.get("manifest"), file_identity(manifest_path),
                                "artifact receipt manifest")
    cases = manifest.get("cases")
    require(isinstance(cases, list) and len(cases) == 6, "artifact case cardinality changed")
    require(all(isinstance(case, dict) for case in cases),
            "artifact case is malformed")
    identity = {(case.get("origin"), case.get("format")) for case in cases}
    require(identity == {(origin_name, fmt) for origin_name in
                         ("generated-harness-corpus", "caller-named-real-file")
                         for fmt in ("DOCX", "XLSX", "PPTX")},
            "artifact corpus identity set changed")
    require({case.get("case_id") for case in cases} == {
        "generated-docx-medium", "generated-xlsx-medium", "generated-pptx-medium",
        "real-000-docx", "real-001-xlsx", "real-002-pptx",
    }, "artifact case identifiers changed")
    for case in cases:
        require(isinstance(case, dict), "artifact case is malformed")
        require(isinstance(case.get("case_id"), str) and case.get("case_id")
                and case.get("format") in {"DOCX", "XLSX", "PPTX"}
                and case.get("origin") in {"generated-harness-corpus", "caller-named-real-file"},
                "artifact case identity changed")
        require(case.get("edit_admitted") is True
                and case.get("edit_outcome") == "admitted",
                "artifact edit outcome changed before failed admission")
        positive_int(case.get("source_archive_bytes"), "artifact source bytes")
        require(is_sha(case.get("source_archive_sha256")),
                "artifact source digest changed")
        positive_int(case.get("published_bytes"), "artifact published bytes")
        require(is_sha(case.get("published_sha256")),
                "artifact published digest changed")
        source_descriptor = case.get("source_archive")
        source_path = artifact_in_directory(source_descriptor, manifest_path.parent,
                                            "artifact source archive")
        require(source_path is not None
                and source_path.stat().st_size == case["source_archive_bytes"]
                and sha256(source_path) == case["source_archive_sha256"],
                "artifact source archive identity changed")
        if case["origin"] == "caller-named-real-file":
            require(case.get("input_path") in
                    {str(ROOT / name) for name in custody_value["corpus"]},
                    "real artifact input path changed")
        else:
            require(case.get("input_path") is None,
                    "generated artifact unexpectedly names a real input")
        policies = case.get("policy_outputs")
        require(isinstance(policies, list) and len(policies) == 4,
                f"artifact policy count changed: {case.get('case_id')}")
        require(isinstance(case.get("stream_output"), dict), "artifact stream output missing")
        for output in [*policies, case["stream_output"]]:
            require(isinstance(output, dict) and output.get("matches_reference") is True
                    and output.get("source_unchanged") is True
                    and output.get("reopen_admitted") is True
                    and isinstance(output.get("matches_source"), bool),
                    "artifact exporter gate changed")
            require(output.get("policy") in {"default", "full", "file-only", "no-sync", "stream"}
                    and output.get("publication") in
                    {"save", "save_with_durability", "sequential_sink"},
                    "artifact policy identity changed")
            expected_publication = {
                "default": "save", "full": "save_with_durability",
                "file-only": "save_with_durability", "no-sync": "save_with_durability",
                "stream": "sequential_sink",
            }[output["policy"]]
            require(output["publication"] == expected_publication,
                    "artifact publication entry point changed")
            descriptor = output.get("output")
            path = artifact_in_directory(descriptor, manifest_path.parent,
                                         "artifact policy output")
            require(path is not None, "artifact policy output missing")
            require(output.get("sha256") == sha256(path)
                    and output.get("bytes") == path.stat().st_size,
                    "artifact policy digest changed")
            require(output.get("bytes") == case["published_bytes"]
                    and output.get("sha256") == case["published_sha256"],
                    "artifact policy differs from published reference")
    # The audit is the admission gate.  This run failed, so the retained
    # report lives under admission-0 and is deliberately consumed as failed
    # evidence.  Do not synthesize a root artifact-audit.json or relabel the
    # deterministic exports as admitted.
    audit_path = PACKET / "admission-0" / "artifact-audit.json"
    admission_receipt_path = PACKET / "admission-0" / "receipt.json"
    audit_receipt = read_json(admission_receipt_path)
    require(isinstance(audit_receipt, dict)
            and audit_receipt.get("schema") ==
            "litchi.performance.0817.admission-attempt.v1"
            and audit_receipt.get("exit_code") == 1,
            "failed artifact admission receipt changed")
    artifact_descriptor_matches(audit_receipt.get("auditor"),
                                file_identity(PACKET / "artifact_audit.py"),
                                "artifact admission auditor")
    artifact_descriptor_matches(audit_receipt.get("manifest"),
                                file_identity(manifest_path),
                                "artifact admission manifest")
    admission_log = artifact(audit_receipt.get("log"), "artifact admission log",
                             packet_bound=True)
    require(admission_log is not None
            and admission_log.resolve() ==
            (PACKET / "admission-0" / "audit.log").resolve(),
            "artifact admission log path changed")
    artifact_descriptor_matches(audit_receipt.get("report"),
                                file_identity(audit_path),
                                "artifact admission report")
    audit = read_json(audit_path)
    require(isinstance(audit, dict)
            and audit.get("schema") == "litchi.performance.0817.artifact-audit.v1"
            and audit.get("ok") is False,
            "independent artifact audit failure witness changed")
    audit_cases = audit.get("cases")
    require(isinstance(audit_cases, list) and len(audit_cases) == 6,
            "artifact audit case cardinality changed")
    require(isinstance(audit.get("errors"), list) and audit.get("errors"),
            "artifact audit failure errors are missing")
    require(audit.get("errors") == [error for case in audit_cases
                                     for error in case.get("errors", [])],
            "artifact audit case errors do not replay the top-level errors")
    require(all(isinstance(case, dict)
                and isinstance(case.get("ok"), bool)
                and isinstance(case.get("errors"), list)
                for case in audit_cases),
            "artifact audit failure case is malformed")
    audit_by_identity = {(case.get("origin"), case.get("format")): case for case in audit_cases}
    require(len(audit_by_identity) == len(audit_cases), "artifact audit case identities are duplicated")
    for case in cases:
        key = (case.get("origin"), case.get("format"))
        audited = audit_by_identity.get(key)
        require(audited is not None, f"artifact audit missing case: {key}")
        require(isinstance(audited.get("source_equals_output"), bool),
                "artifact audit source identity malformed")
        require(audited.get("edit_admitted") == case.get("edit_admitted")
                and audited.get("edit_outcome") == case.get("edit_outcome"),
                f"artifact audit edit outcome differs: {key}")
        manifest_outputs = [*case.get("policy_outputs", []), case["stream_output"]]
        audit_outputs = audited.get("policy_outputs")
        require(isinstance(audit_outputs, list) and len(audit_outputs) == 5,
                f"artifact audit policy cardinality changed: {key}")
        manifest_by_policy = {output.get("policy"): output for output in manifest_outputs}
        audit_by_policy = {output.get("policy"): output for output in audit_outputs}
        require(set(manifest_by_policy) == {"default", "full", "file-only", "no-sync", "stream"}
                and set(audit_by_policy) == set(manifest_by_policy),
                f"artifact audit policy identity changed: {key}")
        for policy, output in manifest_by_policy.items():
            audited_output = audit_by_policy[policy]
            require(audited_output.get("bytes") == output.get("bytes")
                    and audited_output.get("sha256") == output.get("sha256")
                    and audited_output.get("matches_source") == output.get("matches_source")
                    and resolve_path(audited_output.get("path")).resolve() ==
                    path_for_manifest_output(output, manifest_path).resolve(),
                    f"artifact audit policy digest changed: {key}/{policy}")
        if case.get("origin") == "caller-named-real-file":
            source_name = next((name for name, value in custody_value["corpus"].items()
                                if value.get("sha256") == case.get("source_archive_sha256")), None)
            require(source_name is not None, f"real artifact source is not a packet input: {key}")
            require(case.get("source_archive_bytes") == custody_value["corpus"][source_name]["bytes"],
                    f"real artifact source byte count changed: {key}")
    return {
        "complete": file_identity(complete_path), "manifest": file_identity(manifest_path),
        "receipt": file_identity(PACKET / "artifacts-receipt.json"),
        "audit": file_identity(audit_path),
        "admission_receipt": file_identity(admission_receipt_path),
        "audit_ok": False, "audit_errors": list(audit["errors"]),
        "audit_case_status": [
            {
                "case_id": case["case_id"], "origin": case["origin"],
                "format": case["format"], "ok": case["ok"],
                "errors": list(case["errors"]),
                "changed_members": [row.get("name")
                                    for row in case.get("changed_members", [])],
                "policy_outputs": len(case.get("policy_outputs", [])),
                "source_equals_output": case.get("source_equals_output"),
            }
            for case in audit_cases
        ],
        "cases": len(cases),
        "policy_outputs": sum(len(case["policy_outputs"]) + 1 for case in cases),
        "audit_cases": len(audit_cases),
    }


def path_for_manifest_output(output: dict[str, Any], manifest_path: Path) -> Path:
    """Resolve one manifest output without duplicating descriptor checks."""
    value = output.get("output")
    require(isinstance(value, dict), "artifact policy output descriptor is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, "artifact policy output path is missing")
    candidate = Path(raw)
    if not candidate.is_absolute():
        candidate = manifest_path.parent / candidate
    return candidate.resolve(strict=False)


def load_admissions(p: dict[str, Any], artifacts: dict[str, Any]) -> dict[str, Any]:
    """Replay the failed admission boundary and prove timing was not run."""
    require(not (PACKET / "artifact-admission.json").exists()
            and not (PACKET / "qualification-admission.json").exists()
            and not (PACKET / "artifact-audit.json").exists(),
            "success-only admission files must remain absent")
    attempt_dirs = sorted(path.name for path in PACKET.glob("admission-*") if path.is_dir())
    require(attempt_dirs == ["admission-0"], "unexpected artifact admission attempts retained")
    for lane in ("qualification", "native", "observer"):
        require(not (PACKET / lane).exists(), f"timing lane was unexpectedly retained: {lane}")
    receipt_path = PACKET / "admission-0" / "receipt.json"
    report_path = PACKET / "admission-0" / "artifact-audit.json"
    receipt = read_json(receipt_path)
    report = read_json(report_path)
    require(receipt.get("exit_code") == 1 and report.get("ok") is False,
            "artifact admission failure status changed")
    require(report.get("schema") == "litchi.performance.0817.artifact-audit.v1"
            and len(report.get("cases", [])) == 6,
            "artifact admission audit identity changed")
    case_status = artifacts.get("audit_case_status")
    require(isinstance(case_status, list) and len(case_status) == 6,
            "artifact admission case summary changed")
    return {
        "schema": ADMISSION_SUMMARY_SCHEMA,
        "accepted": False,
        "stage": "artifacts",
        "reason": ADMISSION_FAILURE_REASON,
        "attempt": "admission-0",
        "exit_code": receipt["exit_code"],
        "receipt": file_identity(receipt_path),
        "audit": file_identity(report_path),
        "audit_ok": False,
        "audit_error_count": len(report["errors"]),
        "audit_errors": list(report["errors"]),
        "cases": case_status,
        "case_count": artifacts["cases"],
        "policy_output_count": artifacts["policy_outputs"],
        "artifact_admission_file": None,
        "qualification_admission_file": None,
        "timing": {
            "native_reports": 0, "native_samples": 0,
            "observer_reports": 0, "observer_samples": 0,
            "qualification_reports": 0, "qualification_samples": 0,
        },
    }


def nearest_rank(values: Iterable[float], quantile: float) -> float:
    ordered = sorted(values)
    require(ordered, "nearest rank received no values")
    index = max(0, min(len(ordered) - 1, math.ceil(quantile * len(ordered)) - 1))
    return ordered[index]


def descriptive_workload(source_bytes: Any, published_bytes: Any,
                         elapsed_p50_ns: float, phase: str) -> dict[str, Any]:
    """Return archive-byte rates with an explicit descriptive-only scope.

    The byte counts come from the ordinary-save phase's frozen corpus summary,
    while the denominator is the phase's own elapsed p50.  These rates are a
    way to state the workload size beside the timing; they are not physical
    I/O, compression, memory-bandwidth, or additive phase-throughput claims.
    """
    positive_int(source_bytes, "source logical bytes")
    positive_int(published_bytes, "published logical bytes")
    finite_number(elapsed_p50_ns, "workload elapsed p50")
    require(float(elapsed_p50_ns) > 0, "workload elapsed p50 is not positive")
    factor = 1_000_000_000.0 / float(elapsed_p50_ns)
    return {
        "source_logical_bytes": source_bytes,
        "published_logical_bytes": published_bytes,
        "source_logical_bytes_per_second": float(source_bytes) * factor,
        "published_logical_bytes_per_second": float(published_bytes) * factor,
        "rate_basis": "phase elapsed p50; archive logical byte counts from this phase summary",
        "phase": phase,
        "publication_timed": phase != "edit",
        "claim": "descriptive logical workload rate only; no physical-I/O or memory-bandwidth claim",
    }


def midpoint(left: int | float, right: int | float) -> float:
    return (float(left) + float(right)) / 2.0


def integer_midpoint(left: int, right: int) -> int:
    """Match the harness' overflow-safe integer p50 midpoint."""
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def bootstrap_absolute(values: list[float]) -> dict[str, Any]:
    require(len(values) == 6, "absolute p50 bootstrap requires six blocks")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates = [statistics.median(rng.choice(values) for _ in values)
                 for _ in range(BOOTSTRAP_RESAMPLES)]
    ordered = sorted(estimates)
    return {
        "estimate": statistics.median(values),
        "lower": ordered[BOOTSTRAP_LOW_RANK],
        "upper": ordered[BOOTSTRAP_HIGH_RANK],
        "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
        "confidence": BOOTSTRAP_CONFIDENCE,
        "low_rank": BOOTSTRAP_LOW_RANK, "high_rank": BOOTSTRAP_HIGH_RANK,
        "statistic": "median of six process p50 values",
        "units": "ns",
    }


def report_path_from_receipt(row: dict[str, Any], label: str) -> Path:
    result = artifact(row.get("report"), f"{label} report", packet_bound=True)
    require(result is not None, f"{label} report missing")
    return result


def parse_rss(row: dict[str, Any], label: str) -> tuple[Path, int]:
    path = artifact(row.get("rss"), f"{label} RSS", packet_bound=True)
    require(path is not None, f"{label} RSS missing")
    text = path.read_text(encoding="utf-8").strip()
    require(text.isdigit() and int(text) > 0, f"{label} RSS is invalid")
    return path, int(text)


def exact_vector(value: Any, count: int, label: str) -> list[Any]:
    require(isinstance(value, dict), f"{label} metric vector is malformed")
    values = value.get("values")
    require(isinstance(values, list) and len(values) == count,
            f"{label} metric vector cardinality changed")
    require(all(isinstance(item, int) and not isinstance(item, bool) and item >= 0
                for item in values), f"{label} metric vector value changed")
    return list(values)


def validate_process_delta(value: Any, label: str) -> None:
    require(value is None or isinstance(value, dict), f"{label} delta is malformed")
    if value is None:
        return
    require(set(value) == set(PROCESS_DELTA_NAMES), f"{label} delta fields changed")
    require(all(isinstance(value[name], int) and not isinstance(value[name], bool)
                and value[name] >= 0 for name in PROCESS_DELTA_NAMES),
            f"{label} delta values changed")


def copy_metric_group(group: Any, count: int, label: str,
                      *, required_status: str | None = None) -> dict[str, Any]:
    require(isinstance(group, dict), f"{label} metrics are missing")
    if required_status is not None:
        require(group.get("status") == required_status,
                f"{label} status changed")
    # Preserve every raw field, while checking every vector exposed by the
    # operation-metrics schema.  No diagnostic vector is subtracted or pooled.
    for key, value in group.items():
        if isinstance(value, dict) and "values" in value:
            values = value.get("values")
            require(isinstance(values, list) and len(values) == count,
                    f"{label}.{key} vector cardinality changed")
    return json.loads(json.dumps(group, sort_keys=True))


def validate_operation_metrics(result: dict[str, Any], count: int, lane: str,
                               sample_order: list[int], label: str) -> dict[str, Any]:
    metrics = result.get("operation_metrics")
    require(isinstance(metrics, dict), f"{label} operation_metrics missing")
    require(metrics.get("sample_count") == count
            and isinstance(metrics.get("sample_indices"), list)
            and metrics["sample_indices"] == sample_order
            and sorted(metrics["sample_indices"]) == list(range(count)),
            f"{label} operation metric sample identity changed")
    require(metrics.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{label} operation metric alignment changed")
    allocation = metrics.get("allocation")
    process = metrics.get("process")
    if lane == "native":
        require(isinstance(allocation, dict) and allocation.get("status") == "unavailable",
                f"{label} native allocation instrumentation changed")
        require(isinstance(process, dict) and process.get("status") == "unavailable",
                f"{label} native process instrumentation changed")
        require(metrics.get("latency_claim") == "comparable_timed_operation",
                f"{label} native latency claim changed")
    else:
        require(isinstance(allocation, dict) and allocation.get("status") == "measured",
                f"{label} observer allocation metrics unavailable")
        require(isinstance(process, dict) and process.get("status") == "measured",
                f"{label} observer process metrics unavailable")
        require(metrics.get("latency_claim") == "allocator_instrumented_elapsed_not_latency_claim",
                f"{label} observer latency claim changed")
    require(allocation.get("scope") == "operation_global_system_allocator",
            f"{label} allocator scope changed")
    if lane != "native":
        for name in ALLOCATION_VECTOR_NAMES:
            exact_vector(allocation.get(name), count, f"{label}.allocation.{name}")
        for name in PROCESS_VECTOR_NAMES:
            exact_vector(process.get(name), count, f"{label}.process.{name}")
    return {
        "allocation": copy_metric_group(allocation, count, f"{label}.allocation"),
        "process": copy_metric_group(process, count, f"{label}.process"),
        "operation_metrics": json.loads(json.dumps(metrics, sort_keys=True)),
    }


def validate_report(report_path: Path, receipt: dict[str, Any], case: dict[str, Any],
                    lane: str, p: dict[str, Any], binary: dict[str, Any],
                    rss_kib: int, audit_case: dict[str, Any]) -> dict[str, Any]:
    report = read_json(report_path)
    label = f"{lane}/{case['case']}/block{receipt['block']}"
    require(isinstance(report, dict) and report.get("schema_version") == REPORT_SCHEMA_VERSION,
            f"{label} report schema changed")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{label} tool identity missing")
    expected_binary = "litchi-perf-baseline-alloc" if lane != "native" else "litchi-perf-baseline"
    require(tool.get("binary") == expected_binary and tool.get("profile") == "release",
            f"{label} tool binary changed")
    if lane == "native":
        require(tool.get("instrumentation") == "none",
                f"{label} native instrumentation changed")
    else:
        require(tool.get("instrumentation") ==
                "ordinary_save_procfs_and_system_allocator_operation_scoped",
                f"{label} observer instrumentation changed")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict) and identity.get("binary_sha256") == binary["sha256"]
            and identity.get("binary_bytes") == binary["bytes"]
            and identity.get("executable") is True and identity.get("profile") == "release",
            f"{label} report executable identity changed")
    environment = report.get("environment")
    require(isinstance(environment, dict)
            and environment.get("git_revision") == origin()["base"]
            and environment.get("cpu_affinity") == "12-19",
            f"{label} report environment changed")
    configuration = report.get("configuration")
    lane_plan = p["lanes"][lane]
    require(isinstance(configuration, dict)
            and configuration.get("samples_per_case") == lane_plan["samples"]
            and configuration.get("warmup_iterations_per_case") == lane_plan["warmup"]
            and configuration.get("cases") == [case["case"]]
            and configuration.get("filesystem_cache_states") == ["warm"]
            and configuration.get("filesystem_root_selected") is True,
            f"{label} report configuration changed")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label} result cardinality changed")
    result = results[0]
    require(isinstance(result, dict) and result.get("case") == case["case"],
            f"{label} result selector changed")
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict) and elapsed.get("unit") == "ns",
            f"{label} elapsed statistics missing")
    samples = elapsed.get("samples")
    sample_order = elapsed.get("sample_order")
    count = lane_plan["samples"]
    require(isinstance(samples, list) and len(samples) == count,
            f"{label} sample cardinality changed")
    require(all(isinstance(value, int) and not isinstance(value, bool) and value > 0
                for value in samples), f"{label} elapsed sample changed")
    require(isinstance(sample_order, list) and sorted(sample_order) == list(range(count)),
            f"{label} elapsed sample order changed")
    # The harness retains elapsed samples sorted by (value, original index).
    require(samples == sorted(samples), f"{label} elapsed samples are not sorted")
    expected_p50 = integer_midpoint(samples[(count - 1) // 2], samples[count // 2])
    expected_p95 = int(nearest_rank(samples, 0.95))
    expected_p99 = int(nearest_rank(samples, 0.99))
    require(elapsed.get("min") == samples[0] and elapsed.get("p50") == expected_p50
            and elapsed.get("max") == samples[-1]
            and elapsed.get("p95") == expected_p95 and elapsed.get("p99") == expected_p99,
            f"{label} report quantiles changed")
    finite_number(elapsed.get("mean"), f"{label} mean")
    finite_number(elapsed.get("standard_deviation"), f"{label} standard deviation")
    require(float(elapsed["standard_deviation"]) >= 0,
            f"{label} standard deviation is negative")
    confidence = elapsed.get("confidence_interval_95")
    require(isinstance(confidence, dict)
            and confidence.get("method") == "two-sided Student's t interval for the mean",
            f"{label} confidence interval metadata changed")
    finite_number(confidence.get("lower"), f"{label} confidence lower")
    finite_number(confidence.get("upper"), f"{label} confidence upper")
    require(float(confidence["lower"]) >= 0
            and float(confidence["upper"]) >= float(confidence["lower"]),
            f"{label} confidence interval bounds changed")
    require(abs(float(elapsed["mean"]) - statistics.fmean(samples)) < 1e-9,
            f"{label} report mean changed")
    ordinary = (result.get("source") or {}).get("ordinary_save")
    require(isinstance(ordinary, dict), f"{label} ordinary-save evidence missing")
    require(ordinary.get("format") == case["format"].upper()
            and ordinary.get("origin") == "caller-named-real-file",
            f"{label} corpus origin changed")
    expected_phase = {
        "lifecycle": "open+edit+save", "edit": "edit",
        "atomic_publish": "save-to-path", "counting_publish": "serialize-to-counting-sink",
    }[case["phase"]]
    require(ordinary.get("phase") == expected_phase,
            f"{label} phase changed")
    summary = ordinary.get("corpus")
    require(isinstance(summary, dict), f"{label} ordinary corpus evidence missing")
    real_file = summary.get("real_file")
    expected_input = custody_value_corpus(case["input"])
    require(isinstance(real_file, dict) and real_file.get("path") == str(ROOT / case["input"])
            and real_file.get("bytes") == expected_input["bytes"]
            and real_file.get("sha256") == expected_input["sha256"],
            f"{label} real input identity changed")
    source_bytes = summary.get("source_archive_bytes")
    source_sha = summary.get("source_archive_sha256")
    published_bytes = summary.get("published_bytes")
    published_sha = summary.get("published_sha256")
    positive_int(source_bytes, f"{label} source logical bytes")
    positive_int(published_bytes, f"{label} published logical bytes")
    require(source_bytes == expected_input["bytes"] and source_sha == expected_input["sha256"],
            f"{label} source archive identity changed")
    result_corpus = result.get("corpus")
    require(isinstance(result_corpus, dict)
            and result_corpus.get("archive_bytes") == source_bytes
            and result_corpus.get("archive_sha256") == source_sha
            and result_corpus.get("archive_member_count") == summary.get("source_member_count"),
            f"{label} corpus manifest identity changed")
    positive_int(summary.get("source_member_count"), f"{label} source member count")
    require(ordinary.get("publications_identical") is True
            and ordinary.get("edit_outcomes_identical") is True,
            f"{label} outcome determinism changed")
    require(summary.get("edit_admitted") is True
            and summary.get("repeated_cycles_identical") is True
            and summary.get("repeated_saves_identical") is True
            and is_sha(summary.get("repeated_cycle_sha256"))
            and is_sha(summary.get("repeated_save_sha256"))
            and summary.get("repeated_cycle_sha256") == published_sha
            and summary.get("repeated_save_sha256") == published_sha,
            f"{label} qualification outcome provenance changed")
    outcome = summary.get("edit_outcome")
    require(outcome == "admitted", f"{label} edit outcome was not admitted")
    outcome_digests = ordinary.get("edit_outcome_sha256")
    require(isinstance(outcome_digests, list) and len(outcome_digests) == count
            and all(value == hashlib.sha256(outcome.encode()).hexdigest() for value in outcome_digests),
            f"{label} edit outcome vector changed")
    published = ordinary.get("published_sha256")
    require(isinstance(published, list), f"{label} publication vector missing")
    if case["phase"] == "edit":
        require(len(published) == 0, f"{label} edit publication vector is nonempty")
    else:
        require(len(published) == count and all(isinstance(value, str) and is_sha(value)
                                                for value in published),
                f"{label} publication vector cardinality changed")
    audit_outputs = audit_case.get("policy_outputs")
    require(isinstance(audit_outputs, list) and len(audit_outputs) == 5,
            f"{label} artifact oracle policy identities missing")
    audit_sha = audit_case.get("published_sha256")
    if audit_sha is None:
        audit_sha = audit_outputs[0].get("sha256")
    require(all(isinstance(output, dict) and output.get("sha256") == audit_sha
                for output in audit_outputs),
            f"{label} artifact oracle policy identities differ")
    require(isinstance(audit_sha, str) and is_sha(audit_sha)
            and published_sha == audit_sha,
            f"{label} output identity differs from artifact oracle")
    audit_bytes = audit_outputs[0].get("bytes")
    positive_int(audit_bytes, f"{label} artifact oracle published bytes")
    require(published_bytes == audit_bytes,
            f"{label} published logical byte count differs from artifact oracle")
    split = summary.get("byte_split")
    require(isinstance(split, dict), f"{label} byte split evidence missing")
    for name in (
        "output_total_bytes", "payload_bytes_deflated", "payload_bytes_stored",
        "payload_bytes_identical_to_source", "payload_bytes_regenerated",
        "uncompressed_payload_bytes_compressed", "uncompressed_payload_bytes_regenerated",
        "framing_bytes", "output_member_count", "deflate_member_count",
        "stored_member_count", "members_identical_to_source", "members_regenerated",
    ):
        nonnegative_int(split.get(name), f"{label}.byte_split.{name}")
    require(split.get("output_total_bytes") == published_bytes,
            f"{label} byte split output identity changed")
    sample_split = ordinary.get("sample_byte_split")
    if case["phase"] == "counting_publish":
        require(isinstance(sample_split, dict), f"{label} counting byte split is missing")
        require(sample_split.get("output_total_bytes") == published_bytes,
                f"{label} counting byte split output identity changed")
    else:
        require(sample_split is None, f"{label} non-counting byte split is present")
    if case["phase"] != "edit":
        require(all(value == audit_sha for value in published)
                and result.get("output_sha256") == audit_sha,
                f"{label} output identity changed")
    else:
        require(result.get("output_sha256") is None,
                f"{label} edit output identity is present")
    metrics = validate_operation_metrics(result, count, lane, sample_order, label)
    process_probe = ordinary.get("process_probe")
    if lane == "native":
        require(process_probe is None, f"{label} native procfs probe is present")
    else:
        require(isinstance(process_probe, dict)
                and process_probe.get("phase") == ordinary.get("phase")
                and process_probe.get("timing_scope") == ordinary.get("timing_scope")
                and process_probe.get("fixed_count") == OBSERVER_CONTROL_COUNT
                and process_probe.get("scope") ==
                "ordinary_save_phase_interval_same_process_counters_including_procfs_probe_overhead"
                and process_probe.get("latency_claim") ==
                "diagnostic_only_procfs_probe_instrumentation_latency"
                and process_probe.get("control_scope") ==
                "fixed_32_empty_adjacent_procfs_snapshot_pairs_acquired_before_warmups_and_never_subtracted"
                and isinstance(process_probe.get("empty_adjacent_snapshot_controls"), list)
                and len(process_probe["empty_adjacent_snapshot_controls"]) == OBSERVER_CONTROL_COUNT
                and isinstance(process_probe.get("sample_deltas"), list)
                and len(process_probe["sample_deltas"]) == count,
                f"{label} procfs probe evidence changed")
        for index, control in enumerate(process_probe["empty_adjacent_snapshot_controls"]):
            validate_process_delta(control, f"{label}.procfs_control[{index}]")
        for index, delta in enumerate(process_probe["sample_deltas"]):
            validate_process_delta(delta, f"{label}.procfs_sample[{index}]")
    _, raw_rss = parse_rss(receipt, label)
    require(raw_rss == rss_kib, f"{label} RSS changed while parsing")
    workload = descriptive_workload(source_bytes, published_bytes,
                                    nearest_rank(samples, 0.50), case["phase"])
    return {
        "lane": lane, "block": receipt["block"], "order": receipt["order"],
        "case": case["case"], "format": case["format"], "phase": case["phase"],
        "input": case["input"], "samples": samples, "sample_order": sample_order,
        "rss_kib": rss_kib, "reported": {
            "min": elapsed["min"], "p50": elapsed["p50"], "p95": elapsed["p95"],
            "p99": elapsed["p99"], "max": elapsed["max"], "mean": elapsed["mean"],
        },
        "nearest_rank": {
            "p50_ns": nearest_rank(samples, 0.50),
            "p95_ns": nearest_rank(samples, 0.95),
            "p99_ns": nearest_rank(samples, 0.99),
            "mean_ns": statistics.fmean(samples),
        },
        "source_sha256": real_file["sha256"], "source_logical_bytes": source_bytes,
        "published_logical_bytes": published_bytes,
        "edit_admitted": summary.get("edit_admitted"),
        "edit_outcome": outcome, "published_sha256": audit_sha,
        "workload": workload,
        "operation_metrics": metrics["operation_metrics"],
        "process_probe": None if process_probe is None
        else json.loads(json.dumps(process_probe, sort_keys=True)),
        "report": file_identity(report_path),
    }


_CUSTODY_CORPUS: dict[str, dict[str, Any]] = {}


def custody_value_corpus(name: str) -> dict[str, Any]:
    require(name in _CUSTODY_CORPUS, f"unknown corpus input: {name}")
    return _CUSTODY_CORPUS[name]


def expected_rows(p: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    cases = p["cases"]
    for block, order in enumerate(p["lanes"][lane]["orders"]):
        ordered = cases if order == "forward" else list(reversed(cases))
        result.extend({"lane": lane, "block": block, "order": order, **case}
                      for case in ordered)
    return result


def load_lane(lane: str, p: dict[str, Any], build: dict[str, Any],
              audit_by_format: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    directory = PACKET / lane
    receipts_path = directory / "receipts.json"
    receipts = read_json(receipts_path)
    require(isinstance(receipts, list), f"{lane} receipts are malformed")
    wanted = expected_rows(p, lane)
    require(len(receipts) == len(wanted), f"{lane} report cardinality changed")
    binary_name = p["lanes"][lane]["binary"]
    binary = build["binaries"][binary_name]["artifact"]
    result: list[dict[str, Any]] = []
    for index, (receipt, expected) in enumerate(zip(receipts, wanted)):
        label = f"{lane} receipt {index}"
        require(isinstance(receipt, dict)
                and receipt.get("schema") == "litchi.performance.0817.capture-receipt.v1",
                f"{label} schema changed")
        for key, value in expected.items():
            require(receipt.get(key) == value, f"{label} {key} changed")
        require(receipt.get("exit_code") == 0 and receipt.get("binary") == binary,
                f"{label} execution identity changed")
        started = receipt.get("started")
        ended = receipt.get("ended")
        finite_number(started, f"{label} start time")
        finite_number(ended, f"{label} end time")
        require(float(started) <= float(ended), f"{label} receipt time order changed")
        stem = f"{receipt['block']:02d}-{expected['format']}-{expected['phase']}"
        expected_command = [
            "/usr/bin/time", "-f", "%M", "-o", str(directory / f"{stem}.rss"),
            "taskset", "-c", ",".join(str(cpu) for cpu in p["affinity"]),
            binary["artifact"]["path"], "--warmup", str(p["lanes"][lane]["warmup"]),
            "--samples", str(p["lanes"][lane]["samples"]), "--case", expected["case"],
            "--json", str(directory / f"{stem}.json"), "--filesystem-root",
            str(custody.SCRATCH), "--ooxml-file", str(ROOT / expected["input"]),
        ]
        require(receipt.get("command") == expected_command,
                f"{label} command changed")
        require(receipt.get("root_inputs") == custody_value_root("root_inputs")
                and receipt.get("locks") == custody_value_root("locks")
                and receipt.get("architecture") == custody_value_root("architecture")
                and receipt.get("corpus") == custody_value_root("corpus")
                and receipt.get("host") == custody_value_root("host_sha256")
                and receipt.get("packet") == custody.packet_hashes()
                and receipt.get("drivers") == custody.driver_hashes()
                and receipt.get("unrelated") == custody_value_root("unrelated"),
                f"{label} custody changed")
        source_descriptor = receipt.get("source")
        source_path = artifact(source_descriptor, f"{label} source", packet_bound=True)
        require(source_path is not None
                and source_path.resolve() == (directory / "source.json").resolve(),
                f"{label} source witness changed")
        require(read_json(source_path) == {
                    "production": build["source"]["production"],
                    "tool": build["source"]["tool"],
                }, f"{label} source contents changed")
        scratch_marker = receipt.get("scratch_marker")
        require(isinstance(scratch_marker, dict)
                and resolve_path(scratch_marker.get("path"), packet_bound=False).resolve() ==
                (custody.SCRATCH / ".litchi-performance-0817-owned").resolve(),
                f"{label} scratch marker identity changed")
        artifact(scratch_marker, f"{label} scratch marker", allow_missing=True)
        report_path = report_path_from_receipt(receipt, label)
        _, rss_kib = parse_rss(receipt, label)
        log_path = artifact(receipt.get("log"), f"{label} log", packet_bound=True)
        require(log_path is not None, f"{label} log missing")
        rss_path = resolve_path(receipt["rss"]["path"], packet_bound=True)
        require(report_path.resolve() == (directory / f"{stem}.json").resolve()
                and rss_path.resolve() == (directory / f"{stem}.rss").resolve()
                and log_path.resolve() == (directory / f"{stem}.log").resolve(),
                f"{label} raw artifact paths changed")
        audit_case = audit_by_format[expected["format"].upper()]
        parsed = validate_report(report_path, receipt, expected, lane, p, binary, rss_kib,
                                 audit_case)
        parsed["log"] = rel(log_path)
        result.append(parsed)
    complete_path = directory / "complete.json"
    complete = read_json(complete_path)
    require(isinstance(complete, dict)
            and complete.get("schema") == f"litchi.performance.0817.{lane}.complete.v1"
            and complete.get("blocks") == p["lanes"][lane]["blocks"]
            and complete.get("children") == len(wanted)
            and complete.get("reports") == len(wanted)
            and complete.get("samples") == len(wanted) * p["lanes"][lane]["samples"],
            f"{lane} completion receipt changed")
    require(complete.get("plan_sha256") == sha256(PLAN_PATH)
            and complete.get("build_sha256") == sha256(PACKET / "build.json")
            and complete.get("quality_sha256") == sha256(PACKET / "quality.json"),
            f"{lane} completion provenance changed")
    complete_source = artifact(complete.get("source"), f"{lane} completion source",
                               packet_bound=True)
    require(complete_source is not None
            and complete_source.resolve() == (directory / "source.json").resolve(),
            f"{lane} completion source identity changed")
    receipts_descriptor = artifact(complete.get("receipts"), f"{lane} completion receipts",
                                   packet_bound=True)
    require(receipts_descriptor is not None
            and receipts_descriptor.resolve() == receipts_path.resolve(),
            f"{lane} completion receipts identity changed")
    return result


def custody_value_root(key: str) -> Any:
    require(key in _CUSTODY_ROOT, f"missing custody value: {key}")
    return _CUSTODY_ROOT[key]


_CUSTODY_ROOT: dict[str, Any] = {}


def native_summaries(native: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[str, list[dict[str, Any]]] = {}
    for row in native:
        groups.setdefault(row["case"], []).append(row)
    require(set(groups) == {case["case"] for case in _PLAN["cases"]},
            "native selector set changed")
    output: list[dict[str, Any]] = []
    for case in _PLAN["cases"]:
        rows = sorted(groups[case["case"]], key=lambda row: row["block"])
        require([row["block"] for row in rows] == list(range(6)),
                f"native block coverage changed: {case['case']}")
        blocks = []
        p50s: list[float] = []
        p95s: list[float] = []
        p99s: list[float] = []
        means: list[float] = []
        source_bytes: list[int] = []
        published_bytes: list[int] = []
        for row in rows:
            q = row["nearest_rank"]
            block = {
                "block": row["block"], "order": row["order"],
                "samples": list(row["samples"]), "sample_order": list(row["sample_order"]),
                "p50_ns": q["p50_ns"], "p95_ns": q["p95_ns"],
                "p99_ns": q["p99_ns"], "mean_ns": q["mean_ns"],
                "raw_reported": dict(row["reported"]),
                "quantile_definition": "nearest rank within this process block",
                "raw_report_quantile_definition":
                    "harness p50 integer midpoint; p95/p99 nearest rank",
                "rss_kib": row["rss_kib"], "report": row["report"],
                "workload": row["workload"],
            }
            blocks.append(block)
            p50s.append(float(q["p50_ns"]))
            p95s.append(float(q["p95_ns"]))
            p99s.append(float(q["p99_ns"]))
            means.append(float(q["mean_ns"]))
            source_bytes.append(row["workload"]["source_logical_bytes"])
            published_bytes.append(row["workload"]["published_logical_bytes"])
        require(len(set(source_bytes)) == 1 and len(set(published_bytes)) == 1,
                f"native workload byte identity changed: {case['case']}")
        spread_values = {
            "p50": max(p50s) / min(p50s) - 1.0,
            "p95": max(p95s) / min(p95s) - 1.0,
            "p99": max(p99s) / min(p99s) - 1.0,
            "mean": max(means) / min(means) - 1.0,
        }
        tail = statistics.median(p99s) / statistics.median(p50s)
        summary = {
            "case": case["case"], "format": case["format"], "phase": case["phase"],
            "input": case["input"], "blocks": 6, "samples_per_report": 30,
            "p50_ns": statistics.median(p50s), "p95_ns": statistics.median(p95s),
            "p99_ns": statistics.median(p99s), "mean_ns": statistics.median(means),
            "p50_block_values_ns": p50s, "p95_block_values_ns": p95s,
            "p99_block_values_ns": p99s, "mean_block_values_ns": means,
            "spread_ratios": spread_values,
            "spread_flags": [name for name, value in spread_values.items() if value > 0.05],
            "spread_flag": any(value > 0.05 for value in spread_values.values()),
            "p99_to_p50_ratio": tail, "tail_flag": tail > 1.05,
            "p50_bootstrap_ci95_ns": bootstrap_absolute(p50s),
            "block_raw_quantiles": blocks,
            "quantile_definition":
                "nearest rank within each process; median over six process blocks",
            "raw_report_quantile_definition":
                "harness p50 integer midpoint; p95/p99 nearest rank",
            "workload": descriptive_workload(source_bytes[0], published_bytes[0],
                                              statistics.median(p50s), case["phase"]),
            "timing_scope": "native elapsed evidence for one real-file ordinary-save selector",
            "historical_timing_comparison": False,
        }
        output.append(summary)
    require(len(output) == 12, "native summary cardinality changed")
    return output


NATIVE_FIELDS = [
    "case", "format", "phase", "input", "p50_ns", "p95_ns", "p99_ns", "mean_ns",
    "p50_bootstrap_ci95_ns", "spread_ratios", "spread_flags", "spread_flag",
    "p99_to_p50_ratio", "tail_flag", "historical_timing_comparison",
    "source_logical_bytes", "published_logical_bytes",
    "source_logical_bytes_per_second", "published_logical_bytes_per_second",
]


def csv_text(rows: list[dict[str, Any]]) -> str:
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=NATIVE_FIELDS, lineterminator="\n")
    writer.writeheader()
    writer.writerows({
        **{field: row[field] for field in NATIVE_FIELDS
           if field not in {"source_logical_bytes", "published_logical_bytes",
                            "source_logical_bytes_per_second",
                            "published_logical_bytes_per_second"}},
        "source_logical_bytes": row["workload"]["source_logical_bytes"],
        "published_logical_bytes": row["workload"]["published_logical_bytes"],
        "source_logical_bytes_per_second":
            row["workload"]["source_logical_bytes_per_second"],
        "published_logical_bytes_per_second":
            row["workload"]["published_logical_bytes_per_second"],
    } for row in rows)
    return output.getvalue()


def render_markdown(analysis: dict[str, Any]) -> str:
    lines = [
        "# 0817 real-file ordinary-save baseline",
        "",
        "This is an offline replay of retained real-file ordinary-save receipts.",
        "The packet is a baseline only; native elapsed values are not compared with historical runs.",
        "Observer allocation and procfs values remain diagnostic evidence and are never pooled with native timing.",
        "Production and runtime harness sources match the base; the declared test-only recovery file is retained in custody.",
        "Analysis quantiles use nearest rank per process; raw report p50 uses the harness integer midpoint.",
        "",
        f"- Native reports/samples: {analysis['counts']['native_reports']} / {analysis['counts']['native_samples']}",
        f"- Observer reports/samples: {analysis['counts']['observer_reports']} / {analysis['counts']['observer_samples']}",
        f"- Qualification reports/samples: {analysis['counts']['qualification_reports']} / {analysis['counts']['qualification_samples']}",
        "",
        "| Selector | p50 ns | p95 ns | p99 ns | mean ns | p50 CI95 ns | source logical B/s | published logical B/s | Spread | Tail |",
        "|---|---:|---:|---:|---:|---|---:|---:|---|---|",
    ]
    for row in analysis["native"]:
        ci = row["p50_bootstrap_ci95_ns"]
        lines.append("| " + " | ".join([
            row["case"], f"{row['p50_ns']:.6g}", f"{row['p95_ns']:.6g}",
            f"{row['p99_ns']:.6g}", f"{row['mean_ns']:.6g}",
            f"[{ci['lower']:.6g}, {ci['upper']:.6g}]",
            f"{row['workload']['source_logical_bytes_per_second']:.6g}",
            f"{row['workload']['published_logical_bytes_per_second']:.6g}",
            ",".join(row["spread_flags"]) or "-", "yes" if row["tail_flag"] else "-",
        ]) + " |")
    lines.extend(["", "No speedup, adoption, optimization, or historical timing claim is made.", ""])
    return "\n".join(lines)


def build_analysis() -> dict[str, Any]:
    global _CUSTODY_CORPUS, _CUSTODY_ROOT, _PLAN
    custody_value = load_packet_custody()
    _CUSTODY_CORPUS = custody_value["corpus"]
    _CUSTODY_ROOT = custody_value
    _PLAN = custody_value["plan"]
    p = _PLAN
    build = load_build(custody_value, p)
    quality = load_quality(custody_value, build)
    artifacts = load_artifacts(custody_value, p, build, quality)
    admissions = load_admissions(p, artifacts)
    # The audit failed before the two admission receipts could be written.
    # Keep the timing shape explicit and empty: no qualification, native, or
    # observer process was authorized, so there are no quantiles or workload
    # rates to derive.
    native: list[dict[str, Any]] = []
    observer: list[dict[str, Any]] = []
    qualification: list[dict[str, Any]] = []
    analysis = {
        "schema": ANALYSIS_SCHEMA,
        "plan_schema": p["schema"], "report_schema_version": REPORT_SCHEMA_VERSION,
        "scope": p["scope"], "regression_policy": p["regression_policy"],
        "historical_timing_comparison": False,
        "status": "admission_failed",
        "timing_status": {
            "accepted": False,
            "reason": ADMISSION_FAILURE_REASON,
            "reports": 0,
            "samples": 0,
            "lanes": {
                "qualification": {"reports": 0, "samples": 0},
                "native": {"reports": 0, "samples": 0},
                "observer": {"reports": 0, "samples": 0},
            },
        },
        "custody": {
            "base": origin()["base"], "production_file_count": PRODUCTION_FILES,
            "architecture_file_count": ARCHITECTURE_FILES,
            "production_matches_base": True, "harness_matches_base": False,
            "runtime_harness_matches_base": True,
            "tool_allowlist": list(TOOL_ALLOWLIST),
            "allowlisted_test_change": ALLOWLISTED_TEST_SOURCE,
            "corpus_inputs": custody_value["corpus"],
            "provenance": file_identity(PACKET / "provenance.json"),
            "unrelated": custody_value["unrelated"],
        },
        "build": build, "quality": quality, "artifacts": artifacts,
        "admissions": admissions,
        "counts": {
            "reports": 0, "samples": 0,
            "native_reports": 0, "native_samples": 0,
            "observer_reports": 0, "observer_samples": 0,
            "qualification_reports": 0, "qualification_samples": 0,
            "artifact_cases": artifacts["cases"],
            "artifact_policy_outputs": artifacts["policy_outputs"],
        },
        "bootstrap": {
            "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
            "confidence": BOOTSTRAP_CONFIDENCE, "low_rank": BOOTSTRAP_LOW_RANK,
            "high_rank": BOOTSTRAP_HIGH_RANK,
            "statistic": "median of six process p50 values", "units": "ns",
            "status": "not_run",
        },
        "native": [],
        "observer": {
            "timings_pooled_with_native": False,
            "latency_claim": "diagnostic_only; observer elapsed values are not latency evidence",
            "operation_metrics_subtracted": False,
            "procfs_controls_subtracted": False,
            "reports": [],
        },
        "qualification": [],
        "raw_reports": {"native": native, "observer": observer,
                        "qualification": qualification},
        "verification": {
            "packet_custody_checked": True, "production_source_checked": True,
            "harness_source_checked": True, "runtime_harness_source_checked": True,
            "allowlisted_test_source_checked": True,
            "quality_checked": True, "quality_recovery_checked": True,
            "build_driver_recovery_checked": True,
            "build_checked": True, "artifact_export_checked": True,
            "independent_artifact_oracle_checked": True,
            "artifact_admission_checked": True, "qualification_admission_checked": True,
            "report_schema_checked": True, "binary_identity_checked": True,
            "corpus_identity_checked": True, "outcome_identity_checked": True,
            "sample_cardinality_checked": True, "timing_absence_checked": True,
            "observer_diagnostics_retained": False,
            "native_timing_separated": True, "historical_comparison_omitted": True,
            "bootstrap_checked": False, "logical_workload_descriptors_checked": False,
        },
    }
    return analysis


def write_outputs(analysis: dict[str, Any]) -> None:
    outputs = {
        "analysis.json": json.dumps(analysis, indent=2, sort_keys=True) + "\n",
        "admission-summary.json": json.dumps(analysis["admissions"], indent=2,
                                              sort_keys=True) + "\n",
    }
    for name in outputs:
        candidate = PACKET / name
        require(not candidate.exists() and not candidate.is_symlink(),
                f"refusing to overwrite retained {name}")
    for name, text in outputs.items():
        (PACKET / name).write_text(text, encoding="utf-8")


def check_outputs(analysis: dict[str, Any]) -> None:
    path = PACKET / "analysis.json"
    retained = read_json(path)
    require(retained == analysis, "analysis.json does not replay deterministically")
    require(path.read_text(encoding="utf-8") ==
            json.dumps(analysis, indent=2, sort_keys=True) + "\n",
            "analysis.json formatting does not replay deterministically")
    expected = {
        "admission-summary.json": json.dumps(analysis["admissions"], indent=2,
                                              sort_keys=True) + "\n",
    }
    for name, text in expected.items():
        candidate = PACKET / name
        require(candidate.is_file() and not candidate.is_symlink(), f"{name} is missing")
        require(candidate.read_text(encoding="utf-8") == text,
                f"{name} does not replay deterministically")


def analyze(*, write: bool = False, check: bool = False) -> dict[str, Any]:
    result = build_analysis()
    if write:
        write_outputs(result)
    if check:
        check_outputs(result)
    return result


def main(argv: list[str] | None = None) -> int:
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true", help="write immutable derived outputs")
    mode.add_argument("--check", action="store_true", help="replay retained outputs")
    args = parser.parse_args(argv)
    try:
        analyze(write=args.write or not args.check, check=args.check)
    except (ReplayError, AssertionError, OSError, ValueError, KeyError, TypeError,
            IndexError) as error:
        print(f"0817 replay failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
