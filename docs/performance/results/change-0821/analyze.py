"""Fail-closed offline replay for the 0821 real-file save-durability packet.

This reader consumes retained JSON, logs, archives, binary receipts, and the
committed 8312 blobs used for the five rolling 0820 index witnesses. It never
starts Cargo, the exporter, a benchmark child, or a workload. The same replay
builds the derived analysis before cleanup and after cleanup; the ``--check``
mode refuses any changed derived output.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import math
import random
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable

import custody


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN_SCHEMA = "litchi.performance.0821.plan.v1"
ANALYSIS_SCHEMA = "litchi.performance.0821.save-durability-analysis.v1"
REPORT_SCHEMA_VERSION = 1
CAPTURE_SCHEMA = "litchi.performance.0821.capture-receipt.v1"
BOOTSTRAP_SEED = 821821
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_LOW_RANK = 250
BOOTSTRAP_HIGH_RANK = 9749
BOOTSTRAP_CONFIDENCE = 0.95
PRODUCTION_FILES = 9_197
TOOL_FILES = 87
ARCHITECTURE_FILES = 35
OBSERVER_CONTROLS = 32
BASE_REVISION = "8312aaa29b59f73d2a7409cb501828989320a8c9"
PRIOR_REPAIR_REVISION = "096b810f23cc66fe01ec8c96c74f36888f954eff"
QUALITY_REUSE_SCHEMA = "litchi.performance.0821.quality-reuse.v1"
QUALITY_SCHEMA = "litchi.performance.0821.quality.v1"
REPAIR_QUALITY_SCHEMA = "litchi.performance.0820.repair-quality.v1"
REPAIR_TESTS_SCHEMA = "litchi.performance.0820.repair-tests.v1"
REPAIR_COMMAND_SCHEMA = "litchi.performance.0820.repair-command.v1"
REPAIR_INPUTS_SCHEMA = "litchi.performance.0820.repair-inputs.v1"
PRIOR_SEAL_SCHEMA = "litchi.performance.0820.seal.v1"
PRIOR_SEAL_MUTABLE_INDEXES = frozenset({
    "docs/performance/BASELINE.md",
    "docs/performance/CRUD_COVERAGE.md",
    "docs/performance/GOAL_AUDIT.md",
    "docs/performance/HOTSPOTS.md",
    "docs/performance/REPORT.md",
})
HEX = frozenset("0123456789abcdefABCDEF")
EDIT_OUTCOME_SHA = hashlib.sha256(b"admitted").hexdigest()
SAVE_ENTRY_POINTS = {
    "DOCX": "litchi_docx::Package::save",
    "XLSX": "litchi_xlsx::Workbook::save",
    "PPTX": "litchi_pptx::Package::save",
}
SINK_ENTRY_POINTS = {
    "DOCX": "litchi_docx::Package::to_stream",
    "XLSX": "litchi_xlsx::Workbook::write_to",
    "PPTX": "litchi_pptx::Package::to_bytes",
}
FULL_ATOMIC_STEPS = (
    "litchi_opc::atomic::replace_with: destination permission probe, "
    "sibling temporary creation in the destination's own directory, the "
    "publication write, permission preservation, sync_all on the temporary, "
    "persist (rename) over the destination, parent-directory sync"
)
ATOMIC_STEPS_BY_POLICY = {
    "default": FULL_ATOMIC_STEPS,
    "full": FULL_ATOMIC_STEPS,
    "file-only": (
        "litchi_opc::atomic::replace_with_durability(FileOnly): destination permission probe, "
        "sibling temporary creation in the destination's own directory, the publication write, "
        "permission preservation, sync_all on the temporary, persist (rename) over the "
        "destination; no parent-directory sync"
    ),
    "no-sync": (
        "litchi_opc::atomic::replace_with_durability(NoSync): destination permission probe, "
        "sibling temporary creation in the destination's own directory, the publication write, "
        "permission preservation, persist (rename) over the destination; no sync_all and no "
        "parent-directory sync"
    ),
}
POLICIES = ("default", "full", "file-only", "no-sync")
POLICY_SEMANTICS = {
    "default": "documented save; full durability",
    "full": "explicit save_with_durability Full; temporary file sync_all and parent-directory sync",
    "file-only": "explicit FileOnly; temporary file sync_all, no parent-directory sync",
    "no-sync": "explicit NoSync; neither temporary-file nor parent-directory sync",
}
GENERIC_TIMING_SCOPES = {
    "lifecycle": (
        "open the path through the documented reader, make one semantic edit, save to a "
        "path; destination preparation, readback, digest and cleanup are outside the clock"
    ),
    "atomic_publish": (
        "one documented save-to-path only: sibling creation in the destination directory, "
        "the publication write, permission preservation, the temporary's data sync, the "
        "rename that replaces the destination, and the parent-directory sync; the open, the "
        "edit, the readback and the cleanup are outside the clock"
    ),
}
BYTE_SPLIT_SCOPE = (
    "derived from the published archive and the source archive, not from production counters: "
    "litchi-opc's OpcOperationAccounting excludes PartWriter and the topology publishers, "
    "which is the path a documented save takes. `payload_bytes_identical_to_source` is an "
    "upper bound on what a copy-through publisher could have avoided re-deflating, not an "
    "observation that this writer copied anything."
)


class ReplayError(RuntimeError):
    """Evidence is missing, stale, malformed, or contradictory."""


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


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def resolve_path(raw: Any, *, packet_bound: bool = False) -> Path:
    require(isinstance(raw, str) and raw, f"invalid artifact path: {raw!r}")
    value = Path(raw)
    candidates: list[Path] = []
    if value.is_absolute():
        candidates.append(value)
    else:
        candidates.append(PACKET / value)
        prefix = "docs/performance/results/change-0821/"
        if raw.startswith(prefix):
            candidates.insert(0, PACKET / raw[len(prefix):])
        candidates.append(ROOT / value)
    for candidate in candidates:
        candidate = candidate.resolve(strict=False)
        if packet_bound and not candidate.is_relative_to(PACKET.resolve()):
            continue
        if candidate.is_file() and not candidate.is_symlink():
            return candidate
    candidate = candidates[0].resolve(strict=False)
    if packet_bound:
        require(candidate.is_relative_to(PACKET.resolve()),
                f"artifact path escaped packet: {raw}")
    return candidate


def descriptor(value: Any, label: str, *, packet_bound: bool = False,
               allow_missing: bool = False) -> Path | None:
    require(isinstance(value, dict), f"{label} is not an artifact descriptor")
    path = resolve_path(value.get("path"), packet_bound=packet_bound)
    nonnegative_int(value.get("bytes"), f"{label}.bytes")
    require(is_sha(value.get("sha256")), f"{label}.sha256 is invalid")
    if not path.is_file():
        require(allow_missing, f"missing {label}: {path}")
        return None
    require(not path.is_symlink(), f"{label} is a symlink: {path}")
    require(path.stat().st_size == value["bytes"], f"{label}.bytes changed")
    require(sha256(path) == value["sha256"], f"{label}.sha256 changed")
    return path


def identity(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    return {"path": relative(path), "bytes": path.stat().st_size,
            "sha256": sha256(path)}


def same_descriptor(value: Any, expected: dict[str, Any], label: str,
                    *, packet_bound: bool = True) -> Path:
    path = descriptor(value, label, packet_bound=packet_bound)
    expected_path = resolve_path(expected.get("path"), packet_bound=packet_bound)
    require(path is not None and path.resolve() == expected_path.resolve(),
            f"{label}.path changed")
    require(value.get("bytes") == expected.get("bytes")
            and value.get("sha256") == expected.get("sha256"),
            f"{label} identity changed")
    return path


def origin() -> dict[str, Any]:
    value = read_json(PACKET / "origin.json")
    require(value.get("schema") == "litchi.performance.0821.origin.v1",
            "origin schema changed")
    require(value.get("base") == BASE_REVISION,
            "origin base is invalid")
    require(value.get("production_changed") is False
            and value.get("runtime_harness_changed") is False
            and value.get("tool_changed") is False
            and value.get("tool_allowlist") == [],
            "source-change witness changed")
    require(value.get("unrelated") == custody.UNRELATED,
            "unrelated-file witness changed")
    reuse = value.get("quality_reuse")
    require(isinstance(reuse, dict)
            and reuse.get("schema") == QUALITY_REUSE_SCHEMA
            and reuse.get("mode") == "committed-receipt-replay"
            and reuse.get("cargo_commands_executed") is False
            and reuse.get("previous_seal_checked") is True
            and reuse.get("source_hashes_checked") is True
            and reuse.get("result_and_log_descriptors_checked") is True
            and reuse.get("required_gate_count") == 6
            and reuse.get("required_test_counts") == {
                "failed": 0, "ignored": 1, "passed": 641, "suites": 28,
            }, "quality reuse boundary changed")
    return value


def plan() -> dict[str, Any]:
    value = read_json(PACKET / "plan.json")
    require(value.get("schema") == PLAN_SCHEMA, "plan schema changed")
    origin_value = origin()
    require(value.get("base") == origin_value["base"] == BASE_REVISION,
            "plan base changed")
    require(value.get("target") == str(custody.TARGET)
            and value.get("scratch") == str(custody.SCRATCH),
            "owned build roots changed")
    require(value.get("affinity") == list(range(12, 20)), "capture affinity changed")
    require(value.get("scope") ==
            "warm filesystem/provider caches; paired durability-policy configuration attribution; default/full contract retained; phases independently prepared and not additive; native and observer latencies never pooled",
            "analysis scope changed")
    require(value.get("admission") ==
            "fresh release export of six corpora/five policy outputs; independent XML and ZIP preservation gates before qualification; all24selectors bound to admitted source and policy hash; retain typed refusals and failures",
            "admission contract changed")
    require(value.get("binaries") == {
        "native": {"cargo_bin": "litchi-perf-baseline", "features": [],
                    "latency": "native"},
        "artifacts": {"cargo_bin": "ordinary_save_artifacts", "features": [],
                       "latency": "untimed"},
        "observer": {"cargo_bin": "litchi-perf-baseline-alloc",
                      "features": ["allocator-metrics", "ordinary-save-process-metrics"],
                      "latency": "diagnostic_only"},
    }, "binary contract changed")
    cases = custody.plan_cases(value)
    require(value.get("expected") == {
        "artifact_cases": 6, "artifact_policy_outputs_per_case": 5,
        "native_reports": 144, "native_samples": 4320,
        "observer_reports": 48, "observer_samples": 144,
        "qualification_reports": 24, "qualification_samples": 24,
        "total_reports": 216, "total_samples": 4488,
    }, "plan cardinality changed")
    require(value.get("bootstrap") == {
        "confidence": BOOTSTRAP_CONFIDENCE, "high_rank": BOOTSTRAP_HIGH_RANK,
        "low_rank": BOOTSTRAP_LOW_RANK, "resamples": BOOTSTRAP_RESAMPLES,
        "seed": BOOTSTRAP_SEED, "statistic": "median of six process p50 values",
        "paired_statistic": "median of six matched process-block policy/default p50 ratios",
    }, "bootstrap contract changed")
    require(value.get("quantile") ==
            "nearest rank within each process; median over six process blocks",
            "analysis quantile contract changed")
    require(value.get("spread_flag") == "max/min block metric > 1.05"
            and value.get("tail_flag") ==
            "median process p99 / median process p50 > 1.05",
            "diagnostic flag contract changed")
    require(value.get("regression_policy") ==
            "configuration attribution only; no production optimization, historical timing comparisons, or weakened default adoption",
            "regression policy changed")
    require(value.get("comparison") ==
            "within-block policy/default ratios for the same format and phase; full/default controls explicit API route; weaker-policy differences change durability semantics",
            "comparison policy changed")
    require(value.get("policies") == list(POLICIES)
            and value.get("policy_semantics") == POLICY_SEMANTICS,
            "policy semantics changed")
    require(value.get("source") == {
        "harness_changed": False,
        "lock_graph": "existing tracked perf-baseline Cargo.lock; exact differences from workspace lock retained",
        "production_changed": False,
        "runtime_harness_changed": False,
    }, "source plan changed")
    require(value.get("lanes") == {
        "native": {"binary": "native", "blocks": 6,
                    "orders": ["forward", "reverse", "forward", "reverse", "reverse", "forward"],
                    "policy_orders": [
                        ["default", "full", "file-only", "no-sync"],
                        ["no-sync", "file-only", "full", "default"],
                        ["full", "file-only", "no-sync", "default"],
                        ["default", "no-sync", "file-only", "full"],
                        ["file-only", "no-sync", "default", "full"],
                        ["full", "default", "no-sync", "file-only"],
                    ], "samples": 30, "warmup": 3},
        "observer": {"binary": "observer", "blocks": 2,
                      "orders": ["forward", "reverse"],
                      "policy_orders": [
                          ["default", "full", "file-only", "no-sync"],
                          ["no-sync", "file-only", "full", "default"],
                      ], "samples": 3, "warmup": 0},
        "qualification": {"binary": "observer", "blocks": 1,
                           "orders": ["forward"],
                           "policy_orders": [["default", "full", "file-only", "no-sync"]],
                           "samples": 1, "warmup": 0},
    }, "lane plan changed")
    require(value.get("build") == {
        "codegen_units": 1, "debug": 1, "incremental": False, "jobs": 2,
        "lto": "thin", "opt_level": 3, "panic": "unwind", "rustflags": None,
    }, "build plan changed")
    require(value.get("policies") == list(POLICIES), "policy list changed")
    require(value.get("quality_reuse") == origin_value["quality_reuse"],
            "quality reuse plan changed")
    require(len(cases) == 24, "case cardinality changed")
    return value


def load_custody() -> dict[str, Any]:
    p = plan()
    root_inputs = custody.assert_root_inputs()
    locks = custody.lock_identity()
    corpus = custody.assert_corpus_inputs()
    provenance = custody.assert_provenance(corpus)
    architecture = custody.architecture_hashes()
    require(len(architecture) == ARCHITECTURE_FILES, "architecture census changed")
    host = custody.assert_host()
    unrelated = custody.assert_unrelated()
    current = custody.source()
    require(len(current["files"]) == PRODUCTION_FILES,
            "production source census changed")
    tool = custody.tool_source()
    require(current.get("revision") == BASE_REVISION,
            "production source revision changed")
    require(len(tool) == TOOL_FILES, "tool source census changed")
    return {"plan": p, "origin": origin(), "root_inputs": root_inputs,
            "locks": locks, "corpus": corpus, "provenance": provenance,
            "architecture": architecture, "host": host, "unrelated": unrelated,
            "current_source": current, "current_tool": tool,
            "packet": custody.packet_hashes(), "drivers": custody.driver_hashes()}


def source_witness(value: Any, custody_value: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(value, dict)
            and isinstance(value.get("production"), dict)
            and isinstance(value.get("tool"), dict), f"{label} source is malformed")
    production = value["production"]
    require(production.get("revision") == custody_value["origin"]["base"],
            f"{label} base revision changed")
    require(isinstance(production.get("files"), dict)
            and len(production["files"]) == PRODUCTION_FILES
            and production["files"] == custody_value["current_source"]["files"],
            f"{label} production source hashes changed")
    require(value["tool"] == custody_value["current_tool"],
            f"{label} tool source hashes changed")
    return {"production": production, "tool": dict(value["tool"])}


def cleanup_value() -> dict[str, Any] | None:
    path = PACKET / "cleanup.json"
    if not path.exists():
        return None
    value = read_json(path)
    require(set(value) == {
        "schema", "target_removed", "scratch_removed",
        "binaries_verified_before_removal", "removed_binaries", "source",
        "removed", "started", "ended",
    }, "cleanup keys changed")
    require(value.get("schema") == "litchi.performance.0821.cleanup.v1",
            "cleanup schema changed")
    require(value.get("target_removed") is True
            and value.get("scratch_removed") is True
            and value.get("binaries_verified_before_removal") is True,
            "cleanup status changed")
    binaries = value.get("removed_binaries")
    require(isinstance(binaries, list) and len(binaries) == 3,
            "cleanup binary cardinality changed")
    require(all(isinstance(item, dict) and set(item) == {"path", "bytes", "sha256"}
                and nonnegative_int(item.get("bytes"), "cleanup bytes") is None
                and is_sha(item.get("sha256")) for item in binaries),
            "cleanup binary descriptors changed")
    removed = value.get("removed")
    require(isinstance(removed, list) and len(removed) == 2
            and all(isinstance(row, dict)
                    and set(row) == {"path", "files", "logical_bytes"}
                    and isinstance(row.get("path"), str)
                    and nonnegative_int(row.get("files"), "cleanup file count") is None
                    and nonnegative_int(row.get("logical_bytes"), "cleanup logical bytes") is None
                    for row in removed), "cleanup directory rows changed")
    require({row.get("path") for row in removed} ==
            {str(custody.TARGET), str(custody.SCRATCH)},
            "cleanup directory witness changed")
    finite(value.get("started"), "cleanup start")
    finite(value.get("ended"), "cleanup end")
    require(value["started"] <= value["ended"], "cleanup timestamps reversed")
    return value


def cleanup_has_binary(cleanup: dict[str, Any] | None, expected: dict[str, Any]) -> bool:
    return isinstance(cleanup, dict) and expected in cleanup.get("removed_binaries", [])


def load_build(cv: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "build.json"
    build = read_json(path)
    require(build.get("schema") == "litchi.performance.0821.build.v1",
            "build schema changed")
    source_path = descriptor(build.get("source"), "build source", packet_bound=True)
    require(source_path is not None, "build source missing")
    source = source_witness(read_json(source_path), cv, "build")
    for key in ("root_inputs", "locks", "architecture", "corpus", "provenance",
                "host", "unrelated", "packet", "drivers"):
        expected = {"host": cv["host"]}.get(key, cv.get(key))
        if key == "provenance":
            expected = cv["provenance"]
        require(build.get(key) == expected, f"build {key} witness changed")
    frozen = descriptor(build.get("frozen_inputs"), "build frozen inputs", packet_bound=True)
    require(frozen is not None, "build frozen inputs missing")
    frozen_value = read_json(frozen)
    require(frozen_value.get("root_inputs") == cv["root_inputs"]
            and frozen_value.get("locks") == cv["locks"]
            and frozen_value.get("architecture") == cv["architecture"]
            and frozen_value.get("corpus") == cv["corpus"]
            and frozen_value.get("provenance") == cv["provenance"]
            and frozen_value.get("host") == cv["host"]
            and frozen_value.get("unrelated") == cv["unrelated"],
            "build frozen custody changed")
    expected_rows = {
        "native": ("litchi-perf-baseline", []),
        "artifacts": ("ordinary_save_artifacts", []),
        "observer": ("litchi-perf-baseline-alloc",
                      ["allocator-metrics", "ordinary-save-process-metrics"]),
    }
    plan_build = cv["plan"]["build"]
    expected_environment = {
        "CARGO_TARGET_DIR": str(custody.TARGET),
        "CARGO_BUILD_JOBS": str(plan_build["jobs"]),
        "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_OPT_LEVEL": str(plan_build["opt_level"]),
        "CARGO_PROFILE_RELEASE_DEBUG": str(plan_build["debug"]),
        "CARGO_PROFILE_RELEASE_LTO": str(plan_build["lto"]),
        "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": str(plan_build["codegen_units"]),
        "CARGO_PROFILE_RELEASE_INCREMENTAL": str(plan_build["incremental"]).lower(),
        "CARGO_PROFILE_RELEASE_PANIC": str(plan_build["panic"]),
    }
    require(build.get("profile") == {key: plan_build[key] for key in
                                      ("opt_level", "debug", "lto", "codegen_units",
                                       "incremental", "panic")},
            "build profile changed")
    require(build.get("environment") == expected_environment,
            "build environment changed")
    rows = build.get("rows")
    require(isinstance(rows, list) and len(rows) == 3, "build rows changed")
    seen = set()
    for row in rows:
        name = row.get("binary")
        require(name in expected_rows and name not in seen, "build row identity changed")
        seen.add(name)
        cargo_name, features = expected_rows[name]
        require(row.get("cargo_binary") == cargo_name and row.get("features") == features
                and row.get("exit_code") == 0, f"build result changed: {name}")
        expected_command = ["cargo", "build", "--offline", "--locked", "--release",
                            "--manifest-path", str(custody.TOOL / "Cargo.toml"),
                            "--bin", cargo_name]
        if features:
            expected_command.extend(["--features", ",".join(features)])
        require(row.get("command") == expected_command,
                f"build command changed: {name}")
        require(row.get("environment") == expected_environment,
                f"build row environment changed: {name}")
        log = descriptor(row.get("log"), f"build {name} log", packet_bound=True)
        require(log is not None, f"build {name} log missing")
        finite(row.get("started"), f"build {name} start")
        finite(row.get("ended"), f"build {name} end")
        require(row["started"] <= row["ended"], f"build {name} timestamps reversed")
    require(seen == set(expected_rows), "build rows incomplete")
    cleanup = cleanup_value()
    if cleanup is not None:
        same_descriptor(cleanup.get("source"), build.get("source"),
                        "cleanup source witness")
    binaries: dict[str, Any] = {}
    for name, (cargo_name, features) in expected_rows.items():
        entry = build.get("binaries", {}).get(name)
        require(isinstance(entry, dict) and entry.get("cargo_name") == cargo_name
                and entry.get("features") == features, f"build binary metadata changed: {name}")
        raw = entry.get("artifact")
        path_value = descriptor(raw, f"build {name} binary", allow_missing=True)
        if path_value is None:
            require(cleanup_has_binary(cleanup, {k: raw.get(k) for k in ("path", "bytes", "sha256")}),
                    f"missing {name} binary lacks cleanup witness")
        binaries[name] = {"cargo_name": cargo_name, "features": list(features),
                          "artifact": {k: raw[k] for k in ("path", "bytes", "sha256")}}
    return {"receipt": identity(path), "source": source, "source_path": relative(source_path),
            "frozen_inputs": identity(frozen), "binaries": binaries, "raw": build,
            "cleanup": cleanup}


def root_path(raw: Any, label: str) -> Path:
    """Resolve a retained path without allowing it to escape the workspace."""
    require(isinstance(raw, str) and raw, f"{label} path is invalid")
    value = Path(raw)
    path = value if value.is_absolute() else ROOT / value
    path = path.resolve(strict=False)
    require(path.is_relative_to(ROOT.resolve()), f"{label} escaped workspace")
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    return path


def referenced_artifact(raw: Any, expected: Path, label: str) -> Path:
    """Check either a path string or the normal path/bytes/hash descriptor."""
    expected = expected.resolve()
    if isinstance(raw, dict):
        path = descriptor(raw, label)
    else:
        path = root_path(raw, label)
    require(path.resolve() == expected, f"{label} path changed")
    if isinstance(raw, dict):
        require(raw.get("bytes") == expected.stat().st_size
                and raw.get("sha256") == sha256(expected),
                f"{label} identity changed")
    return expected


def repair_reference(provenance: dict[str, Any], receipt: dict[str, Any],
                     key: str, label: str) -> Path:
    expected = root_path(provenance.get(key), f"quality reuse {label}")
    return referenced_artifact(receipt.get(key), expected, f"quality reuse {label}")


def check_previous_seal(path: Path) -> dict[str, Any]:
    value = read_json(path)
    require(value.get("schema") == PRIOR_SEAL_SCHEMA,
            "previous quality seal schema changed")
    files = value.get("files")
    require(isinstance(files, dict) and files, "previous quality seal file map missing")
    for name, expected_hash in files.items():
        require(isinstance(name, str) and is_sha(expected_hash),
                f"previous quality seal entry malformed: {name!r}")
        file_path = root_path(name, f"previous sealed file {name}")
        # These five rolling indexes are intentionally updated by the 0821
        # report.  Their prior hashes remain evidence in the committed 0820
        # seal; all packet/source payloads and the committed repair blob must
        # still match the retained seal exactly.
        if name not in PRIOR_SEAL_MUTABLE_INDEXES:
            require(sha256(file_path) == expected_hash,
                    f"previous sealed payload changed: {name}")
        else:
            # The five rolling indexes may be updated by the 0821 report.  Use
            # the committed 8312 blob as their prior witness instead of
            # reading the current live index, whose bytes are expected to
            # change after capture.
            try:
                blob = subprocess.check_output(
                    ["git", "show", f"{BASE_REVISION}:{name}"], cwd=ROOT
                )
            except (OSError, subprocess.CalledProcessError) as error:
                fail(f"cannot read committed previous sealed payload {name}: {error}")
            require(hashlib.sha256(blob).hexdigest() == expected_hash,
                    f"committed previous sealed payload changed: {name}")
    return value


def check_repair_command_result(path: Path, label: str) -> dict[str, Any]:
    value = read_json(path)
    require(value.get("schema") == REPAIR_COMMAND_SCHEMA
            and value.get("exit_code") == 0,
            f"{label} result changed")
    finite(value.get("started"), f"{label} start")
    finite(value.get("ended"), f"{label} end")
    require(value["started"] <= value["ended"], f"{label} timestamps reversed")
    log = value.get("log")
    require(isinstance(log, dict), f"{label} log descriptor missing")
    log_path = descriptor(log, f"{label} log")
    require(log_path.is_file(), f"{label} log missing")
    return value


def check_quality_reuse(cv: dict[str, Any]) -> dict[str, Any]:
    """Replay the committed 0820 six-gate result through the fresh adapter.

    The 0821 packet must carry an adapter receipt, but it must not run Cargo
    again.  This routine verifies every retained result/log descriptor and the
    old source census against the live committed 8312 source hashes.
    """
    path = PACKET / "quality.json"
    value = read_json(path)
    require(value.get("schema") == QUALITY_SCHEMA
            and value.get("status") == "pass"
            and value.get("gate_count") == 6
            and value.get("mode") == "committed-receipt-replay",
            "quality reuse adapter changed")
    reuse = cv["origin"]["quality_reuse"]
    plan_reuse = cv["plan"].get("quality_reuse")
    require(plan_reuse == reuse, "quality reuse plan/origin differs")
    adapter_reuse = value.get("quality_reuse")
    require(isinstance(adapter_reuse, dict)
            and adapter_reuse.get("schema") == QUALITY_REUSE_SCHEMA
            and adapter_reuse.get("mode") == "committed-receipt-replay"
            and adapter_reuse.get("cargo_commands_executed") is False
            and isinstance(adapter_reuse.get("source_witness"), dict)
            and adapter_reuse["source_witness"].get("current_revision") == BASE_REVISION
            and adapter_reuse["source_witness"].get("prior_revision") ==
            PRIOR_REPAIR_REVISION
            and adapter_reuse["source_witness"].get("production_files") == PRODUCTION_FILES
            and adapter_reuse["source_witness"].get("tool_files") == TOOL_FILES
            and isinstance(adapter_reuse.get("seal"), dict)
            and adapter_reuse["seal"].get("worktree_checked") is True,
            "quality reuse adapter metadata changed")
    require(value.get("reuse_provenance") == reuse,
            "quality reuse provenance changed")

    quality_path = repair_reference(reuse, adapter_reuse, "quality",
                                    "previous quality receipt")
    tests_path = repair_reference(reuse, adapter_reuse, "test_summary",
                                  "previous test summary")
    inputs_path = repair_reference(reuse, adapter_reuse, "inputs",
                                   "previous repair inputs")
    source_path = repair_reference(reuse, adapter_reuse, "source",
                                   "previous repair source")
    seal_path = repair_reference(reuse, adapter_reuse, "seal",
                                 "previous quality seal")
    source_witness = adapter_reuse["source_witness"]
    witness_source = descriptor(source_witness.get("source"),
                                "quality reuse source witness")
    require(witness_source.resolve() == source_path.resolve(),
            "quality reuse source witness changed")
    seal_meta = adapter_reuse["seal"]
    require(seal_meta.get("base") == BASE_REVISION
            and seal_meta.get("schema") == PRIOR_SEAL_SCHEMA
            and seal_meta.get("files") == 80
            and seal_meta.get("worktree_checked") is True,
            "quality reuse seal metadata changed")

    previous = read_json(quality_path)
    require(previous.get("schema") == REPAIR_QUALITY_SCHEMA
            and previous.get("status") == "pass"
            and previous.get("gate_count") == 6,
            "previous repair quality receipt changed")
    gates = previous.get("gates")
    require(isinstance(gates, list) and len(gates) == 6,
            "previous repair gate cardinality changed")
    gate_names: list[str] = []
    for index, item in enumerate(gates, 1):
        gate_path = descriptor(item, f"previous repair gate {index}")
        result = check_repair_command_result(gate_path, f"previous repair gate {index}")
        require(result.get("name") in {"fmt", "check", "tests", "clippy", "rustdoc", "boundaries"}
                and result.get("name") not in gate_names,
                f"previous repair gate name changed: {index}")
        gate_names.append(result["name"])
    require(gate_names == ["fmt", "check", "tests", "clippy", "rustdoc", "boundaries"],
            "previous repair gate order changed")
    focused = descriptor(previous.get("focused"), "previous focused repair result")
    check_repair_command_result(focused, "previous focused repair")
    adapter_gates = adapter_reuse.get("gates")
    require(isinstance(adapter_gates, list) and len(adapter_gates) == 6,
            "quality reuse gate descriptors changed")
    for index, item in enumerate(adapter_gates, 1):
        require(isinstance(item, dict), f"quality reuse gate {index} malformed")
        result_path = descriptor(item.get("result"),
                                 f"quality reuse gate {index} result")
        log_path = descriptor(item.get("log"), f"quality reuse gate {index} log")
        require(result_path.resolve() == descriptor(gates[index - 1],
                                                    f"previous repair gate {index}").resolve(),
                f"quality reuse gate {index} result changed")
        result = check_repair_command_result(result_path, f"quality reuse gate {index}")
        require(log_path.resolve() == descriptor(result.get("log"),
                                                 f"quality reuse gate {index} result log").resolve(),
                f"quality reuse gate {index} log changed")
    adapter_focused = adapter_reuse.get("focused")
    require(isinstance(adapter_focused, dict), "quality reuse focused descriptors missing")
    require(descriptor(adapter_focused.get("result"), "quality reuse focused result").resolve()
            == focused.resolve(), "quality reuse focused result changed")
    require(descriptor(adapter_focused.get("log"), "quality reuse focused log").resolve()
            == descriptor(check_repair_command_result(focused, "previous focused repair").get("log"),
                           "previous focused repair log").resolve(),
            "quality reuse focused log changed")
    old_source = descriptor(previous.get("source"), "previous repair source")
    require(old_source.resolve() == source_path.resolve(),
            "previous repair source descriptor changed")
    old_inputs = descriptor(previous.get("inputs"), "previous repair inputs")
    require(old_inputs.resolve() == inputs_path.resolve(),
            "previous repair inputs descriptor changed")

    source_value = read_json(source_path)
    require(isinstance(source_value, dict)
            and isinstance(source_value.get("production"), dict)
            and isinstance(source_value.get("tool"), dict),
            "previous repair source malformed")
    require(source_value["production"].get("revision") == PRIOR_REPAIR_REVISION
            and source_value["production"].get("files") == cv["current_source"]["files"]
            and source_value["tool"] == cv["current_tool"],
            "committed repair source hashes changed")
    require(len(source_value["production"]["files"]) == PRODUCTION_FILES
            and len(source_value["tool"]) > 0,
            "committed repair source census changed")

    inputs_value = read_json(inputs_path)
    require(inputs_value.get("schema") == REPAIR_INPUTS_SCHEMA,
            "previous repair inputs schema changed")
    for key in ("runner", "origin", "before", "source", "original_frozen_inputs"):
        descriptor(inputs_value.get(key), f"previous repair input {key}")
    require(descriptor(inputs_value["source"], "previous repair input source").resolve()
            == source_path.resolve(), "previous repair input source changed")
    original_frozen_path = descriptor(inputs_value["original_frozen_inputs"],
                                      "previous original frozen inputs")
    original_frozen = read_json(original_frozen_path)
    require(original_frozen.get("schema") == "litchi.performance.0820.frozen-inputs.v1",
            "previous original frozen-input schema changed")
    require(original_frozen.get("root_inputs") == cv["root_inputs"]
            and original_frozen.get("architecture") == cv["architecture"]
            and original_frozen.get("corpus") == cv["corpus"]
            and original_frozen.get("unrelated") == cv["unrelated"],
            "previous original custody changed")
    previous_host = ROOT / "docs/performance/results/change-0820/host.json"
    require(previous_host.is_file() and not previous_host.is_symlink()
            and original_frozen.get("host") == sha256(previous_host),
            "previous original host witness changed")
    old_locks = original_frozen.get("locks")
    require(isinstance(old_locks, dict), "previous original lock witness missing")
    for lock_name in ("root", "tool", "packet_tool"):
        require(isinstance(old_locks.get(lock_name), dict)
                and old_locks[lock_name].get("bytes") == cv["locks"][lock_name]["bytes"]
                and old_locks[lock_name].get("sha256") == cv["locks"][lock_name]["sha256"],
                f"previous original lock changed: {lock_name}")
    old_packet = original_frozen.get("packet")
    old_drivers = original_frozen.get("drivers")
    require(isinstance(old_packet, dict) and isinstance(old_drivers, dict),
            "previous original packet witness missing")
    previous_packet = ROOT / "docs/performance/results/change-0820"
    for name, expected_hash in {**old_packet, **old_drivers}.items():
        previous_path = previous_packet / name
        require(previous_path.is_file() and not previous_path.is_symlink()
                and sha256(previous_path) == expected_hash,
                f"previous original packet changed: {name}")

    tests = read_json(tests_path)
    require(tests.get("schema") == REPAIR_TESTS_SCHEMA
            and tests.get("suites") == 28
            and tests.get("passed") == 641
            and tests.get("failed") == 0
            and tests.get("ignored") == 1,
            "previous repair test summary changed")
    test_log = descriptor(tests.get("log"), "previous repair test log")
    require(test_log.is_file(), "previous repair test log missing")
    require(descriptor(adapter_reuse.get("test_summary_log"),
                       "quality reuse test summary log").resolve() == test_log.resolve(),
            "quality reuse test summary log changed")
    previous_seal = check_previous_seal(seal_path)
    require(seal_meta.get("bytes") == seal_path.stat().st_size
            and seal_meta.get("sha256") == sha256(seal_path)
            and isinstance(previous_seal.get("files"), dict)
            and len(previous_seal["files"]) == seal_meta.get("files"),
            "quality reuse seal identity changed")

    fresh_source_path = descriptor(value.get("source"), "quality current source",
                                   packet_bound=True)
    fresh_source_value = read_json(fresh_source_path)
    require(fresh_source_value == {
        "production": {"revision": cv["origin"]["base"],
                        "files": cv["current_source"]["files"]},
        "tool": dict(cv["current_tool"]),
    }, "quality current source witness changed")
    checks_path = descriptor(value.get("checks"), "quality reuse checks",
                             packet_bound=True)
    checks = read_json(checks_path)
    require(checks.get("schema") == "litchi.performance.0821.quality-reuse-checks.v1"
            and checks.get("mode") == "committed-receipt-replay"
            and checks.get("cargo_commands_executed") is False,
            "quality reuse checks changed")
    check_rows = checks.get("rows")
    require(isinstance(check_rows, list) and len(check_rows) == 6
            and [row.get("name") for row in check_rows] ==
            ["fmt", "check", "tests", "clippy", "rustdoc", "boundaries"]
            and all(row.get("exit_code") == 0 and row.get("reused") is True
                    for row in check_rows),
            "quality reuse check rows changed")
    for index, row in enumerate(check_rows):
        result_descriptor = adapter_reuse["gates"][index]["result"]
        result_path = descriptor(result_descriptor,
                                  f"quality reuse check row {index + 1} result")
        result = read_json(result_path)
        require(row.get("gate") == index + 1
                and row.get("name") == result.get("name")
                and row.get("command") == result.get("command")
                and row.get("started") == result.get("started")
                and row.get("ended") == result.get("ended")
                and row.get("log") == result.get("log")
                and row.get("reused_result") == result_descriptor,
                f"quality reuse check row {index + 1} changed")
    frozen_path = descriptor(value.get("frozen_inputs"),
                             "quality reuse frozen inputs", packet_bound=True)
    frozen = read_json(frozen_path)
    require(frozen.get("schema") == "litchi.performance.0821.quality-reuse-inputs.v1"
            and frozen.get("quality_reuse") == adapter_reuse,
            "quality reuse frozen inputs changed")

    # A fresh adapter must also bind the current packet custody to the reused
    # source.  The adapter may carry descriptors or the custody objects; the
    # exact live hashes are checked here rather than trusted from metadata.
    for key in ("root_inputs", "locks", "architecture", "corpus", "host", "unrelated"):
        if key in value:
            require(value[key] == cv[key], f"quality reuse {key} witness changed")
    return {"receipt": identity(path), "raw": value,
            "previous_quality": identity(quality_path),
            "previous_tests": identity(tests_path),
            "previous_inputs": identity(inputs_path),
            "previous_source": identity(source_path),
            "previous_seal": identity(seal_path),
            "fresh_source": identity(fresh_source_path),
            "checks": identity(checks_path),
            "frozen_inputs": identity(frozen_path),
            "source": fresh_source_value }


def load_quality(cv: dict[str, Any], build: dict[str, Any]) -> dict[str, Any]:
    value = check_quality_reuse(cv)
    # Build and capture receipts still use the current committed source
    # witness.  The reused 0820 source is byte-identical, but its historical
    # revision label is intentionally retained in the previous packet.
    current = {"production": {"revision": cv["origin"]["base"],
                               "files": cv["current_source"]["files"]},
               "tool": dict(cv["current_tool"])}
    require(build.get("source") == current, "quality current source differs from build source")
    build_quality = build.get("raw", {}).get("quality")
    require(isinstance(build_quality, dict), "build quality reuse witness missing")
    same_descriptor(build_quality, identity(PACKET / "quality.json"),
                    "build quality reuse witness")
    require(build.get("raw", {}).get("quality_reuse") == value["raw"].get("quality_reuse"),
            "build quality reuse witness changed")
    return {"receipt": value["receipt"], "source": current,
            "source_path": None, "attempt": None,
            "rows": [], "reuse": {k: v for k, v in value.items() if k != "raw"},
            "raw": value["raw"]}


def manifest_path(value: Any, directory: Path, label: str) -> Path:
    require(isinstance(value, dict) and isinstance(value.get("path"), str),
            f"{label} descriptor malformed")
    path = Path(value["path"])
    if not path.is_absolute():
        path = directory / path
    path = path.resolve(strict=False)
    require(path.is_relative_to(directory.resolve()), f"{label} escaped export directory")
    descriptor({"path": str(path), "bytes": value.get("bytes"),
                "sha256": value.get("sha256")}, label)
    return path


def load_admission_attempts(admission: dict[str, Any], manifest: Path,
                            audit_path: Path, auditor: Path,
                            directory: Path) -> dict[str, Any]:
    """Retain every independent artifact-audit attempt and its parent error."""
    attempts: list[dict[str, Any]] = []
    for path in sorted(PACKET.glob("admission-*"), key=lambda item: item.name):
        if not path.is_dir() or not path.name.removeprefix("admission-").isdigit():
            continue
        suffix = path.name.removeprefix("admission-")
        receipt_path = path / "receipt.json"
        receipt = read_json(receipt_path)
        require(receipt.get("schema") ==
                "litchi.performance.0821.admission-attempt.v1"
                and receipt.get("exit_code") == 0,
                f"artifact admission attempt {suffix} changed")
        finite(receipt.get("started"), f"artifact admission attempt {suffix} start")
        finite(receipt.get("ended"), f"artifact admission attempt {suffix} end")
        require(receipt["started"] <= receipt["ended"],
                f"artifact admission attempt {suffix} timestamps reversed")
        command = receipt.get("command")
        report_path = descriptor(receipt.get("report"),
                                 f"artifact admission attempt {suffix} report",
                                 packet_bound=True)
        require(report_path is not None, f"artifact admission attempt {suffix} report missing")
        require(isinstance(command, list) and len(command) == 7
                and isinstance(command[0], str)
                and Path(command[0]).name in {"python", "python3"}
                and command[1] == "-B"
                and command[2] == str(PACKET / "artifact_audit.py")
                and command[3] == "--artifacts"
                and command[4] == str(directory)
                and command[5] == "--report"
                and command[6] == str(report_path),
                f"artifact admission attempt {suffix} command changed")
        same_descriptor(receipt.get("manifest"), identity(manifest),
                        f"artifact admission attempt {suffix} manifest")
        same_descriptor(receipt.get("auditor"), identity(auditor),
                        f"artifact admission attempt {suffix} auditor")
        log_path = descriptor(receipt.get("log"),
                              f"artifact admission attempt {suffix} log",
                              packet_bound=True)
        require(log_path is not None, f"artifact admission attempt {suffix} log missing")
        report = read_json(report_path)
        require(report.get("schema") == "litchi.performance.0821.artifact-audit.v1"
                and report.get("ok") is True and report.get("errors") == []
                and len(report.get("cases", [])) == 6
                and all(row.get("ok") is True for row in report["cases"]),
                f"artifact admission attempt {suffix} audit changed")
        require(sha256(report_path) == sha256(audit_path),
                f"artifact admission attempt {suffix} audit differs from accepted audit")
        attempts.append({
            "index": int(suffix),
            "receipt": identity(receipt_path),
            "report": identity(report_path),
            "log": identity(log_path),
            "started": receipt["started"],
            "ended": receipt["ended"],
        })
    require(attempts, "artifact admission attempts missing")

    error_path = PACKET / "execution-errors.json"
    errors = read_json(error_path)
    require(isinstance(errors, list) and errors,
            "artifact admission execution-error witness missing")
    by_index = {row["index"]: row for row in attempts}
    retained: list[str] = []
    for error in errors:
        require(isinstance(error, dict)
                and isinstance(error.get("retained_attempt"), str)
                and error.get("retained_attempt", "").startswith("admission-")
                and error.get("exit_code") == 1
                and error.get("artifact_audit_exit_code") == 0
                and error.get("timing_started") is False,
                "artifact admission execution error changed")
        retained_name = error["retained_attempt"]
        suffix = retained_name.removeprefix("admission-")
        require(suffix.isdigit() and int(suffix) in by_index,
                f"retained artifact admission attempt missing: {retained_name}")
        retained.append(retained_name)

    accepted = [row for row in attempts if row["started"] == admission.get("started")]
    require(len(accepted) == 1, "accepted artifact admission attempt is ambiguous")
    accepted_attempt = accepted[0]
    require(all(name != f"admission-{accepted_attempt['index']}" for name in retained),
            "accepted artifact admission attempt was retained as a failure")
    return {
        "accepted_attempt": accepted_attempt["index"],
        "retained_attempts": retained,
        "attempts": attempts,
        "execution_errors": identity(error_path),
        "errors": errors,
    }


def load_artifacts(cv: dict[str, Any], build: dict[str, Any], quality: dict[str, Any]) -> dict[str, Any]:
    directory = PACKET / "artifacts"
    complete_path = PACKET / "artifacts.complete.json"
    receipt_path = PACKET / "artifacts-receipt.json"
    complete = read_json(complete_path)
    receipt = read_json(receipt_path)
    require(receipt.get("schema") == "litchi.performance.0821.artifact-receipt.v1",
            "artifact receipt schema changed")
    require(complete.get("schema") == "litchi.performance.0821.artifacts.complete.v1"
            and complete.get("cases") == 6 and complete.get("reports") == 0
            and complete.get("samples") == 0, "artifact completion changed")
    manifest = manifest_path(complete.get("manifest"), directory, "artifact manifest")
    same_descriptor(complete.get("receipt"), identity(receipt_path),
                    "artifact completion receipt")
    manifest_value = read_json(manifest)
    cases = manifest_value.get("cases")
    require(isinstance(cases, list) and len(cases) == 6, "artifact case cardinality changed")
    require(sum(x.get("origin") == "generated-harness-corpus" for x in cases) == 3
            and sum(x.get("origin") == "caller-named-real-file" for x in cases) == 3,
            "artifact origins changed")
    by_format: dict[str, dict[str, Any]] = {}
    for case in cases:
        require(isinstance(case, dict) and case.get("format") in {"DOCX", "XLSX", "PPTX"},
                "artifact case malformed")
        outputs = list(case.get("policy_outputs", []))
        stream = case.get("stream_output")
        require(len(outputs) == 4 and isinstance(stream, dict),
                f"artifact policy outputs changed: {case.get('case_id')}")
        source = manifest_path(case.get("source_archive"), directory, "artifact source")
        output_descriptors = [x.get("output") for x in outputs] + [stream.get("output")]
        output_paths = [manifest_path(x, directory, "artifact output") for x in output_descriptors]
        digests = [x.get("sha256") for x in output_descriptors]
        require(len(set(digests)) == 1 and all(is_sha(x) for x in digests),
                f"policy output equality changed: {case.get('case_id')}")
        require({item.get("policy") for item in outputs} == set(POLICIES)
                and stream.get("policy") == "stream"
                and all(item.get("matches_reference") is True
                        and item.get("source_unchanged") is True
                        and item.get("reopen_admitted") is True
                        and item.get("publication") ==
                        ("save" if item.get("policy") == "default"
                         else "save_with_durability") for item in outputs),
                f"artifact policy admission changed: {case.get('case_id')}")
        require(case.get("source_archive_sha256") == sha256(source)
                and case.get("source_archive_bytes") == source.stat().st_size,
                f"artifact source identity changed: {case.get('case_id')}")
        if case.get("origin") == "caller-named-real-file":
            fmt = case["format"].lower()
            input_name = next(x["input"] for x in cv["plan"]["cases"] if x["format"] == fmt)
            require(case["source_archive_sha256"] == cv["corpus"][input_name]["sha256"],
                    f"real corpus source changed: {fmt}")
            require(case.get("edit_admitted") is True and case.get("edit_outcome") == "admitted",
                    f"real corpus admission changed: {fmt}")
            by_format[case["format"]] = {
                "manifest": case, "source": source,
                "output": output_descriptors[0],
                "output_paths": output_paths,
                "policy_digests": {
                    item["policy"]: item["output"]["sha256"]
                    for item in outputs
                },
            }
    require(set(by_format) == {"DOCX", "XLSX", "PPTX"}, "real artifact formats incomplete")
    admission_path = PACKET / "artifact-admission.json"
    admission = read_json(admission_path)
    require(admission.get("schema") == "litchi.performance.0821.artifact-admission.v1"
            and admission.get("accepted") is True
            and admission.get("plan_sha256") == sha256(PACKET / "plan.json"),
            "artifact admission changed")
    same_descriptor(admission.get("artifact_complete"), identity(complete_path),
                    "artifact admission completion")
    same_descriptor(admission.get("manifest"), identity(manifest), "artifact admission manifest")
    audit_path = same_descriptor(admission.get("audit"),
                                 identity(resolve_path(admission["audit"]["path"], packet_bound=True)),
                                 "artifact admission audit")
    audit = read_json(audit_path)
    require(audit.get("schema") == "litchi.performance.0821.artifact-audit.v1"
            and audit.get("ok") is True and audit.get("errors") == []
            and len(audit.get("cases", [])) == 6
            and all(row.get("ok") is True for row in audit["cases"]),
            "independent artifact audit changed")
    auditor_path = same_descriptor(admission.get("auditor"),
                                   identity(PACKET / "artifact_audit.py"),
                                   "artifact auditor")
    attempts = load_admission_attempts(admission, manifest, audit_path,
                                       auditor_path, directory)
    zip_desc = admission.get("zip_preservation")
    zip_path = same_descriptor(zip_desc, identity(PACKET / "zip-preservation.json"),
                               "ZIP preservation witness")
    zip_value = read_json(zip_path)
    require(zip_value.get("schema") == "litchi.performance.0821.zip-preservation.v1"
            and len(zip_value.get("cases", [])) == 6,
            "ZIP preservation witness changed")
    audit_by_key = {(x.get("origin"), x.get("format")): x for x in audit["cases"]}
    zip_by_id = {x.get("case_id"): x for x in zip_value["cases"]}
    for fmt, row in by_format.items():
        key = ("caller-named-real-file", fmt)
        audited = audit_by_key.get(key)
        require(isinstance(audited, dict), f"audit real case missing: {fmt}")
        audit_outputs = audited.get("policy_outputs")
        require(isinstance(audit_outputs, list) and len(audit_outputs) == 5,
                f"audit policy cardinality changed: {fmt}")
        audit_by_policy = {x.get("policy"): x for x in audit_outputs
                           if isinstance(x, dict)}
        expected_policies = {"default", "full", "file-only", "no-sync", "stream"}
        audit_digests = [x.get("sha256") for x in audit_outputs
                         if isinstance(x, dict)]
        require(set(audit_by_policy) == expected_policies
                and len(audit_digests) == 5
                and all(is_sha(x) for x in audit_digests)
                and len(set(audit_digests)) == 1
                and audit_digests[0] == row["output"]["sha256"]
                and audited.get("source_equals_output") is False,
                f"audit output identity changed: {fmt}")
        require(all(audit_by_policy[policy].get("sha256") ==
                    row["manifest"]["published_sha256"]
                    for policy in expected_policies),
                f"audit policy digest identity changed: {fmt}")
        zipped = zip_by_id.get(row["manifest"].get("case_id"))
        require(isinstance(zipped, dict) and zipped.get("source_sha256") == row["manifest"]["source_archive_sha256"]
                and zipped.get("output_sha256") == row["output"]["sha256"]
                and zipped.get("member_order_equal") is True
                and zipped.get("archive_comment_equal") is True,
                f"ZIP preservation identity changed: {fmt}")
        require(isinstance(zipped.get("untouched_members"), list),
                f"ZIP untouched member witness missing: {fmt}")
        if fmt == "DOCX":
            require(any(x.get("name") == "word/_rels/document.xml.rels"
                        for x in zipped["untouched_members"]),
                    "DOCX untouched relationship witness missing")
    selectors = admission.get("selectors")
    require(isinstance(selectors, list) and len(selectors) == len(cv["plan"]["cases"]),
            "artifact admission selector cardinality changed")
    expected = {x["id"]: x for x in cv["plan"]["cases"]}
    observed = set()
    for selector in selectors:
        require(selector.get("id") in expected and selector["id"] not in observed,
                "artifact admission selector changed")
        case = expected[selector["id"]]
        require(selector.get("format") == case["format"]
                and selector.get("phase") == case["phase"]
                and selector.get("input") == case["input"]
                and selector.get("case") == case["case"]
                and selector.get("policy") == case["policy"],
                f"artifact selector fields changed: {selector.get('id')}")
        real = by_format[case["format"].upper()]["manifest"]
        require(selector.get("source_sha256") == real["source_archive_sha256"]
                and selector.get("source_bytes") == real["source_archive_bytes"]
                and selector.get("published_sha256") == real["published_sha256"]
                and selector.get("policy_output_sha256") ==
                by_format[case["format"].upper()]["policy_digests"][case["policy"]]
                and selector.get("policy_output_bytes") ==
                next(item["output"]["bytes"] for item in
                     by_format[case["format"].upper()]["manifest"]["policy_outputs"]
                     if item["policy"] == case["policy"])
                and selector.get("policy_output_sha256") == selector.get("published_sha256")
                and selector.get("policy_publication") ==
                ("save" if case["policy"] == "default" else "save_with_durability")
                and is_sha(selector.get("source_sha256"))
                and is_sha(selector.get("published_sha256"))
                and selector.get("edit_outcome") == "admitted",
                f"artifact selector identity changed: {selector.get('id')}")
        observed.add(selector["id"])
    require(observed == set(expected), "artifact selector set changed")
    return {"complete": identity(complete_path), "receipt": identity(receipt_path),
            "manifest": identity(manifest), "audit": identity(audit_path),
            "zip_preservation": identity(zip_path), "cases": 6,
            "policy_outputs": 30, "audit_cases": 6, "audit_ok": True,
            "admission": {"receipt": identity(admission_path), "selectors": selectors,
                           "audit": audit, "zip": zip_value,
                           "attempts": attempts}}


def load_qualification_admission(artifacts: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "qualification-admission.json"
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0821.qualification-admission.v1"
            and value.get("accepted") is True
            and value.get("plan_sha256") == sha256(PACKET / "plan.json")
            and value.get("reports") == 24 and value.get("samples") == 24,
            "qualification admission changed")
    same_descriptor(value.get("artifact_admission"),
                    artifacts["admission"]["receipt"], "qualification artifact admission")
    complete_path = PACKET / "qualification" / "complete.json"
    same_descriptor(value.get("qualification_complete"), identity(complete_path),
                    "qualification completion")
    return {"receipt": identity(path), "reports": 24, "samples": 24,
            "artifact_admission": artifacts["admission"]["receipt"],
            "qualification_complete": identity(complete_path)}


def nearest_rank(values: Iterable[float], quantile: float) -> float:
    ordered = sorted(values)
    require(ordered, "nearest rank received no values")
    index = max(0, min(len(ordered) - 1, math.ceil(quantile * len(ordered)) - 1))
    return ordered[index]


def integer_midpoint(left: int, right: int) -> int:
    require(isinstance(left, int) and isinstance(right, int), "midpoint inputs are not integers")
    return (left + right) // 2


def rust_welford_mean(values: Iterable[int | float]) -> float:
    """Match the sorted-sample Welford mean used by the Rust harness."""
    mean = 0.0
    for index, value in enumerate(values):
        next_count = float(index + 1)
        delta = float(value) - mean
        mean += delta / next_count
    return mean


def workload(source_bytes: int, published_bytes: int, p50_ns: float, phase: str) -> dict[str, Any]:
    positive_int(source_bytes, "source logical bytes")
    positive_int(published_bytes, "published logical bytes")
    finite(p50_ns, "workload p50")
    require(p50_ns > 0, "workload p50 is not positive")
    factor = 1_000_000_000.0 / float(p50_ns)
    return {
        "source_logical_bytes": source_bytes,
        "published_logical_bytes": published_bytes,
        "source_logical_bytes_per_second": source_bytes * factor,
        "published_logical_bytes_per_second": published_bytes * factor,
        "rate_basis": "phase elapsed p50; archive logical byte counts from this phase summary",
        "phase": phase, "publication_timed": phase != "edit",
        "claim": "descriptive logical workload rate only; no physical-I/O or memory-bandwidth claim",
    }


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(len(values) == 6, "absolute bootstrap requires six process blocks")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates = [statistics.median(rng.choice(values) for _ in values)
                 for _ in range(BOOTSTRAP_RESAMPLES)]
    ordered = sorted(estimates)
    return {"estimate": statistics.median(values), "lower": ordered[BOOTSTRAP_LOW_RANK],
            "upper": ordered[BOOTSTRAP_HIGH_RANK], "seed": BOOTSTRAP_SEED,
            "resamples": BOOTSTRAP_RESAMPLES, "confidence": BOOTSTRAP_CONFIDENCE,
            "low_rank": BOOTSTRAP_LOW_RANK, "high_rank": BOOTSTRAP_HIGH_RANK,
            "statistic": "median of six process p50 values", "units": "ns"}


def paired_bootstrap(values: list[float]) -> dict[str, Any]:
    """Bootstrap the six within-block policy/default p50 ratios."""
    require(len(values) == 6, "paired bootstrap requires six process blocks")
    require(all(float(value) > 0 and math.isfinite(float(value)) for value in values),
            "paired bootstrap received a non-positive ratio")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates = [statistics.median(rng.choice(values) for _ in values)
                 for _ in range(BOOTSTRAP_RESAMPLES)]
    ordered = sorted(estimates)
    return {"estimate": statistics.median(values), "lower": ordered[BOOTSTRAP_LOW_RANK],
            "upper": ordered[BOOTSTRAP_HIGH_RANK], "seed": BOOTSTRAP_SEED,
            "resamples": BOOTSTRAP_RESAMPLES, "confidence": BOOTSTRAP_CONFIDENCE,
            "low_rank": BOOTSTRAP_LOW_RANK, "high_rank": BOOTSTRAP_HIGH_RANK,
            "statistic": "median of six matched process-block policy/default p50 ratios",
            "units": "ratio"}


def validate_delta(value: Any, label: str) -> None:
    # Procfs can legitimately report an unavailable sample on a shared host;
    # retain the null exactly and never replace it with a subtraction or zero.
    if value is None:
        return
    require(isinstance(value, dict), f"{label} is malformed")
    for key, item in value.items():
        if key in {"status", "scope", "phase", "timing_scope", "latency_claim",
                   "control_scope", "alignment"}:
            continue
        if isinstance(item, (int, float)):
            require(item >= 0 and math.isfinite(float(item)), f"{label}.{key} changed")


def validate_metrics(result: dict[str, Any], count: int, lane: str,
                     sample_order: list[int], label: str) -> dict[str, Any]:
    metrics = result.get("operation_metrics")
    require(isinstance(metrics, dict)
            and metrics.get("sample_count") == count
            and metrics.get("sample_indices") == sample_order
            and metrics.get("alignment") ==
            "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{label} operation metric alignment changed")
    allocation, process = metrics.get("allocation"), metrics.get("process")
    if lane == "native":
        require(isinstance(allocation, dict) and allocation.get("status") == "unavailable"
                and isinstance(process, dict) and process.get("status") == "unavailable"
                and metrics.get("latency_claim") == "comparable_timed_operation",
                f"{label} native instrumentation changed")
    else:
        require(isinstance(allocation, dict) and allocation.get("status") == "measured"
                and isinstance(process, dict) and process.get("status") == "measured"
                and metrics.get("latency_claim") ==
                "allocator_instrumented_elapsed_not_latency_claim",
                f"{label} observer instrumentation changed")
        for group_name, group in (("allocation", allocation), ("process", process)):
            for key, item in group.items():
                if isinstance(item, dict) and "values" in item:
                    require(isinstance(item["values"], list) and len(item["values"]) == count,
                            f"{label}.{group_name}.{key} vector cardinality changed")
    return json.loads(json.dumps(metrics, sort_keys=True))


def validate_report(path: Path, receipt: dict[str, Any], case: dict[str, Any],
                    lane: str, cv: dict[str, Any], build: dict[str, Any],
                    selector: dict[str, Any]) -> dict[str, Any]:
    report = read_json(path)
    label = f"{lane}/{case['case']}/block{receipt['block']}"
    require(report.get("schema_version") == REPORT_SCHEMA_VERSION,
            f"{label} report schema changed")
    binary_name = cv["plan"]["lanes"][lane]["binary"]
    binary = build["binaries"][binary_name]["artifact"]
    tool = report.get("tool")
    expected_binary = "litchi-perf-baseline" if lane == "native" else "litchi-perf-baseline-alloc"
    require(isinstance(tool, dict) and tool.get("binary") == expected_binary
            and tool.get("profile") == "release", f"{label} tool identity changed")
    expected_instrumentation = "none" if lane == "native" else \
        "ordinary_save_procfs_and_system_allocator_operation_scoped"
    require(tool.get("instrumentation") == expected_instrumentation,
            f"{label} instrumentation identity changed")
    identity_value = report.get("binary_identity")
    require(isinstance(identity_value, dict)
            and identity_value.get("binary_sha256") == binary["sha256"]
            and identity_value.get("binary_bytes") == binary["bytes"]
            and identity_value.get("executable") is True
            and identity_value.get("profile") == "release",
            f"{label} binary identity changed")
    environment = report.get("environment")
    require(isinstance(environment, dict)
            and environment.get("git_revision") == cv["origin"]["base"]
            and environment.get("cpu_affinity") == "12-19",
            f"{label} environment changed")
    lane_plan = cv["plan"]["lanes"][lane]
    configuration = report.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("samples_per_case") == lane_plan["samples"]
            and configuration.get("warmup_iterations_per_case") == lane_plan["warmup"]
            and configuration.get("cases") == [case["case"]]
            and configuration.get("filesystem_cache_states") == ["warm", "cold-requested"]
            and configuration.get("filesystem_fresh_child_per_sample") is True
            and configuration.get("filesystem_process_isolated") is True
            and configuration.get("filesystem_root_selected") is True,
            f"{label} configuration changed")
    result_list = report.get("results")
    require(isinstance(result_list, list) and len(result_list) == 1,
            f"{label} result cardinality changed")
    result = result_list[0]
    elapsed = result.get("elapsed_ns")
    count = lane_plan["samples"]
    require(result.get("case") == case["case"] and isinstance(elapsed, dict)
            and elapsed.get("unit") == "ns", f"{label} selector changed")
    samples = elapsed.get("samples")
    order = elapsed.get("sample_order")
    require(isinstance(samples, list) and len(samples) == count
            and all(isinstance(x, int) and x > 0 for x in samples)
            and samples == sorted(samples)
            and isinstance(order, list) and sorted(order) == list(range(count)),
            f"{label} elapsed samples changed")
    require(elapsed.get("min") == samples[0]
            and elapsed.get("max") == samples[-1]
            and elapsed.get("p50") == integer_midpoint(samples[(count - 1) // 2], samples[count // 2])
            and elapsed.get("p95") == nearest_rank(samples, .95)
            and elapsed.get("p99") == nearest_rank(samples, .99),
            f"{label} raw quantiles changed")
    finite(elapsed.get("mean"), f"{label} mean")
    require(abs(float(elapsed["mean"]) - rust_welford_mean(samples)) < 1e-12,
            f"{label} mean changed")
    ordinary = (result.get("source") or {}).get("ordinary_save")
    require(isinstance(ordinary, dict)
            and ordinary.get("format") == case["format"].upper()
            and ordinary.get("origin") == "caller-named-real-file",
            f"{label} ordinary-save evidence changed")
    phase_name = {"lifecycle": "open+edit+save",
                  "atomic_publish": "save-to-path"}[case["phase"]]
    require(ordinary.get("phase") == phase_name, f"{label} phase changed")
    timing_scope = ordinary.get("timing_scope")
    require(timing_scope == GENERIC_TIMING_SCOPES[case["phase"]],
            f"{label} generic timing boundary changed")
    policy = case["policy"]
    if policy == "default":
        require("save_durability" not in ordinary,
                f"{label} default durability unexpectedly explicit")
    else:
        require(ordinary.get("save_durability") == policy,
                f"{label} durability route changed")
    atomic_steps = ordinary.get("atomic_publication_steps")
    require(atomic_steps == ATOMIC_STEPS_BY_POLICY[policy],
            f"{label} atomic publication boundary changed")
    summary = ordinary.get("corpus")
    require(isinstance(summary, dict), f"{label} corpus evidence missing")
    real = summary.get("real_file") if isinstance(summary, dict) else None
    expected_input = cv["corpus"][case["input"]]
    require(isinstance(real, dict) and real.get("bytes") == expected_input["bytes"]
            and real.get("sha256") == expected_input["sha256"],
            f"{label} real input identity changed")
    require(summary.get("source_archive_bytes") == expected_input["bytes"]
            and summary.get("source_archive_sha256") == expected_input["sha256"]
            and summary.get("published_bytes") > 0
            and summary.get("published_sha256") == selector["published_sha256"]
            and summary.get("save_entry_point") == SAVE_ENTRY_POINTS[case["format"].upper()]
            and summary.get("sink_entry_point") == SINK_ENTRY_POINTS[case["format"].upper()]
            and summary.get("edit_admitted") is True
            and summary.get("edit_outcome") == selector["edit_outcome"] == "admitted"
            and summary.get("repeated_cycles_identical") is True
            and summary.get("repeated_saves_identical") is True
            and is_sha(summary.get("repeated_cycle_sha256"))
            and summary.get("repeated_cycle_sha256") == selector["published_sha256"]
            and summary.get("repeated_save_sha256") == selector["published_sha256"],
            f"{label} admitted corpus outcome changed")
    require(ordinary.get("publications_identical") is True
            and ordinary.get("edit_outcomes_identical") is True,
            f"{label} determinism changed")
    # SaveEvidence carries the one untimed reference digest; the enclosing
    # OrdinarySaveSummary carries the per-sample publication vector.
    published = ordinary.get("published_sha256")
    outcome_hashes = ordinary.get("edit_outcome_sha256")
    require(isinstance(outcome_hashes, list) and len(outcome_hashes) == count
            and all(x == EDIT_OUTCOME_SHA for x in outcome_hashes),
            f"{label} edit outcome vector changed")
    require(isinstance(published, list) and len(published) == count
            and all(x == selector["published_sha256"] for x in published)
            and selector.get("policy_output_sha256") == selector["published_sha256"]
            and result.get("output_sha256") == selector["published_sha256"],
            f"{label} publication identity changed")
    split = summary.get("byte_split")
    require(isinstance(split, dict)
            and split.get("accounting_scope") == BYTE_SPLIT_SCOPE,
            f"{label} byte split scope changed")
    for key in ("output_total_bytes", "payload_bytes_deflated", "payload_bytes_stored",
                "payload_bytes_identical_to_source", "payload_bytes_regenerated",
                "uncompressed_payload_bytes_compressed", "uncompressed_payload_bytes_regenerated",
                "framing_bytes", "output_member_count", "deflate_member_count",
                "stored_member_count", "members_identical_to_source", "members_regenerated"):
        nonnegative_int(split.get(key), f"{label}.byte_split.{key}")
    require(split["output_total_bytes"] == summary["published_bytes"],
            f"{label} output bytes changed")
    sample_split = ordinary.get("sample_byte_split")
    require(sample_split is None, f"{label} non-counting byte split present")
    metrics = validate_metrics(result, count, lane, order, label)
    probe = ordinary.get("process_probe")
    if lane == "native":
        require(probe is None, f"{label} native procfs probe present")
    else:
        require(isinstance(probe, dict)
                and probe.get("phase") == ordinary.get("phase")
                and probe.get("timing_scope") == timing_scope
                and probe.get("fixed_count") == OBSERVER_CONTROLS
                and isinstance(probe.get("empty_adjacent_snapshot_controls"), list)
                and len(probe["empty_adjacent_snapshot_controls"]) == OBSERVER_CONTROLS
                and isinstance(probe.get("sample_deltas"), list)
                and len(probe["sample_deltas"]) == count
                and "never_subtracted" in str(probe.get("control_scope", "")),
                f"{label} procfs probe evidence changed")
        for index, item in enumerate(probe["empty_adjacent_snapshot_controls"]):
            validate_delta(item, f"{label}.control[{index}]")
        for index, item in enumerate(probe["sample_deltas"]):
            validate_delta(item, f"{label}.sample_delta[{index}]")
    return {
        "lane": lane, "block": receipt["block"], "order": receipt["order"],
        "id": case["id"], "case": case["case"], "format": case["format"],
        "phase": case["phase"], "policy": policy,
        "input": case["input"], "samples": list(samples), "sample_order": list(order),
        "reported": {k: elapsed[k] for k in ("min", "p50", "p95", "p99", "max", "mean")},
        "nearest_rank": {"p50_ns": nearest_rank(samples, .50),
                          "p95_ns": nearest_rank(samples, .95),
                          "p99_ns": nearest_rank(samples, .99),
                          "mean_ns": rust_welford_mean(samples)},
        "rss_kib": read_rss(receipt, label),
        "source_sha256": expected_input["sha256"],
        "source_logical_bytes": summary["source_archive_bytes"],
        "published_logical_bytes": summary.get("published_bytes", expected_input["bytes"]),
        "published_sha256": selector["published_sha256"],
        "edit_outcome": "admitted", "workload": None,
        "operation_metrics": metrics,
        "process_probe": json.loads(json.dumps(probe, sort_keys=True)) if probe else None,
        # Keep the complete report payload so allocation/process/procfs
        # diagnostics remain auditable after the compact summaries are built.
        "raw_report": json.loads(json.dumps(report, sort_keys=True)),
        "report": identity(path),
    }


def read_rss(receipt: dict[str, Any], label: str) -> int:
    value = receipt.get("rss")
    path = descriptor(value, f"{label} RSS", packet_bound=True)
    require(path is not None, f"{label} RSS missing")
    text = path.read_text(encoding="utf-8").strip()
    require(text.isdigit() and int(text) > 0, f"{label} RSS is invalid")
    return int(text)


def expected_rows(p: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    rows = []
    for block, order in enumerate(p["lanes"][lane]["orders"]):
        ordered = custody.ordered_cases(p, lane, block)
        rows.extend({"lane": lane, "block": block, "order": order, **case} for case in ordered)
    return rows


def load_lane(lane: str, cv: dict[str, Any], build: dict[str, Any],
              selectors: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    directory = PACKET / lane
    receipts_path = directory / "receipts.json"
    receipts = read_json(receipts_path)
    wanted = expected_rows(cv["plan"], lane)
    require(isinstance(receipts, list) and len(receipts) == len(wanted),
            f"{lane} receipt cardinality changed")
    binary_name = cv["plan"]["lanes"][lane]["binary"]
    binary = build["binaries"][binary_name]["artifact"]
    result = []
    for receipt, expected in zip(receipts, wanted):
        label = f"{lane}/{expected['case']}/block{expected['block']}"
        require(receipt.get("schema") == CAPTURE_SCHEMA and receipt.get("exit_code") == 0,
                f"{label} receipt changed")
        for key, value in expected.items():
            require(receipt.get(key) == value, f"{label} {key} changed")
        require(receipt.get("binary") == binary, f"{label} binary receipt changed")
        finite(receipt.get("started"), f"{label} start")
        finite(receipt.get("ended"), f"{label} end")
        require(receipt["started"] <= receipt["ended"], f"{label} timestamps reversed")
        source_path = descriptor(receipt.get("source"), f"{label} source", packet_bound=True)
        require(source_path is not None and read_json(source_path) == build["source"],
                f"{label} source witness changed")
        marker = receipt.get("scratch_marker")
        marker_path = descriptor(marker, f"{label} scratch marker", allow_missing=True)
        if marker_path is None:
            cleanup = cleanup_value()
            require(isinstance(cleanup, dict) and cleanup.get("scratch_removed") is True,
                    f"{label} missing scratch marker lacks cleanup witness")
        else:
            require(marker_path.resolve() ==
                    (custody.SCRATCH / ".litchi-performance-0821-owned").resolve(),
                    f"{label} scratch marker changed")
        for key in ("root_inputs", "locks", "architecture", "corpus", "host",
                    "packet", "drivers", "unrelated"):
            expected_value = cv["host"] if key == "host" else cv[key]
            require(receipt.get(key) == expected_value, f"{label} custody changed")
        report_path = descriptor(receipt.get("report"), f"{label} report", packet_bound=True)
        log_path = descriptor(receipt.get("log"), f"{label} log", packet_bound=True)
        require(report_path is not None and log_path is not None, f"{label} raw evidence missing")
        expected_command = [
            "/usr/bin/time", "-f", "%M", "-o",
            str(directory / f"{receipt['block']:02d}-{expected['id']}.rss"),
            "taskset", "-c", ",".join(str(cpu) for cpu in cv["plan"]["affinity"]),
            binary["path"], "--warmup", str(cv["plan"]["lanes"][lane]["warmup"]),
            "--samples", str(cv["plan"]["lanes"][lane]["samples"]),
            "--case", expected["case"], "--json",
            str(directory / f"{receipt['block']:02d}-{expected['id']}.json"),
            "--filesystem-root", str(custody.SCRATCH), "--ooxml-file",
            str(ROOT / expected["input"]),
        ]
        if expected["policy"] != "default":
            expected_command.extend(["--save-durability", expected["policy"]])
        require(receipt.get("command") == expected_command, f"{label} command changed")
        selector = selectors[expected["id"]]
        parsed = validate_report(report_path, receipt, expected, lane, cv, build, selector)
        parsed["log"] = relative(log_path)
        result.append(parsed)
    complete = read_json(directory / "complete.json")
    require(complete.get("schema") == f"litchi.performance.0821.{lane}.complete.v1"
            and complete.get("blocks") == cv["plan"]["lanes"][lane]["blocks"]
            and complete.get("children") == len(wanted)
            and complete.get("reports") == len(wanted)
            and complete.get("samples") == len(wanted) * cv["plan"]["lanes"][lane]["samples"],
            f"{lane} completion changed")
    require(complete.get("plan_sha256") == sha256(PACKET / "plan.json")
            and complete.get("build_sha256") == sha256(PACKET / "build.json")
            and complete.get("quality_sha256") == sha256(PACKET / "quality.json"),
            f"{lane} completion provenance changed")
    complete_source = descriptor(complete.get("source"), f"{lane} completion source",
                                 packet_bound=True)
    require(complete_source is not None and read_json(complete_source) == build["source"],
            f"{lane} completion source changed")
    same_descriptor(complete.get("receipts"), identity(receipts_path),
                    f"{lane} completion receipts")
    return result


def native_summary(rows: list[dict[str, Any]], p: dict[str, Any]) -> list[dict[str, Any]]:
    result = []
    require(len(rows) == p["expected"]["native_reports"],
            "native report cardinality changed")
    for case in p["cases"]:
        selected = sorted((x for x in rows if x["id"] == case["id"]),
                          key=lambda x: x["block"])
        require([x["block"] for x in selected] == list(range(6)),
                f"native block coverage changed: {case['id']}")
        defaults = sorted((x for x in rows
                           if x["case"] == case["case"] and x["policy"] == "default"),
                          key=lambda x: x["block"])
        require(len(defaults) == 6 and [x["block"] for x in defaults] == list(range(6)),
                f"native default control coverage changed: {case['case']}")
        p50 = [x["nearest_rank"]["p50_ns"] for x in selected]
        p95 = [x["nearest_rank"]["p95_ns"] for x in selected]
        p99 = [x["nearest_rank"]["p99_ns"] for x in selected]
        means = [x["nearest_rank"]["mean_ns"] for x in selected]
        source = {x["source_logical_bytes"] for x in selected}
        published = {x["published_logical_bytes"] for x in selected}
        require(len(source) == len(published) == 1, f"{case['case']} byte identity changed")
        spread = {name: max(values) / min(values) for name, values in
                  (("p50", p50), ("p95", p95), ("p99", p99), ("mean", means))}
        spread = {name: value for name, value in spread.items()}
        tail = statistics.median(p99) / statistics.median(p50)
        median_p50 = statistics.median(p50)
        default_p50 = [x["nearest_rank"]["p50_ns"] for x in defaults]
        ratios = [numerator / denominator
                  for numerator, denominator in zip(p50, default_p50)]
        ratio_bootstrap = paired_bootstrap(ratios)
        blocks = []
        for row in selected:
            blocks.append({"block": row["block"], "order": row["order"],
                           "samples": row["samples"], "sample_order": row["sample_order"],
                           "p50_ns": row["nearest_rank"]["p50_ns"],
                           "p95_ns": row["nearest_rank"]["p95_ns"],
                           "p99_ns": row["nearest_rank"]["p99_ns"],
                           "mean_ns": row["nearest_rank"]["mean_ns"],
                           "raw_reported": row["reported"], "rss_kib": row["rss_kib"],
                           "raw_report": row["raw_report"],
                           "report": row["report"],
                           "quantile_definition": "nearest rank within this process block",
                           "raw_report_quantile_definition":
                           "harness p50 integer midpoint; p95/p99 nearest rank"})
        item = {"id": case["id"], "case": case["case"], "format": case["format"],
                "phase": case["phase"], "policy": case["policy"],
                "input": case["input"], "blocks": 6, "samples_per_report": 30,
                "p50_ns": median_p50, "p95_ns": statistics.median(p95),
                "p99_ns": statistics.median(p99), "mean_ns": statistics.median(means),
                "p50_block_values_ns": p50, "p95_block_values_ns": p95,
                "p99_block_values_ns": p99, "mean_block_values_ns": means,
                "spread_ratios": spread,
                "spread_flags": [name for name, value in spread.items() if value > 1.05],
                "spread_flag": any(value > 1.05 for value in spread.values()),
                "p99_to_p50_ratio": tail, "tail_flag": tail > 1.05,
                "p50_bootstrap_ci95_ns": bootstrap(p50), "block_raw_quantiles": blocks,
                "quantile_definition":
                "nearest rank within each process; median over six process blocks",
                "raw_report_quantile_definition":
                "harness p50 integer midpoint; p95/p99 nearest rank",
                "workload": workload(next(iter(source)), next(iter(published)), median_p50, case["phase"]),
                "timing_scope": "native elapsed evidence for one real-file save-durability selector",
                "historical_timing_comparison": False,
                "policy_semantics": POLICY_SEMANTICS[case["policy"]],
                "paired_ratio_to_default": {
                    "control_policy": "default",
                    "control_id": f"{case['case']}__default",
                    "definition": "matched process-block policy p50 / default p50 for the same format and phase",
                    "block_p50_ratios": ratios,
                    "ratio_median": ratio_bootstrap["estimate"],
                    "bootstrap_ci95": ratio_bootstrap,
                    "configuration_attribution_only": True,
                    "default_full_control": case["policy"] in {"default", "full"},
                    "weaker_policy_semantics_changed": case["policy"] in {"file-only", "no-sync"},
                },
                "generic_timing_scope_note":
                "The producer's Phase::timing_scope string describes the full durability boundary for every policy; policy-specific synchronization is bound by save_durability and atomic_publication_steps."}
        result.append(item)
    return result


def csv_text(rows: list[dict[str, Any]]) -> str:
    fields = ["id", "case", "format", "phase", "policy", "input", "p50_ns", "p95_ns", "p99_ns",
              "mean_ns", "p50_bootstrap_ci95_ns", "spread_ratios", "spread_flags",
              "spread_flag", "p99_to_p50_ratio", "tail_flag",
              "paired_ratio_to_default",
              "historical_timing_comparison", "source_logical_bytes",
              "published_logical_bytes", "source_logical_bytes_per_second",
              "published_logical_bytes_per_second"]
    stream = io.StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=fields, lineterminator="\n")
    writer.writeheader()
    for row in rows:
        item = {key: row[key] for key in fields if key not in {
            "source_logical_bytes", "published_logical_bytes",
            "source_logical_bytes_per_second", "published_logical_bytes_per_second"}}
        for key in ("p50_bootstrap_ci95_ns", "spread_ratios", "paired_ratio_to_default"):
            item[key] = json.dumps(row[key], sort_keys=True, separators=(",", ":"))
        item.update({"source_logical_bytes": row["workload"]["source_logical_bytes"],
                     "published_logical_bytes": row["workload"]["published_logical_bytes"],
                     "source_logical_bytes_per_second": row["workload"]["source_logical_bytes_per_second"],
                     "published_logical_bytes_per_second": row["workload"]["published_logical_bytes_per_second"]})
        writer.writerow(item)
    return stream.getvalue()


def build_analysis() -> dict[str, Any]:
    cv = load_custody()
    build = load_build(cv)
    quality = load_quality(cv, build)
    artifacts = load_artifacts(cv, build, quality)
    qualification_admission = load_qualification_admission(artifacts)
    selectors = {x["id"]: x for x in artifacts["admission"]["selectors"]}
    qualification = load_lane("qualification", cv, build, selectors)
    native = load_lane("native", cv, build, selectors)
    observer = load_lane("observer", cv, build, selectors)
    native_rows = native_summary(native, cv["plan"])
    for row in qualification + observer:
        row["workload"] = workload(row["source_logical_bytes"], row["published_logical_bytes"],
                                    row["reported"]["p50"], row["phase"])
    counts = cv["plan"]["expected"]
    admissions = dict(artifacts["admission"])
    admissions["qualification"] = qualification_admission
    return {
        "schema": ANALYSIS_SCHEMA, "plan_schema": PLAN_SCHEMA,
        "report_schema_version": REPORT_SCHEMA_VERSION, "scope": cv["plan"]["scope"],
        "regression_policy": cv["plan"]["regression_policy"],
        "historical_timing_comparison": False, "status": "accepted",
        "timing_status": {"accepted": True, "reports": 216, "samples": 4488,
                           "lanes": {"qualification": {"reports": 24, "samples": 24},
                                     "native": {"reports": 144, "samples": 4320},
                                     "observer": {"reports": 48, "samples": 144}}},
        "custody": {"base": cv["origin"]["base"], "production_file_count": PRODUCTION_FILES,
                    "architecture_file_count": ARCHITECTURE_FILES,
                    "production_matches_base": True, "harness_matches_base": True,
                    "runtime_harness_matches_base": True, "tool_file_count": TOOL_FILES,
                    "tool_matches_repair_source": True, "tool_allowlist": [],
                    "corpus_inputs": cv["corpus"], "provenance": identity(PACKET / "provenance.json"),
                    "unrelated": cv["unrelated"]},
        # Cleanup is intentionally validated as a live/removal witness but is
        # omitted from derived analysis so pre- and post-cleanup replay bytes
        # remain identical.
        "build": {k: v for k, v in build.items() if k not in {"raw", "cleanup"}},
        "quality": {k: v for k, v in quality.items() if k != "raw"},
        "artifacts": {k: v for k, v in artifacts.items() if k != "admission"},
        "admissions": admissions,
        "counts": {"reports": counts["total_reports"], "samples": counts["total_samples"],
                   "native_reports": counts["native_reports"], "native_samples": counts["native_samples"],
                   "observer_reports": counts["observer_reports"], "observer_samples": counts["observer_samples"],
                   "qualification_reports": counts["qualification_reports"],
                   "qualification_samples": counts["qualification_samples"],
                   "artifact_cases": 6, "artifact_policy_outputs": 30},
        "bootstrap": {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                      "confidence": BOOTSTRAP_CONFIDENCE, "low_rank": BOOTSTRAP_LOW_RANK,
                      "high_rank": BOOTSTRAP_HIGH_RANK,
                      "statistic": "median of six process p50 values",
                      "paired_statistic": "median of six matched process-block policy/default p50 ratios",
                      "units": "ns",
                      "status": "computed"},
        "native": native_rows,
        "comparison": cv["plan"]["comparison"],
        "policies": list(POLICIES),
        "policy_semantics": POLICY_SEMANTICS,
        "paired_comparisons": {
            "control_policy": "default",
            "self_controls": 6,
            "nondefault_comparisons": 18,
            "total_rows": 24,
            "default_full_control_rows": 6,
            "weaker_policy_rows": 12,
            "bootstrap": {
                "resamples": BOOTSTRAP_RESAMPLES,
                "seed": BOOTSTRAP_SEED,
                "low_rank": BOOTSTRAP_LOW_RANK,
                "high_rank": BOOTSTRAP_HIGH_RANK,
                "statistic":
                "median of six matched process-block policy/default p50 ratios",
            },
            "scope": "configuration attribution only; no optimization or default weakening",
        },
        "observer": {"timings_pooled_with_native": False,
                      "latency_claim": "diagnostic_only; observer elapsed values are not latency evidence",
                      "operation_metrics_subtracted": False, "procfs_controls_subtracted": False,
                      "reports": observer},
        "qualification": qualification,
        "raw_reports": {"native": native, "observer": observer, "qualification": qualification},
        "verification": {
            "packet_custody_checked": True, "production_source_checked": True,
            "harness_source_checked": True, "runtime_harness_source_checked": True,
            "quality_checked": True, "build_checked": True,
            "artifact_export_checked": True, "independent_artifact_oracle_checked": True,
            "zip_preservation_checked": True, "artifact_admission_checked": True,
            "admission_attempts_checked": True,
            "qualification_admission_checked": True, "report_schema_checked": True,
            "binary_identity_checked": True, "corpus_identity_checked": True,
            "outcome_identity_checked": True, "sample_cardinality_checked": True,
            "observer_diagnostics_retained": True, "native_timing_separated": True,
            "historical_comparison_omitted": True, "bootstrap_checked": True,
            "logical_workload_descriptors_checked": True,
            "policy_route_checked": True, "paired_comparisons_checked": True,
            "policy_semantics_retained": True},
    }


def render_markdown(value: dict[str, Any]) -> str:
    lines = ["# 0821 real-file save durability attribution", "",
             "Offline replay of admitted real-file save-durability receipts.",
             "Native elapsed values describe configuration attribution; no production optimization, historical timing, or default-weakening claim is made.",
             "Default save and explicit full durability are matched controls. File-only and no-sync change durability semantics.",
             "Observer allocation and procfs counters remain diagnostic and are never pooled with native latency.",
             "Analysis p50/p95/p99 use nearest rank within each process block; raw harness p50 is its integer midpoint. Ratios pair each policy with default in the same block, format, and phase.",
             "", f"Native reports/samples: {value['counts']['native_reports']} / {value['counts']['native_samples']}",
             f"Observer reports/samples: {value['counts']['observer_reports']} / {value['counts']['observer_samples']}",
             f"Qualification reports/samples: {value['counts']['qualification_reports']} / {value['counts']['qualification_samples']}",
             "", "| Selector | policy | p50 ns | p95 ns | p99 ns | p50 CI95 ns | ratio/default | ratio CI95 | source B/s | published B/s | spread | tail |",
             "|---|---|---:|---:|---:|---|---:|---|---:|---:|---|---|"]
    for row in value["native"]:
        ci = row["p50_bootstrap_ci95_ns"]
        ratio = row["paired_ratio_to_default"]
        lines.append("| " + " | ".join((row["case"], row["policy"], f"{row['p50_ns']:.6g}",
            f"{row['p95_ns']:.6g}", f"{row['p99_ns']:.6g}",
            f"[{ci['lower']:.6g}, {ci['upper']:.6g}]",
            f"{ratio['ratio_median']:.6g}",
            f"[{ratio['bootstrap_ci95']['lower']:.6g}, {ratio['bootstrap_ci95']['upper']:.6g}]",
            f"{row['workload']['source_logical_bytes_per_second']:.6g}",
            f"{row['workload']['published_logical_bytes_per_second']:.6g}",
            ",".join(row["spread_flags"]) or "-", "yes" if row["tail_flag"] else "-")) + " |")
    lines.extend(["", "No optimization, adoption, or historical comparison is inferred.", ""])
    return "\n".join(lines)


def output_text(value: dict[str, Any]) -> dict[str, str]:
    return {"analysis.json": json.dumps(value, indent=2, sort_keys=True) + "\n",
            "native.csv": csv_text(value["native"]),
            "observer.json": json.dumps(value["observer"], indent=2, sort_keys=True) + "\n",
            "analysis.md": render_markdown(value)}


def analyze(*, write: bool = False, check: bool = False) -> dict[str, Any]:
    value = build_analysis()
    outputs = output_text(value)
    if write:
        for name in outputs:
            require(not (PACKET / name).exists(), f"refusing to overwrite retained {name}")
        for name, text in outputs.items():
            (PACKET / name).write_text(text, encoding="utf-8")
    if check:
        for name, text in outputs.items():
            path = PACKET / name
            require(path.is_file() and path.read_text(encoding="utf-8") == text,
                    f"{name} does not replay deterministically")
    return value


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true", help="write derived outputs")
    mode.add_argument("--check", action="store_true", help="replay retained outputs")
    args = parser.parse_args(argv)
    try:
        value = analyze(write=args.write or not args.check, check=args.check)
        print(json.dumps({"status": value["status"], "reports": value["counts"]["reports"],
                          "samples": value["counts"]["samples"]}, sort_keys=True))
    except (ReplayError, AssertionError, OSError, ValueError, KeyError, TypeError, IndexError) as error:
        print(f"0821 analysis failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
