"""Fail-closed offline replay for the 0820 real-file save-durability packet.

This reader consumes retained JSON, logs, archives, and binary receipts only.
It never starts Cargo, the exporter, a benchmark child, or a workload.  The
same replay builds the derived analysis before cleanup and after cleanup; the
``--check`` mode refuses any changed derived output.
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
import sys
from pathlib import Path
from typing import Any, Iterable

import custody


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN_SCHEMA = "litchi.performance.0820.plan.v1"
ANALYSIS_SCHEMA = "litchi.performance.0820.save-durability-analysis.v1"
REPORT_SCHEMA_VERSION = 1
CAPTURE_SCHEMA = "litchi.performance.0820.capture-receipt.v1"
BOOTSTRAP_SEED = 820820
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_LOW_RANK = 250
BOOTSTRAP_HIGH_RANK = 9749
BOOTSTRAP_CONFIDENCE = 0.95
PRODUCTION_FILES = 9_197
ARCHITECTURE_FILES = 35
OBSERVER_CONTROLS = 32
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
        prefix = "docs/performance/results/change-0820/"
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
    require(value.get("schema") == "litchi.performance.0820.origin.v1",
            "origin schema changed")
    require(isinstance(value.get("base"), str) and len(value["base"]) == 40,
            "origin base is invalid")
    require(value.get("production_changed") is False
            and value.get("runtime_harness_changed") is False
            and value.get("tool_changed") is False
            and value.get("tool_allowlist") == [],
            "source-change witness changed")
    require(value.get("unrelated") == custody.UNRELATED,
            "unrelated-file witness changed")
    return value


def plan() -> dict[str, Any]:
    value = read_json(PACKET / "plan.json")
    require(value.get("schema") == PLAN_SCHEMA, "plan schema changed")
    require(value.get("base") == origin()["base"], "plan base changed")
    require(value.get("affinity") == list(range(12, 20)), "capture affinity changed")
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
    require(tool, "tool source census is empty")
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
    require(value.get("schema") == "litchi.performance.0820.cleanup.v1",
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
    require(build.get("schema") == "litchi.performance.0820.build.v1",
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


def load_quality(cv: dict[str, Any], build: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "quality.json"
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0820.quality.v1"
            and value.get("status") == "pass" and value.get("gate_count") == 6,
            "quality receipt changed")
    source_path = descriptor(value.get("source"), "quality source", packet_bound=True)
    require(source_path is not None, "quality source missing")
    source = source_witness(read_json(source_path), cv, "quality")
    require(source == build["source"], "quality source differs from build source")
    for key in ("root_inputs", "locks", "architecture", "corpus", "provenance",
                "host", "unrelated"):
        require(value.get(key) == cv[key], f"quality {key} witness changed")
    require(value.get("packet") == cv["packet"] and value.get("drivers") == cv["drivers"],
            "quality packet/driver witness changed")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 6, "quality gate cardinality changed")
    manifest = str(custody.TOOL / "Cargo.toml")
    expected_commands = [
        ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        ["cargo", "check", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--all-targets"],
        ["cargo", "test", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--", "--test-threads=2"],
        ["cargo", "clippy", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--no-deps"],
        ["python3", "-B", "tools/check_crate_boundaries.py"],
    ]
    checks_path = source_path.parent / "checks.json"
    same_descriptor(value.get("checks"), identity(checks_path), "quality checks")
    for index, row in enumerate(rows, 1):
        require(row.get("gate") == index and row.get("exit_code") == 0,
                f"quality gate {index} failed")
        require(row.get("command") == expected_commands[index - 1],
                f"quality command changed: gate {index}")
        log = descriptor(row.get("log"), f"quality gate {index} log", packet_bound=True)
        require(log is not None, f"quality gate {index} log missing")
        finite(row.get("started"), f"quality gate {index} start")
        finite(row.get("ended"), f"quality gate {index} end")
        require(row["started"] <= row["ended"], f"quality gate {index} timestamps reversed")
    return {"receipt": identity(path), "source": source, "source_path": relative(source_path),
            "attempt": value.get("attempt"), "rows": rows, "raw": value}


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


def load_artifacts(cv: dict[str, Any], build: dict[str, Any], quality: dict[str, Any]) -> dict[str, Any]:
    directory = PACKET / "artifacts"
    complete_path = PACKET / "artifacts.complete.json"
    receipt_path = PACKET / "artifacts-receipt.json"
    complete = read_json(complete_path)
    receipt = read_json(receipt_path)
    require(receipt.get("schema") == "litchi.performance.0820.artifact-receipt.v1",
            "artifact receipt schema changed")
    require(complete.get("schema") == "litchi.performance.0820.artifacts.complete.v1"
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
    require(admission.get("schema") == "litchi.performance.0820.artifact-admission.v1"
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
    require(audit.get("schema") == "litchi.performance.0820.artifact-audit.v1"
            and audit.get("ok") is True and audit.get("errors") == []
            and len(audit.get("cases", [])) == 6
            and all(row.get("ok") is True for row in audit["cases"]),
            "independent artifact audit changed")
    same_descriptor(admission.get("auditor"), identity(PACKET / "artifact_audit.py"),
                    "artifact auditor")
    zip_desc = admission.get("zip_preservation")
    zip_path = same_descriptor(zip_desc, identity(PACKET / "zip-preservation.json"),
                               "ZIP preservation witness")
    zip_value = read_json(zip_path)
    require(zip_value.get("schema") == "litchi.performance.0820.zip-preservation.v1"
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
                and selector.get("published_sha256") == real["published_sha256"]
                and selector.get("policy_output_sha256") ==
                by_format[case["format"].upper()]["policy_digests"][case["policy"]]
                and selector.get("policy_output_sha256") == selector.get("published_sha256")
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
                           "audit": audit, "zip": zip_value}}


def load_qualification_admission(artifacts: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "qualification-admission.json"
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0820.qualification-admission.v1"
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
                    (custody.SCRATCH / ".litchi-performance-0820-owned").resolve(),
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
    require(complete.get("schema") == f"litchi.performance.0820.{lane}.complete.v1"
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
    observer_rows = qualification + observer
    for row in observer_rows:
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
                    "runtime_harness_matches_base": True, "tool_allowlist": [],
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
        "observer": {"timings_pooled_with_native": False,
                      "latency_claim": "diagnostic_only; observer elapsed values are not latency evidence",
                      "operation_metrics_subtracted": False, "procfs_controls_subtracted": False,
                      "reports": observer_rows},
        "qualification": qualification,
        "raw_reports": {"native": native, "observer": observer, "qualification": qualification},
        "verification": {
            "packet_custody_checked": True, "production_source_checked": True,
            "harness_source_checked": True, "runtime_harness_source_checked": True,
            "quality_checked": True, "build_checked": True,
            "artifact_export_checked": True, "independent_artifact_oracle_checked": True,
            "zip_preservation_checked": True, "artifact_admission_checked": True,
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
    lines = ["# 0820 real-file save durability attribution", "",
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
        print(f"0820 analysis failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
