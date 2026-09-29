#!/usr/bin/env python3
"""Independent custody and numeric audit for the 0836 OPC profile packet.

The driver, workload readers, and profile decoders are separate programs.  This
module deliberately does not import any of them, invoke a command, run Git, or
read a live target directory except through a recorded descriptor.  It checks
the retained source/build/command boundary, validates the report cardinalities,
and recomputes the small numeric comparison that the packet permits.  All
comparisons are descriptive diagnostics; this packet does not authorize a
production optimization claim.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import math
import random
import re
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
TARGET = ROOT.parent / "litchi-target-0836"
SCRATCH = ROOT.parent / "litchi-fs-0836"
BASE = "5bb91a4de42a403cd1dfcb6e10058e3fe62d2228"
TOOL = "tools/perf-baseline/Cargo.toml"
CASES = ("opc_file_source_one_part_atomic_save",)
STAGES = ("baseline", "fp")
STATES = ("warm", "cold-verified")
QUALIFICATION_STATES = ("warm", "cold-verified")
ALLOWED_SOURCES = (
    "tools/perf-baseline/src/filesystem.rs",
    "tools/perf-baseline/src/filesystem/aligned_zip.rs",
    "tools/perf-baseline/README.md",
)
PLAN_SCHEMA = "litchi.0836.opc-filesystem-profile-plan.v1"
AUDIT_SCHEMA = "litchi.0836.independent-custody-audit.v1"
EXPECTED_SCOPE = "Operation-scoped OPC source-backed filesystem save profiling on committed source; iWork excluded"
EXPECTED_COUNTS = {
    "native_reports": 24,
    "native_samples": 288,
    "perf_reports": 4,
    "perf_samples": 120,
    "qualification_reports": 2,
    "qualification_samples": 4,
    "reports": 34,
    "samples": 444,
    "trace_reports": 4,
    "trace_samples": 32,
}
SEED = 836083
BOOTSTRAP = 10_000
ENDPOINTS = (249, 9749)
TRACE_SYSCALLS = ("fsync", "fdatasync", "rename", "renameat", "renameat2")
SHARED_PROFILE_COUNTS = {"reports": 8, "samples": 152,
                         "cpu_reports": 4, "cpu_samples": 120,
                         "trace_reports": 4, "trace_samples": 32}


class AuditError(RuntimeError):
    """The retained packet is incomplete or internally contradictory."""


def fail(message: str) -> None:
    raise AuditError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON evidence {path}: {error}")
    _finite(value, str(path))
    return value


def _finite(value: Any, label: str) -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{label}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            require(isinstance(key, str), f"{label}: non-string key")
            _finite(child, f"{label}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{label}[{index}]")


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


def valid_commit(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 40 and all(
        char in "0123456789abcdef" for char in value)


def path_inside(path: Path, root: Path, label: str) -> Path:
    resolved = path.resolve(strict=False)
    require(resolved.is_relative_to(root.resolve()), f"{label} escaped {root}: {path}")
    return resolved


def packet_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}: missing path")
    candidate = Path(raw)
    return path_inside(candidate if candidate.is_absolute() else PACKET / candidate,
                       PACKET, label)


def cleanup_witness() -> dict[str, Any]:
    path = PACKET / "cleanup.json"
    require(path.is_file() and not path.is_symlink(), "cleanup witness is missing")
    value = read_json(path)
    require(value.get("status") == "pass", "cleanup witness is not pass")
    return value


def contains_descriptor(value: Any, expected: dict[str, Any]) -> bool:
    if isinstance(value, dict):
        if all(value.get(key) == expected.get(key)
               for key in ("path", "bytes", "sha256")):
            return True
        return any(contains_descriptor(child, expected) for child in value.values())
    if isinstance(value, list):
        return any(contains_descriptor(child, expected) for child in value)
    return False


def descriptor_value(value: Any, label: str, *, allow_missing: bool = False,
                     nonempty: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: descriptor is malformed")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label}.path: missing")
    size = value.get("bytes")
    require(type(size) is int and size >= (1 if nonempty else 0),
            f"{label}.bytes: invalid")
    digest = value.get("sha256")
    require(valid_sha(digest), f"{label}.sha256: invalid")
    candidate = Path(raw_path)
    if not candidate.is_absolute():
        candidate = PACKET / candidate
    candidate = candidate.resolve(strict=False)
    require(not candidate.is_symlink(), f"{label}: symlink descriptor")
    if candidate.is_file():
        require(candidate.stat().st_size == size, f"{label}: byte count changed")
        require(sha256(candidate) == digest, f"{label}: hash changed")
    else:
        require(allow_missing, f"{label}: missing evidence {candidate}")
        expected = {"path": raw_path, "bytes": size, "sha256": digest}
        require(contains_descriptor(cleanup_witness(), expected),
                f"{label}: cleanup witness does not retain descriptor")
    return {"path": raw_path, "bytes": size, "sha256": digest}


def descriptor(path: Path, label: str, *, allow_missing: bool = False,
               nonempty: bool = False) -> dict[str, Any]:
    regular = path.is_file() and not path.is_symlink()
    value = {"path": str(path),
             "bytes": path.stat().st_size if regular else (1 if nonempty else 0),
             "sha256": sha256(path) if regular else "0" * 64}
    return descriptor_value(value, label, allow_missing=allow_missing,
                           nonempty=nonempty)


def same_descriptor(left: Any, right: Any, label: str) -> None:
    require(isinstance(left, dict) and isinstance(right, dict),
            f"{label}: descriptor missing")
    for key in ("path", "bytes", "sha256"):
        require(left.get(key) == right.get(key), f"{label}.{key}: differs")


def source_descriptor(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    return {"path": str(path), "bytes": path.stat().st_size,
            "sha256": sha256(path)}


def root_file(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw and not Path(raw).is_absolute(),
            f"{label}: invalid repository path")
    return path_inside(ROOT / raw, ROOT, label)


def load_origin() -> dict[str, Any]:
    path = PACKET / "origin.json"
    value = read_json(path)
    require(value.get("base") == BASE, "origin base changed")
    require(value.get("scope") == EXPECTED_SCOPE, "origin scope changed")
    normative = value.get("normative")
    unrelated = value.get("unrelated")
    require(isinstance(normative, dict) and normative,
            "origin normative custody map missing")
    require(isinstance(unrelated, dict) and unrelated,
            "origin unrelated custody map missing")
    for raw, expected in (*normative.items(), *unrelated.items()):
        require(valid_sha(expected), f"origin hash malformed: {raw}")
        source = root_file(raw, "origin source")
        require(source.is_file() and sha256(source) == expected,
                f"origin source changed: {raw}")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "base": BASE,
            "normative_count": len(normative), "unrelated_count": len(unrelated)}


def source_inventory(freeze: dict[str, Any], label: str) -> dict[str, str]:
    source = freeze.get("source")
    require(isinstance(source, dict) and len(source) >= 9000,
            f"{label}: source inventory missing or too small")
    for raw, digest in source.items():
        require(isinstance(raw, str) and not Path(raw).is_absolute()
                and valid_sha(digest), f"{label}: malformed source entry")
        live = root_file(raw, f"{label} source")
        require(live.is_file() and sha256(live) == digest,
                f"{label}: live source changed: {raw}")
    return source


def validate_freeze(stage: str, *, required: bool = True) -> dict[str, Any] | None:
    path = PACKET / f"freeze-{stage}.json"
    if not path.is_file():
        require(not required, f"missing {stage} source freeze")
        return None
    value = read_json(path)
    require(value.get("stage") == stage, f"{stage} freeze stage changed")
    source = source_inventory(value, stage)
    archive_root = PACKET / "sources" / stage
    require(archive_root.is_dir() and not archive_root.is_symlink(),
            f"missing {stage} source archive")
    archived: dict[str, Path] = {}
    for item in archive_root.rglob("*"):
        require(not item.is_symlink(), f"{stage} source archive contains symlink")
        if item.is_file():
            archived[str(item.relative_to(archive_root))] = item
    require(set(archived) == set(ALLOWED_SOURCES),
            f"{stage} source archive boundary changed")
    for raw, item in archived.items():
        require(sha256(item) == source[raw], f"{stage} archived source changed: {raw}")
    same_descriptor(value.get("driver"), source_descriptor(PACKET / "driver.py",
                                                            f"{stage} driver"),
                    f"{stage} freeze driver")
    same_descriptor(value.get("origin"), source_descriptor(PACKET / "origin.json",
                                                            f"{stage} origin"),
                    f"{stage} freeze origin")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "stage": stage, "source_count": len(source),
            "source": source, "driver": value["driver"], "origin": value["origin"]}


def validate_host() -> dict[str, Any]:
    path = PACKET / "host.json"
    value = read_json(path)
    require(value.get("base") == BASE, "host base changed")
    affinity = value.get("affinity")
    require(isinstance(affinity, list) and 12 in affinity,
            "host CPU affinity does not include CPU 12")
    for key in ("rustc", "cargo", "cpu", "memory"):
        require(isinstance(value.get(key), str) and value[key],
                f"host field missing: {key}")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "base": BASE, "affinity": affinity}


def validate_quality_reuse() -> dict[str, Any]:
    path = PACKET / "quality-reuse.json"
    value = read_json(path)
    require(value.get("status") == "pass" and value.get("reused") is True
            and value.get("source_matches_exactly") is True
            and value.get("source_count") == 9389,
            "quality reuse receipt changed")
    prior = value.get("prior_packet")
    require(isinstance(prior, str), "quality reuse prior packet missing")
    prior_path = Path(prior).resolve(strict=False)
    require(prior_path == (PACKET.parent / "change-0834").resolve(),
            "quality reuse prior packet changed")
    inputs = value.get("inputs")
    require(isinstance(inputs, dict) and inputs, "quality reuse inputs missing")
    for raw, binding in inputs.items():
        require(isinstance(raw, str) and not Path(raw).is_absolute(),
                "quality reuse input path invalid")
        actual = prior_path / raw
        same_descriptor(binding, descriptor(actual, f"quality reuse input {raw}"),
                        f"quality reuse input {raw}")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "status": "pass", "source_count": 9389,
            "input_count": len(inputs), "prior_packet": prior}


def expected_build_argv() -> list[str]:
    return ["cargo", "build", "--offline", "--locked", "--release",
            "--manifest-path", TOOL, "--bin", "litchi-perf-baseline"]


def expected_stage_rustflags(stage: str) -> list[str]:
    return [] if stage == "baseline" else ["-C", "force-frame-pointers=yes"]


def validate_build_environment(stage: str, binary: dict[str, Any]) -> dict[str, Any]:
    """Validate additive Cargo fingerprint custody for one build stage.

    The original build receipts predate environment fields.  The separate
    build-environment receipt therefore carries the actual Cargo fingerprint
    descriptors while the frozen driver remains the source of the intended
    stage distinction.  Original target files may be witnessed by cleanup;
    packet-side retained copies must remain readable.
    """

    path = PACKET / "build-environment.json"
    value = read_json(path)
    require(value.get("status") == "pass", "build environment receipt is not pass")
    stages = value.get("stages")
    require(isinstance(stages, dict) and set(stages) >= set(STAGES),
            "build environment stage coverage changed")
    item = stages.get(stage)
    require(isinstance(item, dict), f"build environment {stage} is malformed")
    same_descriptor(item.get("binary"), binary,
                    f"build environment {stage} binary")
    environment = item.get("environment")
    require(isinstance(environment, dict),
            f"build environment {stage} environment is missing")
    flags = environment.get("RUSTFLAGS", environment.get("rustflags"))
    expected_flags = expected_stage_rustflags(stage)
    if stage == "baseline":
        require(flags in (None, "", []),
                "baseline build environment RUSTFLAGS changed")
    else:
        require(flags == "-C force-frame-pointers=yes",
                "fp build environment RUSTFLAGS changed")

    driver_path = PACKET / "driver.py"
    try:
        driver_text = driver_path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        fail(f"cannot read frozen driver for environment binding: {error}")
    require("assert not e.get('RUSTFLAGS') and not e.get('CARGO_ENCODED_RUSTFLAGS')"
            in driver_text,
            "driver environment guard changed")
    if stage == "fp":
        require("RUSTFLAGS='-C force-frame-pointers=yes'" in driver_text,
                "driver fp RUSTFLAGS binding changed")

    fingerprints = item.get("fingerprints")
    require(isinstance(fingerprints, list) and len(fingerprints) >= 2,
            f"build environment {stage} fingerprints are missing")
    checked: list[dict[str, Any]] = []
    for index, fingerprint in enumerate(fingerprints):
        label = f"build-environment.{stage}.fingerprints[{index}]"
        require(isinstance(fingerprint, dict), f"{label}: malformed fingerprint")
        original = descriptor_value(fingerprint.get("original"), f"{label}.original",
                                    allow_missing=True, nonempty=True)
        retained = descriptor_value(fingerprint.get("retained"), f"{label}.retained",
                                    nonempty=True)
        fingerprint_name = Path(retained["path"]).name
        require(fingerprint_name.endswith(".json")
                and ("litchi-perf-baseline" in fingerprint_name
                     or "litchi_perf_baseline" in fingerprint_name),
                f"{label}: retained fingerprint filename changed")
        require(original["bytes"] == retained["bytes"]
                and original["sha256"] == retained["sha256"],
                f"{label} original/retained content differs")
        recorded = fingerprint.get("rustflags")
        require(recorded == expected_flags,
                f"{label}: rustflags metadata changed")
        retained_path = packet_path(retained["path"], f"{label}.retained")
        fingerprint_json = read_json(retained_path)
        require(fingerprint_json.get("rustflags") == expected_flags,
                f"{label}: Cargo fingerprint rustflags changed")
        checked.append({"original": original, "retained": retained,
                        "rustflags": recorded})
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "stage": stage,
            "environment": environment, "fingerprints": checked}


def expected_workload_argv(binary: str, case: str, state: str, samples: int,
                           warmup: int, report: str) -> list[str]:
    return ["taskset", "-c", "12", binary, "--case", case,
            "--samples", str(samples), "--warmup", str(warmup),
            "--filesystem-cache", state, "--filesystem-root", str(SCRATCH),
            "--json", report]


def expected_capture_argv(kind: str, raw: str, workload: list[str]) -> list[str]:
    if kind == "profile":
        return ["perf", "record", "--no-buildid-cache", "-e", "cycles:u",
                "-F", "997", "--call-graph", "fp", "-o", raw, "--",
                *workload]
    if kind == "trace":
        return ["strace", "-f", "-qq", "-ttt", "-T", "-yy", "-e",
                "trace=" + ",".join(TRACE_SYSCALLS), "-o", raw, "--",
                *workload]
    fail(f"unknown capture kind: {kind}")


def _number(value: Any, label: str, *, positive: bool = False) -> float:
    require(type(value) in (int, float) and not isinstance(value, bool)
            and math.isfinite(float(value))
            and (not positive or value > 0), f"{label}: invalid number")
    return float(value)


def validate_command(label: str, *, stage: str | None = None,
                     expected_argv: list[str] | None = None,
                     expected_exit: int = 0) -> dict[str, Any]:
    root = PACKET / "commands" / label
    require(root.is_dir() and not root.is_symlink(), f"missing command directory: {label}")
    started_path, receipt_path, log_path = (root / name for name in
                                             ("started.json", "receipt.json", "output.log"))
    started, receipt = read_json(started_path), read_json(receipt_path)
    argv = expected_argv if expected_argv is not None else receipt.get("argv")
    require(isinstance(argv, list) and all(isinstance(item, str) for item in argv),
            f"{label}: argv missing")
    require(started.get("argv") == argv and receipt.get("argv") == argv,
            f"{label}: argv differs between receipts")
    require(started.get("cwd") == str(ROOT), f"{label}: cwd changed")
    start = started.get("started_unix")
    finish = receipt.get("finished_unix")
    _number(start, f"{label}.started_unix")
    _number(finish, f"{label}.finished_unix")
    require(receipt.get("started_unix") == start and finish >= start,
            f"{label}: command chronology changed")
    require(receipt.get("exit_code") == expected_exit and receipt.get("error") is None,
            f"{label}: exit/error changed")
    # Every root-owned command is stage-bound by its freeze descriptor.  Older
    # runners did not copy a separate ``stage`` field into the receipt, so
    # derive it from the immutable freeze path and require both records to
    # agree.  This also makes the all-command chronology check meaningful for
    # symbol/decode auxiliary commands, not just workload rows.
    freeze_values = [item.get("freeze") for item in (started, receipt)
                     if item.get("freeze") is not None]
    inferred: set[str] = set()
    for freeze_value in freeze_values:
        require(isinstance(freeze_value, dict), f"{label}: malformed freeze binding")
        raw_path = freeze_value.get("path")
        require(isinstance(raw_path, str), f"{label}: freeze path missing")
        matches = [candidate for candidate in STAGES
                   if Path(raw_path).name == f"freeze-{candidate}.json"]
        require(len(matches) == 1, f"{label}: freeze stage is ambiguous")
        inferred.add(matches[0])
    if stage is None:
        raw_stage = receipt.get("stage") or started.get("stage")
        if raw_stage is not None:
            require(raw_stage in STAGES, f"{label}: invalid stage")
            stage = raw_stage
        elif inferred:
            require(len(inferred) == 1, f"{label}: command has multiple stages")
            stage = next(iter(inferred))
    if stage is not None:
        require(stage in STAGES, f"{label}: invalid stage")
        if inferred:
            require(inferred == {stage}, f"{label}: stage/freeze binding changed")
        freeze = PACKET / f"freeze-{stage}.json"
        require(freeze.is_file(), f"{label}: missing stage freeze")
        expected_freeze = descriptor(freeze, f"{label} freeze")
        require(freeze_values, f"{label}: freeze binding missing")
        for item, name in ((started, "start"), (receipt, "receipt")):
            require(item.get("freeze") is not None,
                    f"{label} {name} freeze binding missing")
            same_descriptor(item["freeze"], expected_freeze,
                            f"{label} {name} freeze")
    require(receipt.get("log") is not None, f"{label}: log descriptor missing")
    same_descriptor(receipt["log"], descriptor(log_path, f"{label} log"),
                    f"{label} log")
    # Drivers may expose a source hash or a source descriptor in either receipt.
    for item, source in ((started, f"{label} start"), (receipt, f"{label} receipt")):
        for key in ("source_sha256", "source_hash", "source_inventory_sha256"):
            if key in item:
                require(valid_sha(item[key]), f"{source}.{key}: invalid hash")
    environments = [item.get("environment") for item in (started, receipt)
                    if isinstance(item.get("environment"), dict)]
    if environments:
        require(all(item == environments[0] for item in environments),
                f"{label}: command environment differs between receipts")
    return {"name": label, "argv": argv, "exit_code": expected_exit,
            "started_unix": start, "finished_unix": finish,
            "started": descriptor(started_path, f"{label} started"),
            "receipt": descriptor(receipt_path, f"{label} receipt"),
            "log": descriptor(log_path, f"{label} log"), "stage": stage,
            "environment": environments[0] if environments else None}


def _walk_descriptors(value: Any, label: str) -> None:
    """Check every explicit {path, bytes, sha256} object in retained metadata."""
    if isinstance(value, dict):
        if all(key in value for key in ("path", "bytes", "sha256")):
            descriptor_value(value, label, allow_missing=True)
            return
        for key, child in value.items():
            _walk_descriptors(child, f"{label}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _walk_descriptors(child, f"{label}[{index}]")


def validate_build(stage: str, freeze: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    path = PACKET / f"build-{stage}.json"
    value = read_json(path)
    _walk_descriptors(value, str(path))
    binary_value = value.get("binary")
    if binary_value is None and isinstance(value.get("binaries"), dict):
        binary_value = value["binaries"].get("baseline" if stage == "baseline" else "fp")
    binary = descriptor_value(binary_value, f"{stage} binary", allow_missing=True,
                              nonempty=True)
    binary_path = Path(binary["path"]).resolve(strict=False)
    require(binary_path.name == "litchi-perf-baseline"
            and binary_path.parent.name == stage and binary_path.is_relative_to(TARGET),
            f"{stage} binary path changed")
    command = validate_command(f"build-{stage}", stage=stage,
                               expected_argv=expected_build_argv())
    receipt = value.get("receipt")
    require(receipt is not None, f"{stage} build receipt missing")
    same_descriptor(receipt, command["receipt"], f"{stage} build receipt")
    source = value.get("source") or value.get("source_inventory")
    if source is not None:
        if isinstance(source, dict) and "sha256" in source:
            require(valid_sha(source["sha256"]), f"{stage} build source hash invalid")
        elif isinstance(source, dict):
            for raw, digest in source.items():
                require(freeze["source"].get(raw) == digest,
                        f"{stage} build source identity changed: {raw}")
    environment = validate_build_environment(stage, binary)
    return ({"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
             "sha256": sha256(path), "stage": stage, "binary": binary,
             "receipt": receipt, "environment": environment}, command)


def midpoint(left: int | float, right: int | float) -> int | float:
    if type(left) is int and type(right) is int:
        return left // 2 + right // 2 + ((left % 2 + right % 2) // 2)
    return (left + right) / 2


def nearest_rank(values: list[int | float], quantile: float) -> int | float:
    ordered = sorted(values)
    require(ordered, "empty sample vector")
    return ordered[max(1, math.ceil(len(ordered) * quantile)) - 1]


def summary(values: list[int | float]) -> dict[str, Any]:
    require(values and all(type(value) in (int, float) and not isinstance(value, bool)
                           and math.isfinite(float(value)) and value > 0 for value in values),
            "invalid positive sample vector")
    ordered = sorted(values)
    middle = (ordered[len(ordered) // 2] if len(ordered) % 2 else
              midpoint(ordered[len(ordered) // 2 - 1], ordered[len(ordered) // 2]))
    return {"min": ordered[0], "p50": middle,
            "p95": nearest_rank(ordered, .95), "p99": nearest_rank(ordered, .99),
            "max": ordered[-1], "mean": statistics.mean(ordered)}


def _filesystem_samples(value: dict[str, Any], label: str,
                        expected_case: str, expected_state: str,
                        expected_count: int) -> list[dict[str, Any]]:
    """Return only the measured samples in the report's declared evidence.

    The workload report has several arrays named ``samples`` in its result and
    diagnostic sections.  Only ``filesystem_evidence[].samples`` is a custody
    boundary for measured child processes.  Requiring the single admitted case,
    its declared cache states, and exact per-state cardinality prevents a
    result summary or a nested diagnostic vector from satisfying this audit.
    """

    states = _state_values(expected_state)
    evidence = value.get("filesystem_evidence")
    require(isinstance(evidence, list) and len(evidence) == 1,
            f"{label}: filesystem evidence record count changed")
    record = evidence[0]
    require(isinstance(record, dict), f"{label}: filesystem evidence is malformed")
    require(record.get("case") == expected_case,
            f"{label}: filesystem evidence case changed")
    declared_states = record.get("cache_states")
    require(isinstance(declared_states, list)
            and tuple(declared_states) == states,
            f"{label}: filesystem evidence cache states changed")
    samples = record.get("samples")
    require(isinstance(samples, list),
            f"{label}: filesystem evidence samples are missing")
    require(expected_count % len(states) == 0,
            f"{label}: expected sample count is not divisible by cache states")
    per_state = expected_count // len(states)
    require(len(samples) == expected_count,
            f"{label}: expected {expected_count} filesystem evidence samples")
    require(all(isinstance(item, dict) for item in samples),
            f"{label}: filesystem evidence sample is malformed")
    counts = {state: 0 for state in states}
    for index, item in enumerate(samples):
        state = item.get("cache_state")
        require(state in counts,
                f"{label}.filesystem_evidence.samples[{index}] cache state invalid")
        # The actual reader schema places this scalar directly on each
        # evidence sample.  Do not accept a nested result statistic here.
        require(type(item.get("elapsed_ns")) is int and item["elapsed_ns"] > 0,
                f"{label}.filesystem_evidence.samples[{index}].elapsed_ns invalid")
        counts[state] += 1
    require(all(counts[state] == per_state for state in states),
            f"{label}: filesystem evidence state cardinality changed")
    sample_count = record.get("sample_count")
    require(type(sample_count) is int and sample_count == per_state,
            f"{label}: filesystem evidence sample_count changed")
    return samples


def _sample_metric(sample: dict[str, Any], key: str) -> Any:
    if key in sample:
        return sample[key]
    for holder in (sample.get("process_metrics"), sample.get("metrics"),
                   sample.get("resource_usage")):
        if isinstance(holder, dict):
            if key in holder:
                return holder[key]
            aliases = {"peak_rss_bytes": ("peak_rss", "max_rss_bytes", "rss_bytes"),
                       "elapsed_ns": ("elapsed", "duration_ns")}
            for alias in aliases.get(key, ()):
                if alias in holder:
                    return holder[alias]
    return None


def _binary_sha(value: Any) -> str | None:
    if isinstance(value, dict):
        for key in ("binary_sha256", "sha256"):
            if key in value and valid_sha(value[key]):
                return value[key]
        for child in value.values():
            result = _binary_sha(child)
            if result is not None:
                return result
    elif isinstance(value, list):
        for child in value:
            result = _binary_sha(child)
            if result is not None:
                return result
    return None


def _state_values(expected_state: str) -> tuple[str, ...]:
    values = tuple(part for part in expected_state.split(",") if part)
    require(values and all(value in STATES for value in values),
            f"unknown cache state: {expected_state}")
    require(len(set(values)) == len(values),
            f"duplicate cache state: {expected_state}")
    return values


def _check_result_statistics(value: dict[str, Any], samples: list[dict[str, Any]],
                             expected_state: str, label: str) -> None:
    """Recompute each retained result's scalar timing summary from samples."""

    results = value.get("results")
    require(isinstance(results, list) and results,
            f"{label}: timed result list is missing")
    states = _state_values(expected_state)
    checked = 0
    for result in results:
        if not isinstance(result, dict) or result.get("cache_state") not in states:
            continue
        state = result["cache_state"]
        state_samples = [sample for sample in samples
                         if sample.get("cache_state") == state]
        require(state_samples, f"{label}: result {state} has no evidence samples")
        elapsed = [sample["elapsed_ns"] for sample in state_samples]
        stats = result.get("elapsed_ns")
        require(isinstance(stats, dict),
                f"{label}: result {state} timing statistics are missing")
        ordered = sorted(elapsed)
        sorted_values = stats.get("samples")
        require(sorted_values == ordered,
                f"{label}: result {state} sorted elapsed vector differs")
        sample_order = stats.get("sample_order")
        if sample_order is not None:
            by_index = {sample.get("sample_index"): sample["elapsed_ns"]
                        for sample in state_samples
                        if type(sample.get("sample_index")) is int}
            require(len(by_index) == len(state_samples),
                    f"{label}: result {state} sample indices are missing")
            require(sorted(sample_order) == sorted(by_index),
                    f"{label}: result {state} sample order is not a permutation")
            require([by_index[index] for index in sample_order] == sorted_values,
                    f"{label}: result {state} sample order does not match evidence")
        expected = summary(elapsed)
        for field in ("min", "p50", "p95", "p99", "max"):
            require(stats.get(field) == expected[field],
                    f"{label}: result {state} {field} differs from evidence")
        require(math.isclose(float(stats.get("mean")), float(expected["mean"]),
                             rel_tol=0.0, abs_tol=1e-6),
                f"{label}: result {state} mean differs from evidence")
        checked += 1
    require(checked == len(states),
            f"{label}: result state coverage differs from evidence")


def report_samples(path: Path, label: str, expected_case: str,
                   expected_state: str, expected_count: int,
                   expected_binary: dict[str, Any], seen_pids: set[int],
                   *, strict_count: bool = True) -> dict[str, Any]:
    value = read_json(path)
    states_expected = _state_values(expected_state)
    configuration = value.get("configuration")
    if isinstance(configuration, dict):
        if "samples_per_case" in configuration:
            per_state = expected_count // len(states_expected)
            require(configuration["samples_per_case"] in (expected_count, per_state),
                    f"{label}: sample configuration changed")
        if "warmup_iterations_per_case" in configuration:
            require(configuration["warmup_iterations_per_case"] >= 0,
                    f"{label}: warmup configuration invalid")
        states = configuration.get("filesystem_cache_states")
        if states is not None:
            require(isinstance(states, (list, tuple))
                    and tuple(states) == states_expected,
                    f"{label}: cache state changed")
    samples = _filesystem_samples(value, label, expected_case, expected_state,
                                  expected_count)
    require(len(samples) == expected_count or not strict_count,
            f"{label}: expected {expected_count} samples")
    pids: list[int] = []
    elapsed: list[int] = []
    rss: list[int] = []
    for index, item in enumerate(samples):
        raw_elapsed = item.get("elapsed_ns")
        require(type(raw_elapsed) is int and raw_elapsed > 0,
                f"{label}.samples[{index}].elapsed_ns invalid")
        elapsed.append(raw_elapsed)
        cache_state = item.get("cache_state")
        require(cache_state in states_expected,
                f"{label}.samples[{index}] cache state invalid")
        sample_index = item.get("sample_index")
        if sample_index is not None:
            require(type(sample_index) is int and sample_index >= 0,
                    f"{label}.samples[{index}] sample index invalid")
        pid = item.get("child_process_id", item.get("pid", item.get("process_id")))
        if pid is not None:
            require(type(pid) is int and pid > 0, f"{label}.samples[{index}] pid invalid")
            require(pid not in seen_pids, f"{label}: duplicate child PID {pid}")
            seen_pids.add(pid)
            pids.append(pid)
        raw_rss = _sample_metric(item, "peak_rss_bytes")
        if raw_rss is not None:
            require(type(raw_rss) is int and raw_rss > 0,
                    f"{label}.samples[{index}] peak RSS invalid")
            rss.append(raw_rss)
        if cache_state == "cold-verified":
            proof = item["cold_verified"]
            require(isinstance(proof, dict) and proof.get("status") == "eligible",
                    f"{label}.samples[{index}] cold proof missing")
    require(len(pids) == len(samples), f"{label}: process identity is missing")
    identity = value.get("binary_identity")
    require(isinstance(identity, dict), f"{label}: binary identity is missing")
    binary = identity.get("binary_sha256")
    require(valid_sha(binary) and binary == expected_binary["sha256"],
            f"{label}: binary identity changed")
    if expected_case and isinstance(value.get("results"), list):
        for result in value["results"]:
            if isinstance(result, dict) and result.get("case") is not None:
                require(result.get("case") == expected_case,
                        f"{label}: report case changed")
    _check_result_statistics(value, samples, expected_state, label)
    return {"samples": len(samples), "pids": pids, "elapsed": elapsed, "rss": rss,
            "binary_sha256": binary, "raw": value}


def report_binding(row: dict[str, Any], label: str) -> tuple[Path, Path]:
    report_value = row.get("report")
    receipt_value = row.get("receipt")
    if isinstance(report_value, dict) and "descriptor" in report_value:
        report_value = report_value["descriptor"]
    if isinstance(receipt_value, dict) and "descriptor" in receipt_value:
        receipt_value = receipt_value["descriptor"]
    require(report_value is not None and receipt_value is not None,
            f"{label}: report/receipt descriptor missing")
    report = descriptor_value(report_value, f"{label} report")
    receipt = descriptor_value(receipt_value, f"{label} receipt")
    report_path = packet_path(report["path"], f"{label} report")
    receipt_path = packet_path(receipt["path"], f"{label} receipt")
    return report_path, receipt_path


def validate_manifest_row(row: Any, label: str, expected_stage: str,
                          expected_state: str, expected_samples: int,
                          expected_warmup: int, expected_case: str,
                          binaries: dict[str, Any], seen_pids: set[int],
                          *, strict_report: bool = True,
                          capture_kind: str | None = None) -> dict[str, Any]:
    require(isinstance(row, dict), f"{label}: row malformed")
    require(row.get("label", label) == label, f"{label}: label changed")
    require(row.get("stage") == expected_stage and row.get("state") == expected_state
            and row.get("case") == expected_case
            and row.get("exit_code") == 0, f"{label}: manifest identity changed")
    # The qualification runner predates the standard row schema and records
    # one CLI sample for each of two cache states without copying samples and
    # warmup into the row.  Formal native/profile/trace rows must carry both.
    combined_qualification = expected_state == ",".join(QUALIFICATION_STATES)
    require(row.get("samples") == expected_samples,
            f"{label}: sample count changed")
    if not combined_qualification or "warmup" in row:
        require(row.get("warmup") == expected_warmup,
                f"{label}: warmup changed")
    report, receipt = report_binding(row, label)
    require(report.is_file() and not report.is_symlink(), f"{label}: report missing")
    require(receipt.is_file() and not receipt.is_symlink(), f"{label}: receipt missing")
    receipt_value = read_json(receipt)
    require(receipt_value.get("exit_code") == 0 and receipt_value.get("error") is None,
            f"{label}: command receipt failed")
    cli_states = expected_state
    cli_samples = expected_samples
    workload_argv = expected_workload_argv(
        binaries[expected_stage]["path"], expected_case, cli_states, cli_samples,
        expected_warmup, str(report))
    command_argv = workload_argv
    if capture_kind is not None:
        raw_value = row.get("raw")
        if isinstance(raw_value, dict) and "descriptor" in raw_value:
            raw_value = raw_value["descriptor"]
        raw = descriptor_value(raw_value, f"{label} raw", allow_missing=True,
                               nonempty=True)
        command_argv = expected_capture_argv(capture_kind, raw["path"], workload_argv)
    require(receipt_value.get("argv") == command_argv,
            f"{label}: command argv changed")
    expected_binary = binaries[expected_stage]
    report_count = expected_samples * len(_state_values(expected_state))
    parsed = report_samples(report, label, expected_case, expected_state,
                             report_count, expected_binary, seen_pids,
                             strict_count=strict_report)
    report_environment = parsed["raw"].get("environment")
    require(isinstance(report_environment, dict),
            f"{label}: report environment is missing")
    report_flags = report_environment.get("rustflags",
                                         report_environment.get("RUSTFLAGS"))
    expected_flags = None if expected_stage == "baseline" else "-C force-frame-pointers=yes"
    require(report_flags in (None, "") if expected_stage == "baseline"
            else report_flags == expected_flags,
            f"{label}: report RUSTFLAGS differs from stage")
    # Explicit PID bindings in profile/trace manifests must equal the report's
    # observed child set.  Different harmless spellings are accepted because
    # old helper scripts used both names.
    for key in ("pids", "child_pids", "report_pids", "measured_pids"):
        if key in row:
            raw = row[key]
            require(isinstance(raw, list) and all(type(pid) is int for pid in raw),
                    f"{label}.{key}: malformed PID binding")
            require(sorted(raw) == sorted(parsed["pids"]),
                    f"{label}.{key}: does not equal report PIDs")
    for key in ("counter_fields", "process_metric_fields", "metrics"):
        if key in row:
            raw = row[key]
            require(isinstance(raw, list) and raw
                    and all(isinstance(field, str) and field for field in raw)
                    and len(set(raw)) == len(raw),
                    f"{label}.{key}: malformed counter-field list")
    return {"label": label, "stage": expected_stage, "state": expected_state,
            "case": expected_case, "samples": expected_samples,
            "warmup": expected_warmup, "report": report, "receipt": receipt,
            "argv": workload_argv, "command_argv": command_argv,
            "parsed": parsed}


def _gzip_digest(path: Path) -> tuple[int, str]:
    digest = hashlib.sha256()
    size = 0
    try:
        with gzip.open(path, "rb") as stream:
            for block in iter(lambda: stream.read(1 << 20), b""):
                size += len(block)
                digest.update(block)
    except (OSError, EOFError, gzip.BadGzipFile) as error:
        fail(f"invalid retained gzip profile artifact {path}: {error}")
    return size, digest.hexdigest()


def _command_label_from_receipt(value: Any, label: str) -> tuple[str, dict[str, Any]]:
    receipt = descriptor_value(value, f"{label} receipt")
    path = packet_path(receipt["path"], f"{label} receipt")
    require(path.name == "receipt.json" and path.parent.parent == PACKET / "commands",
            f"{label}: command receipt path changed")
    return path.parent.name, receipt


def validate_profile_artifacts(row: dict[str, Any], label: str,
                               expected_stage: str, kind: str) -> None:
    """Bind capture, decoder and compressed witnesses to exact commands.

    The large perf ``.data`` file may be removed by cleanup.  Its packet-side
    gzip is the durable witness, so the audit checks the uncompressed byte
    count and digest against the original descriptor and lets
    ``descriptor_value`` obtain the removed descriptor from cleanup.json.
    """

    raw_value = row.get("raw")
    gzip_value = row.get("raw_gzip")
    if isinstance(raw_value, dict) and "descriptor" in raw_value:
        raw_value = raw_value["descriptor"]
    if isinstance(gzip_value, dict) and "descriptor" in gzip_value:
        gzip_value = gzip_value["descriptor"]
    require(raw_value is not None and gzip_value is not None,
            f"{label}: perf raw/gzip descriptors are missing")
    raw = descriptor_value(raw_value, f"{label} perf raw", allow_missing=True)
    compressed = descriptor_value(gzip_value, f"{label} perf gzip")
    compressed_path = packet_path(compressed["path"], f"{label} perf gzip")
    require(compressed_path.is_file() and compressed_path.suffix == ".gz",
            f"{label}: perf gzip witness is missing")
    size, digest = _gzip_digest(compressed_path)
    require(size == raw["bytes"] and digest == raw["sha256"],
            f"{label}: gzip witness does not reproduce perf raw data")

    if kind == "trace":
        require(not any(key in row for key in ("decoded", "decoded_plain",
                                               "decode_receipt", "buildid",
                                               "buildid_receipt")),
                f"{label}: trace row has profile-only artifacts")
        return
    require(kind == "profile", f"{label}: unknown artifact kind")
    decoded_value = row.get("decoded")
    decoded_plain_value = row.get("decoded_plain")
    decode_receipt_value = row.get("decode_receipt")
    buildid_value = row.get("buildid")
    buildid_receipt_value = row.get("buildid_receipt")
    require(all(value is not None for value in
                (decoded_value, decoded_plain_value, decode_receipt_value,
                 buildid_value, buildid_receipt_value)),
            f"{label}: profile decoder/build-id artifacts are missing")
    decoded = descriptor_value(decoded_value, f"{label} decoded", nonempty=True)
    decoded_plain = descriptor_value(decoded_plain_value,
                                     f"{label} decoded plain",
                                     allow_missing=True, nonempty=True)
    decoded_path = packet_path(decoded["path"], f"{label} decoded")
    require(decoded_path.suffix == ".gz",
            f"{label}: decoded profile witness is not gzip")
    size, digest = _gzip_digest(decoded_path)
    require(size == decoded_plain["bytes"] and digest == decoded_plain["sha256"],
            f"{label}: decoded gzip witness does not reproduce plain script")
    decode_label, decode_descriptor = _command_label_from_receipt(
        decode_receipt_value, f"{label} decode")
    decode_command = validate_command(
        decode_label, stage=expected_stage,
        expected_argv=["perf", "script", "--no-inline", "--ns",
                       "--show-lost-events", "-i", raw["path"]])
    same_descriptor(decode_descriptor, decode_command["receipt"],
                    f"{label} decode receipt")
    buildid_label, buildid_descriptor = _command_label_from_receipt(
        buildid_receipt_value, f"{label} build-id")
    buildid_command = validate_command(
        buildid_label, stage=expected_stage,
        expected_argv=["perf", "buildid-list", "-i", raw["path"]])
    buildid_output = descriptor_value(buildid_value, f"{label} build-id output")
    same_descriptor(buildid_output, buildid_command["log"],
                    f"{label} build-id output")
    same_descriptor(buildid_descriptor, buildid_command["receipt"],
                    f"{label} build-id receipt")


def validate_mappings(binaries: dict[str, Any]) -> dict[str, Any]:
    """Validate the optional retained mmap/task-event decode admission.

    Mapping decodes are auxiliary to CPU attribution, but their command and
    gzip/plain custody still belongs to the packet.  The mapping tool reads
    only the retained profile records and always uses the frame-pointer binary.
    """

    path = PACKET / "mappings.json"
    value = read_json(path)
    require(value.get("status") == "pass", "mapping admission is not pass")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 4,
            "mapping row count changed")
    same_descriptor(value.get("binary"), binaries["fp"], "mapping binary")
    elf = descriptor_value(value.get("elf"), "mapping ELF output")
    elf_label, elf_receipt = _command_label_from_receipt(
        value.get("elf_receipt"), "mapping ELF")
    elf_command = validate_command(
        elf_label, stage="fp",
        expected_argv=["readelf", "-lW", binaries["fp"]["path"]])
    same_descriptor(elf, elf_command["log"], "mapping ELF output")
    same_descriptor(elf_receipt, elf_command["receipt"], "mapping ELF receipt")

    profiles = read_json(PACKET / "profiles.json")
    profile_rows = profiles.get("rows")
    require(isinstance(profile_rows, list) and len(profile_rows) == 4,
            "mapping profile source rows are missing")
    profile_by_label = {row.get("label"): row for row in profile_rows
                        if isinstance(row, dict)}
    seen: set[str] = set()
    checked: list[dict[str, Any]] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"mapping row {index}: malformed")
        label = row.get("label")
        require(isinstance(label, str) and label in profile_by_label and label not in seen,
                f"mapping row {index}: profile label changed")
        seen.add(label)
        source = profile_by_label[label].get("raw")
        if isinstance(source, dict) and "descriptor" in source:
            source = source["descriptor"]
        raw_value = row.get("raw")
        if isinstance(raw_value, dict) and "descriptor" in raw_value:
            raw_value = raw_value["descriptor"]
        raw = descriptor_value(raw_value, f"mapping {label} raw",
                               allow_missing=True, nonempty=True)
        source_raw = descriptor_value(source, f"mapping {label} profile raw",
                                      allow_missing=True, nonempty=True)
        same_descriptor(raw, source_raw, f"mapping {label} raw")
        decoded = descriptor_value(row.get("decoded"),
                                   f"mapping {label} decoded", nonempty=True)
        decoded_plain = descriptor_value(row.get("decoded_plain"),
                                         f"mapping {label} decoded plain",
                                         allow_missing=True, nonempty=True)
        decoded_path = packet_path(decoded["path"], f"mapping {label} decoded")
        require(decoded_path.suffix == ".gz",
                f"mapping {label}: decoded witness is not gzip")
        size, digest = _gzip_digest(decoded_path)
        require(size == decoded_plain["bytes"]
                and digest == decoded_plain["sha256"],
                f"mapping {label}: decoded gzip witness differs")
        command_label, receipt = _command_label_from_receipt(
            row.get("receipt"), f"mapping {label}")
        command = validate_command(
            command_label, stage="fp",
            expected_argv=["perf", "script", "--no-inline", "--ns",
                           "--show-mmap-events", "--show-task-events", "-i",
                           raw["path"]])
        same_descriptor(receipt, command["receipt"],
                        f"mapping {label} receipt")
        checked.append({"label": label, "raw": raw, "decoded": decoded,
                        "decoded_plain": decoded_plain})
    require(seen == set(profile_by_label), "mapping profile coverage changed")
    _walk_descriptors(value, "mappings")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "status": "pass", "reports": len(checked),
            "samples": 0, "rows": checked}


def validate_qualification(binaries: dict[str, Any], seen_pids: set[int]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    path = PACKET / "qualification.json"
    value = read_json(path)
    require(value.get("status") in ("commands_pass", "pass")
            and isinstance(value.get("rows"), list)
            and len(value["rows"]) == 2, "qualification manifest changed")
    if "report_count" in value:
        require(value["report_count"] == EXPECTED_COUNTS["qualification_reports"],
                "qualification report cardinality changed")
    if "sample_count" in value:
        require(value["sample_count"] == EXPECTED_COUNTS["qualification_samples"],
                "qualification sample cardinality changed")
    commands: list[dict[str, Any]] = []
    parsed: list[dict[str, Any]] = []
    for index, stage in enumerate(STAGES):
        row = value["rows"][index]
        label = row.get("label")
        if not isinstance(label, str) or not label:
            raw_receipt = row.get("receipt")
            if isinstance(raw_receipt, dict) and "descriptor" in raw_receipt:
                raw_receipt = raw_receipt["descriptor"]
            if isinstance(raw_receipt, dict) and isinstance(raw_receipt.get("path"), str):
                label = Path(raw_receipt["path"]).parent.name
        label = label or f"qualification-{stage}"
        item = validate_manifest_row(row, label, stage, "warm,cold-verified", 1, 0,
                                     CASES[0], binaries, seen_pids)
        parsed.append(item)
        commands.append(validate_command(label, stage=stage,
                                         expected_argv=item["argv"]))
    require(sum(item["parsed"]["samples"] for item in parsed) ==
            EXPECTED_COUNTS["qualification_samples"],
            "qualification sample cardinality changed")
    return ({"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
             "sha256": sha256(path), "status": value["status"], "reports": 2,
             "samples": 4}, commands)


def expected_native_rows(plan: dict[str, Any]) -> list[dict[str, Any]]:
    native = plan.get("native")
    require(isinstance(native, list) and len(native) == 24,
            "plan native rows changed")
    expected: list[dict[str, Any]] = []
    for item in native:
        require(isinstance(item, dict), "plan native row malformed")
        row = dict(item)
        row["case"] = CASES[0]
        expected.append(row)
    return expected


def validate_native(plan: dict[str, Any], binaries: dict[str, Any],
                    seen_pids: set[int]) -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    path = PACKET / "native.json"
    value = read_json(path)
    require(value.get("status") in ("commands_pass", "pass")
            and isinstance(value.get("rows"), list)
            and len(value["rows"]) == 24, "native manifest changed")
    if "report_count" in value:
        require(value["report_count"] == EXPECTED_COUNTS["native_reports"],
                "native report cardinality changed")
    if "sample_count" in value:
        require(value["sample_count"] == EXPECTED_COUNTS["native_samples"],
                "native sample cardinality changed")
    expected = expected_native_rows(plan)
    commands: list[dict[str, Any]] = []
    records: list[dict[str, Any]] = []
    for index, expected_row in enumerate(expected):
        row = value["rows"][index]
        label = f"native-{index:03}"
        if isinstance(row, dict) and row.get("label"):
            label = row["label"]
        require(isinstance(row, dict), f"{label}: native row malformed")
        for key in ("block", "stage", "state", "samples", "warmup", "case"):
            require(row.get(key) == expected_row[key],
                    f"{label}.{key}: differs from frozen plan")
        item = validate_manifest_row(row, label, expected_row["stage"],
                                     expected_row["state"], expected_row["samples"],
                                     expected_row["warmup"], expected_row["case"],
                                     binaries, seen_pids)
        item["block"] = expected_row["block"]
        records.append(item)
        commands.append(validate_command(label, stage=expected_row["stage"],
                                         expected_argv=item["argv"]))
    require(sum(item["parsed"]["samples"] for item in records) == 288,
            "native sample cardinality changed")
    stats = native_statistics(records, plan)
    return ({"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
             "sha256": sha256(path), "status": value["status"], "reports": 24,
             "samples": 288}, commands, stats)


def native_statistics(records: list[dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    grouped: dict[tuple[int, str, str], dict[str, Any]] = {}
    for record in records:
        key = (record["block"], record["stage"], record["state"])
        require(key not in grouped, f"duplicate native key: {key}")
        parsed = record["parsed"]
        require(len(parsed["elapsed"]) == 12,
                f"native {key}: elapsed vector missing")
        require(len(parsed["rss"]) in (0, 12),
                f"native {key}: incomplete RSS vector")
        grouped[key] = {"block": record["block"], "stage": record["stage"],
                        "state": record["state"], "latency": summary(parsed["elapsed"]),
                        "rss": summary(parsed["rss"]) if parsed["rss"] else None,
                        "pids": parsed["pids"]}
    require(len(grouped) == 24, "native paired coverage changed")
    pair_rows: list[dict[str, Any]] = []
    spread_flags: list[dict[str, Any]] = []
    for state in STATES:
        for stage in STAGES:
            for metric in ("latency", "rss"):
                values = [grouped[(block, stage, state)][metric]["p50"]
                          for block in range(6)
                          if grouped[(block, stage, state)][metric] is not None]
                if values:
                    ratio = max(values) / min(values)
                    if ratio > 1.2:
                        spread_flags.append({"stage": stage, "state": state,
                                             "metric": metric, "max_over_min": ratio,
                                             "reason": "six-block p50 spread exceeds 20%"})
    for state in STATES:
        for block in range(6):
            baseline = grouped[(block, "baseline", state)]
            fp = grouped[(block, "fp", state)]
            for metric in ("latency", "rss"):
                if baseline[metric] is None or fp[metric] is None:
                    continue
                numerator = fp[metric]["p50"]
                denominator = baseline[metric]["p50"]
                require(numerator > 0 and denominator > 0,
                        "native ratio has non-positive denominator")
                pair_rows.append({"block": block, "state": state, "metric": metric,
                                  "baseline_p50": denominator, "fp_p50": numerator,
                                  "fp_over_baseline": numerator / denominator})
    bootstrap_rows: list[dict[str, Any]] = []
    for state in STATES:
        for metric in ("latency", "rss"):
            ratios = [row["fp_over_baseline"] for row in pair_rows
                      if row["state"] == state and row["metric"] == metric]
            if not ratios:
                continue
            require(len(ratios) == 6, f"native {state}/{metric}: paired ratio count changed")
            rng = random.Random(SEED)
            resampled = sorted(statistics.median(rng.choices(ratios, k=6))
                               for _ in range(BOOTSTRAP))
            bootstrap_rows.append({"state": state, "metric": metric,
                                   "ratios": ratios, "median": statistics.median(ratios),
                                   "bootstrap_95": [resampled[ENDPOINTS[0]],
                                                     resampled[ENDPOINTS[1]]],
                                   "bootstrap": {"method": "median of six matched block ratios",
                                                 "seed": SEED, "resamples": BOOTSTRAP,
                                                 "lower_index": ENDPOINTS[0],
                                                 "upper_index": ENDPOINTS[1]}})
    stats = plan.get("statistics", {})
    require(stats.get("seed") == SEED and stats.get("bootstrap_resamples") == BOOTSTRAP
            and stats.get("sorted_endpoints") == list(ENDPOINTS),
            "native statistics plan changed")
    return {"blocks": [grouped[key] for key in sorted(grouped)],
            "paired_ratios": pair_rows, "bootstrap": bootstrap_rows,
            "spread_flags": spread_flags, "seed": SEED,
            "resamples": BOOTSTRAP, "scope": "descriptive fp/baseline control ratio; no optimization effect"}


def _symbol_matches(symbol: Any, owner: str) -> bool:
    return (isinstance(symbol, str)
            and (symbol == owner
                 or re.fullmatch(re.escape(owner) + r"::h[0-9a-f]+", symbol)
                 is not None))


def validate_symbols(plan: dict[str, Any], binaries: dict[str, Any]) -> dict[str, Any]:
    """Validate the exact owner/range admission retained before profiling."""

    path = PACKET / "symbols.json"
    value = read_json(path)
    require(value.get("status") == "pass", "symbol admission is not pass")
    contract = plan.get("perf_contract")
    require(isinstance(contract, dict), "perf contract is missing")
    candidates = contract.get("owner_candidates")
    require(isinstance(candidates, list) and candidates
            and all(isinstance(candidate, str) and candidate for candidate in candidates),
            "perf owner candidates are missing")
    owner = value.get("owner")
    require(owner in candidates, "admitted owner is not a plan candidate")
    scope = value.get("scope")
    require(isinstance(scope, str) and scope,
            "admitted owner scope is missing")
    stages = value.get("stages")
    require(isinstance(stages, dict) and set(stages) == set(STAGES),
            "symbol stage coverage changed")
    checked_ranges: dict[str, int] = {}
    for stage in STAGES:
        stage_value = stages[stage]
        require(isinstance(stage_value, dict), f"symbols {stage}: malformed stage")
        same_descriptor(stage_value.get("binary"), binaries[stage],
                        f"symbols {stage} binary")
        inventory = stage_value.get("inventory")
        require(inventory is not None, f"symbols {stage}: inventory missing")
        inventory_descriptor = descriptor_value(inventory, f"symbols {stage} inventory")
        inventory_path = packet_path(inventory_descriptor["path"],
                                     f"symbols {stage} inventory")
        require(inventory_path.name == f"symbols-{stage}.json",
                f"symbols {stage}: inventory path changed")
        inventory_value = read_json(inventory_path)
        require(inventory_value.get("stage") == stage,
                f"symbols {stage}: inventory stage changed")
        same_descriptor(inventory_value.get("binary"), binaries[stage],
                        f"symbols {stage} inventory binary")
        records = inventory_value.get("records")
        require(isinstance(records, dict) and set(records) >= {"raw", "demangled", "elf"},
                f"symbols {stage}: command inventory missing")
        for record_name, record in records.items():
            require(isinstance(record, dict),
                    f"symbols {stage}/{record_name}: malformed record")
            for field in ("output", "receipt"):
                descriptor_value(record.get(field),
                                 f"symbols {stage}/{record_name}.{field}")
        demangled_output = records["demangled"]["output"]
        demangled_path = packet_path(demangled_output["path"],
                                     f"symbols {stage}/demangled output")
        try:
            demangled_text = demangled_path.read_text(encoding="utf-8", errors="replace")
        except OSError as error:
            fail(f"symbols {stage}: cannot read demangled output: {error}")
        emitted_owner_count = sum(
            1 for line in demangled_text.splitlines()
            if len(line.split()) >= 4 and _symbol_matches(" ".join(line.split()[3:]), owner)
        )
        require(emitted_owner_count > 0,
                f"symbols {stage}: admitted owner is absent from demangled inventory")
        build_id = stage_value.get("build_id")
        require(isinstance(build_id, str) and re.fullmatch(r"[0-9a-fA-F]+", build_id),
                f"symbols {stage}: build ID is malformed")
        ranges = stage_value.get("ranges")
        require(isinstance(ranges, list) and ranges,
                f"symbols {stage}: admitted ranges are missing")
        previous_end = -1
        for index, admitted in enumerate(ranges):
            label = f"symbols.{stage}.ranges[{index}]"
            require(isinstance(admitted, dict), f"{label}: malformed range")
            address = admitted.get("address")
            size = admitted.get("size")
            end = admitted.get("end")
            require(type(address) is int and address >= 0
                    and type(size) is int and size > 0
                    and type(end) is int and end == address + size,
                    f"{label}: range bounds changed")
            require(address >= previous_end, f"{label}: overlapping admitted ranges")
            previous_end = end
            require(_symbol_matches(admitted.get("symbol"), owner),
                    f"{label}: owner symbol changed")
            require(isinstance(admitted.get("raw_symbol"), str)
                    and admitted["raw_symbol"], f"{label}: raw symbol missing")
            raw_descriptor = admitted.get("assembly")
            receipt_descriptor = admitted.get("receipt")
            descriptor_value(raw_descriptor, f"{label}.assembly")
            descriptor_value(receipt_descriptor, f"{label}.receipt")
            assembly_path = packet_path(raw_descriptor["path"], f"{label}.assembly")
            receipt_path = packet_path(receipt_descriptor["path"], f"{label}.receipt")
            require(assembly_path.parent.name == f"symbols-{stage}-owner-{index:02}",
                    f"{label}: assembly command binding changed")
            require(receipt_path.parent == assembly_path.parent
                    and receipt_path.name == "receipt.json",
                    f"{label}: receipt command binding changed")
            if stage == "fp":
                require(admitted.get("frame_pointer_prologue") is True,
                        f"{label}: frame-pointer prologue admission missing")
            else:
                require(admitted.get("frame_pointer_prologue") is False,
                        f"{label}: baseline prologue binding changed")
        checked_ranges[stage] = len(ranges)
    plan_descriptor = value.get("plan")
    same_descriptor(plan_descriptor, descriptor(PACKET / "plan.json", "symbols plan"),
                    "symbols plan")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "status": "pass", "owner": owner,
            "scope": scope, "ranges": checked_ranges}


def _manifest_rows(path: Path, expected_name: str, count: int) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    value = read_json(path)
    require(value.get("status") in ("commands_pass", "pass"),
            f"{expected_name} manifest is not pass")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == count,
            f"{expected_name} manifest row count changed")
    if "report_count" in value:
        require(value["report_count"] == count,
                f"{expected_name} report cardinality changed")
    return value, rows


def validate_profiles(plan: dict[str, Any], binaries: dict[str, Any],
                      seen_pids: set[int]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    path = PACKET / "profiles.json"
    value, rows = _manifest_rows(path, "profiles", 4)
    commands: list[dict[str, Any]] = []
    expected_rows = plan.get("perf")
    require(isinstance(expected_rows, list) and len(expected_rows) == 4,
            "plan profile rows changed")
    parsed_rows = []
    for index, expected in enumerate(expected_rows):
        row = rows[index]
        label = row.get("label", f"profile-{index:02}")
        for key in ("repeat", "stage", "state", "samples", "warmup"):
            require(row.get(key) == expected[key], f"{label}.{key}: profile plan changed")
        item = validate_manifest_row(row, label, expected["stage"], expected["state"],
                                     expected["samples"], expected["warmup"], CASES[0],
                                     binaries, seen_pids, strict_report=True,
                                     capture_kind="profile")
        validate_profile_artifacts(row, label, expected["stage"], "profile")
        _walk_descriptors(row, label)
        parsed_rows.append(item)
        commands.append(validate_command(label, stage=expected["stage"],
                                         expected_argv=item["command_argv"]))
    require(sum(row["samples"] for row in rows) == EXPECTED_COUNTS["perf_samples"],
            "profile sample cardinality changed")
    if "sample_count" in value:
        require(value["sample_count"] == EXPECTED_COUNTS["perf_samples"],
                "profile manifest sample cardinality changed")
    return ({"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
             "sha256": sha256(path), "status": value["status"], "reports": 4,
             "samples": 120, "rows": [{"label": item["label"],
                                        "pids": item["parsed"]["pids"]}
                                       for item in parsed_rows]}, commands)


def validate_traces(plan: dict[str, Any], binaries: dict[str, Any],
                    seen_pids: set[int]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    path = PACKET / "traces.json"
    value, rows = _manifest_rows(path, "traces", 4)
    commands: list[dict[str, Any]] = []
    expected_rows = plan.get("trace")
    require(isinstance(expected_rows, list) and len(expected_rows) == 4,
            "plan trace rows changed")
    parsed_rows = []
    for index, expected in enumerate(expected_rows):
        row = rows[index]
        label = row.get("label", f"trace-{index:02}")
        for key in ("repeat", "stage", "state", "samples", "warmup"):
            require(row.get(key) == expected[key], f"{label}.{key}: trace plan changed")
        item = validate_manifest_row(row, label, expected["stage"], expected["state"],
                                     expected["samples"], expected["warmup"], CASES[0],
                                     binaries, seen_pids, strict_report=True,
                                     capture_kind="trace")
        validate_profile_artifacts(row, label, expected["stage"], "trace")
        _walk_descriptors(row, label)
        parsed_rows.append(item)
        commands.append(validate_command(label, stage=expected["stage"],
                                         expected_argv=item["command_argv"]))
    require(sum(row["samples"] for row in rows) == EXPECTED_COUNTS["trace_samples"],
            "trace sample cardinality changed")
    if "sample_count" in value:
        require(value["sample_count"] == EXPECTED_COUNTS["trace_samples"],
                "trace manifest sample cardinality changed")
    return ({"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
             "sha256": sha256(path), "status": value["status"], "reports": 4,
             "samples": 32, "rows": [{"label": item["label"],
                                        "pids": item["parsed"]["pids"]}
                                       for item in parsed_rows]}, commands)


def validate_analysis_if_present(name: str, expected_reports: int,
                                 expected_samples: int, *,
                                 required: bool = False) -> dict[str, Any] | None:
    path = PACKET / name
    if not path.is_file():
        require(not required, f"missing required analysis: {name}")
        return None
    value = read_json(path)
    require(value.get("status") in ("pass", "ok") or value.get("schema"),
            f"{name}: analysis is not terminal")
    for key, expected in (("reports", expected_reports), ("samples", expected_samples),
                          ("report_count", expected_reports), ("sample_count", expected_samples)):
        if key in value:
            require(value[key] == expected, f"{name}.{key}: cardinality changed")
    # The combined profile analyzer may expose its two populations both as
    # totals and as named counters.  Require whichever representation it
    # provides to agree with the manifest contract.
    for key, expected in (("cpu_reports", SHARED_PROFILE_COUNTS["cpu_reports"]),
                          ("cpu_samples", SHARED_PROFILE_COUNTS["cpu_samples"]),
                          ("profile_reports", SHARED_PROFILE_COUNTS["cpu_reports"]),
                          ("profile_samples", SHARED_PROFILE_COUNTS["cpu_samples"]),
                          ("trace_reports", SHARED_PROFILE_COUNTS["trace_reports"]),
                          ("trace_samples", SHARED_PROFILE_COUNTS["trace_samples"])):
        if key in value:
            require(value[key] == expected, f"{name}.{key}: cardinality changed")
    for holder_name, reports, samples in (("cpu", 4, 120), ("profile", 4, 120),
                                          ("trace", 4, 32), ("syscall", 4, 32)):
        holder = value.get(holder_name)
        if isinstance(holder, dict):
            for key, expected in (("reports", reports), ("report_count", reports),
                                  ("samples", samples), ("sample_count", samples)):
                if key in holder:
                    require(holder[key] == expected,
                            f"{name}.{holder_name}.{key}: cardinality changed")
    if required:
        has_total = (value.get("reports") == expected_reports
                     or value.get("report_count") == expected_reports)
        has_populations = (value.get("cpu_reports", value.get("profile_reports")) == 4
                           and value.get("trace_reports") == 4
                           and value.get("cpu_samples", value.get("profile_samples")) == 120
                           and value.get("trace_samples") == 32)
        nested = (isinstance(value.get("cpu"), dict)
                  and isinstance(value.get("trace"), dict))
        require(has_total or has_populations or nested,
                f"{name}: combined CPU/trace cardinality is missing")
    _walk_descriptors(value, name)
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path)}


def validate_plan() -> dict[str, Any]:
    path = PACKET / "plan.json"
    value = read_json(path)
    require(value.get("schema") == PLAN_SCHEMA and value.get("base") == BASE
            and value.get("case") == CASES[0] and value.get("cpu") == 12
            and value.get("source_changes") is False,
            "profile plan identity changed")
    require(value.get("expected") == EXPECTED_COUNTS, "profile plan counts changed")
    require(value.get("purpose") == "Current-source operation-scoped CPU and durability syscall diagnosis; no optimization adoption",
            "profile plan purpose changed")
    limits = value.get("limits")
    require(isinstance(limits, list) and limits and
            any("No production before/after speedup claim." == item for item in limits),
            "profile plan claim limit missing")
    stats = value.get("statistics")
    require(isinstance(stats, dict) and stats.get("seed") == SEED
            and stats.get("bootstrap_resamples") == BOOTSTRAP
            and stats.get("sorted_endpoints") == list(ENDPOINTS),
            "profile plan statistics changed")
    contract = value.get("perf_contract")
    require(isinstance(contract, dict) and contract.get("event") == "cycles:u"
            and contract.get("frequency_hz") == 997,
            "perf contract changed")
    trace = value.get("trace_contract")
    require(isinstance(trace, dict) and trace.get("syscalls") ==
            ["fsync", "fdatasync", "rename", "renameat", "renameat2"],
            "trace contract changed")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "schema": value["schema"],
            "base": BASE, "expected": EXPECTED_COUNTS}


def validate_admission(name: str, *, required: bool = True) -> dict[str, Any] | None:
    path = PACKET / name
    if not path.is_file():
        require(not required, f"missing {name}")
        return None
    value = read_json(path)
    require(value.get("status") == "pass", f"{name}: not pass")
    inputs = value.get("inputs")
    require(isinstance(inputs, dict) and inputs, f"{name}: inputs missing")
    for raw, binding in inputs.items():
        if isinstance(binding, dict) and "descriptor" in binding:
            binding = binding["descriptor"]
        actual = packet_path(raw, f"{name} input")
        same_descriptor(binding, descriptor(actual, f"{name} input {raw}"),
                        f"{name} input {raw}")
    for key in ("reports", "samples", "validated_sample_count"):
        if key in value:
            require(type(value[key]) is int and value[key] >= 0,
                    f"{name}.{key}: invalid")
    _walk_descriptors(value, name)
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "status": "pass", "input_count": len(inputs),
            "reports": value.get("reports"), "samples": value.get("samples")}


def _command_refs(value: Any, labels: set[str]) -> None:
    if isinstance(value, dict):
        raw_path = value.get("path")
        if isinstance(raw_path, str):
            candidate = Path(raw_path)
            if not candidate.is_absolute():
                candidate = PACKET / candidate
            parts = candidate.parts
            if candidate.is_relative_to(PACKET / "commands") and "commands" in parts:
                index = parts.index("commands")
                if index + 1 < len(parts):
                    labels.add(parts[index + 1])
        for child in value.values():
            _command_refs(child, labels)
    elif isinstance(value, list):
        for child in value:
            _command_refs(child, labels)


def referenced_command_labels() -> set[str]:
    labels: set[str] = set()
    for path in sorted(PACKET.glob("*.json")):
        if path.name == "audit.json":
            continue
        _command_refs(read_json(path), labels)
    # The perf availability probe is intentionally auxiliary: it has no
    # report row, but its exact successful command is part of the retained
    # capture admission and must still be audited.
    if (PACKET / "commands" / "perf-availability").is_dir():
        labels.add("perf-availability")
    return labels


def validate_all_commands(expected_rows: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    command_root = PACKET / "commands"
    require(command_root.is_dir() and not command_root.is_symlink(),
            "command receipt root missing")
    expected_labels = {row["name"] for row in expected_rows}
    actual = sorted(item.name for item in command_root.iterdir()
                    if item.is_dir() and not item.is_symlink())
    referenced = referenced_command_labels()
    require(referenced >= expected_labels,
            "manifest-referenced command receipt is missing")
    require(set(actual) == referenced,
            "command inventory contains an unreferenced or missing command")
    rows: list[dict[str, Any]] = []
    for label in actual:
        expected_argv = None
        if label == "perf-availability":
            expected_argv = ["perf", "stat", "-e", "cycles:u", "--",
                             "taskset", "-c", "12", "true"]
        rows.append(validate_command(label, expected_argv=expected_argv))
    require(rows, "command inventory is empty")
    ordered = sorted(rows, key=lambda row: (row["started_unix"], row["name"]))
    for previous, current in zip(ordered, ordered[1:]):
        require(previous["finished_unix"] <= current["started_unix"],
                f"command intervals overlap: {previous['name']} and {current['name']}")
    return rows


def validate_cleanup_if_present(*, required: bool = False) -> dict[str, Any] | None:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        require(not required, "cleanup witness is missing")
        return None
    value = read_json(path)
    require(value.get("status") == "pass", "cleanup witness is not pass")
    require(value.get("target_absent") is True and value.get("scratch_absent") is True,
            "cleanup root absence witness is incomplete")
    owned = value.get("owned_paths")
    require(isinstance(owned, list)
            and {Path(raw).resolve(strict=False) for raw in owned
                if isinstance(raw, str)} == {TARGET.resolve(), SCRATCH.resolve()}
            and len(owned) == 2,
            "cleanup owned path inventory changed")
    for raw in owned:
        require(isinstance(raw, str) and Path(raw).is_absolute(),
                "cleanup owned path is malformed")
        require(not Path(raw).exists(), f"owned temporary path remains: {raw}")
    roots = value.get("roots")
    require(isinstance(roots, list) and len(roots) == 2,
            "cleanup root inventory changed")
    root_paths: set[Path] = set()
    for index, root in enumerate(roots):
        require(isinstance(root, dict), f"cleanup root {index} is malformed")
        raw = root.get("path")
        require(isinstance(raw, str) and Path(raw).is_absolute(),
                f"cleanup root {index} path is malformed")
        root_path = Path(raw).resolve(strict=False)
        require(root_path in {TARGET.resolve(), SCRATCH.resolve()},
                f"cleanup root {index} is outside owned roots")
        root_paths.add(root_path)
        require(not root_path.exists(), f"cleanup root remains: {raw}")
        marker = root.get("marker", root.get("owner"))
        require(isinstance(marker, dict), f"cleanup root {index} marker is missing")
        descriptor_value(marker, f"cleanup root {index} marker", allow_missing=True)
    require(root_paths == {TARGET.resolve(), SCRATCH.resolve()},
            "cleanup root coverage changed")
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, list) and artifacts,
            "cleanup artifact descriptor inventory is missing")
    for index, item in enumerate(artifacts):
        descriptor_value(item, f"cleanup artifacts[{index}]", allow_missing=True)
    compressed = value.get("compressed")
    require(isinstance(compressed, list) and compressed,
            "cleanup compressed witness inventory is missing")
    for index, item in enumerate(compressed):
        require(isinstance(item, dict), f"cleanup compressed[{index}] is malformed")
        original = descriptor_value(item.get("original"),
                                    f"cleanup compressed[{index}].original",
                                    allow_missing=True, nonempty=True)
        retained = descriptor_value(item.get("retained"),
                                    f"cleanup compressed[{index}].retained",
                                    nonempty=True)
        retained_path = packet_path(retained["path"],
                                    f"cleanup compressed[{index}].retained")
        require(retained_path.suffix == ".gz",
                f"cleanup compressed[{index}]: retained witness is not gzip")
        size, digest = _gzip_digest(retained_path)
        require(size == original["bytes"] and digest == original["sha256"],
                f"cleanup compressed[{index}]: witness differs")
    # This also verifies every target/scratch descriptor retained in the
    # witness, including Cargo fingerprint originals and removed raw captures.
    _walk_descriptors(value, "cleanup")
    return {"path": str(path.relative_to(PACKET)), "bytes": path.stat().st_size,
            "sha256": sha256(path), "status": "pass",
            "artifact_count": len(artifacts), "compressed_count": len(compressed)}


def preflight_components() -> dict[str, Any]:
    """Validate every admission input needed before native/profile capture.

    This is intentionally callable before any formal capture exists.  It
    binds the packet to the committed source, both stage builds and their
    distinct environments, the owner-range admission, and the strict
    qualification reports.  It does not inspect native/profile/trace output
    and never writes an artifact.
    """

    origin = load_origin()
    baseline_freeze = validate_freeze("baseline")
    fp_freeze = validate_freeze("fp")
    require(baseline_freeze is not None and fp_freeze is not None,
            "both baseline and fp freezes are required")
    require(fp_freeze["source"] == baseline_freeze["source"],
            "baseline and fp source inventories differ")
    host = validate_host()
    quality = validate_quality_reuse()
    plan_meta = validate_plan()
    plan = read_json(PACKET / "plan.json")
    builds: dict[str, dict[str, Any]] = {}
    build_commands: list[dict[str, Any]] = []
    for stage, freeze in zip(STAGES, (baseline_freeze, fp_freeze)):
        build, command = validate_build(stage, freeze)
        builds[stage] = build
        build_commands.append(command)
    binaries = {stage: builds[stage]["binary"] for stage in STAGES}
    symbols = validate_symbols(plan, binaries)
    seen_pids: set[int] = set()
    qualification, qualification_commands = validate_qualification(binaries, seen_pids)
    return {"origin": origin, "baseline_freeze": baseline_freeze,
            "fp_freeze": fp_freeze, "host": host, "quality": quality,
            "plan_meta": plan_meta, "plan": plan, "builds": builds,
            "build_commands": build_commands, "binaries": binaries,
            "symbols": symbols, "qualification": qualification,
            "qualification_commands": qualification_commands,
            "seen_pids": seen_pids}


def build_audit(*, require_cleanup: bool = False) -> dict[str, Any]:
    components = preflight_components()
    origin = components["origin"]
    baseline_freeze = components["baseline_freeze"]
    fp_freeze = components["fp_freeze"]
    host = components["host"]
    quality = components["quality"]
    plan_meta = components["plan_meta"]
    plan = components["plan"]
    builds = components["builds"]
    build_commands = components["build_commands"]
    binaries = components["binaries"]
    symbols = components["symbols"]
    qualification = components["qualification"]
    qualification_commands = components["qualification_commands"]
    seen_pids = components["seen_pids"]
    native, native_commands, native_stats = validate_native(plan, binaries, seen_pids)
    profiles, profile_commands = validate_profiles(plan, binaries, seen_pids)
    traces, trace_commands = validate_traces(plan, binaries, seen_pids)
    mappings = validate_mappings(binaries)
    admission = validate_admission("admission.json")
    # Native admission and formal CPU/trace admission are separate gates.  The
    # packet historically called the latter capture-admission.json; accept
    # that spelling as a compatibility alias while preferring the explicit
    # profile-admission.json gate used by this capture.
    profile_admission = validate_admission("profile-admission.json", required=False)
    capture_admission = validate_admission("capture-admission.json", required=False)
    require(profile_admission is not None or capture_admission is not None,
            "formal profile/trace admission is missing")
    # CPU and syscall decoding are one analysis contract.  A separate
    # trace-analysis.json is deliberately not required; retaining one is
    # harmless, but the shared analysis must carry both populations.
    analyses = {
        "profile": validate_analysis_if_present(
            "profile-analysis.json", SHARED_PROFILE_COUNTS["reports"],
            SHARED_PROFILE_COUNTS["samples"], required=True),
        "shared_profile_alias": validate_analysis_if_present(
            "sharedprofile-analysis.json", SHARED_PROFILE_COUNTS["reports"],
            SHARED_PROFILE_COUNTS["samples"], required=False),
        "native": validate_analysis_if_present("native-analysis.json", 24, 288),
    }
    expected_commands = [*build_commands, *qualification_commands, *native_commands,
                         *profile_commands, *trace_commands]
    commands = validate_all_commands(expected_commands)
    require(len(seen_pids) == EXPECTED_COUNTS["samples"],
            "workload reports do not expose one unique PID per sample")
    cleanup = validate_cleanup_if_present(required=require_cleanup)
    return {
        "schema": AUDIT_SCHEMA, "status": "pass", "base": BASE,
        "custody": {"origin": origin, "freeze": {
            "baseline": {key: value for key, value in baseline_freeze.items()
                          if key != "source"},
            "fp": {key: value for key, value in fp_freeze.items() if key != "source"},
            "source_count": len(baseline_freeze["source"]),
            "source_match": True,
        }, "quality_reuse": quality},
        "host": host, "plan": plan_meta, "builds": builds,
        "symbols": symbols, "qualification": qualification,
        "native": {**native, "statistics": native_stats},
        "profiles": profiles, "traces": traces, "mappings": mappings,
        "analysis": analyses, "admission": admission,
        "profile_admission": profile_admission,
        "capture_admission": capture_admission,
        "commands": {"count": len(commands), "serial": True, "rows": commands},
        "counts": EXPECTED_COUNTS,
        "cleanup": cleanup,
        "claims": {"performance_claim": "none",
                    "scope": EXPECTED_SCOPE,
                    "native": "descriptive fp/baseline process timing and RSS controls",
                    "profiles": "descriptive sampled CPU ownership/census",
                    "traces": "descriptive ptrace syscall census",
                    "iwork": "excluded"},
    }


def encoded(value: dict[str, Any]) -> bytes:
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    require(len(data) < 50 * 1024 * 1024, "audit output unexpectedly large")
    return data


def write_or_check(value: dict[str, Any], check: bool, preview: bool) -> None:
    path = PACKET / "audit.json"
    data = encoded(value)
    if preview:
        return
    if check or path.exists():
        require(path.is_file() and not path.is_symlink(), "audit.json missing")
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
    modes.add_argument("--check", action="store_true")
    modes.add_argument("--preview", action="store_true")
    modes.add_argument("--preflight", action="store_true",
                       help="validate source/build/owner/qualification admission only")
    args = parser.parse_args(argv)
    try:
        if args.preflight:
            preflight_components()
        else:
            write_or_check(build_audit(require_cleanup=not args.preview),
                           args.check, args.preview)
    except (AuditError, AssertionError, OSError, UnicodeError, ValueError,
            KeyError, TypeError, IndexError) as error:
        print(f"0836 independent audit failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps({"status": "pass", "mode": "preflight" if args.preflight else "audit",
                      "reports": EXPECTED_COUNTS["reports"],
                      "samples": EXPECTED_COUNTS["samples"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
