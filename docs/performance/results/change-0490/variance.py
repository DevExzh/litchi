#!/usr/bin/env python3
"""Bounded block-paired variance evidence for the smallest OPC route.

This is a self-contained evidence driver for the 0490 follow-up to the 0489
candidate-audit measurement.  It deliberately imports the sealed 0484 route
validator, but never changes that module's globals or writes in its directory.
The only files written by this module are the protocol, captures, summaries,
and validation receipts below ``change-0490`` (plus short-lived replay files
under ``/home/zhuhe/.cache/litchi-goal-0490``).

The formal design has six paired blocks.  Odd blocks run ``before`` then
``after``; even blocks run ``after`` then ``before``.  The role/route order is
reversed in even blocks.  Each block contains both roles and all three route
providers, giving 72 children and 4,320 measured samples.  Analysis first
reduces each child to its within-child p50/p95/p99, then pairs those summaries
within block.  Bootstrap resamples six blocks, never the 60 samples inside a
child, so the interval does not pretend that in-process samples are
independent processes.
"""

from __future__ import annotations

import argparse
from contextlib import contextmanager
import copy
import datetime
import fcntl
import hashlib
import json
import math
import os
from pathlib import Path
import random
import re
import signal
import statistics
import subprocess
import sys
from typing import Any, Iterable, Mapping


# Import the sealed 0484 helper set first.  Its ``measure_routes`` module
# imports ``measure``, ``common``, and ``corpus_oracle`` by basename, so the
# sealed directory must lead sys.path before any of those names are imported.
REPO = Path(__file__).resolve().parents[4]
ROOT = Path(__file__).resolve().parent
ROUTE_ROOT = REPO / "docs" / "performance" / "results" / "change-0484"
BASELINE_ROOT = REPO / "docs" / "performance" / "results" / "change-0487"
CANDIDATE_ROOT = REPO / "docs" / "performance" / "results" / "change-0489"
if str(ROUTE_ROOT) not in sys.path:
    sys.path.insert(0, str(ROUTE_ROOT))
import measure_routes as old_routes  # noqa: E402


SCHEMA = "docx-opc-candidate-audit-variance-v1"
VERSION = 1
PROTOCOL_FILE = "variance-protocol.json"
SUMMARY_FILE = "variance-summary.json"
SUMMARY_SCHEMA = "docx-opc-candidate-audit-variance-summary-v1"
CAPTURE_SCHEMA = "docx-opc-candidate-audit-variance-capture-v1"
TERMINAL_SCHEMA = "docx-opc-candidate-audit-variance-terminal-v1"
REPLAY_CLEANUP_SCHEMA = "docx-opc-candidate-audit-variance-replay-cleanup-v1"

TEMP = Path("/home/zhuhe/.cache/litchi-goal-0490")
CPU_LOCK = Path("/home/zhuhe/.cache/litchi-goal-0484/cpu.lock")
CPU = 2
CASE_LABEL = "s64-a64-short-c64"
ROLES = ("normal", "allocator")
ROUTES = ("deterministic", "memory_store", "file_store")
PHASES = ("before", "after")
BLOCKS = tuple(range(1, 7))
FORMAL_SAMPLES = 60
FORMAL_WARMUPS = 5
CAPTURE_TIMEOUT_SECONDS = 180
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_SEED = 490
REVIEW_THRESHOLD_PERCENT = 5.0
ATTEMPT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")

BEFORE_SOURCE_SHA256 = "cb69be7b5a1f443191a27000fdfab285b707ba28517cccc3ee764d833cde0ba2"
AFTER_SOURCE_SHA256 = "0f5d9a455d25146823a7a8eb3620712f6e8f199abe0d40c32678d100c3994afa"
AFTER_SOURCE_MANIFEST = CANDIDATE_ROOT / "validation-sources" / f"{AFTER_SOURCE_SHA256}.json"

# These are the direct old validator and its imported dependencies.  Their
# hashes are frozen, then rechecked before every capture and analysis.
SEALED_HELPERS = {
    "measure_routes.py": ROUTE_ROOT / "measure_routes.py",
    "measure.py": ROUTE_ROOT / "measure.py",
    "common.py": ROUTE_ROOT / "common.py",
    "corpus_oracle.py": ROUTE_ROOT / "corpus_oracle.py",
}

# Every value here is a per-child observation or an allocator counter.  RSS
# from GNU time is added separately as a one-observation metric per child.
SAMPLE_METRICS = (
    "elapsed_ns",
    "process_rss_bytes",
    "process_peak_rss_bytes",
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
    "operation_peak_increment_bytes",
)
METRICS = SAMPLE_METRICS + ("time_max_rss_bytes",)


class VarianceError(RuntimeError):
    """A protocol, receipt, report, or analysis invariant failed."""


def fail(message: str) -> None:
    raise VarianceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _finite(value: Any, path: str = "value") -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            _finite(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{path}[{index}]")


def read_json(path: Path) -> Any:
    try:
        value = json.loads(
            path.read_text(encoding="utf-8"),
            parse_constant=lambda token: (_ for _ in ()).throw(ValueError(token)),
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        raise VarianceError(f"cannot read JSON {path}: {error}") from error
    _finite(value, str(path))
    return value


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8", newline="\n") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except FileExistsError as error:
        raise VarianceError(f"refusing to replace existing artifact: {path}") from error


def sha(path: Path | str) -> str:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    try:
        with path.open("rb") as stream:
            return hashlib.file_digest(stream, "sha256").hexdigest()
    except OSError as error:
        raise VarianceError(f"cannot hash {path}: {error}") from error


def metadata(path: Path | str, label: str = "file") -> dict[str, Any]:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"{label}: missing regular file {path}")
    try:
        details = {"bytes": path.stat().st_size, "sha256": sha(path)}
    except OSError as error:
        raise VarianceError(f"{label}: cannot stat {path}: {error}") from error
    require(details["bytes"] >= 0, f"{label}: negative size")
    return {"path": str(path), **details}


def now() -> str:
    # Avoid importing the old common module's clock so this helper remains
    # independently hashable and its receipts stay self-describing.
    import datetime

    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def _attempt(value: str) -> str:
    require(ATTEMPT_RE.fullmatch(value) is not None, "attempt must be a path-safe token")
    return value


def _sha_field(value: Any, path: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None, f"{path}: invalid SHA-256")
    return value


def _int(value: Any, path: str, *, positive: bool = False) -> int:
    require(type(value) is int and (value > 0 if positive else value >= 0), f"{path}: invalid integer")
    return value


def _timestamp(value: Any, path: str) -> datetime.datetime:
    require(isinstance(value, str) and value, f"{path}: timestamp is missing")
    try:
        parsed = datetime.datetime.fromisoformat(value)
    except ValueError as error:
        raise VarianceError(f"{path}: malformed timestamp") from error
    require(parsed.tzinfo is not None, f"{path}: timestamp must include a timezone")
    return parsed


def _obj(value: Any, path: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{path}: expected object")
    return value


def _path_relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def _sealed_helper_bindings() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for name, path in SEALED_HELPERS.items():
        result[name] = {"path": str(path), "sha256": sha(path), "bytes": path.stat().st_size}
    return result


def _local_bindings() -> dict[str, Any]:
    result: dict[str, Any] = {"variance.py": {"path": str(Path(__file__)), "sha256": sha(Path(__file__))}}
    for name in ("machine.json", "provenance.json", "adr-refresh.json"):
        path = ROOT / name
        require(path.is_file() and not path.is_symlink(), f"required custody record is missing: {path}")
        result[name] = {"path": _path_relative(path), "sha256": sha(path), "bytes": path.stat().st_size}
    return result


def _environment() -> dict[str, str]:
    # The benchmark uses the exact 0484 route environment.  This deliberately
    # excludes the 0489 Cargo target directory: no build is performed here.
    return {key: old_routes.base.ENV[key] for key in old_routes.base.ENV_KEYS}


def source_names() -> list[str]:
    """Return the same source inventory used by the 0489 custody helper."""

    try:
        output = subprocess.check_output(
            ["git", "ls-files", "-co", "--exclude-standard", "-z"], cwd=REPO,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        raise VarianceError(f"cannot enumerate source files: {error}") from error
    selected: list[str] = []
    for name in sorted(set(output.decode("utf-8").split("\0"))):
        if not name:
            continue
        path = Path(name)
        is_source = path.suffix in {".rs", ".toml", ".lock"}
        is_template = name.startswith("crates/") and "/src/" in name and path.suffix == ".xml"
        if (is_source or is_template) and (REPO / path).is_file():
            selected.append(name)
    return selected


def current_source_manifest() -> dict[str, Any]:
    values = {name: sha(REPO / name) for name in source_names()}
    encoded = (json.dumps(values, indent=2, sort_keys=True) + "\n").encode("utf-8")
    digest = hashlib.sha256(encoded).hexdigest()
    return {"sha256": digest, "files": len(values), "entries": values}


def _manifest_record(path: Path, expected_sha: str, expected_files: int) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"source manifest is missing: {path}")
    require(sha(path) == expected_sha, f"source manifest hash differs: {path}")
    value = read_json(path)
    require(isinstance(value, dict), f"source manifest must be an object: {path}")
    require(len(value) == expected_files, f"source manifest file count differs: {path}")
    for name, digest in value.items():
        require(isinstance(name, str) and isinstance(digest, str) and SHA256_RE.fullmatch(digest) is not None, f"malformed source manifest entry {name!r}")
    return {"path": str(path), "sha256": expected_sha, "files": expected_files}


def _verify_current_source(expected_sha: str = AFTER_SOURCE_SHA256) -> dict[str, Any]:
    manifest = current_source_manifest()
    require(manifest["sha256"] == expected_sha, f"current source manifest differs from sealed 0489 source: {manifest['sha256']}")
    require(manifest["files"] == 7157, f"current source manifest file count differs: {manifest['files']}")
    expected = read_json(AFTER_SOURCE_MANIFEST)
    require(manifest["entries"] == expected, "current source entries differ from 0489 manifest")
    return {"path": str(AFTER_SOURCE_MANIFEST), "sha256": expected_sha, "files": manifest["files"]}


def _machine_binding() -> dict[str, Any]:
    path = ROOT / "machine.json"
    value = _obj(read_json(path), str(path))
    require(value.get("schema") == "docx-stream-route-machine-v1", f"{path}: machine schema differs")
    require(value.get("selected_cpu") == CPU, f"{path}: selected CPU differs")
    affinity = value.get("coordinator_affinity")
    require(isinstance(affinity, list) and CPU in affinity, f"{path}: selected CPU is outside coordinator affinity")
    scratch = _obj(value.get("scratch"), f"{path}.scratch")
    require(scratch.get("path") == str(TEMP.resolve()), f"{path}: scratch path differs")
    commands = _obj(value.get("commands"), f"{path}.commands")
    for name in ("cpu", "rustc", "cargo", "time", "scratch_mount"):
        command = _obj(commands.get(name), f"{path}.commands.{name}")
        require(command.get("status") == "pass" and command.get("exit_code") == 0, f"{path}: {name} observation failed")
    return {"path": _path_relative(path), "sha256": sha(path), "bytes": path.stat().st_size}


def _provenance_binding() -> dict[str, Any]:
    result: dict[str, Any] = {}
    for name in ("provenance.json", "adr-refresh.json"):
        path = ROOT / name
        _obj(read_json(path), str(path))
        result[name] = {"path": _path_relative(path), "sha256": sha(path), "bytes": path.stat().st_size}
    return result


def _build_record_path(phase: str, role: str) -> Path:
    root = BASELINE_ROOT if phase == "before" else CANDIDATE_ROOT
    return root / f"build-{role}.json"


def _expected_source_for_phase(phase: str) -> str:
    return BEFORE_SOURCE_SHA256 if phase == "before" else AFTER_SOURCE_SHA256


def _validate_build_record(path: Path, phase: str, role: str) -> dict[str, Any]:
    value = _obj(read_json(path), str(path))
    require(value.get("schema") == old_routes.base.BUILD_SCHEMA, f"{path}: build schema differs")
    require(value.get("role") == role and value.get("attempt") == "after1", f"{path}: build identity differs")
    require(value.get("source_unchanged") is True and value.get("source_before") == value.get("source_after"), f"{path}: source custody failed")
    source = _obj(value.get("source_after"), f"{path}.source_after")
    require(source.get("sha256") == _expected_source_for_phase(phase) and source.get("files") == 7157, f"{path}: source binding differs")
    source_path = Path(str(source.get("path", "")))
    if not source_path.is_absolute():
        source_path = path.parent / source_path
    _manifest_record(source_path, str(source["sha256"]), int(source["files"]))
    gate = _obj(value.get("gate"), f"{path}.gate")
    gate_path = Path(str(gate.get("path", "")))
    require(gate_path.is_file() and not gate_path.is_symlink(), f"{path}: build gate receipt is missing")
    require(_sha_field(gate.get("sha256"), f"{path}.gate.sha256") == sha(gate_path), f"{path}: build gate receipt hash changed")
    gate_value = _obj(read_json(gate_path), str(gate_path))
    require(gate_value.get("schema") == "docx-stream-append-gate-v1", f"{path}: build gate schema differs")
    require(gate_value.get("exit_code") == 0, f"{path}: build gate did not pass")
    require(gate_value.get("source_unchanged") is True, f"{path}: build gate source custody failed")
    require(gate_value.get("source_before") == value.get("source_before"), f"{path}: build gate source-before differs")
    require(gate_value.get("source_after") == value.get("source_after"), f"{path}: build gate source-after differs")
    require(gate_value.get("source_before") == gate_value.get("source_after"), f"{path}: build gate source changed")
    require(gate_value.get("cwd") == str(REPO), f"{path}: build gate cwd differs")
    require(gate_value.get("argv") == value.get("command"), f"{path}: build gate command differs")
    require(gate_value.get("environment") == value.get("environment"), f"{path}: build gate environment differs")
    binary = _obj(value.get("binary"), f"{path}.binary")
    binary_path = Path(str(binary.get("path", "")))
    actual = metadata(binary_path, f"{path}.binary")
    expected = {"path": str(binary_path), "bytes": binary.get("bytes"), "sha256": binary.get("sha256"), "executable": True}
    require(binary == expected and actual["bytes"] == binary["bytes"] and actual["sha256"] == binary["sha256"], f"{path}: binary metadata changed")
    require(os.access(binary_path, os.X_OK), f"{path}: binary is not executable")
    return value


def _build_bindings() -> dict[str, dict[str, dict[str, Any]]]:
    result: dict[str, dict[str, dict[str, Any]]] = {"before": {}, "after": {}}
    for phase in PHASES:
        for role in ROLES:
            path = _build_record_path(phase, role)
            value = _validate_build_record(path, phase, role)
            binary = _obj(value["binary"], f"{path}.binary")
            result[phase][role] = {
                "phase": phase,
                "role": role,
                "record_path": str(path),
                "record_sha256": sha(path),
                "record_bytes": path.stat().st_size,
                "source": copy.deepcopy(value["source_after"]),
                "gate": copy.deepcopy(value["gate"]),
                "binary": copy.deepcopy(binary),
            }
    for role in ROLES:
        require(result["before"][role]["source"]["sha256"] == BEFORE_SOURCE_SHA256, "before source binding differs")
        require(result["after"][role]["source"]["sha256"] == AFTER_SOURCE_SHA256, "after source binding differs")
    return result


def _base_pairs() -> list[dict[str, str]]:
    return [{"role": role, "route": route} for role in ROLES for route in ROUTES]


def formal_runs() -> list[dict[str, Any]]:
    runs: list[dict[str, Any]] = []
    index = 0
    for block in BLOCKS:
        phases = PHASES if block % 2 else tuple(reversed(PHASES))
        pairs = _base_pairs() if block % 2 else list(reversed(_base_pairs()))
        for phase in phases:
            for pair in pairs:
                index += 1
                label = f"b{block:02d}-{phase}-{pair['role']}-{pair['route']}-{CASE_LABEL}"
                runs.append({
                    "index": index,
                    "block": block,
                    "phase": phase,
                    "role": pair["role"],
                    "route": pair["route"],
                    "case": CASE_LABEL,
                    "label": label,
                })
    return runs


def pilot_runs() -> list[dict[str, Any]]:
    return []


def _protocol_value(*, freeze_source: bool) -> dict[str, Any]:
    source = _verify_current_source()
    builds = _build_bindings()
    machine = _machine_binding()
    provenance = _provenance_binding()
    value: dict[str, Any] = {
        "schema": SCHEMA,
        "version": VERSION,
        "change": 490,
        "claim_authorized": False,
        "performance_claim": "none_until_block_paired_review",
        "comparison": "0489 small source/authored file-store tail variance follow-up",
        "scope": "one exact s64-a64-short-c64 workload over deterministic, memory-store, and data-synced file-store routes",
        "cpu": CPU,
        "case": dict(old_routes.ROUTE_CASE_BY_LABEL[CASE_LABEL]),
        "routes": list(ROUTES),
        "roles": list(ROLES),
        "phases": list(PHASES),
        "blocks": list(BLOCKS),
        "formal": {"samples": FORMAL_SAMPLES, "warmups": FORMAL_WARMUPS},
        "capture_timeout_seconds": CAPTURE_TIMEOUT_SECONDS,
        "expected_formal_processes": len(formal_runs()),
        "expected_formal_samples": len(formal_runs()) * FORMAL_SAMPLES,
        "formal_runs": formal_runs(),
        "pilot_runs": pilot_runs(),
        "ordering": {
            "odd_blocks": {"phase_order": ["before", "after"], "pair_order": _base_pairs()},
            "even_blocks": {"phase_order": ["after", "before"], "pair_order": list(reversed(_base_pairs()))},
        },
        "bootstrap": {
            "unit": "block",
            "paired_blocks": len(BLOCKS),
            "resamples": BOOTSTRAP_RESAMPLES,
            "seed": BOOTSTRAP_SEED,
            "confidence": 0.95,
            "method": "ordinary_resampling_with_replacement_of_six_block_differences",
            "sample_pseudoreplication": False,
        },
        "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
        "environment": _environment(),
        "scratch": {"path": str(TEMP.resolve()), "replay_policy": "one-directory-per-child-empty-then-remove"},
        "cpu_lock": str(CPU_LOCK),
        "source": {
            "sealed_0489_manifest": source,
            "before_manifest_sha256": BEFORE_SOURCE_SHA256,
            "after_manifest_sha256": AFTER_SOURCE_SHA256,
        },
        "builds": builds,
        "machine": machine,
        "custody": provenance,
        "helpers": {
            "local": _local_bindings(),
            "sealed_0484": _sealed_helper_bindings(),
        },
    }
    if freeze_source:
        value["frozen_source"] = source
    return value


def protocol_path() -> Path:
    return ROOT / PROTOCOL_FILE


def _assert_binding_file(record: Mapping[str, Any], label: str) -> None:
    path = Path(str(record.get("path", "")))
    if not path.is_absolute():
        path = ROOT / path
    require(path.is_file() and not path.is_symlink(), f"{label}: bound file is missing")
    require(sha(path) == record.get("sha256"), f"{label}: bound file changed")


def _validate_helper_bindings(protocol: Mapping[str, Any]) -> None:
    helpers = _obj(protocol.get("helpers"), "protocol.helpers")
    local = _obj(helpers.get("local"), "protocol.helpers.local")
    require(local == _local_bindings(), "local helper/custody bindings changed after freeze")
    for name, record in local.items():
        _assert_binding_file(_obj(record, f"protocol.helpers.local.{name}"), f"protocol.helpers.local.{name}")
    sealed = _obj(helpers.get("sealed_0484"), "protocol.helpers.sealed_0484")
    current = _sealed_helper_bindings()
    require(sealed == current, "sealed 0484 helper bindings changed after freeze")


def _validate_protocol(value: Mapping[str, Any]) -> None:
    require(value.get("schema") == SCHEMA and value.get("version") == VERSION, "variance protocol schema/version differs")
    require(value.get("claim_authorized") is False, "variance protocol must remain claim-disabled")
    require(value.get("performance_claim") == "none_until_block_paired_review", "variance claim policy differs")
    case = _obj(value.get("case"), "protocol.case")
    require(case.get("label") == CASE_LABEL, "variance case differs")
    require(value.get("routes") == list(ROUTES) and value.get("roles") == list(ROLES), "variance route/role matrix differs")
    require(value.get("phases") == list(PHASES) and value.get("blocks") == list(BLOCKS), "variance phase/block matrix differs")
    formal = _obj(value.get("formal"), "protocol.formal")
    require(formal.get("samples") == FORMAL_SAMPLES and formal.get("warmups") == FORMAL_WARMUPS, "formal sample contract differs")
    require(value.get("capture_timeout_seconds") == CAPTURE_TIMEOUT_SECONDS, "capture timeout differs")
    runs = value.get("formal_runs")
    require(runs == formal_runs(), "formal run inventory/order differs")
    require(value.get("expected_formal_processes") == 72 and value.get("expected_formal_samples") == 4320, "formal cardinality differs")
    bootstrap = _obj(value.get("bootstrap"), "protocol.bootstrap")
    require(bootstrap.get("unit") == "block" and bootstrap.get("paired_blocks") == 6, "bootstrap unit differs")
    require(bootstrap.get("resamples") == BOOTSTRAP_RESAMPLES and bootstrap.get("seed") == BOOTSTRAP_SEED, "bootstrap settings differ")
    require(bootstrap.get("sample_pseudoreplication") is False, "sample pseudoreplication was enabled")
    require(value.get("environment") == _environment(), "capture environment differs from frozen environment")
    source = _obj(value.get("source"), "protocol.source")
    require(source.get("after_manifest_sha256") == AFTER_SOURCE_SHA256, "after source manifest binding differs")
    require(source.get("sealed_0489_manifest") == _verify_current_source(AFTER_SOURCE_SHA256), "sealed 0489 source manifest binding differs")
    _validate_helper_bindings(value)
    machine = _obj(value.get("machine"), "protocol.machine")
    require(machine == _machine_binding(), "machine binding changed after freeze")
    _assert_binding_file(machine, "protocol.machine")
    custody = _obj(value.get("custody"), "protocol.custody")
    require(custody == _provenance_binding(), "provenance/ADR bindings changed after freeze")
    for name in ("provenance.json", "adr-refresh.json"):
        _assert_binding_file(_obj(custody.get(name), f"protocol.custody.{name}"), f"protocol.custody.{name}")
    builds = _obj(value.get("builds"), "protocol.builds")
    current_builds = _build_bindings()
    require(builds == current_builds, "exact before/after build or binary binding changed")


def load_protocol() -> tuple[dict[str, Any], str]:
    path = protocol_path()
    require(path.is_file() and not path.is_symlink(), f"variance protocol is missing: {path}")
    value = _obj(read_json(path), str(path))
    _validate_protocol(value)
    return value, sha(path)


def freeze() -> None:
    path = protocol_path()
    require(not path.exists(), f"refusing to replace existing protocol: {path}")
    value = _protocol_value(freeze_source=True)
    write_json(path, value)
    print(f"wrote {path} with {len(value['formal_runs'])} formal processes and {value['expected_formal_samples']} samples")


def _run_directory(attempt: str, run: Mapping[str, Any]) -> Path:
    path = ROOT / "captures" / attempt / f"block-{int(run['block']):02d}" / str(run["label"])
    require(not path.exists(), f"refusing to replace capture directory: {path}")
    path.mkdir(parents=True)
    return path


def _binary_for(protocol: Mapping[str, Any], phase: str, role: str) -> dict[str, Any]:
    builds = _obj(protocol.get("builds"), "protocol.builds")
    phase_builds = _obj(builds.get(phase), f"protocol.builds.{phase}")
    binding = _obj(phase_builds.get(role), f"protocol.builds.{phase}.{role}")
    binary = copy.deepcopy(_obj(binding.get("binary"), f"protocol.builds.{phase}.{role}.binary"))
    actual = metadata(Path(str(binary.get("path", ""))), "capture binary")
    require(actual["bytes"] == binary.get("bytes") and actual["sha256"] == binary.get("sha256") and os.access(Path(binary["path"]), os.X_OK), "capture binary changed")
    return binary


def _argv_for(run: Mapping[str, Any], binary: Mapping[str, Any], directory: Path, replay_dir: Path | None) -> list[str]:
    case = dict(old_routes.ROUTE_CASE_BY_LABEL[CASE_LABEL])
    spec = old_routes.ROUTE_BY_NAME[str(run["route"])]
    return old_routes._route_argv(
        dict(binary), case, spec,
        samples=FORMAL_SAMPLES,
        warmups=FORMAL_WARMUPS,
        report=directory / "report.json",
        resource=directory / "resource.txt",
        replay_dir=replay_dir,
    )


@contextmanager
def capture_lock() -> Iterable[None]:
    TEMP.mkdir(parents=True, exist_ok=True)
    CPU_LOCK.parent.mkdir(parents=True, exist_ok=True)
    with CPU_LOCK.open("a") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        yield


def _replay_cleanup(replay_dir: Path | None, run_dir: Path) -> dict[str, Any]:
    if replay_dir is None:
        value = {"schema": REPLAY_CLEANUP_SCHEMA, "status": "not_applicable", "path": None, "entries": [], "removed": False}
    else:
        require(replay_dir.is_dir() and not replay_dir.is_symlink(), f"replay directory is missing: {replay_dir}")
        entries = sorted(item.name for item in replay_dir.iterdir())
        require(not entries, f"file-store replay directory is not empty: {replay_dir}")
        replay_dir.rmdir()
        value = {"schema": REPLAY_CLEANUP_SCHEMA, "status": "removed", "path": str(replay_dir), "entries": entries, "removed": True}
    write_json(run_dir / "replay-cleanup.json", value)
    return value


def _resource_rss(path: Path) -> int:
    prefix = "Maximum resident set size (kbytes):"
    values: list[int] = []
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        raise VarianceError(f"cannot read resource receipt {path}: {error}") from error
    for line in lines:
        if line.strip().startswith(prefix):
            raw = line.strip()[len(prefix):].strip()
            require(raw.isdigit(), f"{path}: malformed RSS")
            values.append(int(raw) * 1024)
    require(len(values) == 1, f"{path}: expected one maximum RSS observation")
    return values[0]


def _identity(observed: Mapping[str, Any], label: str) -> dict[str, Any]:
    source = _obj(observed.get("source"), f"{label}.source")
    authored = _obj(observed.get("authored"), f"{label}.authored")
    oracle = _obj(observed.get("oracle"), f"{label}.oracle")
    source_id = {
        "archive_bytes": _int(source.get("archive_bytes"), f"{label}.source.archive_bytes", positive=True),
        "archive_sha256": _sha_field(source.get("archive_sha256"), f"{label}.source.archive_sha256"),
        "main_xml_bytes": _int(source.get("main_xml_bytes"), f"{label}.source.main_xml_bytes", positive=True),
        "main_xml_sha256": _sha_field(source.get("main_xml_sha256"), f"{label}.source.main_xml_sha256"),
    }
    authored_id = {
        "encoded_xml_bytes": _int(authored.get("encoded_xml_bytes"), f"{label}.authored.encoded_xml_bytes", positive=True),
        "expected_event_sha256": _sha_field(authored.get("expected_event_sha256"), f"{label}.authored.expected_event_sha256"),
        "expected_encoded_sha256": _sha_field(authored.get("expected_encoded_sha256"), f"{label}.authored.expected_encoded_sha256"),
    }
    candidate_id = {
        "archive_bytes": _int(oracle.get("candidate_archive_bytes"), f"{label}.oracle.candidate_archive_bytes", positive=True),
        "archive_sha256": _sha_field(oracle.get("candidate_archive_sha256"), f"{label}.oracle.candidate_archive_sha256"),
        "main_xml_bytes": _int(oracle.get("candidate_main_xml_bytes"), f"{label}.oracle.candidate_main_xml_bytes", positive=True),
        "main_xml_sha256": _sha_field(oracle.get("candidate_main_xml_sha256"), f"{label}.oracle.candidate_main_xml_sha256"),
    }
    for name in (
        "candidate_xml_exact", "candidate_semantic_exact", "untouched_member_metadata_exact",
        "untouched_raw_members_preserved", "physical_order_exact", "opaque_member_exact",
        "source_unchanged", "inverse_exact",
    ):
        require(oracle.get(name) is True, f"{label}.oracle.{name}: oracle failed")
    return {"source": source_id, "authored": authored_id, "candidate": candidate_id}


def _sample_values(report: Mapping[str, Any], resource: Path, role: str, label: str) -> tuple[dict[str, Any], dict[str, list[int]], int]:
    cases = report.get("cases")
    require(isinstance(cases, list) and len(cases) == 1, f"{label}: report case cardinality differs")
    observed = _obj(cases[0], f"{label}.cases[0]")
    identities = _identity(observed, label)
    samples = observed.get("samples")
    require(isinstance(samples, list) and len(samples) == FORMAL_SAMPLES, f"{label}: expected {FORMAL_SAMPLES} samples")
    values: dict[str, list[int]] = {name: [] for name in SAMPLE_METRICS}
    for index, raw in enumerate(samples):
        sample = _obj(raw, f"{label}.samples[{index}]")
        require(sample.get("sample") == index, f"{label}.samples[{index}]: sample index differs")
        values["elapsed_ns"].append(_int(sample.get("elapsed_ns"), f"{label}.samples[{index}].elapsed_ns", positive=True))
        process = _obj(sample.get("process"), f"{label}.samples[{index}].process")
        values["process_rss_bytes"].append(_int(process.get("rss_bytes"), f"{label}.samples[{index}].process.rss_bytes"))
        values["process_peak_rss_bytes"].append(_int(process.get("peak_rss_bytes"), f"{label}.samples[{index}].process.peak_rss_bytes"))
        require(values["process_peak_rss_bytes"][-1] >= values["process_rss_bytes"][-1], f"{label}.samples[{index}]: process peak below RSS")
        allocation = sample.get("allocation")
        if role == "normal":
            require(allocation is None, f"{label}.samples[{index}]: normal run has allocator counters")
        else:
            alloc = _obj(allocation, f"{label}.samples[{index}].allocation")
            require(alloc.get("status") == "measured" and alloc.get("scope") == "operation_global_system_allocator", f"{label}.samples[{index}]: allocator identity differs")
            for name in SAMPLE_METRICS[3:14]:
                values[name].append(_int(alloc.get(name), f"{label}.samples[{index}].allocation.{name}"))
            operation = values["region_peak_live_bytes"][-1] - values["live_bytes_before"][-1]
            require(operation >= 0, f"{label}.samples[{index}]: negative operation peak increment")
            values["operation_peak_increment_bytes"].append(operation)
        sink = _obj(sample.get("sink"), f"{label}.samples[{index}].sink")
        require(sink.get("accepted_bytes") == identities["candidate"]["archive_bytes"], f"{label}.samples[{index}]: candidate archive bytes differ")
        require(sink.get("sha256") == identities["candidate"]["archive_sha256"], f"{label}.samples[{index}]: candidate archive hash differs")
    if role == "normal":
        for name in SAMPLE_METRICS[3:]:
            values.pop(name)
    rss = _resource_rss(resource)
    return identities, values, rss


def _stats(values: Iterable[int | float]) -> dict[str, int | float]:
    numeric = [float(value) for value in values]
    require(numeric and all(math.isfinite(value) for value in numeric), "cannot summarize empty/non-finite vector")
    ordered = sorted(numeric)
    return {
        "n": len(ordered),
        "min": ordered[0],
        "max": ordered[-1],
        "mean": statistics.fmean(ordered),
        "p50": _percentile(ordered, 0.50),
        "p95": _percentile(ordered, 0.95),
        "p99": _percentile(ordered, 0.99),
    }


def _percentile(values: list[float], fraction: float) -> float:
    require(values, "cannot calculate percentile of empty vector")
    ordered = sorted(values)
    if len(ordered) == 1:
        return float(ordered[0])
    position = (len(ordered) - 1) * fraction
    low = math.floor(position)
    high = math.ceil(position)
    if low == high:
        return float(ordered[low])
    ratio = position - low
    return float(ordered[low] + (ordered[high] - ordered[low]) * ratio)


def _percent_change(before: float, after: float) -> float | None:
    if before == 0:
        return 0.0 if after == 0 else None
    return (after - before) / abs(before) * 100.0


def _bootstrap(values: list[float]) -> dict[str, Any] | None:
    if not values:
        return None
    rng = random.Random(BOOTSTRAP_SEED)
    resampled: list[float] = []
    for _ in range(BOOTSTRAP_RESAMPLES):
        resampled.append(statistics.fmean(values[rng.randrange(len(values))] for _ in values))
    resampled.sort()
    return {
        "statistic": "mean",
        "n_blocks": len(values),
        "seed": BOOTSTRAP_SEED,
        "resamples": BOOTSTRAP_RESAMPLES,
        "ci95": [_percentile(resampled, 0.025), _percentile(resampled, 0.975)],
    }


def _artifact_inventory(run_dir: Path) -> dict[str, dict[str, Any]]:
    # terminal.json records this inventory; including its own hash would make
    # the receipt recursively self-referential.
    names = ("report.json", "resource.txt", "stdout.txt", "stderr.txt", "started.json", "replay-cleanup.json")
    result: dict[str, dict[str, Any]] = {}
    for name in names:
        result[name] = metadata(run_dir / name, f"capture artifact {name}")
    return result


def _validate_terminal_custody(
    started: Mapping[str, Any],
    terminal: Mapping[str, Any],
    *,
    attempt: str,
    run: Mapping[str, Any],
    protocol_hash: str,
    protocol: Mapping[str, Any],
    binary: Mapping[str, Any],
    argv: list[str],
) -> tuple[datetime.datetime, datetime.datetime]:
    """Validate the immutable header and successful terminal outcome.

    Keeping this check independent makes a malformed or timed-out terminal
    receipt testable without constructing a full 72-child capture tree.
    """

    require(started.get("schema") == CAPTURE_SCHEMA, "started receipt schema differs")
    require(terminal.get("schema") == TERMINAL_SCHEMA, "terminal receipt schema differs")
    require(started.get("version") == 1, "started receipt version differs")
    require(terminal.get("version") == 1, "terminal receipt version differs")
    require(started.get("status") == "running", "started receipt status differs")
    require(terminal.get("status") == "pass", "terminal receipt status differs")
    for value, name in ((started, "started"), (terminal, "terminal")):
        require(value.get("attempt") == attempt and value.get("run") == dict(run), f"{name} run identity differs")
        require(value.get("protocol") == {"path": PROTOCOL_FILE, "sha256": protocol_hash}, f"{name} protocol binding differs")
        require(value.get("cwd") == str(REPO), f"{name} cwd differs")
        require(value.get("environment") == _environment(), f"{name} environment differs")
        require(value.get("machine") == protocol["machine"], f"{name} machine binding differs")
    expected_source = {"sha256": _expected_source_for_phase(str(run["phase"])), "files": 7157}
    require(started.get("source") == expected_source and terminal.get("source") == expected_source, "capture source binding differs")
    require(started.get("binary") == binary and terminal.get("binary") == binary, "capture binary binding differs")
    require(started.get("argv") == argv and terminal.get("argv") == argv, "capture argv binding differs")
    require(terminal.get("exit_code") == 0, "terminal exit code is not zero")
    require(terminal.get("timed_out") is False, "terminal records a timeout")
    require("launch_error" not in terminal and "validation_error" not in terminal, "terminal contains a failure detail")
    require(terminal.get("failure") is None, "terminal failure field is not empty")
    require(terminal.get("started_utc") == started.get("started_utc"), "terminal start timestamp differs")
    started_at = _timestamp(started.get("started_utc"), "started.started_utc")
    finished_at = _timestamp(terminal.get("finished_utc"), "terminal.finished_utc")
    require(finished_at > started_at, "terminal finished before capture started")
    return started_at, finished_at


def capture_one(protocol: Mapping[str, Any], protocol_hash: str, attempt: str, run: Mapping[str, Any]) -> Path:
    directory = _run_directory(attempt, run)
    phase = str(run["phase"])
    role = str(run["role"])
    route = str(run["route"])
    binary = _binary_for(protocol, phase, role)
    replay_dir = TEMP / "replay" / attempt / str(run["label"]) if route == "file_store" else None
    if replay_dir is not None:
        replay_dir.parent.mkdir(parents=True, exist_ok=True)
        require(not replay_dir.exists(), f"refusing to replace replay directory: {replay_dir}")
        replay_dir.mkdir()
    argv = _argv_for(run, binary, directory, replay_dir)
    started = {
        "schema": CAPTURE_SCHEMA,
        "version": 1,
        "status": "running",
        "attempt": attempt,
        "run": dict(run),
        "protocol": {"path": PROTOCOL_FILE, "sha256": protocol_hash},
        "source": {"sha256": _expected_source_for_phase(phase), "files": 7157},
        "binary": copy.deepcopy(binary),
        "argv": argv,
        "cwd": str(REPO),
        "environment": _environment(),
        "machine": copy.deepcopy(protocol["machine"]),
        "started_utc": now(),
    }
    started_path = directory / "started.json"
    write_json(started_path, started)
    launch_error: str | None = None
    timed_out = False
    process: subprocess.Popen[bytes] | None = None
    try:
        with (directory / "stdout.txt").open("x", encoding="utf-8") as stdout, (directory / "stderr.txt").open("x", encoding="utf-8") as stderr:
            process = subprocess.Popen(
                argv,
                cwd=REPO,
                env=old_routes.base.ENV,
                stdout=stdout,
                stderr=stderr,
                start_new_session=True,
            )
            try:
                process.wait(timeout=CAPTURE_TIMEOUT_SECONDS)
            except subprocess.TimeoutExpired:
                timed_out = True
                try:
                    os.killpg(process.pid, signal.SIGTERM)
                except ProcessLookupError:
                    pass
                try:
                    process.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    try:
                        os.killpg(process.pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                    try:
                        process.wait(timeout=5)
                    except subprocess.TimeoutExpired as error:
                        launch_error = f"TimeoutExpired: process group did not terminate: {error}"
            exit_code = process.returncode
    except (OSError, subprocess.SubprocessError) as error:
        exit_code = None
        launch_error = f"{type(error).__name__}: {error}"
    validation_error: str | None = None
    cleanup: dict[str, Any] | None = None
    if exit_code == 0:
        try:
            case = dict(old_routes.ROUTE_CASE_BY_LABEL[CASE_LABEL])
            spec = old_routes.ROUTE_BY_NAME[route]
            old_routes.check_route_report(
                directory / "report.json", role, case, spec,
                samples=FORMAL_SAMPLES, warmups=FORMAL_WARMUPS,
                binary=binary, argv=argv, replay_dir=replay_dir,
            )
            if replay_dir is not None:
                require(replay_dir.is_dir() and not any(replay_dir.iterdir()), f"{replay_dir}: replay directory is not empty")
            cleanup = _replay_cleanup(replay_dir, directory)
        except (VarianceError, old_routes.base.MeasureError, OSError, ValueError) as error:
            validation_error = str(error)
            if replay_dir is not None and replay_dir.exists() and replay_dir.is_dir() and not any(replay_dir.iterdir()):
                cleanup = _replay_cleanup(replay_dir, directory)
    if cleanup is None and not (directory / "replay-cleanup.json").exists():
        cleanup = _replay_cleanup(replay_dir, directory) if replay_dir is not None and replay_dir.exists() and replay_dir.is_dir() and not any(replay_dir.iterdir()) else {"schema": REPLAY_CLEANUP_SCHEMA, "status": "failed", "path": None if replay_dir is None else str(replay_dir), "entries": [], "removed": False}
        if not (directory / "replay-cleanup.json").exists():
            write_json(directory / "replay-cleanup.json", cleanup)
    artifacts: dict[str, Any] = {}
    for name in ("report.json", "resource.txt", "stdout.txt", "stderr.txt", "started.json", "replay-cleanup.json"):
        path = directory / name
        if path.is_file():
            artifacts[name] = metadata(path, name)
    passed = exit_code == 0 and validation_error is None and all((directory / name).is_file() for name in ("report.json", "resource.txt", "stdout.txt", "stderr.txt", "replay-cleanup.json"))
    terminal = {
        "schema": TERMINAL_SCHEMA,
        "version": 1,
        "status": "pass" if passed else "failed",
        "attempt": attempt,
        "run": dict(run),
        "protocol": {"path": PROTOCOL_FILE, "sha256": protocol_hash},
        "started": {"path": "started.json", "sha256": sha(started_path)},
        "source": {"sha256": _expected_source_for_phase(phase), "files": 7157},
        "binary": copy.deepcopy(binary),
        "argv": argv,
        "cwd": str(REPO),
        "environment": _environment(),
        "machine": copy.deepcopy(protocol["machine"]),
        "exit_code": exit_code,
        "timed_out": timed_out,
        "started_utc": started["started_utc"],
        "finished_utc": now(),
        "artifacts": artifacts,
    }
    if launch_error is not None:
        terminal["launch_error"] = launch_error
    if validation_error is not None:
        terminal["validation_error"] = validation_error
    if passed:
        terminal["failure"] = None
    else:
        if timed_out:
            failure_kind = "timeout"
        elif launch_error is not None:
            failure_kind = "launch"
        elif validation_error is not None:
            failure_kind = "validation"
        else:
            failure_kind = "exit"
        terminal["failure"] = {"kind": failure_kind, "message": launch_error or validation_error or f"exit {exit_code}"}
    terminal_path = directory / "terminal.json"
    write_json(terminal_path, terminal)
    if not passed:
        fail(f"{run['label']} failed; terminal receipt retained")
    return terminal_path


def capture(attempt: str, block: int | None = None) -> None:
    attempt = _attempt(attempt)
    protocol, protocol_hash = load_protocol()
    runs = [run for run in formal_runs() if block is None or run["block"] == block]
    require(runs, "capture selection is empty")
    with capture_lock():
        for run in runs:
            capture_one(protocol, protocol_hash, attempt, run)
    print(f"captured {len(runs)} variance children for {attempt}")


def _expected_run_directory(attempt: str, run: Mapping[str, Any]) -> Path:
    return ROOT / "captures" / attempt / f"block-{int(run['block']):02d}" / str(run["label"])


def _validate_capture_run(protocol: Mapping[str, Any], protocol_hash: str, attempt: str, run: Mapping[str, Any]) -> dict[str, Any]:
    directory = _expected_run_directory(attempt, run)
    require(directory.is_dir() and not directory.is_symlink(), f"missing run directory: {directory}")
    allowed = {"report.json", "resource.txt", "stdout.txt", "stderr.txt", "started.json", "terminal.json", "replay-cleanup.json"}
    actual_names = {item.name for item in directory.iterdir()}
    require(actual_names == allowed, f"{directory}: retained artifact inventory differs")
    started = _obj(read_json(directory / "started.json"), str(directory / "started.json"))
    terminal = _obj(read_json(directory / "terminal.json"), str(directory / "terminal.json"))
    binary = _binary_for(protocol, str(run["phase"]), str(run["role"]))
    argv = _argv_for(run, binary, directory, TEMP / "replay" / attempt / str(run["label"]) if run["route"] == "file_store" else None)
    started_at, finished_at = _validate_terminal_custody(
        started,
        terminal,
        attempt=attempt,
        run=run,
        protocol_hash=protocol_hash,
        protocol=protocol,
        binary=binary,
        argv=argv,
    )
    started_binding = _obj(terminal.get("started"), f"{directory}/terminal.started")
    require(_sha_field(started_binding.get("sha256"), f"{directory}/terminal.started.sha256") == sha(directory / "started.json"), f"{directory}: started receipt hash differs")
    artifacts = _artifact_inventory(directory)
    recorded = _obj(terminal.get("artifacts"), f"{directory}/terminal.artifacts")
    require(set(recorded) == set(artifacts), f"{directory}: terminal artifact inventory differs")
    for name, details in artifacts.items():
        record = _obj(recorded.get(name), f"{directory}/terminal.artifacts.{name}")
        require(record.get("bytes") == details["bytes"] and record.get("sha256") == details["sha256"], f"{directory}: artifact {name} changed")
    cleanup = _obj(read_json(directory / "replay-cleanup.json"), f"{directory}/replay-cleanup.json")
    if run["route"] == "file_store":
        require(cleanup.get("status") == "removed" and cleanup.get("removed") is True, f"{directory}: file replay cleanup failed")
        replay_path = Path(str(cleanup.get("path", "")))
        expected_replay = TEMP / "replay" / attempt / str(run["label"])
        require(replay_path == expected_replay, f"{directory}: replay scratch path differs")
        require(not replay_path.exists() and not replay_path.is_symlink(), f"{directory}: replay scratch remains")
    else:
        require(cleanup.get("status") == "not_applicable", f"{directory}: unexpected replay cleanup")
    binary_record = _obj(_obj(protocol["builds"][run["phase"]], f"protocol.builds.{run['phase']}")[run["role"]], "build")
    require(sha(Path(str(binary_record["record_path"]))) == binary_record["record_sha256"], f"{directory}: build receipt changed")
    case = dict(old_routes.ROUTE_CASE_BY_LABEL[CASE_LABEL])
    spec = old_routes.ROUTE_BY_NAME[str(run["route"])]
    try:
        report = old_routes.check_route_report(
            directory / "report.json", str(run["role"]), case, spec,
            samples=FORMAL_SAMPLES, warmups=FORMAL_WARMUPS,
            binary=binary, argv=argv,
            replay_dir=TEMP / "replay" / attempt / str(run["label"]) if run["route"] == "file_store" else None,
        )
    except old_routes.base.MeasureError as error:
        fail(f"{directory}: sealed report validator rejected report: {error}")
    identities, values, rss = _sample_values(report, directory / "resource.txt", str(run["role"]), str(directory / "report.json"))
    row = {
        "run": dict(run),
        "directory": _path_relative(directory),
        "protocol": {"path": PROTOCOL_FILE, "sha256": protocol_hash},
        "binary": binary,
        "identities": identities,
        "metrics": {name: _stats(vector) for name, vector in values.items()},
        "sample_values": values,
        "time_max_rss_bytes": rss,
        "timing": {"started_utc": started_at.isoformat(), "finished_utc": finished_at.isoformat()},
        "raw_statistics": {name: _stats(vector) for name, vector in values.items()},
    }
    row["metrics"]["time_max_rss_bytes"] = _stats([rss])
    row["raw_statistics"]["time_max_rss_bytes"] = _stats([rss])
    return row


def _validate_inventory(attempt: str, expected: list[Mapping[str, Any]]) -> None:
    root = ROOT / "captures" / attempt
    require(root.is_dir() and not root.is_symlink(), f"formal capture root is missing: {root}")
    expected_blocks = {f"block-{block:02d}" for block in BLOCKS}
    root_children = list(root.iterdir())
    actual_blocks = {item.name for item in root_children if item.is_dir() and not item.is_symlink()}
    require({item.name for item in root_children} == expected_blocks, "formal capture root contains extra or non-directory entries")
    require(actual_blocks == expected_blocks, "formal block inventory differs")
    expected_dirs = {_expected_run_directory(attempt, run).relative_to(root).as_posix() for run in expected}
    actual_dirs = {
        item.relative_to(root).as_posix()
        for item in root.rglob("*")
        if item.is_dir() and not item.is_symlink() and len(item.relative_to(root).parts) == 2
    }
    require(actual_dirs == expected_dirs, "formal child inventory differs")


def _identity_key(identity: Mapping[str, Any]) -> tuple[Any, ...]:
    return (
        identity["archive_bytes"], identity["archive_sha256"],
        identity["main_xml_bytes"], identity["main_xml_sha256"],
    )


def _validate_timeline(rows: list[Mapping[str, Any]], expected: list[Mapping[str, Any]]) -> list[str]:
    require(len(rows) == len(expected), "timeline row cardinality differs")
    ordered = sorted(
        rows,
        key=lambda row: _timestamp(row["timing"]["started_utc"], f"{row['run']['label']}.timing.started_utc"),
    )
    actual_labels = [str(row["run"]["label"]) for row in ordered]
    expected_labels = [str(run["label"]) for run in expected]
    require(actual_labels == expected_labels, "actual capture chronology differs from frozen formal run order")
    for previous, current in zip(ordered, ordered[1:]):
        previous_finished = _timestamp(previous["timing"]["finished_utc"], f"{previous['run']['label']}.timing.finished_utc")
        current_started = _timestamp(current["timing"]["started_utc"], f"{current['run']['label']}.timing.started_utc")
        require(previous_finished <= current_started, f"capture intervals overlap: {previous['run']['label']} and {current['run']['label']}")
    return actual_labels


def _block_comparisons(rows: list[Mapping[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    by_key = {(row["run"]["block"], row["run"]["phase"], row["run"]["role"], row["run"]["route"]): row for row in rows}
    comparisons: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    for role in ROLES:
        for route in ROUTES:
            for metric in METRICS:
                block_changes: list[dict[str, Any]] = []
                for block in BLOCKS:
                    before = by_key[(block, "before", role, route)]
                    after = by_key[(block, "after", role, route)]
                    before_stats = before["metrics"].get(metric)
                    after_stats = after["metrics"].get(metric)
                    if before_stats is None or after_stats is None:
                        continue
                    quantile_changes: dict[str, Any] = {}
                    for quantile in ("p50", "p95", "p99"):
                        before_value = float(before_stats[quantile])
                        after_value = float(after_stats[quantile])
                        percent = _percent_change(before_value, after_value)
                        item = {
                            "before": before_value,
                            "after": after_value,
                            "absolute_change": after_value - before_value,
                            "percent_change": percent,
                            "adverse_over_threshold": percent is not None and percent > REVIEW_THRESHOLD_PERCENT,
                        }
                        quantile_changes[quantile] = item
                        if item["adverse_over_threshold"]:
                            adverse.append({"block": block, "role": role, "route": route, "metric": metric, "quantile": quantile, **item})
                    block_changes.append({"block": block, "quantiles": quantile_changes})
                if not block_changes:
                    continue
                aggregate: dict[str, Any] = {}
                for quantile in ("p50", "p95", "p99"):
                    absolute = [item["quantiles"][quantile]["absolute_change"] for item in block_changes]
                    percentages = [item["quantiles"][quantile]["percent_change"] for item in block_changes if item["quantiles"][quantile]["percent_change"] is not None]
                    aggregate[quantile] = {
                        "absolute_change": _stats(absolute),
                        "percent_change": _stats(percentages) if percentages else None,
                        "bootstrap_absolute_mean_ci95": _bootstrap([float(value) for value in absolute]),
                        "bootstrap_percent_mean_ci95": _bootstrap([float(value) for value in percentages]) if len(percentages) == len(block_changes) else None,
                    }
                comparisons.append({"role": role, "route": route, "metric": metric, "blocks": block_changes, "aggregate": aggregate})
    return comparisons, adverse


def analyze(attempt: str) -> dict[str, Any]:
    attempt = _attempt(attempt)
    protocol, protocol_hash = load_protocol()
    expected = formal_runs()
    _validate_inventory(attempt, expected)
    rows = [_validate_capture_run(protocol, protocol_hash, attempt, run) for run in expected]
    require(len(rows) == 72, "formal row cardinality differs")
    capture_order = _validate_timeline(rows, expected)
    # Every route/role/phase uses the same exact tiny fixture and authored
    # payload.  This check catches an oracle or archive substitution even if
    # each individual report is internally self-consistent.
    identity_reference = rows[0]["identities"]
    for row in rows[1:]:
        require(row["identities"] == identity_reference, f"{row['run']['label']}: source/authored/candidate identity differs")
    comparisons, adverse = _block_comparisons(rows)
    return {
        "schema": SUMMARY_SCHEMA,
        "version": VERSION,
        "protocol": {"path": PROTOCOL_FILE, "sha256": protocol_hash, "attempt": attempt},
        "analyzer": {"path": _path_relative(Path(__file__)), "sha256": sha(Path(__file__))},
        "inventory": {
            "processes": len(rows),
            "samples": len(rows) * FORMAL_SAMPLES,
            "blocks": len(BLOCKS),
            "samples_per_process": FORMAL_SAMPLES,
            "warmups_per_process": FORMAL_WARMUPS,
            "roles": list(ROLES),
            "routes": list(ROUTES),
            "phases": list(PHASES),
            "pilots_included": False,
        },
        "identity": {
            "source": identity_reference["source"],
            "authored": identity_reference["authored"],
            "candidate": identity_reference["candidate"],
            "all_rows_exact": True,
        },
        "timeline": {"ordered_labels": capture_order, "non_overlapping": True},
        "rows": rows,
        "comparisons": comparisons,
        "adverse_over_threshold": adverse,
        "bootstrap": {
            "unit": "paired block",
            "blocks": len(BLOCKS),
            "resamples": BOOTSTRAP_RESAMPLES,
            "seed": BOOTSTRAP_SEED,
            "confidence": 0.95,
            "sample_pseudoreplication": False,
        },
        "claims": {
            "performance_claim": "none",
            "causal_speedup": False,
            "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
            "rss_scope": "GNU time maximum is one observation per child and is not a bounded-memory proof",
            "uncertainty": "block-paired p50/p95/p99 changes with six-block bootstrap; in-process samples are not independent blocks",
        },
    }


def _canonical(value: Any) -> str:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)


def verify(summary_path: Path) -> None:
    summary = _obj(read_json(summary_path), str(summary_path))
    require(summary.get("schema") == SUMMARY_SCHEMA and summary.get("version") == VERSION, "variance summary schema differs")
    protocol = _obj(summary.get("protocol"), "summary.protocol")
    require(isinstance(protocol.get("attempt"), str), "summary attempt is missing")
    expected = analyze(str(protocol["attempt"]))
    require(_canonical(summary) == _canonical(expected), "variance summary does not reproduce from retained evidence")
    print(f"verified {summary_path}")


def parser() -> argparse.ArgumentParser:
    command = argparse.ArgumentParser(description=__doc__)
    sub = command.add_subparsers(dest="command", required=True)
    sub.add_parser("freeze", help="freeze the six-block variance protocol")
    capture_parser = sub.add_parser("capture", help="capture all formal children under the shared CPU lock")
    capture_parser.add_argument("--attempt", required=True)
    capture_parser.add_argument("--block", type=int, choices=BLOCKS, help="capture one block for resumable orchestration")
    analyze_parser = sub.add_parser("analyze", help="validate retained children and write the variance summary")
    analyze_parser.add_argument("--attempt", required=True)
    analyze_parser.add_argument("--summary", type=Path, default=ROOT / SUMMARY_FILE)
    verify_parser = sub.add_parser("verify", help="recompute and verify an existing summary")
    verify_parser.add_argument("--summary", type=Path, default=ROOT / SUMMARY_FILE)
    return command


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    try:
        if args.command == "freeze":
            freeze()
        elif args.command == "capture":
            capture(args.attempt, args.block)
        elif args.command == "analyze":
            result = analyze(args.attempt)
            write_json(args.summary.resolve(), result)
            print(f"analyzed {result['inventory']['processes']} children and {result['inventory']['samples']} samples")
        else:
            verify(args.summary.resolve())
    except (VarianceError, old_routes.base.MeasureError) as error:
        print(f"variance.py: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
