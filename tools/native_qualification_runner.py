#!/usr/bin/env python3
"""Fail-closed runner for XLSB, XLSX raw, W5 and legacy native ABBA lanes.

The runner deliberately takes all source, target, binary, fixture, and output
paths from a JSON manifest.  It does not know a wave commit and therefore
cannot silently continue using an old source checkout.  A manifest records the
source identity it observes at execution time instead.

There are two phases:

``build``
    Runs the manifest's Cargo commands one at a time, injecting ``--offline``
    and ``--locked`` only when those flags are absent before Cargo's ``--``
    separator.  It records source/lock/probe/fixture and binary hashes.

``abba``
    Requires the build receipt, then runs every program for the exact
    ``before, after, after, before`` order.  A child failure, missing output,
    source mutation, target overrun, or unrelated Litchi builder causes an
    immediate failed receipt.  Failed child logs are retained.

This is orchestration only.  It never deletes a target or an output and it
does not interpret benchmark numbers.  Keep large receipts and binaries in an
external directory; only reduced summaries belong in the repository.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import errno
import fcntl
import hashlib
import json
import os
from pathlib import Path
import re
import resource
import signal
import subprocess
import sys
import time
import tomllib
from typing import Any, Callable, Iterable, Mapping, Sequence

from native_qualification.prepare_xlsx_raw_inputs import (
    BASELINE_REVISION,
    validate_preparation_receipt,
)


SCHEMA = "litchi.native-qualification.v1"
ABBA_ORDER = ("before", "after", "after", "before")
# The v2 DrawingML validator and its imports must come from the pinned
# protocol checkout, never from either source leg.  Keep this list explicit so
# a candidate cannot replace one of the policy or compile modules through its
# own tree.
DRAWINGML_V2_PROTOCOL_TOOL_PATHS = (
    "tools/native_qualification_runner.py",
    "tools/native_qualification/drawingml_raw_paired_runner_v2.py",
    "tools/native_qualification/drawingml_raw_validator_v2.py",
    "tools/native_qualification/drawingml_compile_driver_v2.py",
    "tools/native_qualification/drawingml_compile_driver.py",
    "tools/native_qualification/prepare_drawingml_raw_inputs_v2.py",
    "tools/native_qualification/prepare_drawingml_raw_inputs.py",
    "tools/native_qualification/prepare_xlsx_raw_inputs.py",
    "tools/native_qualification/drawingml_raw_paired_runner.py",
    "tools/native_qualification/drawingml_raw_validator.py",
)
ROUNDS = 31
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_LIMIT_BYTES = 20_000_000_000
MEMORY_HARD_LIMIT_BYTES = 20_000_000_000
BUILD_PERMIT_SCHEMA = "litchi.native-build-permit.v1"
BUILD_EXECUTION_CLASSES = frozenset({"exclusive", "compile_only"})
PEER_COMMAND_CLASSES = frozenset(
    {
        "cargo-build",
        "cargo-check",
        "cargo-clippy",
        "cargo-test",
    }
)
DISALLOWED_DESCENDANT_NAMES = frozenset(
    {
        "cargo-fuzz",
        "heaptrack",
        "perf",
        "rustdoc",
        "rustfmt",
        "valgrind",
    }
)
HEAVY_NAMES = frozenset(
    {
        "cargo",
        "cargo-fuzz",
        "rustc",
        "rustdoc",
        "rustfmt",
        "clippy-driver",
        "valgrind",
        "perf",
        "heaptrack",
    }
)
LITCHI_ROOTS = (
    Path("/home/zhuhe/code/litchi"),
    Path("/home/zhuhe/code/litchi-worktrees"),
)
_TOKEN = re.compile(r"\{([a-z_][a-z0-9_]*)\}")


class RunnerError(RuntimeError):
    """A fail-closed manifest, guard, or child-command error."""


class ProcMonitorError(RunnerError):
    """The process monitor could not establish a trustworthy /proc scan."""


def _utc_now() -> _dt.datetime:
    return _dt.datetime.now(_dt.timezone.utc)


def _now() -> str:
    return _utc_now().isoformat(timespec="milliseconds")


def _parse_utc(value: Any, label: str) -> _dt.datetime:
    if not isinstance(value, str) or not value:
        raise RunnerError(f"{label} must be an ISO-8601 timestamp")
    try:
        parsed = _dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        raise RunnerError(f"{label} is not a valid ISO-8601 timestamp") from error
    if parsed.tzinfo is None:
        raise RunnerError(f"{label} must carry a timezone")
    return parsed.astimezone(_dt.timezone.utc)


def _sha256_shape(value: Any, label: str) -> str:
    if (
        not isinstance(value, str)
        or len(value) != 64
        or value != value.lower()
        or any(character not in "0123456789abcdef" for character in value)
    ):
        raise RunnerError(f"{label} must be a lowercase SHA-256")
    return value


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink():
        raise RunnerError(f"refusing to follow symlinked file: {path}")
    if not path.is_file():
        raise RunnerError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while True:
            block = stream.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _drawingml_v2_protocol_tool_identities(root: Path) -> dict[str, dict[str, Any]]:
    root = root.resolve(strict=False)
    identities: dict[str, dict[str, Any]] = {}
    for relative in DRAWINGML_V2_PROTOCOL_TOOL_PATHS:
        path = root / relative
        digest, size = _sha256_file(path)
        identities[relative] = {
            "path": str(path.resolve()),
            "sha256": digest,
            "bytes": size,
        }
    return identities


_UNTRACKED_UPSTREAM_COMPONENT = "3rd" + "party"


def _reject_untracked_upstream_checkout(path: Path) -> None:
    if _UNTRACKED_UPSTREAM_COMPONENT in path.parts:
        raise RunnerError(f"untracked upstream checkout path is forbidden: {path}")


def _resolve_path(value: str | os.PathLike[str], *, base: Path | None = None) -> Path:
    path = Path(value)
    if not path.is_absolute():
        if base is None:
            raise RunnerError(f"path must be absolute: {value!r}")
        path = base / path
    probe = Path(path.anchor)
    for part in path.parts[1:]:
        probe /= part
        if probe.is_symlink():
            raise RunnerError(f"symlink path components are forbidden: {path}")
    path = path.expanduser().resolve(strict=False)
    _reject_untracked_upstream_checkout(path)
    return path


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _overlap(left: Path, right: Path) -> bool:
    return left == right or _under(left, right) or _under(right, left)


def _ensure_external(path: Path, source_roots: Iterable[Path], label: str) -> None:
    for root in source_roots:
        if _under(path, root):
            raise RunnerError(f"{label} must be outside source worktrees: {path}")


def _ensure_disjoint(paths: Mapping[str, Path]) -> None:
    items = list(paths.items())
    for index, (left_name, left) in enumerate(items):
        for right_name, right in items[index + 1 :]:
            if _overlap(left, right):
                raise RunnerError(f"paths must be pairwise disjoint: {left_name}={left}, {right_name}={right}")


def _directory_size(path: Path) -> int:
    """Return a conservative size without following symlinks."""

    if not path.exists():
        return 0
    if path.is_symlink():
        raise RunnerError(f"target/output may not be a symlink: {path}")
    if path.is_file():
        return path.stat().st_size
    total = 0
    for root, dirs, files in os.walk(path, followlinks=False):
        dirs[:] = [name for name in dirs if not (Path(root) / name).is_symlink()]
        for name in files:
            item = Path(root) / name
            if item.is_symlink():
                continue
            try:
                total += item.stat().st_size
            except FileNotFoundError:
                # A compiler can replace a file during the walk.  The next
                # check will see it; never turn a missing file into deletion.
                continue
    return total


def _hash_paths(paths: Iterable[Path]) -> dict[str, Any]:
    """Hash files/directories deterministically, including relative names."""

    entries: list[tuple[str, Path]] = []
    for path in paths:
        if path.is_symlink():
            raise RunnerError(f"identity path may not be a symlink: {path}")
        path = path.resolve(strict=False)
        _reject_untracked_upstream_checkout(path)
        if not path.exists():
            raise RunnerError(f"identity path does not exist: {path}")
        if path.is_file():
            entries.append((path.name, path))
            continue
        for root, dirs, files in os.walk(path, followlinks=False):
            dirs[:] = sorted(name for name in dirs if not (Path(root) / name).is_symlink())
            for name in sorted(files):
                item = Path(root) / name
                if item.is_symlink():
                    raise RunnerError(f"identity directory contains symlink: {item}")
                entries.append((str(item.relative_to(path)), item))
    digest = hashlib.sha256()
    records: list[dict[str, Any]] = []
    for relative, item in sorted(entries, key=lambda pair: (str(pair[1].parent), pair[0])):
        file_hash, size = _sha256_file(item)
        digest.update(relative.encode("utf-8"))
        digest.update(b"\0")
        digest.update(str(size).encode("ascii"))
        digest.update(b"\0")
        digest.update(bytes.fromhex(file_hash))
        records.append({"path": str(item), "relative": relative, "bytes": size, "sha256": file_hash})
    return {"sha256": digest.hexdigest(), "files": records}


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise RunnerError(f"git {' '.join(args)} failed in {root}: {result.stderr.strip()}")
    return result.stdout


def _source_identity(spec: Mapping[str, Any], shared: Mapping[str, Any]) -> dict[str, Any]:
    root = _resolve_path(str(spec["source_root"]))
    if not root.is_dir():
        raise RunnerError(f"source root is not a directory: {root}")
    top = _resolve_path(_git(root, "rev-parse", "--show-toplevel").strip())
    if top != root:
        raise RunnerError(f"source_root is not the worktree top-level: {root} != {top}")
    head = _git(root, "rev-parse", "HEAD").strip()
    tree = _git(root, "rev-parse", "HEAD^{tree}").strip()
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    if status and not bool(spec.get("allow_dirty", False)):
        raise RunnerError(f"source worktree is dirty: {root}\n{status}")
    status_paths: list[str] = []
    allowed_dirty = spec.get("allowed_dirty_paths")
    if status and isinstance(allowed_dirty, list):
        allowed = {str(value) for value in allowed_dirty}
        for line in status.splitlines():
            if len(line) < 4:
                raise RunnerError(f"malformed git status entry in {root}: {line!r}")
            value = line[3:].strip()
            if " -> " in value:
                value = value.split(" -> ", 1)[1]
            status_paths.append(value)
            if value not in allowed:
                raise RunnerError(f"unreviewed dirty path in {root}: {value}")
    diff = _git(root, "diff", "--no-ext-diff", "--binary")
    lock_values = spec.get("lock_paths", ["Cargo.lock"])
    if not isinstance(lock_values, list) or not lock_values:
        raise RunnerError("lock_paths must be a non-empty list")
    locks = [_resolve_path(str(value), base=root) for value in lock_values]
    for lock in locks:
        if not lock.is_file():
            raise RunnerError(f"lock file does not exist: {lock}")
    probe_values = [*shared.get("probe_paths", []), *spec.get("probe_paths", [])]
    fixture_values = [*shared.get("fixture_paths", []), *spec.get("fixture_paths", [])]
    probes = [_resolve_path(str(value), base=root) for value in probe_values]
    fixtures = [_resolve_path(str(value), base=root) for value in fixture_values]
    preparation_value = shared.get("preparation_receipt")
    preparation = None
    if preparation_value is not None:
        preparation = _resolve_path(str(preparation_value))
        if not preparation.is_file() or preparation.is_symlink():
            raise RunnerError(f"preparation receipt is missing or symlinked: {preparation}")
    identity_paths = [*locks, *probes, *fixtures]
    if preparation is not None:
        identity_paths.append(preparation)
    files = _hash_paths(identity_paths) if identity_paths else {"sha256": _sha256_bytes(b""), "files": []}
    identity = {
        "source_root": str(root),
        "head": head,
        "tree": tree,
        "source_tree_sha256": _sha256_bytes((tree + "\0" + status + "\0" + diff).encode("utf-8")),
        "clean": not bool(status),
        "status_paths": status_paths,
        "status_sha256": _sha256_bytes(status.encode("utf-8")),
        "diff_sha256": _sha256_bytes(diff.encode("utf-8")),
        "lock_paths": [str(path) for path in locks],
        "probe_paths": [str(path) for path in probes],
        "fixture_paths": [str(path) for path in fixtures],
        "identity_sha256": files["sha256"],
        "identity_files": files["files"],
    }
    if preparation is not None:
        identity["preparation_receipt"] = str(preparation)
    return identity


def _expand_environment(value: Any) -> Any:
    if isinstance(value, str):
        expanded = os.path.expandvars(value)
        if "$" in expanded:
            raise RunnerError(f"unexpanded environment variable in manifest value: {value!r}")
        return expanded
    if isinstance(value, list):
        return [_expand_environment(item) for item in value]
    if isinstance(value, dict):
        return {key: _expand_environment(item) for key, item in value.items()}
    return value


def load_manifest(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise RunnerError(f"cannot read manifest {path}: {error}") from error
    if not isinstance(data, dict):
        raise RunnerError("manifest root must be an object")
    data = _expand_environment(data)
    if data.get("schema") != SCHEMA:
        raise RunnerError(f"manifest schema must be {SCHEMA!r}")
    allowed_lanes = {
        "w4b-xlsb",
        "w4-xlsx-raw",
        "w5-legacy",
        "drawingml-raw",
        "legacy-rtf",
        "legacy-ole-properties",
        "legacy-xls-properties",
    }
    if data.get("lane") not in allowed_lanes:
        raise RunnerError(f"lane must be one of {sorted(allowed_lanes)!r}")
    if data.get("abba_order") != list(ABBA_ORDER):
        raise RunnerError(f"abba_order must be {list(ABBA_ORDER)!r}")
    if data.get("rounds") != ROUNDS:
        raise RunnerError(f"rounds is fixed at {ROUNDS}")
    blocks = data.get("abba_blocks", 1)
    if not isinstance(blocks, int) or not 1 <= blocks <= 5:
        raise RunnerError("abba_blocks must be an integer from 1 through 5")
    if "target_limit_bytes" in data or "memory_limit_bytes" in data:
        raise RunnerError("target and memory limits are fixed by the runner, not manifest-configurable")
    if set(data.get("legs", {})) != {"before", "after"}:
        raise RunnerError("manifest legs must contain exactly before and after")
    if not isinstance(data.get("shared", {}), dict):
        raise RunnerError("shared must be an object")
    output = _resolve_path(str(data.get("output_dir", "")))
    preparation_receipt: Path | None = None
    if data.get("lane") in {"w4-xlsx-raw", "drawingml-raw"}:
        value = data["shared"].get("preparation_receipt")
        if not isinstance(value, str) or not value:
            raise RunnerError(f"{data['lane']} needs a preparation receipt")
        preparation_receipt = _resolve_path(value)
        if _under(preparation_receipt, output) or preparation_receipt == output:
            raise RunnerError(f"{data['lane']} preparation receipt must be outside the fresh output directory")
    elif str(data.get("lane", "")).startswith("legacy-"):
        value = data.get("preparation_receipt")
        if not isinstance(value, str) or not value:
            raise RunnerError("legacy lane must bind a source-preparation receipt")
        preparation_receipt = _resolve_path(value)
    source_roots: list[Path] = []
    for leg, spec in data["legs"].items():
        if not isinstance(spec, dict):
            raise RunnerError(f"legs.{leg} must be an object")
        source = _resolve_path(str(spec.get("source_root", "")))
        target = _resolve_path(str(spec.get("target_dir", "")))
        source_roots.append(source)
        if target == source or _under(target, source):
            raise RunnerError(f"legs.{leg}.target_dir must be external to source_root")
        if target == output or _under(target, output) or _under(output, target):
            raise RunnerError("target_dir and output_dir must be distinct trees")
        if not isinstance(spec.get("build", []), list) or not spec["build"]:
            raise RunnerError(f"legs.{leg}.build must be a non-empty list")
        for command in spec["build"]:
            _validate_command(command, f"legs.{leg}.build")
        programs = spec.get("programs", [])
        if not isinstance(programs, list) or not programs:
            raise RunnerError(f"legs.{leg}.programs must be a non-empty list")
        ids: set[str] = set()
        for program in programs:
            if not isinstance(program, dict) or not isinstance(program.get("id"), str):
                raise RunnerError(f"legs.{leg}.programs entries need an id")
            if program["id"] in ids:
                raise RunnerError(f"duplicate program id {program['id']!r} in {leg}")
            ids.add(program["id"])
            if not isinstance(program.get("argv"), list):
                raise RunnerError(f"program {leg}/{program['id']} needs argv")
            if "binary" not in program:
                raise RunnerError(f"program {leg}/{program['id']} needs binary")
            if not isinstance(program.get("outputs", []), list):
                raise RunnerError(f"program {leg}/{program['id']}.outputs must be a list")
            for stream in ("stdout", "stderr"):
                if stream in program and not isinstance(program[stream], str):
                    raise RunnerError(f"program {leg}/{program['id']}.{stream} must be a path template")
    validators = data.get("validators", [])
    if not isinstance(validators, list) or not validators:
        raise RunnerError("validators must contain at least one lane correctness/analyzer command")
    validator_ids: set[str] = set()
    for validator in validators:
        _validate_validator(validator, "validators")
        if validator["id"] in validator_ids:
            raise RunnerError(f"duplicate validator id {validator['id']!r}")
        validator_ids.add(validator["id"])
    before_ids = [program["id"] for program in data["legs"]["before"]["programs"]]
    after_ids = [program["id"] for program in data["legs"]["after"]["programs"]]
    if before_ids != after_ids:
        raise RunnerError("before and after program order must match")
    targets = [_resolve_path(str(data["legs"][leg]["target_dir"])) for leg in ("before", "after")]
    disjoint_paths = {
        "source_before": source_roots[0],
        "source_after": source_roots[1],
        "target_before": targets[0],
        "target_after": targets[1],
        "output": output,
    }
    if preparation_receipt is not None:
        disjoint_paths["preparation_receipt"] = preparation_receipt
        _ensure_external(preparation_receipt, source_roots, "preparation receipt")
    _ensure_disjoint(disjoint_paths)
    shared = data["shared"]
    for key in ("probe_paths", "fixture_paths"):
        if not isinstance(shared.get(key, []), list):
            raise RunnerError(f"shared.{key} must be a list")
    _validate_build_execution(data, output=output, source_roots=source_roots)
    _validate_lane_workload(data)
    return data


def _validate_command(command: Any, location: str) -> None:
    if not isinstance(command, dict):
        raise RunnerError(f"{location} command must be an object")
    argv = command.get("argv")
    if not isinstance(argv, list) or not argv or not all(isinstance(item, str) for item in argv):
        raise RunnerError(f"{location} command argv must be a non-empty string list")
    executable = Path(argv[0]).name
    if not isinstance(command.get("id"), str) or not command["id"]:
        raise RunnerError(f"{location} command needs a non-empty id")
    if executable != "cargo":
        raise RunnerError(f"{location} command must execute Cargo directly, not {argv[0]!r}")
    if "shell" in command or command.get("shell"):
        raise RunnerError(f"{location} may not use a shell")
    artifacts = command.get("artifacts", [])
    if (
        not isinstance(artifacts, list)
        or not artifacts
        or not all(isinstance(item, str) and item for item in artifacts)
    ):
        raise RunnerError(f"{location}.artifacts must be a non-empty string list")


def _cargo_command_class(argv: Sequence[str]) -> str | None:
    """Classify only the Cargo subcommands that a peer lease may admit."""

    if not argv or Path(argv[0]).name != "cargo":
        return None
    index = 1
    if index < len(argv) and argv[index].startswith("+"):
        index += 1
    if index >= len(argv):
        return None
    return {
        "build": "cargo-build",
        "check": "cargo-check",
        "clippy": "cargo-clippy",
        "test": "cargo-test",
    }.get(argv[index])


def _validate_build_execution(
    data: Mapping[str, Any], *, output: Path, source_roots: Sequence[Path]
) -> None:
    execution = data.get("build_execution")
    if execution is None:
        return
    if not isinstance(execution, dict):
        raise RunnerError("build_execution must be an object")
    if execution.get("class") == "exclusive":
        return
    if execution.get("class") != "compile_only":
        raise RunnerError("build_execution.class must be exclusive or compile_only")
    permit_value = execution.get("permit_path")
    if not isinstance(permit_value, str) or not permit_value:
        raise RunnerError("compile_only build_execution needs permit_path")
    permit = _resolve_path(permit_value)
    if _under(permit, output) or permit == output:
        raise RunnerError("compile_only permit_path must be outside output_dir")
    _ensure_external(permit, source_roots, "compile_only permit_path")
    for leg, spec in data["legs"].items():
        for command in spec["build"]:
            command_class = _cargo_command_class(command["argv"])
            if command_class != "cargo-build":
                raise RunnerError(
                    f"compile_only build_execution admits only cargo build commands ({leg}/{command['id']})"
                )


def _validate_validator(command: Any, location: str) -> None:
    if not isinstance(command, dict) or not isinstance(command.get("id"), str) or not command["id"]:
        raise RunnerError(f"{location} entries need a non-empty id")
    argv = command.get("argv")
    if not isinstance(argv, list) or not argv or not all(isinstance(item, str) for item in argv):
        raise RunnerError(f"{location}.{command['id']} argv must be a non-empty string list")
    if command.get("shell"):
        raise RunnerError(f"{location}.{command['id']} may not use a shell")
    if not isinstance(command.get("script"), str) or not command["script"]:
        raise RunnerError(f"{location}.{command['id']} needs a real validator script")
    outputs = command.get("outputs", [])
    if not isinstance(outputs, list) or not all(isinstance(item, str) for item in outputs):
        raise RunnerError(f"{location}.{command['id']}.outputs must be a string list")


def _validate_lane_workload(data: Mapping[str, Any]) -> None:
    if str(data["lane"]).startswith("legacy-"):
        _validate_legacy_native_workload(data)
        return
    expected: dict[str, list[tuple[str, list[str]]]] = {
        "w4b-xlsb": [
            ("probe", ["{rounds}"]),
            ("sheet", []),
            (
                "crud",
                [
                    "--case",
                    "all",
                    "--backend",
                    "owned",
                    "--fixture",
                    "{fixture}",
                    "--warmup",
                    "3",
                    "--samples",
                    "{rounds}",
                    "--json",
                    "{output}",
                ],
            ),
        ],
        "w5-legacy": [("source", ["{rounds}"]), ("ops", ["{rounds}"])],
    }
    lane = str(data["lane"])
    if lane == "w4-xlsx-raw":
        _validate_xlsx_raw_workload(data)
        return
    if lane == "drawingml-raw":
        _validate_drawingml_raw_workload(data)
        return
    expected_programs = expected[lane]
    for leg, spec in data["legs"].items():
        actual = [(program["id"], program["argv"]) for program in spec["programs"]]
        if actual != expected_programs:
            raise RunnerError(f"{leg} does not preserve the exact {data['lane']} workload")
        expected_stdout = {
            "w4b-xlsb": {
                "probe": "{output_dir}/probe-{leg}-{index}.json",
                "sheet": "{output_dir}/sheet-{leg}-{index}.txt",
                "crud": "{output_dir}/crud-{leg}-{index}.stdout",
            },
            "w5-legacy": {
                "source": "{output_dir}/source-{leg}-{index}.json",
                "ops": "{output_dir}/ops-{leg}-{index}.json",
            },
        }[lane]
        for program in spec["programs"]:
            if program.get("stdout") != expected_stdout[program["id"]]:
                raise RunnerError(f"{leg}/{program['id']} does not use the historical analyzer filename")
    if data["lane"] == "w5-legacy":
        environment = data.get("environment", {})
        if environment.get("LITCHI_PROBE_BATCH_ROUNDS") != "7":
            raise RunnerError("W5 must set LITCHI_PROBE_BATCH_ROUNDS=7")
        for spec in data["legs"].values():
            if "environment" in spec and spec["environment"].get("LITCHI_PROBE_BATCH_ROUNDS") != "7":
                raise RunnerError("W5 leg environment must preserve LITCHI_PROBE_BATCH_ROUNDS=7")
    validators = data.get("validators", [])
    if len(validators) != 1 or validators[0].get("id") != "historical-analyze":
        raise RunnerError("each lane needs exactly one historical-analyze validator")
    if validators[0].get("outputs") != ["{output_dir}/summary.json"]:
        raise RunnerError("historical analyzer must write summary.json in the output directory")


def _validate_xlsx_raw_workload(data: Mapping[str, Any]) -> None:
    """Validate the paired allocator/native XLSX raw-identity workload.

    The source probe itself owns the four-corpus lifecycle and fidelity
    checks. This runner owns only the build identity and quiet ABBA envelope,
    so both its allocation and native binaries must execute in every leg.
    """

    manifest_path = "tools/perf-baseline/unified-xlsx-raw-identity/Cargo.toml"
    expected_build_features = {
        "before": ["--no-default-features", "--features", "allocator-metrics,legacy-xlsx-raw-opc"],
        "after": ["--features", "allocator-metrics,current-xlsx-raw-opc"],
    }
    expected_artifacts = [
        "{target_dir}/release/xlsx-raw-identity-alloc",
        "{target_dir}/release/xlsx-raw-identity-native",
    ]
    expected_build_ids = ["xlsx_raw_identity"]
    expected_program_ids = ["raw-alloc", "raw-native"]
    for leg, spec in data["legs"].items():
        commands = spec["build"]
        if [command["id"] for command in commands] != expected_build_ids:
            raise RunnerError(f"{leg} does not preserve the XLSX raw build command identity")
        command = commands[0]
        argv = command["argv"]
        expected_prefix = ["cargo", "build", "--release", "--locked"]
        if argv[:4] != expected_prefix or "--manifest-path" not in argv:
            raise RunnerError(f"{leg} does not use the locked standalone XLSX raw manifest")
        manifest_index = argv.index("--manifest-path")
        if argv[manifest_index + 1] != manifest_path:
            raise RunnerError(f"{leg} does not use the standalone XLSX raw manifest")
        if argv[manifest_index + 2 :] != [*expected_build_features[leg], "--bin", "xlsx-raw-identity-alloc", "--bin", "xlsx-raw-identity-native"]:
            raise RunnerError(f"{leg} does not preserve the XLSX raw feature route")
        if command["artifacts"] != expected_artifacts:
            raise RunnerError(f"{leg} does not declare both XLSX raw binaries")

        programs = spec["programs"]
        if [program["id"] for program in programs] != expected_program_ids:
            raise RunnerError(f"{leg} does not preserve paired XLSX raw programs")
        expected_variant = "original" if leg == "before" else "candidate"
        leg_revision: str | None = None
        for route, program in zip(("alloc", "native"), programs):
            expected_argv = [
                "xlsx-raw-identity",
                "--api-variant",
                expected_variant,
                "--source-revision",
                "$XLSX_BEFORE_REVISION" if leg == "before" else "$XLSX_AFTER_REVISION",
                "--iterations",
                "5",
                "--warmups",
                "2",
                "--json",
                "{output}",
            ]
            # Environment placeholders are expanded before validation; only
            # the revision's 40-hex shape is intentionally left flexible for
            # the final assembled candidate.
            actual = program["argv"]
            if len(actual) != len(expected_argv) or actual[:4] != expected_argv[:4] or actual[5:] != expected_argv[5:]:
                raise RunnerError(f"{leg}/{route} does not preserve the XLSX raw probe arguments")
            revision = actual[4]
            if len(revision) != 40 or not all(char in "0123456789abcdefABCDEF" for char in revision):
                raise RunnerError(f"{leg}/{route} must carry a 40-character source revision")
            if leg == "before" and revision != BASELINE_REVISION:
                raise RunnerError(f"{leg}/{route} must report the original 0bef revision")
            if leg_revision is None:
                leg_revision = revision
            elif revision != leg_revision:
                raise RunnerError(f"{leg}/{route} does not use the same source revision as its paired route")
            if program.get("outputs") != [f"{{output_dir}}/xlsx-raw-{route}-{{leg}}-{{index}}.json"]:
                raise RunnerError(f"{leg}/{route} does not use its dedicated XLSX raw JSON output")
            if program.get("stdout") != f"{{output_dir}}/xlsx-raw-{route}-{{leg}}-{{index}}.stdout":
                raise RunnerError(f"{leg}/{route} does not use its dedicated stdout log")
            if program.get("stderr") != f"{{output_dir}}/xlsx-raw-{route}-{{leg}}-{{index}}.stderr":
                raise RunnerError(f"{leg}/{route} does not use its dedicated stderr log")
            if program.get("require_stdout") is not False:
                raise RunnerError(f"{leg}/{route} must write its report through --json, not stdout")

    validators = data.get("validators", [])
    preparation_receipt = data.get("shared", {}).get("preparation_receipt")
    if not isinstance(preparation_receipt, str) or not preparation_receipt:
        raise RunnerError("XLSX raw lane needs a preparation receipt")
    if len(validators) != 1 or validators[0].get("id") != "xlsx-raw-identity":
        raise RunnerError("XLSX raw lane needs exactly one xlsx-raw-identity validator")
    validator = validators[0]
    if validator.get("outputs") != ["{output_dir}/summary.json"]:
        raise RunnerError("XLSX raw validator must write summary.json in the output directory")
    expected_validator_argv = [
        "python3",
        str(validator.get("script")),
        "{output_dir}",
        data["legs"]["before"]["programs"][0]["argv"][4],
        data["legs"]["after"]["programs"][0]["argv"][4],
        data["legs"]["before"]["target_dir"],
        data["legs"]["after"]["target_dir"],
        data["legs"]["before"]["source_root"],
        data["legs"]["after"]["source_root"],
        preparation_receipt,
    ]
    if validator.get("argv") != expected_validator_argv:
        raise RunnerError("XLSX raw validator must bind its reports to both source legs and the preparation receipt")
    if validator.get("cwd") != "{output_dir}":
        raise RunnerError("XLSX raw validator must run from the output directory")


def _validate_drawingml_v2_policy(data: Mapping[str, Any]) -> None:
    """Recheck the v2 pair before allowing direct generic-runner execution.

    The paired coordinator performs this check too, but the generic runner is
    an exposed executable boundary.  Without this guard a caller could feed a
    hand-written protocol-v2 profile directly to ``SafeRunner`` and bypass
    candidate binding, trusted-source ordering, role locks, and probe-route
    receipt validation.
    """

    shared = data.get("shared")
    legs = data.get("legs")
    if not isinstance(shared, Mapping) or not isinstance(legs, Mapping):
        raise RunnerError("DrawingML v2 policy needs shared and legs objects")
    protocol_root_value = data.get("drawingml_protocol_root")
    protocol_tools = data.get("drawingml_protocol_tools")
    if not isinstance(protocol_root_value, str) or not protocol_root_value:
        raise RunnerError("DrawingML v2 policy needs its pinned protocol root")
    if not isinstance(protocol_tools, Mapping):
        raise RunnerError("DrawingML v2 policy needs pinned protocol tool identities")
    protocol_root = _resolve_path(protocol_root_value)
    if not protocol_root.is_dir():
        raise RunnerError(f"DrawingML v2 protocol root is not a directory: {protocol_root}")
    expected_generic = (protocol_root / "tools/native_qualification_runner.py").resolve()
    if expected_generic != Path(__file__).resolve():
        raise RunnerError("DrawingML v2 generic runner is not executed from the bound protocol root")
    from native_qualification import prepare_drawingml_raw_inputs_v2 as preparation
    try:
        expected_tools = _drawingml_v2_protocol_tool_identities(protocol_root)
    except RunnerError as error:
        raise RunnerError(f"DrawingML v2 protocol tool identity failed: {error}") from error
    if dict(protocol_tools) != expected_tools:
        raise RunnerError("DrawingML v2 protocol tool identities differ")
    receipt_value = shared.get("preparation_receipt")
    trusted_root = data.get("drawingml_trusted_root")
    before = legs.get("before")
    after = legs.get("after")
    if not isinstance(before, Mapping) or not isinstance(after, Mapping):
        raise RunnerError("DrawingML v2 policy needs before and after legs")
    if set(legs) != {"before", "after"}:
        raise RunnerError("DrawingML v2 policy needs exactly before and after legs")
    if shared.get("probe_paths") != list(preparation.PROBE_COPY_PATHS):
        raise RunnerError("DrawingML v2 policy probe paths differ from the preparation receipt")
    if shared.get("fixture_paths") != [
        "tools/native_qualification/fixtures/drawingml_raw_fixtures.json"
    ]:
        raise RunnerError("DrawingML v2 policy fixture paths differ from the preparation receipt")
    values = (
        receipt_value,
        trusted_root,
        before.get("source_root"),
        after.get("source_root"),
        before.get("source_revision"),
        after.get("source_revision"),
    )
    if not all(isinstance(value, str) and value for value in values):
        raise RunnerError("DrawingML v2 policy is missing a bound source/receipt identity")
    for leg_name, leg in (("before", before), ("after", after)):
        leg_root = _resolve_path(str(leg["source_root"]))
        if _overlap(protocol_root, leg_root):
            raise RunnerError(
                f"DrawingML v2 protocol root overlaps the {leg_name} source root"
            )
    for leg_name, leg in (("before", before), ("after", after)):
        if leg.get("allowed_dirty_paths") != list(preparation.PROBE_COPY_PATHS):
            raise RunnerError(f"DrawingML v2 {leg_name} dirty-path policy differs")
        if leg.get("lock_paths") != list(preparation.ROLE_LOCK_PATHS):
            raise RunnerError(f"DrawingML v2 {leg_name} lock-path policy differs")
    try:
        preparation.validate_preparation_receipt(
            Path(receipt_value),
            baseline_root=Path(str(before["source_root"])),
            candidate_root=Path(str(after["source_root"])),
            baseline_revision=str(before["source_revision"]),
            candidate_revision=str(after["source_revision"]),
            trusted_root=Path(str(trusted_root)),
        )
    except (OSError, ValueError) as error:
        raise RunnerError(f"DrawingML v2 policy refused direct generic execution: {error}") from error


def _validate_drawingml_raw_workload(data: Mapping[str, Any]) -> None:
    """Validate one source-fenced DrawingML raw profile.

    The outer DrawingML runner materializes one instance of this generic
    manifest for each of the five production parents.  Keeping the workload
    checks here makes every profile inherit the same Cargo path, binary,
    fresh-process, ABBA, and validator guards as the established native
    lanes.
    """

    manifest_path = "tools/perf-baseline/Cargo.toml"
    protocol_version = data.get("drawingml_protocol_version", 1)
    if protocol_version not in {1, 2}:
        raise RunnerError("DrawingML raw protocol version is not reviewed")
    if protocol_version == 2:
        _validate_drawingml_v2_policy(data)
    validator_relative = (
        "tools/native_qualification/drawingml_raw_validator_v2.py"
        if protocol_version == 2
        else "tools/native_qualification/drawingml_raw_validator.py"
    )
    expected_build_ids = ["drawingml_raw_native", "drawingml_raw_allocator"]
    expected_program_ids = ["raw-native", "raw-allocator"]
    expected_bins = ["drawingml-raw-native", "drawingml-raw-allocator"]
    profile_lane = data.get("drawingml_profile")
    drawingml_lane = data.get("drawingml_lane", profile_lane)
    endpoint = data.get("drawingml_endpoint", "raw")
    owner_route = data.get("drawingml_owner_route")
    raw_scope = data.get("drawingml_raw_scope")
    owner_lanes = {
        "drawingml-chart": "drawingml-chart-raw",
        "drawingml-theme-family": "drawingml-theme-ink",
        "drawingml-ink-actions": "drawingml-theme-ink",
        "drawingml-transform": "drawingml-geometry-color-formatcode2-transform",
    }
    raw_lanes = {
        "drawingml-chart-raw",
        "drawingml-diagram-raw",
        "drawingml-text-media-raw",
        "drawingml-theme-ink",
        "drawingml-geometry-color-formatcode2-transform",
    }
    if not isinstance(profile_lane, str) or not profile_lane:
        raise RunnerError("drawingml raw profile name is required")
    if not isinstance(drawingml_lane, str) or not drawingml_lane:
        raise RunnerError("drawingml raw lane is required")
    if drawingml_lane not in raw_lanes:
        raise RunnerError("drawingml raw lane is not reviewed")
    if endpoint not in {"raw", "owner"}:
        raise RunnerError("drawingml endpoint must be raw or owner")
    if endpoint == "owner":
        if owner_route not in owner_lanes or owner_lanes[owner_route] != drawingml_lane:
            raise RunnerError("drawingml owner route is not reviewed for this lane")
        if raw_scope is not None:
            raise RunnerError("drawingml owner endpoint carries a raw scope")
    else:
        if owner_route is not None:
            raise RunnerError("drawingml raw endpoint carries an owner route")
        if raw_scope not in {None, "bulk", "family"}:
            raise RunnerError("drawingml raw scope is not reviewed")
    for field in (
        "baseline_workspace_lock_sha256",
        "candidate_workspace_lock_sha256",
    ):
        value = data.get(field)
        if not isinstance(value, str) or len(value) != 64 or value != value.lower() or any(
            character not in "0123456789abcdef" for character in value
        ):
            raise RunnerError(f"{field} must carry a lowercase workspace-lock SHA-256")
    for field in ("baseline_workspace_lock_bytes", "candidate_workspace_lock_bytes"):
        value = data.get(field)
        if not isinstance(value, int) or value <= 0:
            raise RunnerError(f"{field} must carry a positive workspace-lock byte count")
    for leg, spec in data["legs"].items():
        commands = spec["build"]
        if [command["id"] for command in commands] != expected_build_ids:
            raise RunnerError(f"{leg} does not preserve the DrawingML raw build command order")
        for command, binary_name in zip(commands, expected_bins):
            argv = command["argv"]
            expected_prefix = ["cargo", "build", "--release", "--locked"]
            if argv[:4] != expected_prefix or "--manifest-path" not in argv:
                raise RunnerError(f"{leg}/{command['id']} does not use the locked perf-baseline manifest")
            manifest_index = argv.index("--manifest-path")
            if argv[manifest_index + 1] != manifest_path:
                raise RunnerError(f"{leg}/{command['id']} uses a different manifest")
            if argv[manifest_index + 2 :].count("--bin") != 1:
                raise RunnerError(f"{leg}/{command['id']} changes the reviewed DrawingML raw binary route")
            bin_index = argv.index("--bin", manifest_index + 2)
            if argv[bin_index + 1 :] != [binary_name]:
                raise RunnerError(f"{leg}/{command['id']} changes the reviewed DrawingML raw binary route")
            feature_args = argv[manifest_index + 2 : bin_index]
            feature_index = 0
            while feature_index < len(feature_args):
                option = feature_args[feature_index]
                if option == "--no-default-features":
                    feature_index += 1
                    continue
                if option == "--features":
                    if (
                        feature_index + 1 >= len(feature_args)
                        or not feature_args[feature_index + 1]
                        or feature_args[feature_index + 1].startswith("--")
                    ):
                        raise RunnerError(f"{leg}/{command['id']} has an empty Cargo feature set")
                    feature_index += 2
                    continue
                raise RunnerError(f"{leg}/{command['id']} has an unreviewed Cargo feature argument")
            if feature_args.count("--features") > 1 or feature_args.count("--no-default-features") > 1:
                raise RunnerError(f"{leg}/{command['id']} repeats a Cargo feature argument")
            expected_feature_names = data.get(
                "drawingml_baseline_features" if leg == "before" else "drawingml_candidate_features",
                [],
            )
            if not isinstance(expected_feature_names, list) or any(
                not isinstance(value, str) or not value for value in expected_feature_names
            ):
                raise RunnerError(f"{leg}/{command['id']} has invalid reviewed feature metadata")
            expected_feature_args = ["--no-default-features"] if leg == "before" else []
            if expected_feature_names:
                expected_feature_args.extend(["--features", ",".join(expected_feature_names)])
            if feature_args != expected_feature_args:
                raise RunnerError(f"{leg}/{command['id']} changed the reviewed Cargo feature set")
            if command["artifacts"] != [f"{{target_dir}}/release/{binary_name}"]:
                raise RunnerError(f"{leg}/{command['id']} artifact path differs")
        programs = spec["programs"]
        if [program["id"] for program in programs] != expected_program_ids:
            raise RunnerError(f"{leg} does not preserve the paired DrawingML raw program order")
        revision = str(spec.get("source_revision", ""))
        if len(revision) != 40 or revision != revision.lower() or any(char not in "0123456789abcdef" for char in revision):
            raise RunnerError(f"{leg} must carry the observed lowercase source revision")
        expected_role = "baseline" if leg == "before" else "candidate"
        for program, mode, binary_name in zip(programs, ("normal", "allocator"), expected_bins):
            expected_argv = [
                "--lane",
                drawingml_lane,
                "--endpoint",
                endpoint,
            ]
            if owner_route is not None:
                expected_argv.extend(["--owner-route", str(owner_route)])
            if raw_scope is not None:
                expected_argv.extend(["--raw-scope", str(raw_scope)])
            expected_argv.extend([
                "--role",
                expected_role,
                "--mode",
                mode,
                "--fixture",
                "{fixture}",
                "--source-revision",
                revision,
                "--iterations",
                "1",
                "--warmups",
                "3",
                "--blocks",
                "5",
                "--samples",
                "20",
                "--target-ms",
                "500",
                "--minimum-ms",
                "150",
                "--json",
                "{output}",
            ])
            if program["argv"] != expected_argv:
                raise RunnerError(f"{leg}/{program['id']} changes the DrawingML raw probe arguments")
            if program["binary"] != f"{{target_dir}}/release/{binary_name}":
                raise RunnerError(f"{leg}/{program['id']} binary path differs")
            output_name = f"{{output_dir}}/dml-{'native' if mode == 'normal' else 'allocator'}-{{leg}}-{{index}}.json"
            if program.get("outputs") != [output_name]:
                raise RunnerError(f"{leg}/{program['id']} output path differs")
            if program.get("stdout") != output_name.removesuffix(".json") + ".stdout":
                raise RunnerError(f"{leg}/{program['id']} stdout path differs")
            if program.get("stderr") != output_name.removesuffix(".json") + ".stderr":
                raise RunnerError(f"{leg}/{program['id']} stderr path differs")
            if program.get("require_stdout") is not False:
                raise RunnerError(f"{leg}/{program['id']} must write its report through --json")

    controls = data.get("controls")
    if not isinstance(controls, list) or len(controls) != 1:
        raise RunnerError("DrawingML raw profile needs exactly one fixture-preflight control")
    control = controls[0]
    if not isinstance(control, dict) or control.get("id") != "drawingml-fixture-preflight":
        raise RunnerError("DrawingML raw fixture-preflight control id differs")
    if control.get("compare") != [
        "fixture_sha256",
        "fixture_bytes",
        "fixture_lane_identity_sha256",
        "generated_fixture_sha256",
        "generated_fixture_parts",
    ]:
        raise RunnerError("DrawingML raw fixture-preflight comparison fields differ")
    control_programs = control.get("programs")
    if not isinstance(control_programs, dict) or set(control_programs) != {"before", "after"}:
        raise RunnerError("DrawingML raw fixture-preflight must have before and after programs")
    for leg in ("before", "after"):
        program = control_programs[leg]
        if not isinstance(program, dict) or program.get("id") != "fixture-preflight":
            raise RunnerError(f"{leg} fixture-preflight program id differs")
        expected_role = "baseline" if leg == "before" else "candidate"
        expected_endpoint_argv = [
            "--lane",
            drawingml_lane,
            "--endpoint",
            endpoint,
        ]
        if owner_route is not None:
            expected_endpoint_argv.extend(["--owner-route", str(owner_route)])
        if raw_scope is not None:
            expected_endpoint_argv.extend(["--raw-scope", str(raw_scope)])
        expected_endpoint_argv.extend([
            "--role",
            expected_role,
            "--mode",
            "normal",
            "--fixture",
            "{fixture}",
            "--source-revision",
            str(data["legs"][leg]["source_revision"]),
            "--iterations",
            "1",
            "--warmups",
            "3",
            "--blocks",
            "5",
            "--samples",
            "20",
            "--target-ms",
            "500",
            "--minimum-ms",
            "150",
            "--preflight",
            "--json",
            "{output}",
        ])
        if program.get("argv") != expected_endpoint_argv:
            raise RunnerError(f"{leg} fixture-preflight changes the reviewed probe arguments")
        if program.get("binary") != "{target_dir}/release/drawingml-raw-native":
            raise RunnerError(f"{leg} fixture-preflight must use the normal binary")
        output_name = f"{{output_dir}}/controls/dml-preflight-{leg}.json"
        if program.get("outputs") != [output_name]:
            raise RunnerError(f"{leg} fixture-preflight output path differs")
        if program.get("stdout") != output_name.removesuffix(".json") + ".stdout":
            raise RunnerError(f"{leg} fixture-preflight stdout path differs")
        if program.get("stderr") != output_name.removesuffix(".json") + ".stderr":
            raise RunnerError(f"{leg} fixture-preflight stderr path differs")
        if program.get("require_stdout") is not False:
            raise RunnerError(f"{leg} fixture-preflight must write its report through --json")

    validators = data.get("validators", [])
    if len(validators) != 1 or validators[0].get("id") != "drawingml-raw-identity":
        raise RunnerError("DrawingML raw profile needs exactly one identity/fidelity validator")
    validator = validators[0]
    if validator.get("outputs") != ["{output_dir}/summary.json"] or validator.get("cwd") != "{output_dir}":
        raise RunnerError("DrawingML raw validator output/cwd differs")
    expected_validator_argv = [
        "python3",
        str(
            _resolve_path(str(data["drawingml_protocol_root"])) / validator_relative
        )
        if protocol_version == 2
        else "{source_root}/" + validator_relative,
        "--output-dir",
        "{output_dir}",
        "--lane",
        drawingml_lane,
        "--endpoint",
        endpoint,
        "--fixture",
        "{fixture}",
        "--baseline-revision",
        str(data["legs"]["before"]["source_revision"]),
        "--candidate-revision",
        str(data["legs"]["after"]["source_revision"]),
        "--source-fence",
        str(data.get("source_fence", "")),
    ]
    if protocol_version == 2:
        shared_receipt = data.get("shared", {}).get("preparation_receipt")
        trusted_root = data.get("drawingml_trusted_root")
        before_root = data["legs"]["before"].get("source_root")
        after_root = data["legs"]["after"].get("source_root")
        if not all(isinstance(value, str) and value for value in (shared_receipt, trusted_root, before_root, after_root)):
            raise RunnerError("DrawingML v2 validator is missing its bound source receipt routes")
        expected_validator_argv.extend([
            "--preparation-receipt", shared_receipt,
            "--baseline-root", before_root,
            "--candidate-root", after_root,
            "--trusted-root", trusted_root,
        ])
    if validator.get("argv") != expected_validator_argv:
        raise RunnerError("DrawingML raw validator is not bound to both source revisions and its fence")
    expected_validator_script = (
        str(_resolve_path(str(data["drawingml_protocol_root"])) / validator_relative)
        if protocol_version == 2
        else "{source_root}/" + validator_relative
    )
    if validator.get("script") != expected_validator_script:
        raise RunnerError("DrawingML raw validator script path differs")
    if validator.get("require_stdout") is not False:
        raise RunnerError("DrawingML raw validator must write summary.json, not claim stdout")


def _validate_xlsx_preparation_manifest(data: Mapping[str, Any]) -> None:
    shared = data["shared"]
    receipt = Path(str(shared["preparation_receipt"]))
    before_root = Path(str(data["legs"]["before"]["source_root"])).resolve(strict=False)
    after_root = Path(str(data["legs"]["after"]["source_root"])).resolve(strict=False)
    candidate_revision = str(data["legs"]["after"]["programs"][0]["argv"][4])
    try:
        validate_preparation_receipt(
            receipt,
            baseline_root=before_root,
            candidate_root=after_root,
            candidate_revision=candidate_revision,
        )
    except (OSError, ValueError) as error:
        raise RunnerError(f"XLSX preparation receipt validation failed: {error}") from error


def _validate_legacy_native_workload(data: Mapping[str, Any]) -> None:
    """Keep the reusable legacy probes on an explicit, reviewed workload.

    The ordinary W4b/W5 lanes intentionally keep their historical hard-coded
    argv.  These three lanes share the same process/identity/cap guard but
    have different standalone manifests and route contracts.  A manifest is
    accepted only when it matches one of those contracts exactly; the
    lane-specific validator then checks the probe output and receipt.
    """

    lane = str(data["lane"])
    preparation_receipt = data.get("preparation_receipt")
    if not isinstance(preparation_receipt, str) or not preparation_receipt:
        raise RunnerError(f"{lane} must bind a source-preparation receipt")
    contracts: dict[str, dict[str, Any]] = {
        "legacy-rtf": {
            "probe": ["tools/perf-baseline/examples/refusal_contract_bench.rs"],
            "lock": "tools/perf-baseline/Cargo.lock",
            "build_id": "rtf_refusal_contract",
            "build_argv": [
                "cargo",
                "build",
                "--release",
                "--manifest-path",
                "tools/perf-baseline/Cargo.toml",
                "--no-default-features",
                "--features",
                "{rtf_features}",
                "--example",
                "refusal_contract_bench",
            ],
            "artifacts": ["{target_dir}/release/examples/refusal_contract_bench"],
            "programs": [("routes", [])],
            "stdout": "{output_dir}/rtf-routes-block-{block}-{leg}-{index}.jsonl",
            "stderr": "{output_dir}/rtf-routes-block-{block}-{leg}-{index}.stderr",
            "blocks": 5,
        },
        "legacy-ole-properties": {
            "probe": ["tools/perf-baseline/examples/ole_w5_probe.rs"],
            "lock": "tools/perf-baseline/Cargo.lock",
            "build_id": "ole_properties_probe",
            "build_argv": [
                "cargo",
                "build",
                "--release",
                "--manifest-path",
                "tools/perf-baseline/Cargo.toml",
                "--no-default-features",
                "--features",
                "{property_features}",
                "--example",
                "ole_w5_probe",
            ],
            "artifacts": ["{target_dir}/release/examples/ole_w5_probe"],
            "programs": [
                (
                    "read",
                    ["--route", "read", "--warmup", "20", "--samples", "500", "--batch", "1"],
                ),
                (
                    "edit",
                    ["--route", "edit", "--warmup", "20", "--samples", "500", "--batch", "1"],
                ),
            ],
            "stdout": "{output_dir}/ole-{program}-block-{block}-{leg}-{index}.csv",
            "stderr": "{output_dir}/ole-{program}-block-{block}-{leg}-{index}.stderr",
            "features": {"before": "current-owner-api,legacy-property-api", "after": "current-owner-api,current-property-api"},
            "blocks": 1,
        },
        "legacy-xls-properties": {
            "probe": ["crates/litchi-ops/examples/ole_metadata_bench.rs"],
            "lock": "Cargo.lock",
            "fixtures": [
                "test-data/ole/doc/FancyFoot.doc",
                "test-data/ole/xls/Simple.xls",
                "test-data/ole/ppt/SampleShow.ppt",
            ],
            "build_id": "xls_metadata_bench",
            "build_argv": [
                "cargo",
                "build",
                "--release",
                "--manifest-path",
                "Cargo.toml",
                "-p",
                "litchi-ops",
                "--no-default-features",
                "--features",
                "doc,xls,ppt",
                "--example",
                "ole_metadata_bench",
            ],
            "artifacts": ["{target_dir}/release/examples/ole_metadata_bench"],
            "programs": [
                ("metadata", []),
            ],
            "stdout": "{output_dir}/xls-properties-block-{block}-{leg}-{index}.jsonl",
            "stderr": "{output_dir}/xls-properties-block-{block}-{leg}-{index}.stderr",
            "blocks": 1,
        },
    }
    contract = contracts[lane]
    if data.get("abba_blocks", 1) != contract["blocks"]:
        raise RunnerError(f"{lane} must use exactly {contract['blocks']} ABBA block(s)")
    shared = data.get("shared", {})
    if shared.get("probe_paths") != contract["probe"] or shared.get("fixture_paths", []) != contract.get("fixtures", []):
        raise RunnerError(f"{lane} does not bind the reviewed reusable probe")
    for leg, spec in data["legs"].items():
        allowed_dirty = spec.get("allowed_dirty_paths")
        if not isinstance(allowed_dirty, list) or not all(
            isinstance(value, str) and value and not Path(value).is_absolute() for value in allowed_dirty
        ):
            raise RunnerError(f"{leg} must declare relative allowed_dirty_paths")
        if spec.get("allow_dirty") is not True:
            raise RunnerError(f"{leg} must retain the reviewed preparation overlay mode")
        if spec.get("lock_paths") != [contract["lock"]]:
            raise RunnerError(f"{leg} does not bind the reviewed {lane} lock path")
        commands = spec["build"]
        if len(commands) != 1 or commands[0].get("id") != contract["build_id"]:
            raise RunnerError(f"{leg} does not preserve the reviewed {lane} build command")
        command = commands[0]
        expected_build = list(contract["build_argv"])
        if lane == "legacy-rtf":
            expected_build[7] = "rtf-legacy-api" if leg == "before" else "rtf-current-api"
        elif lane == "legacy-ole-properties":
            expected_build[7] = contract["features"][leg]
        if command.get("argv") != expected_build or command.get("artifacts") != contract["artifacts"]:
            raise RunnerError(f"{leg} does not preserve the reviewed {lane} build route")
        programs = spec.get("programs", [])
        if [program.get("id") for program in programs] != [item[0] for item in contract["programs"]]:
            raise RunnerError(f"{leg} does not preserve the reviewed {lane} program order")
        for program, (program_id, expected_argv) in zip(programs, contract["programs"]):
            if program.get("binary") != contract["artifacts"][0]:
                raise RunnerError(f"{leg}/{program_id} does not use the built {lane} binary")
            if program.get("argv") != expected_argv:
                raise RunnerError(f"{leg}/{program_id} does not preserve the reviewed route arguments")
            if program.get("outputs") != []:
                raise RunnerError(f"{leg}/{program_id} must use the stdout report, not an unbound output")
            expected_stdout = contract["stdout"].replace("{program}", program_id)
            expected_stderr = contract["stderr"].replace("{program}", program_id)
            if program.get("stdout") != expected_stdout or program.get("stderr") != expected_stderr:
                raise RunnerError(f"{leg}/{program_id} does not use its dedicated report paths")
            if program.get("require_stdout") is not True:
                raise RunnerError(f"{leg}/{program_id} must emit a nonempty report stream")
    validators = data.get("validators", [])
    if len(validators) != 1 or validators[0].get("id") != "legacy-native-contract":
        raise RunnerError(f"{lane} needs exactly one legacy-native-contract validator")
    if validators[0].get("outputs") != ["{output_dir}/summary.json"]:
        raise RunnerError(f"{lane} validator must write summary.json in the output directory")
    validator_argv = validators[0].get("argv")
    if not isinstance(validator_argv, list) or len(validator_argv) != 9:
        raise RunnerError(f"{lane} validator must receive the explicit manifest path")
    if validator_argv[0] != "python3" or validator_argv[-1] != "{manifest_path}":
        raise RunnerError(f"{lane} validator is not bound to the current manifest path")
    validator_script = validators[0].get("script")
    expected_validator_script = (
        Path(__file__).resolve().parent / "native_qualification/legacy_native_validator.py"
    ).resolve()
    if not isinstance(validator_script, str) or _resolve_path(validator_script) != expected_validator_script:
        raise RunnerError(f"{lane} validator script is not bound to the active qualification source")
    if validator_argv[1] != validator_script:
        raise RunnerError(f"{lane} validator argv does not use its bound script")
    if validators[0].get("cwd") != "{output_dir}":
        raise RunnerError(f"{lane} validator cwd is not output-bound")
    environment = data.get("environment", {})
    if environment.get("CARGO_PROFILE_RELEASE_DEBUG") != "0":
        raise RunnerError(f"{lane} must disable release debug information")
    if not isinstance(environment.get("TMPDIR"), str) or not environment["TMPDIR"]:
        raise RunnerError(f"{lane} must bind an external TMPDIR")


def _render(value: str, context: Mapping[str, Any]) -> str:
    def replace(match: re.Match[str]) -> str:
        key = match.group(1)
        if key not in context:
            raise RunnerError(f"unknown manifest template {{{key}}}")
        return str(context[key])

    return _TOKEN.sub(replace, value)


def _render_argv(values: Sequence[str], context: Mapping[str, Any]) -> list[str]:
    return [_render(value, context) for value in values]


def normalize_cargo_argv(argv: Sequence[str]) -> list[str]:
    """Add Cargo flags before ``--`` without duplicating existing flags."""

    values = list(argv)
    if not values or Path(values[0]).name != "cargo":
        raise RunnerError(f"expected a Cargo executable, got {values!r}")
    separator = values.index("--") if "--" in values else len(values)
    before = values[:separator]
    for flag in ("--offline", "--locked"):
        if flag not in before and not any(item.startswith(flag + "=") for item in before):
            before.append(flag)
    return [*before, *values[separator:]]


def _cargo_config_files(cwd: Path) -> list[Path]:
    files: list[Path] = []
    current = cwd
    while True:
        config_dir = current / ".cargo"
        for name in ("config.toml", "config"):
            candidate = config_dir / name
            if candidate.is_symlink():
                raise RunnerError(f"Cargo config may not be a symlink: {candidate}")
            if candidate.exists():
                files.append(candidate)
        if current.parent == current:
            break
        current = current.parent
    return files


def _validate_cargo_paths(argv: Sequence[str], *, cwd: Path, source: Path, target: Path) -> None:
    separator = argv.index("--") if "--" in argv else len(argv)
    before = list(argv[:separator])
    index = 1
    while index < len(before):
        token = before[index]
        if token == "--config" or token.startswith("--config="):
            raise RunnerError("Cargo --config is forbidden; it can redirect the target or manifest")
        if token in {"--target-dir", "--artifact-dir", "--out-dir"} or token.startswith(
            ("--target-dir=", "--artifact-dir=", "--out-dir=")
        ):
            raise RunnerError("Cargo output-directory overrides are forbidden; CARGO_TARGET_DIR is runner-owned")
        if token == "--manifest-path":
            if index + 1 >= len(before):
                raise RunnerError("Cargo --manifest-path has no value")
            value = before[index + 1]
            manifest = _resolve_path(value, base=cwd)
            if not _under(manifest, source) or not manifest.is_file():
                raise RunnerError(f"Cargo manifest escapes the intended source root: {manifest}")
            index += 2
            continue
        if token.startswith("--manifest-path="):
            manifest = _resolve_path(token.split("=", 1)[1], base=cwd)
            if not _under(manifest, source) or not manifest.is_file():
                raise RunnerError(f"Cargo manifest escapes the intended source root: {manifest}")
        index += 1
    for config in _cargo_config_files(cwd):
        try:
            data = tomllib.loads(config.read_text(encoding="utf-8"))
        except (OSError, tomllib.TOMLDecodeError) as error:
            raise RunnerError(f"cannot parse Cargo config {config}: {error}") from error
        build = data.get("build", {})
        if not isinstance(build, dict) or "target-dir" not in build:
            continue
        configured = _resolve_path(str(build["target-dir"]), base=config.parent)
        if configured != target:
            raise RunnerError(f"Cargo config redirects target outside the runner target: {config} -> {configured}")


def _build_phase_identity(
    manifest: Mapping[str, Any], manifest_path: Path, phase_id: str
) -> tuple[list[str], Path, Path, Path]:
    """Resolve and validate one recorded build phase from its manifest."""

    legs = manifest.get("legs")
    if not isinstance(legs, Mapping):
        raise RunnerError("phase manifest legs are malformed")
    shared = manifest.get("shared", {})
    if not isinstance(shared, Mapping):
        raise RunnerError("phase manifest shared section is malformed")
    output = _resolve_path(str(manifest.get("output_dir", "")))
    for leg, spec in legs.items():
        if not isinstance(leg, str) or not isinstance(spec, Mapping):
            raise RunnerError("phase manifest leg is malformed")
        commands = spec.get("build")
        if not isinstance(commands, list):
            raise RunnerError(f"phase manifest build list is malformed for {leg}")
        for command in commands:
            if not isinstance(command, Mapping) or not isinstance(command.get("id"), str):
                raise RunnerError(f"phase manifest build command is malformed for {leg}")
            if f"build-{leg}-{command['id']}" != phase_id:
                continue
            source = _resolve_path(str(spec.get("source_root", "")))
            target = _resolve_path(str(spec.get("target_dir", "")))
            shared_fixtures = shared.get("fixture_paths", [])
            leg_fixtures = spec.get("fixture_paths", [])
            if not isinstance(shared_fixtures, list) or not isinstance(leg_fixtures, list):
                raise RunnerError(f"phase manifest fixture paths are malformed for {leg}")
            fixture_values = [*shared_fixtures, *leg_fixtures]
            fixture = ""
            if len(fixture_values) == 1:
                fixture = str(_resolve_path(str(fixture_values[0]), base=source))
            context: dict[str, Any] = {
                "source_root": source,
                "target_dir": target,
                "output_dir": output,
                "manifest_path": manifest_path,
                "leg": leg,
                "index": 0,
                "rounds": int(manifest.get("rounds", 31)),
                "block": 1,
                "fixture": fixture,
            }
            argv_value = command.get("argv")
            if not isinstance(argv_value, list) or not all(isinstance(item, str) for item in argv_value):
                raise RunnerError(f"phase manifest argv is malformed for {phase_id}")
            argv = normalize_cargo_argv(_render_argv(argv_value, context))
            cwd = _resolve_path(_render(str(command.get("cwd", ".")), context), base=source)
            if not _under(cwd, source):
                raise RunnerError(f"phase build cwd escapes source root: {cwd}")
            _validate_cargo_paths(argv, cwd=cwd, source=source, target=target)
            return argv, cwd, target, source
    raise RunnerError(f"phase {phase_id!r} is not an admitted manifest build entry")


def _read_proc(pid: int, *, allow_restricted_metadata: bool = False) -> dict[str, Any] | None:
    proc = Path("/proc") / str(pid)
    try:
        comm = (proc / "comm").read_text(encoding="utf-8", errors="replace").strip()
        cmdline = (proc / "cmdline").read_bytes().split(b"\0")
        cmdline = [part.decode("utf-8", "replace") for part in cmdline if part]
        stat = (proc / "stat").read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        # A process can disappear between the directory enumeration and any
        # one of these reads.  Only those two kernel errors mean the identity
        # is expected to be gone; permission, I/O, and filesystem failures
        # invalidate the whole ownership scan and must fail closed.
        if error.errno in (errno.ENOENT, errno.ESRCH):
            return None
        raise ProcMonitorError(f"cannot inspect /proc/{pid}: {error}") from error
    restricted_metadata: list[str] = []
    for name in ("exe", "cwd"):
        try:
            value = os.readlink(proc / name)
        except OSError as error:
            if error.errno in (errno.ENOENT, errno.ESRCH):
                return None
            if allow_restricted_metadata and error.errno in (errno.EACCES, errno.EPERM):
                value = ""
                restricted_metadata.append(name)
            else:
                raise ProcMonitorError(f"cannot inspect /proc/{pid}/{name}: {error}") from error
        if name == "exe":
            exe = value
        else:
            cwd = value
    right = stat.rsplit(") ", 1)
    if len(right) != 2:
        raise ProcMonitorError(f"malformed /proc/{pid}/stat")
    fields = right[1].split()
    try:
        parent = int(fields[1])
        process_group = int(fields[2])
        starttime = fields[19]
    except (IndexError, ValueError) as error:
        raise ProcMonitorError(f"malformed /proc/{pid}/stat fields") from error
    entry = {
        "pid": pid,
        "ppid": parent,
        "pgrp": process_group,
        "starttime": starttime,
        "comm": comm,
        "exe": exe,
        "cwd": cwd,
        "cmdline": cmdline,
    }
    if restricted_metadata:
        entry["restricted_metadata"] = restricted_metadata
    return entry


_SUBREAPER_ENABLED = False


def _enable_child_subreaper() -> None:
    """Make this runner the owner of orphaned descendants on Linux.

    Cargo normally waits for every compiler child, but the process guard must
    also be correct when a helper double-forks or calls ``setsid``.  A process
    group is not an ownership boundary: such a child can leave the group and,
    after its parent exits, be reparented to init before a post-exit scan sees
    it.  Linux's child-subreaper facility keeps those descendants attached to
    this runner so their recorded PID/start-time identities can be checked and
    cleaned up.  There is no safe equivalent in the supported runner
    environment on other platforms, so fail closed there rather than claiming
    process-tree ownership we cannot enforce.
    """

    global _SUBREAPER_ENABLED
    if _SUBREAPER_ENABLED:
        return
    if sys.platform != "linux":
        raise RunnerError("native child lifecycle ownership requires Linux child subreaper support")
    try:
        import ctypes

        libc = ctypes.CDLL(None, use_errno=True)
        prctl = libc.prctl
        prctl.argtypes = [
            ctypes.c_int,
            ctypes.c_ulong,
            ctypes.c_ulong,
            ctypes.c_ulong,
            ctypes.c_ulong,
        ]
        prctl.restype = ctypes.c_int
        # Linux include/uapi/linux/prctl.h: PR_SET_CHILD_SUBREAPER.
        result = prctl(36, 1, 0, 0, 0)
        if result != 0:
            errno = ctypes.get_errno()
            raise OSError(errno, os.strerror(errno))
    except (AttributeError, OSError, ImportError) as error:
        raise RunnerError(f"cannot enable Linux child subreaper: {error}") from error
    _SUBREAPER_ENABLED = True


def _process_record(entry: Mapping[str, Any]) -> dict[str, Any]:
    """Keep the identity needed to avoid signalling a reused PID."""

    return {
        "pid": int(entry["pid"]),
        "starttime": str(entry["starttime"]),
        "ppid": int(entry["ppid"]),
        "pgrp": int(entry["pgrp"]),
        "comm": str(entry.get("comm", "")),
        "exe": str(entry.get("exe", "")),
        "cwd": str(entry.get("cwd", "")),
        "cmdline": [str(value) for value in entry.get("cmdline", [])],
    }


def _same_process_identity(entry: Mapping[str, Any], record: Mapping[str, Any]) -> bool:
    """Compare the kernel start time, the PID-reuse guard's stable identity."""

    return (
        entry.get("pid") == record.get("pid")
        and str(entry.get("starttime")) == str(record.get("starttime"))
    )


def _proc_snapshot() -> dict[int, dict[str, Any]]:
    entries: dict[int, dict[str, Any]] = {}
    try:
        pids = [int(item.name) for item in Path("/proc").iterdir() if item.name.isdigit()]
    except OSError as error:
        raise ProcMonitorError(f"cannot enumerate /proc: {error}") from error
    for pid in pids:
        entry = _read_proc(pid, allow_restricted_metadata=True)
        if entry is not None:
            # ``comm``, the full decoded ``cmdline``, and ``stat`` remain
            # readable when only exe/cwd are restricted.  These fields prove
            # that an opaque entry is neither a heavy executable nor a
            # Litchi-scoped command; otherwise even partial metadata is a
            # monitoring failure rather than an allow-list shortcut.
            if entry.get("restricted_metadata") and (_is_heavy(entry) or _in_litchi_scope(entry)):
                raise ProcMonitorError(
                    f"cannot establish complete process identity for relevant /proc/{pid}"
                )
            entries[pid] = entry
    return entries


def _in_litchi_scope(entry: Mapping[str, Any]) -> bool:
    haystack = [str(entry.get("exe", "")), str(entry.get("cwd", "")), *entry.get("cmdline", [])]
    return any(any(value == str(root) or value.startswith(str(root) + "/") for root in LITCHI_ROOTS) for value in haystack)


def _is_heavy(entry: Mapping[str, Any]) -> bool:
    names = {Path(str(entry.get("exe", ""))).name, str(entry.get("comm", ""))}
    if names & HEAVY_NAMES:
        return True
    first = entry.get("cmdline", [""])[0] if entry.get("cmdline") else ""
    return Path(str(first)).name in HEAVY_NAMES


def _process_command_class(entry: Mapping[str, Any]) -> str | None:
    command = entry.get("cmdline", [])
    if not isinstance(command, list) or not command:
        return None
    executable = Path(str(command[0])).name
    if executable == "cargo":
        return _cargo_command_class([str(value) for value in command])
    return None


def _process_starttime(pid: int) -> str:
    entry = _read_proc(pid)
    if entry is None or not isinstance(entry.get("starttime"), str):
        raise RunnerError(f"cannot read process starttime for pid {pid}")
    return str(entry["starttime"])


def _descendants(snapshot: Mapping[int, Mapping[str, Any]], root: int) -> set[int]:
    allowed = {root}
    changed = True
    while changed:
        changed = False
        for pid, entry in snapshot.items():
            if pid not in allowed and int(entry.get("ppid", -1)) in allowed:
                allowed.add(pid)
                changed = True
    return allowed


class ProcessGuard:
    """Reject unrelated compiler/profiler activity in any Litchi worktree."""

    @staticmethod
    def _check_child_tree(
        snapshot: Mapping[int, Mapping[str, Any]],
        root: int,
        *,
        expected_command_class: str | None,
        label: str,
    ) -> set[int]:
        descendants = _descendants(snapshot, root)
        root_entry = snapshot.get(root)
        if root_entry is None:
            return descendants
        if expected_command_class is not None:
            actual = _process_command_class(root_entry)
            if actual != expected_command_class:
                raise RunnerError(
                    f"{label} command class changed: expected {expected_command_class}, got {actual}"
                )
        for pid in descendants:
            if pid == root:
                continue
            entry = snapshot[pid]
            names = {Path(str(entry.get("exe", ""))).name, str(entry.get("comm", ""))}
            if names & DISALLOWED_DESCENDANT_NAMES:
                raise RunnerError(f"{label} spawned a disallowed compiler/profiler: {entry!r}")
            if "cargo" in names:
                raise RunnerError(f"{label} spawned a nested Cargo process: {entry!r}")
            if "clippy-driver" in names and expected_command_class != "cargo-clippy":
                raise RunnerError(f"{label} spawned clippy-driver outside cargo clippy: {entry!r}")
        return descendants

    def offenders(
        self,
        allowed_root: int | None = None,
        *,
        overlap: "BuildOverlapLease | None" = None,
        expected_command_class: str | None = None,
    ) -> list[dict[str, Any]]:
        snapshot = _proc_snapshot()
        allowed = (
            self._check_child_tree(
                snapshot,
                allowed_root,
                expected_command_class=expected_command_class,
                label="owned child",
            )
            if allowed_root is not None
            else set()
        )
        if overlap is not None:
            allowed.update(overlap.allowed_peer_processes(snapshot))
        return [
            entry
            for pid, entry in snapshot.items()
            if pid not in allowed and _in_litchi_scope(entry) and _is_heavy(entry)
        ]

    def require_quiet(self) -> None:
        offenders = self.offenders()
        if offenders:
            raise RunnerError(f"unrelated Litchi builder/profiler is running: {offenders!r}")

    def require_compile_quiet(self, overlap: "BuildOverlapLease") -> None:
        offenders = self.offenders(overlap=overlap)
        if offenders:
            raise RunnerError(f"unregistered Litchi builder/profiler is running: {offenders!r}")

    @staticmethod
    def loadavg() -> list[float] | None:
        try:
            return [float(value) for value in os.getloadavg()]
        except (AttributeError, OSError):
            return None


class BuildOverlapLease:
    """A short-lived, externally granted lease for compile-only overlap.

    The permit is immutable and coordinator-owned.  Each participating job
    creates its own lease file under the permit's external lease directory;
    the file binds the job group, root PID/starttime, manifest identity and
    exact Cargo command class.  A missing/reused PID, changed permit, stale
    active lease, phase change, or expiry is a hard failure.
    """

    def __init__(
        self,
        manifest: Mapping[str, Any],
        manifest_path: Path,
        *,
        clock: Callable[[], _dt.datetime] | None = None,
    ):
        execution = manifest.get("build_execution")
        if not isinstance(execution, dict) or execution.get("class") != "compile_only":
            raise RunnerError("compile_only build_execution is required for an overlap lease")
        permit_value = execution.get("permit_path")
        if not isinstance(permit_value, str) or not permit_value:
            raise RunnerError("compile_only build_execution needs permit_path")
        self.manifest = manifest
        self.manifest_path = manifest_path.resolve()
        self._clock = clock or _utc_now
        self.manifest_hash, self.manifest_bytes = _sha256_file(self.manifest_path)
        self.permit_path = _resolve_path(permit_value)
        if self.permit_path.is_symlink() or not self.permit_path.is_file():
            raise RunnerError(f"build permit is missing or symlinked: {self.permit_path}")
        self.permit_hash, self.permit_bytes = _sha256_file(self.permit_path)
        try:
            value = json.loads(self.permit_path.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
            raise RunnerError(f"cannot read build permit: {error}") from error
        if not isinstance(value, dict) or value.get("schema") != BUILD_PERMIT_SCHEMA:
            raise RunnerError(f"build permit schema must be {BUILD_PERMIT_SCHEMA!r}")
        self.permit = value
        self.grant_id = self._text("grant_id")
        self.job_group = self._text("job_group")
        self.expires = _parse_utc(value.get("expires_utc"), "build permit expires_utc")
        if self.expires <= self._clock():
            raise RunnerError("build permit is expired")
        max_jobs = value.get("max_jobs")
        if isinstance(max_jobs, bool) or not isinstance(max_jobs, int) or not 1 <= max_jobs <= 4:
            raise RunnerError("build permit max_jobs must be between 1 and 4")
        self.max_jobs = max_jobs
        self.allowed_lanes = self._text_list("allowed_lanes")
        self.allowed_manifests = self._text_list("allowed_manifest_sha256")
        for item in self.allowed_manifests:
            _sha256_shape(item, "build permit allowed_manifest_sha256 entry")
        self.allowed_command_classes = self._text_list("allowed_command_classes")
        if any(item not in PEER_COMMAND_CLASSES for item in self.allowed_command_classes):
            raise RunnerError("build permit contains an unknown command class")
        bindings = value.get("manifest_bindings")
        if not isinstance(bindings, list) or not bindings:
            raise RunnerError("build permit needs exact manifest_bindings")
        self.manifest_bindings: dict[tuple[str, str], dict[str, Any]] = {}
        for binding in bindings:
            if not isinstance(binding, dict):
                raise RunnerError("build permit manifest binding must be an object")
            binding_path_value = binding.get("path")
            if not isinstance(binding_path_value, str) or not binding_path_value:
                raise RunnerError("build permit manifest binding path is malformed")
            binding_path = str(_resolve_path(binding_path_value))
            binding_hash = _sha256_shape(
                binding.get("sha256"), "build permit manifest binding sha256"
            )
            binding_lane = binding.get("lane")
            if not isinstance(binding_lane, str) or binding_lane not in self.allowed_lanes:
                raise RunnerError("build permit manifest binding lane is not admitted")
            binding_roots = self._binding_paths(binding, "source_roots")
            binding_targets = self._binding_paths(binding, "target_dirs")
            if binding_hash not in self.allowed_manifests:
                raise RunnerError("build permit manifest binding is not in allowed_manifest_sha256")
            key = (binding_path, binding_hash)
            if key in self.manifest_bindings:
                raise RunnerError("build permit repeats a manifest binding")
            self.manifest_bindings[key] = {
                "path": binding_path,
                "sha256": binding_hash,
                "lane": binding_lane,
                "source_roots": binding_roots,
                "target_dirs": binding_targets,
            }
        if manifest.get("lane") not in self.allowed_lanes:
            raise RunnerError("build permit does not admit this lane")
        if self.manifest_hash not in self.allowed_manifests:
            raise RunnerError("build permit does not admit this manifest hash")
        manifest_key = (str(self.manifest_path), self.manifest_hash)
        manifest_binding = self.manifest_bindings.get(manifest_key)
        if manifest_binding is None:
            raise RunnerError("build permit does not bind this manifest path and hash")
        source_roots = [
            str(_resolve_path(str(spec["source_root"]))) for spec in manifest.get("legs", {}).values()
        ]
        target_dirs = [
            str(_resolve_path(str(spec["target_dir"]))) for spec in manifest.get("legs", {}).values()
        ]
        if (
            manifest_binding["lane"] != manifest.get("lane")
            or manifest_binding["source_roots"] != source_roots
            or manifest_binding["target_dirs"] != target_dirs
        ):
            raise RunnerError("build permit manifest binding paths differ from the current manifest")
        if "cargo-build" not in self.allowed_command_classes:
            raise RunnerError("build permit does not admit the native cargo-build class")
        lease_value = value.get("lease_dir")
        if not isinstance(lease_value, str) or not lease_value:
            raise RunnerError("build permit needs an external lease_dir")
        self.lease_dir = _resolve_path(lease_value)
        source_roots = [Path(path) for path in source_roots]
        output = _resolve_path(str(manifest.get("output_dir", "")))
        _ensure_external(self.lease_dir, source_roots, "build permit lease_dir")
        if _overlap(self.lease_dir, output):
            raise RunnerError("build permit lease_dir must be disjoint from output_dir")
        if _overlap(self.permit_path, self.lease_dir):
            raise RunnerError("build permit path must be disjoint from lease_dir")
        if self.lease_dir.exists() and not self.lease_dir.is_dir():
            raise RunnerError(f"build permit lease_dir is not a directory: {self.lease_dir}")
        self.root_pid = os.getpid()
        self.root_starttime = _process_starttime(self.root_pid)
        self.lease_path: Path | None = None
        self.lease_id: str | None = None
        self.started_peers: list[dict[str, Any]] = []
        self.ended_peers: list[dict[str, Any]] = []
        self.owned_jobs: list[dict[str, Any]] = []
        self.bound_pid: int | None = None
        self.bound_starttime: str | None = None
        self.bound_phase: str | None = None
        self.phase_argv: list[str] | None = None
        self.phase_cwd: str | None = None
        self.phase_target_dir: str | None = None

    @staticmethod
    def _binding_paths(binding: Mapping[str, Any], key: str) -> list[str]:
        values = binding.get(key)
        if not isinstance(values, list) or not values or any(not isinstance(item, str) for item in values):
            raise RunnerError(f"build permit manifest binding {key} is malformed")
        return [str(_resolve_path(item)) for item in values]

    def _text(self, key: str) -> str:
        value = self.permit.get(key)
        if not isinstance(value, str) or not value or "/" in value:
            raise RunnerError(f"build permit {key} must be a non-empty identifier")
        return value

    def _text_list(self, key: str) -> list[str]:
        value = self.permit.get(key)
        if not isinstance(value, list) or not value or any(not isinstance(item, str) or not item for item in value):
            raise RunnerError(f"build permit {key} must be a non-empty string list")
        return [str(item) for item in value]

    def _lock_path(self) -> Path:
        return self.lease_dir / ".lock"

    def _lease_record(self, *, status: str = "active") -> dict[str, Any]:
        if self.lease_id is None:
            raise RunnerError("build overlap lease has no lease id")
        bound = self.bound_pid is not None
        root_pid = self.bound_pid if bound else self.root_pid
        root_starttime = self.bound_starttime if bound else self.root_starttime
        command_class = "cargo-build" if bound else "controller-reservation"
        return {
            "schema": BUILD_PERMIT_SCHEMA,
            "status": status,
            "lease_id": self.lease_id,
            "grant_id": self.grant_id,
            "job_group": self.job_group,
            "permit_sha256": self.permit_hash,
            "root_pid": root_pid,
            "root_starttime": root_starttime,
            "controller_pid": self.root_pid,
            "controller_starttime": self.root_starttime,
            "lane": self.manifest["lane"],
            "manifest_path": str(self.manifest_path),
            "manifest_sha256": self.manifest_hash,
            "command_class": command_class,
            "allowed_command_classes": list(self.allowed_command_classes),
            "phase": "build",
            "phase_id": self.bound_phase,
            "phase_argv": list(self.phase_argv) if self.phase_argv is not None else None,
            "phase_cwd": self.phase_cwd,
            "phase_target_dir": self.phase_target_dir,
            "source_roots": [
                str(_resolve_path(str(spec["source_root"])))
                for spec in self.manifest["legs"].values()
            ],
            "target_dirs": [
                str(_resolve_path(str(spec["target_dir"])))
                for spec in self.manifest["legs"].values()
            ],
            "created_utc": _now(),
            "expires_utc": self.permit["expires_utc"],
        }

    def _read_records_locked(self) -> list[tuple[Path, dict[str, Any]]]:
        if self.lease_dir.is_symlink():
            raise RunnerError(f"build permit lease_dir is symlinked: {self.lease_dir}")
        self.lease_dir.mkdir(parents=True, exist_ok=True)
        records: list[tuple[Path, dict[str, Any]]] = []
        for path in sorted(self.lease_dir.glob("*.json")):
            if path.is_symlink():
                raise RunnerError(f"build permit lease is symlinked: {path}")
            try:
                value = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
                raise RunnerError(f"cannot read build lease {path}: {error}") from error
            if not isinstance(value, dict) or value.get("schema") != BUILD_PERMIT_SCHEMA:
                raise RunnerError(f"build lease schema mismatch: {path}")
            records.append((path, value))
        return records

    def _phase_manifest(self, path: Path) -> Mapping[str, Any]:
        if path == self.manifest_path:
            return self.manifest
        try:
            value = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
            raise RunnerError(f"cannot read phase manifest {path}: {error}") from error
        if not isinstance(value, dict):
            raise RunnerError(f"phase manifest is not an object: {path}")
        return _expand_environment(value)

    def _validate_record(
        self,
        path: Path,
        value: Mapping[str, Any],
        *,
        snapshot: Mapping[int, Mapping[str, Any]],
    ) -> tuple[int, str] | None:
        if value.get("status") == "released":
            return None
        if value.get("status") != "active":
            raise RunnerError(f"build lease has an unknown status: {path}")
        is_current = path == self.lease_path and value.get("lease_id") == self.lease_id
        if value.get("grant_id") != self.grant_id or value.get("job_group") != self.job_group:
            raise RunnerError(f"build lease belongs to a different grant/group: {path}")
        if value.get("permit_sha256") != self.permit_hash:
            raise RunnerError(f"build lease permit hash differs: {path}")
        pid = value.get("root_pid")
        starttime = value.get("root_starttime")
        if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0 or not isinstance(starttime, str) or not starttime:
            raise RunnerError(f"build lease root identity is malformed: {path}")
        controller_pid = value.get("controller_pid")
        controller_starttime = value.get("controller_starttime")
        if (
            isinstance(controller_pid, bool)
            or not isinstance(controller_pid, int)
            or controller_pid <= 0
            or not isinstance(controller_starttime, str)
            or not controller_starttime
        ):
            raise RunnerError(f"build lease controller identity is malformed: {path}")
        if value.get("lane") not in self.allowed_lanes:
            raise RunnerError(f"build lease lane is not admitted: {path}")
        manifest_path = value.get("manifest_path")
        if not isinstance(manifest_path, str) or not manifest_path:
            raise RunnerError(f"build lease manifest_path is malformed: {path}")
        resolved_manifest_path = _resolve_path(manifest_path)
        manifest_hash = _sha256_shape(value.get("manifest_sha256"), "build lease manifest_sha256")
        if manifest_hash not in self.allowed_manifests:
            raise RunnerError(f"build lease manifest is not admitted: {path}")
        binding = self.manifest_bindings.get((str(resolved_manifest_path), manifest_hash))
        if binding is None:
            raise RunnerError(f"build lease manifest binding is not admitted: {path}")
        if not resolved_manifest_path.is_file() or resolved_manifest_path.is_symlink():
            raise RunnerError(f"build lease manifest is missing or symlinked: {path}")
        actual_manifest_hash, _ = _sha256_file(resolved_manifest_path)
        if actual_manifest_hash != manifest_hash:
            raise RunnerError(f"build lease manifest content changed: {path}")
        if value.get("lane") != binding["lane"]:
            raise RunnerError(f"build lease lane differs from its manifest binding: {path}")
        command_class = value.get("command_class")
        allowed_classes = value.get("allowed_command_classes")
        if (
            not isinstance(allowed_classes, list)
            or any(item not in self.allowed_command_classes for item in allowed_classes)
            or "cargo-build" not in allowed_classes
        ):
            raise RunnerError(f"build lease allowed command classes are not admitted: {path}")
        if value.get("phase") != "build":
            raise RunnerError(f"build lease phase is not build: {path}")
        phase_id = value.get("phase_id")
        phase_argv = value.get("phase_argv")
        phase_cwd = value.get("phase_cwd")
        phase_target_dir = value.get("phase_target_dir")
        if phase_id is None:
            if phase_argv is not None or phase_cwd is not None or phase_target_dir is not None:
                raise RunnerError(f"idle controller build lease carries a phase binding: {path}")
        elif (
            not isinstance(phase_id, str)
            or not phase_id
            or not isinstance(phase_argv, list)
            or not phase_argv
            or any(not isinstance(item, str) or not item for item in phase_argv)
            or _cargo_command_class(phase_argv) is None
            or not isinstance(phase_cwd, str)
            or not phase_cwd
            or not isinstance(phase_target_dir, str)
            or not phase_target_dir
        ):
            raise RunnerError(f"build lease phase binding is malformed: {path}")
        phase_cwd_path = _resolve_path(phase_cwd) if isinstance(phase_cwd, str) else None
        phase_target_path = _resolve_path(phase_target_dir) if isinstance(phase_target_dir, str) else None
        if phase_id is not None and (
            phase_cwd_path is None
            or not any(_under(phase_cwd_path, Path(root)) for root in binding["source_roots"])
            or phase_target_path is None
            or str(phase_target_path) not in binding["target_dirs"]
        ):
            raise RunnerError(f"build lease phase paths differ from their manifest binding: {path}")
        if phase_id is not None:
            phase_manifest = self._phase_manifest(resolved_manifest_path)
            expected_argv, expected_cwd, expected_target, expected_source = _build_phase_identity(
                phase_manifest,
                resolved_manifest_path,
                phase_id,
            )
            if (
                phase_argv != expected_argv
                or phase_cwd != str(expected_cwd)
                or phase_target_dir != str(expected_target)
                or str(expected_source) not in binding["source_roots"]
            ):
                raise RunnerError(f"build lease phase differs from its admitted manifest entry: {path}")
        if command_class == "controller-reservation":
            if phase_id is not None and _cargo_command_class(phase_argv) != "cargo-build":
                raise RunnerError(f"controller build lease phase is not cargo build: {path}")
        elif (
            command_class not in self.allowed_command_classes
            or not isinstance(phase_id, str)
            or not phase_id
            or _cargo_command_class(phase_argv) != command_class
        ):
            raise RunnerError(f"build lease command phase is not admitted: {path}")
        for key in ("source_roots", "target_dirs"):
            paths = value.get(key)
            if not isinstance(paths, list) or not paths or any(not isinstance(item, str) for item in paths):
                raise RunnerError(f"build lease {key} are malformed: {path}")
            normalized = [str(_resolve_path(item)) for item in paths]
            if normalized != binding[key]:
                raise RunnerError(f"build lease {key} differ from their manifest binding: {path}")
        if value.get("expires_utc") != self.permit.get("expires_utc"):
            raise RunnerError(f"build lease expiry differs from its permit: {path}")
        lease_expiry = _parse_utc(value.get("expires_utc"), f"build lease {path} expires_utc")
        if lease_expiry <= _dt.datetime.now(_dt.timezone.utc):
            raise RunnerError(f"active build lease is expired: {path}")
        controller_entry = snapshot.get(controller_pid)
        if controller_entry is None:
            raise RunnerError(f"active build lease controller process is gone: {path}")
        if controller_entry.get("starttime") != controller_starttime:
            raise RunnerError(f"active build lease controller PID was reused: {path}")
        entry = snapshot.get(pid)
        if entry is None:
            # A Cargo child can exit just before its controller rewrites the
            # lease back to a controller reservation.  Keep the slot
            # occupied while that live controller performs the transition;
            # this record never enters the peer process allow-list.
            if not is_current and command_class == "cargo-build":
                return pid, str(value.get("lease_id", ""))
            raise RunnerError(f"active build lease root process is gone: {path}")
        if entry.get("starttime") != starttime:
            raise RunnerError(f"active build lease root PID was reused: {path}")
        if command_class in PEER_COMMAND_CLASSES and entry.get("ppid") != controller_pid:
            raise RunnerError(f"active Cargo build parent changed from its controller: {path}")
        if command_class == "cargo-build" and (
            entry.get("cmdline") != phase_argv or entry.get("cwd") != phase_cwd
        ):
            raise RunnerError(f"active Cargo build phase changed from its published phase: {path}")
        if is_current:
            expected_pid = self.bound_pid if self.bound_pid is not None else self.root_pid
            expected_starttime = self.bound_starttime if self.bound_starttime is not None else self.root_starttime
            expected_class = "cargo-build" if self.bound_pid is not None else "controller-reservation"
            if (
                value.get("manifest_path") != str(self.manifest_path)
                or value.get("lane") != self.manifest.get("lane")
                or value.get("manifest_sha256") != self.manifest_hash
                or pid != expected_pid
                or starttime != expected_starttime
                or command_class != expected_class
                or phase_id != self.bound_phase
                or phase_argv != self.phase_argv
                or phase_cwd != self.phase_cwd
                or phase_target_dir != self.phase_target_dir
                or controller_pid != self.root_pid
                or controller_starttime != self.root_starttime
            ):
                raise RunnerError(f"current build lease identity changed: {path}")
            if self.bound_pid is not None and _process_command_class(entry) != "cargo-build":
                raise RunnerError(f"current Cargo child command phase changed: {path}")
            return pid, str(value.get("lease_id", ""))
        if command_class == "controller-reservation":
            # The controller is a capacity reservation only.  It is never
            # admitted to the process allow-list; reject a record that
            # pretends a live Cargo child is still merely a controller.
            if _process_command_class(entry) is not None:
                raise RunnerError(f"controller build lease root is already a Cargo phase: {path}")
            return pid, str(value.get("lease_id", ""))
        actual = _process_command_class(entry)
        if actual != command_class:
            raise RunnerError(
                f"active build lease command phase changed for {path}: expected {command_class}, got {actual}"
            )
        return pid, str(value.get("lease_id", ""))

    def _active_records_locked(
        self, snapshot: Mapping[int, Mapping[str, Any]]
    ) -> list[tuple[Path, dict[str, Any], int, str]]:
        active: list[tuple[Path, dict[str, Any], int, str]] = []
        for path, value in self._read_records_locked():
            identity = self._validate_record(path, value, snapshot=snapshot)
            if identity is not None:
                active.append((path, value, identity[0], identity[1]))
        return active

    def _with_lock(self, callback: Any) -> Any:
        if self.lease_dir.is_symlink():
            raise RunnerError(f"build permit lease_dir is symlinked: {self.lease_dir}")
        self.lease_dir.mkdir(parents=True, exist_ok=True)
        lock_path = self._lock_path()
        if lock_path.is_symlink():
            raise RunnerError(f"build permit lock is symlinked: {lock_path}")
        with lock_path.open("a+", encoding="utf-8") as stream:
            fcntl.flock(stream.fileno(), fcntl.LOCK_EX)
            try:
                return callback()
            finally:
                fcntl.flock(stream.fileno(), fcntl.LOCK_UN)

    def acquire(self) -> None:
        snapshot = _proc_snapshot()
        if snapshot.get(self.root_pid, {}).get("starttime") != self.root_starttime:
            raise RunnerError("build overlap root PID changed before lease acquisition")
        self.lease_id = f"{self.grant_id}-{self.root_pid}-{self.root_starttime}"
        self.lease_path = self.lease_dir / f"{self.lease_id}.json"

        def create() -> None:
            active = self._active_records_locked(snapshot)
            if len(active) >= self.max_jobs:
                raise RunnerError(f"build permit has no free slot (max_jobs={self.max_jobs})")
            if self.lease_path is None:
                raise RunnerError("build overlap lease path was not prepared")
            if self.lease_path.exists() or self.lease_path.is_symlink():
                raise RunnerError(f"build overlap lease path already exists: {self.lease_path}")
            with self.lease_path.open("x", encoding="utf-8") as stream:
                json.dump(self._lease_record(), stream, indent=2, sort_keys=True)
                stream.write("\n")
                stream.flush()
                os.fsync(stream.fileno())

        self._with_lock(create)
        self.ensure_valid()
        self.started_peers = self.peer_snapshots(_proc_snapshot())

    def _write_current_record(self) -> None:
        if self.lease_path is None or self.lease_path.is_symlink() or not self.lease_path.is_file():
            raise RunnerError("build overlap lease file is missing or symlinked")
        with self.lease_path.open("w", encoding="utf-8") as stream:
            json.dump(self._lease_record(), stream, indent=2, sort_keys=True)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())

    def prepare_phase(
        self,
        *,
        phase_id: str,
        cwd: Path,
        argv: Sequence[str],
        target: Path,
    ) -> None:
        """Publish an exact phase before its Cargo child becomes visible."""

        if not phase_id:
            raise RunnerError("build overlap phase id is empty")
        if _cargo_command_class(argv) != "cargo-build":
            raise RunnerError("overlap phase is not an exact cargo build")
        self.ensure_valid()
        if self.bound_pid is not None or self.bound_phase is not None:
            raise RunnerError("build overlap phase is already bound")
        self.bound_phase = phase_id
        self.phase_argv = list(argv)
        self.phase_cwd = str(_resolve_path(cwd))
        self.phase_target_dir = str(_resolve_path(target))
        try:
            self._with_lock(self._write_current_record)
        except BaseException:
            self.bound_phase = None
            self.phase_argv = None
            self.phase_cwd = None
            self.phase_target_dir = None
            raise

    def bind_child(
        self,
        process: subprocess.Popen[Any],
        *,
        phase_id: str,
        cwd: Path,
        argv: Sequence[str],
        target: Path,
    ) -> None:
        self.ensure_valid()
        entry = _read_proc(process.pid)
        if entry is None or not isinstance(entry.get("starttime"), str):
            raise RunnerError("Cargo child disappeared before overlap binding")
        if _process_command_class(entry) != "cargo-build":
            raise RunnerError("overlap child is not an exact cargo build process")
        if (
            self.bound_phase != phase_id
            or self.phase_argv != list(argv)
            or self.phase_cwd != str(_resolve_path(cwd))
            or self.phase_target_dir != str(_resolve_path(target))
        ):
            raise RunnerError("overlap child does not match its prepared phase")
        if entry.get("cwd") != self.phase_cwd or entry.get("cmdline") != list(argv):
            raise RunnerError("overlap child process identity differs from its prepared phase")
        self.bound_pid = process.pid
        self.bound_starttime = str(entry["starttime"])
        owned = {
            "root_pid": process.pid,
            "root_starttime": self.bound_starttime,
            "command_class": "cargo-build",
            "phase": "build",
            "phase_id": phase_id,
            "cwd": str(cwd),
            "argv": list(argv),
            "target_dir": str(target),
        }
        try:
            self._with_lock(self._write_current_record)
        except BaseException:
            self.bound_pid = None
            self.bound_starttime = None
            raise
        self.owned_jobs.append(owned)

    def unbind_child(self) -> None:
        if self.bound_pid is None and self.bound_phase is None:
            return
        old = (
            self.bound_pid,
            self.bound_starttime,
            self.bound_phase,
            self.phase_argv,
            self.phase_cwd,
            self.phase_target_dir,
        )
        self.bound_pid = None
        self.bound_starttime = None
        self.bound_phase = None
        self.phase_argv = None
        self.phase_cwd = None
        self.phase_target_dir = None
        try:
            self._with_lock(self._write_current_record)
        except BaseException:
            (
                self.bound_pid,
                self.bound_starttime,
                self.bound_phase,
                self.phase_argv,
                self.phase_cwd,
                self.phase_target_dir,
            ) = old
            raise

    def ensure_valid(self) -> None:
        if self.lease_path is None or self.lease_id is None:
            raise RunnerError("build overlap lease is not acquired")
        if self._clock() >= self.expires:
            raise RunnerError("build overlap permit expired")
        current_hash, _ = _sha256_file(self.permit_path)
        if current_hash != self.permit_hash:
            raise RunnerError("build overlap permit changed while active")
        snapshot = _proc_snapshot()
        expected_pid = self.bound_pid if self.bound_pid is not None else self.root_pid
        expected_starttime = self.bound_starttime if self.bound_pid is not None else self.root_starttime
        if snapshot.get(expected_pid, {}).get("starttime") != expected_starttime:
            raise RunnerError("build overlap root PID was reused")

        def check() -> None:
            active = self._active_records_locked(snapshot)
            if len(active) > self.max_jobs:
                raise RunnerError("build permit active-job count exceeds max_jobs")
            if not any(path == self.lease_path for path, _, _, _ in active):
                raise RunnerError("build overlap lease is missing or no longer active")

        self._with_lock(check)

    @staticmethod
    def _peer_summary(entry: Mapping[str, Any], value: Mapping[str, Any]) -> dict[str, Any]:
        return {
            "lease_id": value.get("lease_id"),
            "job_group": value.get("job_group"),
            "root_pid": value.get("root_pid"),
            "root_starttime": value.get("root_starttime"),
            "controller_pid": value.get("controller_pid"),
            "controller_starttime": value.get("controller_starttime"),
            "lane": value.get("lane"),
            "manifest_path": value.get("manifest_path"),
            "manifest_sha256": value.get("manifest_sha256"),
            "command_class": value.get("command_class"),
            "phase": value.get("phase"),
            "phase_id": value.get("phase_id"),
            "source_roots": value.get("source_roots"),
            "target_dirs": value.get("target_dirs"),
            "cwd": entry.get("cwd"),
            "exe": entry.get("exe"),
            "cmdline": entry.get("cmdline"),
        }

    @staticmethod
    def _pending_peer_processes(
        snapshot: Mapping[int, Mapping[str, Any]],
        controller_pid: int,
        value: Mapping[str, Any],
        lease_id: str,
    ) -> set[int]:
        """Allow only a direct child matching the published pending phase."""

        phase_id = value.get("phase_id")
        if phase_id is None:
            return set()
        phase_argv = value.get("phase_argv")
        phase_cwd = value.get("phase_cwd")
        if not isinstance(phase_argv, list) or not isinstance(phase_cwd, str):
            raise RunnerError(f"peer {lease_id} phase binding is malformed")
        children = [
            pid
            for pid, entry in snapshot.items()
            if int(entry.get("ppid", -1)) == controller_pid
            and _process_command_class(entry) == "cargo-build"
        ]
        if len(children) > 1:
            raise RunnerError(f"peer {lease_id} has multiple Cargo phase children")
        if not children:
            return set()
        child_pid = children[0]
        child = snapshot[child_pid]
        if child.get("cmdline") != phase_argv or child.get("cwd") != phase_cwd:
            raise RunnerError(f"peer {lease_id} Cargo child differs from its published phase")
        return ProcessGuard._check_child_tree(
            snapshot,
            child_pid,
            expected_command_class="cargo-build",
            label=f"peer {lease_id}",
        )

    def peer_snapshots(self, snapshot: Mapping[int, Mapping[str, Any]]) -> list[dict[str, Any]]:
        self.ensure_valid()
        summaries: list[dict[str, Any]] = []

        def read() -> None:
            for path, value, pid, lease_id in self._active_records_locked(snapshot):
                if lease_id == self.lease_id:
                    continue
                # A peer may have reserved a slot while its runner is still
                # between lease acquisition and spawning Cargo.  This record
                # is intentionally counted for capacity, but its Python
                # controller is never a permitted process.  Only the exact
                # Cargo child record written by bind_child becomes a peer
                # snapshot/allow-list entry.
                if value.get("command_class") == "controller-reservation":
                    continue
                entry = snapshot.get(pid)
                if entry is None:
                    # The lease remains a counted, non-allowlisted
                    # transitional reservation until its live controller
                    # rewrites it.  A dead controller was already rejected
                    # by _validate_record.
                    continue
                summaries.append(self._peer_summary(entry, value))

        self._with_lock(read)
        summaries.sort(key=lambda item: str(item.get("lease_id")))
        return summaries

    def allowed_peer_processes(self, snapshot: Mapping[int, Mapping[str, Any]]) -> set[int]:
        self.ensure_valid()
        peers: set[int] = set()

        def read() -> None:
            for path, value, pid, lease_id in self._active_records_locked(snapshot):
                if lease_id == self.lease_id:
                    continue
                if value.get("command_class") == "controller-reservation":
                    peers.update(self._pending_peer_processes(snapshot, pid, value, lease_id))
                    continue
                entry = snapshot.get(pid)
                if entry is None:
                    continue
                expected = str(value["command_class"])
                peers.update(
                    ProcessGuard._check_child_tree(
                        snapshot,
                        pid,
                        expected_command_class=expected,
                        label=f"peer {lease_id}",
                    )
                )

        self._with_lock(read)
        return peers

    def event_record(self) -> dict[str, Any]:
        record: dict[str, Any] = {
            "class": "compile_only",
            "permit_path": str(self.permit_path),
            "permit_sha256": self.permit_hash,
            "permit_bytes": self.permit_bytes,
            "grant_id": self.grant_id,
            "job_group": self.job_group,
            "max_jobs": self.max_jobs,
            "expires_utc": self.permit["expires_utc"],
            "lease_path": str(self.lease_path) if self.lease_path is not None else None,
            "lease_id": self.lease_id,
            "root_pid": self.root_pid,
            "root_starttime": self.root_starttime,
            "owned_jobs": list(self.owned_jobs),
            "peer_jobs_start": self.started_peers,
            "peer_jobs_end": self.ended_peers,
        }
        if self.lease_path is not None and self.lease_path.is_file() and not self.lease_path.is_symlink():
            record["lease_sha256"], record["lease_bytes"] = _sha256_file(self.lease_path)
            try:
                lease_value = json.loads(self.lease_path.read_text(encoding="utf-8"))
            except (OSError, UnicodeDecodeError, json.JSONDecodeError):
                lease_value = None
            if isinstance(lease_value, dict):
                record["lease_status"] = lease_value.get("status")
        return record

    def release(self) -> None:
        if self.bound_pid is not None or self.bound_phase is not None:
            try:
                self.unbind_child()
            except (OSError, RunnerError):
                return
        if self.lease_path is None or not self.lease_path.is_file() or self.lease_path.is_symlink():
            return

        def write_release() -> None:
            try:
                value = json.loads(self.lease_path.read_text(encoding="utf-8"))
            except (OSError, UnicodeDecodeError, json.JSONDecodeError):
                return
            if not isinstance(value, dict):
                return
            value["status"] = "released"
            value["released_utc"] = _now()
            with self.lease_path.open("w", encoding="utf-8") as stream:
                json.dump(value, stream, indent=2, sort_keys=True)
                stream.write("\n")
                stream.flush()
                os.fsync(stream.fileno())

        try:
            self._with_lock(write_release)
        except (OSError, RunnerError):
            # The failed build receipt already records the child/guard error;
            # leaving an active external lease is safer than hiding that error.
            return


def _json_hash(path: Path) -> tuple[str, int]:
    # ``Path.exists()`` is false for a dangling symlink.  Check the link first
    # so a replaced log cannot be silently recorded as an empty stream.
    if path.is_symlink():
        raise RunnerError(f"refusing to follow symlinked log: {path}")
    return _sha256_file(path) if path.exists() else (_sha256_bytes(b""), 0)


class SafeRunner:
    def __init__(
        self,
        manifest: Mapping[str, Any],
        manifest_path: Path,
        phase: str,
        cpu: int | None,
        process_observer: Any | None = None,
        clock: Callable[[], _dt.datetime] | None = None,
    ):
        self.manifest = manifest
        self.manifest_path = manifest_path.resolve()
        self.phase = phase
        self.cpu = cpu
        # The default remains deliberately inert for the established W4b/W4/W5
        # lanes.  The DrawingML paired runner supplies a narrow observer only
        # for its measurement children; it never has to intercept helper
        # subprocesses used by this runner's git/config checks.
        self.process_observer = process_observer
        self.clock = clock
        self.guard = ProcessGuard()
        self.build_execution = manifest.get("build_execution")
        if self.build_execution is None:
            self.build_execution = {"class": "exclusive"}
        if not isinstance(self.build_execution, dict):
            raise RunnerError("build_execution must be an object")
        if self.build_execution.get("class") not in BUILD_EXECUTION_CLASSES:
            raise RunnerError("build_execution.class must be exclusive or compile_only")
        if self.build_execution.get("class") == "compile_only":
            _validate_build_execution(
                manifest,
                output=_resolve_path(str(manifest["output_dir"])),
                source_roots=[
                    _resolve_path(str(spec["source_root"]))
                    for spec in manifest.get("legs", {}).values()
                    if isinstance(spec, Mapping) and "source_root" in spec
                ],
            )
        self.overlap_lease: BuildOverlapLease | None = None
        self.target_preflight_limit = TARGET_PREFLIGHT_BYTES
        self.target_limit = TARGET_HARD_LIMIT_BYTES
        self.memory_limit = MEMORY_HARD_LIMIT_BYTES
        self.output = _resolve_path(str(manifest["output_dir"]))
        if self.output.is_symlink():
            raise RunnerError(f"output_dir may not be a symlink: {self.output}")
        if self.output.exists() and not self.output.is_dir():
            raise RunnerError(f"output_dir is not a directory: {self.output}")
        if self.phase == "build" and self.output.exists():
            try:
                existing = list(self.output.iterdir())
            except OSError as error:
                raise RunnerError(f"cannot inspect build output_dir: {self.output}: {error}") from error
            if existing:
                raise RunnerError(
                    f"build output_dir must be a fresh empty directory; choose a new path: {self.output}"
                )
        self.output.mkdir(parents=True, exist_ok=True)
        self.command_dir = self.output / "commands"
        if self.command_dir.is_symlink():
            raise RunnerError(f"command log directory may not be a symlink: {self.command_dir}")
        self.command_dir.mkdir(parents=True, exist_ok=True)
        self.events_path = self.output / f"events-{phase}.jsonl"
        self.receipt_path = self.output / f"{phase}-receipt.json"
        if self.events_path.is_symlink() or self.receipt_path.is_symlink() or self.events_path.exists() or self.receipt_path.exists():
            raise RunnerError(
                f"output already contains {phase} receipt/journal; choose a fresh output_dir: {self.output}"
            )
        self.shared = manifest.get("shared", {})
        if manifest.get("lane") == "w4-xlsx-raw":
            _validate_xlsx_preparation_manifest(manifest)
        self.preparation: dict[str, Any] | None = None
        self.preparation_path: Path | None = None
        if str(manifest.get("lane", "")).startswith("legacy-"):
            self.preparation_path = _resolve_path(str(manifest["preparation_receipt"]))
            self.preparation = self._load_preparation_receipt()
        self.recorded_events: list[dict[str, Any]] = []

    def _event_execution_class(self, event_kind: str) -> str:
        if event_kind == "build" and self.overlap_lease is not None:
            return "compile_only"
        return "exclusive"

    def _build_execution_record(self) -> dict[str, Any]:
        if self.overlap_lease is not None:
            return self.overlap_lease.event_record()
        if self.build_execution.get("class") == "compile_only":
            permit_value = self.build_execution.get("permit_path")
            record: dict[str, Any] = {"class": "compile_only", "status": "not_acquired"}
            if isinstance(permit_value, str) and permit_value:
                try:
                    permit_path = _resolve_path(permit_value)
                    record["permit_path"] = str(permit_path)
                    if permit_path.is_file() and not permit_path.is_symlink():
                        record["permit_sha256"], record["permit_bytes"] = _sha256_file(permit_path)
                except RunnerError as error:
                    record["permit_error"] = str(error)
            return record
        return {"class": "exclusive"}

    def _capture_overlap_end(self) -> None:
        if self.overlap_lease is None:
            return
        self.overlap_lease.ended_peers = self.overlap_lease.peer_snapshots(_proc_snapshot())

    def _acquire_overlap_lease(self) -> None:
        if self.build_execution.get("class") != "compile_only":
            return
        self.overlap_lease = BuildOverlapLease(
            self.manifest, self.manifest_path, clock=self.clock
        )
        self.overlap_lease.acquire()

    def _load_preparation_receipt(self) -> dict[str, Any]:
        if self.preparation_path is None:
            raise RunnerError("legacy lane has no preparation receipt path")
        path = self.preparation_path
        if path.is_symlink() or not path.is_file():
            raise RunnerError(f"legacy preparation receipt is missing or symlinked: {path}")
        try:
            receipt = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            raise RunnerError(f"cannot read legacy preparation receipt: {error}") from error
        if not isinstance(receipt, dict) or receipt.get("schema") != "litchi.legacy-native-preparation.v1":
            raise RunnerError("legacy preparation receipt schema mismatch")
        if receipt.get("lane") != self.manifest.get("lane"):
            raise RunnerError("legacy preparation receipt lane mismatch")
        manifest_record = receipt.get("manifest")
        if not isinstance(manifest_record, dict) or manifest_record.get("path") != str(self.manifest_path):
            raise RunnerError("legacy preparation receipt is bound to a different manifest")
        manifest_hash, manifest_bytes = _sha256_file(self.manifest_path)
        if manifest_record.get("sha256") != manifest_hash or manifest_record.get("bytes", manifest_bytes) != manifest_bytes:
            raise RunnerError("legacy preparation receipt manifest identity changed")
        preparer = receipt.get("preparer")
        preparer_path = Path(__file__).resolve().parent / "native_qualification" / "prepare_legacy_native_inputs.py"
        if not isinstance(preparer, dict) or preparer.get("path") != str(preparer_path):
            raise RunnerError("legacy preparation receipt was produced by a different preparer")
        preparer_hash, _ = _sha256_file(preparer_path)
        if preparer.get("sha256") != preparer_hash:
            raise RunnerError("legacy preparation preparer identity changed")
        identities = receipt.get("identities")
        if not isinstance(identities, dict) or set(identities) != {"before", "after"}:
            raise RunnerError("legacy preparation receipt identities are incomplete")
        reviewed = receipt.get("reviewed_inputs")
        if not isinstance(reviewed, dict):
            raise RunnerError("legacy preparation reviewed inputs are missing")
        for leg in ("before", "after"):
            allowed = self.manifest["legs"][leg].get("allowed_dirty_paths")
            if reviewed.get("allowed_dirty_paths") != allowed:
                raise RunnerError(f"legacy preparation dirty-path allow-list differs for {leg}")
        return receipt

    def _validate_preparation_identity(self, leg: str, identity: Mapping[str, Any]) -> None:
        if self.preparation is None:
            return
        expected = self.preparation.get("identities", {}).get(leg)
        if not isinstance(expected, dict):
            raise RunnerError(f"legacy preparation identity is missing for {leg}")
        for key in (
            "source_root",
            "head",
            "tree",
            "source_tree_sha256",
            "clean",
            "status_paths",
            "status_sha256",
            "diff_sha256",
            "identity_sha256",
            "lock_paths",
            "probe_paths",
            "fixture_paths",
        ):
            if identity.get(key) != expected.get(key):
                raise RunnerError(f"legacy preparation identity differs for {leg}: {key}")
        expected_files = {
            str(item.get("path")): (item.get("bytes"), item.get("sha256"))
            for item in expected.get("identity_files", [])
            if isinstance(item, dict)
        }
        actual_files = {
            str(item.get("path")): (item.get("bytes"), item.get("sha256"))
            for item in identity.get("identity_files", [])
            if isinstance(item, dict)
        }
        if actual_files != expected_files:
            raise RunnerError(f"legacy preparation identity files differ for {leg}")

    def _context(
        self,
        leg: str,
        index: int,
        *,
        output: Path | None = None,
        block: int = 1,
    ) -> dict[str, Any]:
        spec = self.manifest["legs"][leg]
        source = _resolve_path(str(spec["source_root"]))
        target = _resolve_path(str(spec["target_dir"]))
        fixture_values = [*self.shared.get("fixture_paths", []), *spec.get("fixture_paths", [])]
        fixture = ""
        if len(fixture_values) == 1:
            fixture = str(_resolve_path(str(fixture_values[0]), base=source))
        context: dict[str, Any] = {
            "source_root": source,
            "target_dir": target,
            "output_dir": self.output,
            "manifest_path": self.manifest_path,
            "leg": leg,
            "index": index,
            "rounds": int(self.manifest.get("rounds", 31)),
            "block": block,
            "fixture": fixture,
        }
        if output is not None:
            context["output"] = output
        return context

    def _environment(self, spec: Mapping[str, Any], context: Mapping[str, Any]) -> dict[str, str]:
        env = os.environ.copy()
        # Do not inherit Cargo's external configuration channels.  The runner
        # owns CARGO_TARGET_DIR and validates the workspace-local config before
        # starting a build.
        env.pop("CARGO_HOME", None)
        env.pop("CARGO_CONFIG", None)
        # These session switches are never part of a qualification input.
        # A caller that needs one must record it in the reviewed manifest
        # rather than inheriting a process-wide blessing/cap override.
        env.pop("LITCHI_CAPPED", None)
        env.pop("LITCHI_BLESS", None)
        path_prepend = [_render(str(value), context) for value in self.manifest.get("path_prepend", [])]
        if path_prepend:
            for path in path_prepend:
                resolved = _resolve_path(path)
                if not resolved.is_dir():
                    raise RunnerError(f"path_prepend entry is not a directory: {resolved}")
            env["PATH"] = os.pathsep.join([*path_prepend, env.get("PATH", "")])
        custom = spec.get("environment", self.manifest.get("environment", {}))
        if not isinstance(custom, dict):
            raise RunnerError("environment must be an object")
        for key, value in custom.items():
            if key in {"CARGO_HOME", "CARGO_CONFIG"}:
                raise RunnerError(f"{key} is forbidden because it can inject Cargo path configuration")
            env[str(key)] = _render(str(value), context)
        target = str(context["target_dir"])
        if env.get("CARGO_TARGET_DIR", target) != target:
            raise RunnerError("manifest cannot override the explicit CARGO_TARGET_DIR")
        env["CARGO_TARGET_DIR"] = target
        env["CARGO_BUILD_JOBS"] = "1"
        env["CARGO_INCREMENTAL"] = "0"
        # The shim's documented value is an explicit 20G cap.  Do not derive
        # a binary-GiB suffix from the decimal byte ceiling (20e9 would round
        # down to a misleading 19G).
        env["LITCHI_MEM_MAX"] = "20G"
        env.pop("LITCHI_PROBE_ONLY", None)
        return env

    @staticmethod
    def _recorded_environment(env: Mapping[str, str]) -> dict[str, str]:
        keys = (
            "PATH",
            "CARGO_TARGET_DIR",
            "CARGO_BUILD_JOBS",
            "CARGO_INCREMENTAL",
            "CARGO_PROFILE_RELEASE_DEBUG",
            "LITCHI_MEM_MAX",
            "LITCHI_PROBE_BATCH_ROUNDS",
            "TMPDIR",
        )
        return {key: env[key] for key in keys if key in env}

    def _identity(self, leg: str) -> dict[str, Any]:
        identity = _source_identity(self.manifest["legs"][leg], self.shared)
        identity["target_dir"] = str(_resolve_path(str(self.manifest["legs"][leg]["target_dir"])))
        identity["output_dir"] = str(self.output)
        return identity

    def _build_artifact_paths(
        self, leg: str, command: Mapping[str, Any], context: Mapping[str, Any]
    ) -> list[Path]:
        target = _resolve_path(str(self.manifest["legs"][leg]["target_dir"]))
        paths: list[Path] = []
        for value in command.get("artifacts", []):
            path = _resolve_path(_render(str(value), context))
            if not _under(path, target) or path == target:
                raise RunnerError(f"build artifact escapes its leg target: {path}")
            if path in paths:
                raise RunnerError(f"build command declares an artifact more than once: {path}")
            paths.append(path)
        return paths

    def _check_identity(self, leg: str, expected: Mapping[str, Any]) -> dict[str, Any]:
        actual = self._identity(leg)
        for key in ("head", "tree", "source_tree_sha256", "clean", "status_sha256", "diff_sha256", "identity_sha256"):
            if actual.get(key) != expected.get(key):
                raise RunnerError(f"source identity changed for {leg}: {key}")
        return actual

    def _check_target(self, target: Path, *, preflight: bool = False) -> int:
        if target.is_symlink():
            raise RunnerError(f"target_dir may not be a symlink: {target}")
        if target.exists() and not target.is_dir():
            raise RunnerError(f"target_dir is not a directory: {target}")
        size = _directory_size(target)
        limit = self.target_preflight_limit if preflight else self.target_limit
        if size > limit:
            label = "preflight" if preflight else "hard"
            raise RunnerError(f"target exceeds {label} byte limit: {target} ({size} > {limit})")
        return size

    def _require_fresh_target(self, target: Path) -> None:
        if target.is_symlink():
            raise RunnerError(f"fresh target may not be a symlink: {target}")
        if not target.exists():
            return
        if not target.is_dir():
            raise RunnerError(f"fresh target is not a directory: {target}")
        try:
            next(target.iterdir())
        except StopIteration:
            return
        raise RunnerError(f"fresh target is not empty; refusing stale artifacts: {target}")

    def _append_event(self, event: Mapping[str, Any]) -> None:
        if self.events_path.is_symlink():
            raise RunnerError(f"event journal may not be a symlink: {self.events_path}")
        mode = "x" if not self.events_path.exists() else "a"
        with self.events_path.open(mode, encoding="utf-8") as stream:
            stream.write(json.dumps(event, sort_keys=True) + "\n")
            stream.flush()
            os.fsync(stream.fileno())

    @staticmethod
    def _terminate(process: subprocess.Popen[Any]) -> None:
        try:
            os.killpg(process.pid, signal.SIGTERM)
        except (ProcessLookupError, PermissionError):
            process.terminate()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except (ProcessLookupError, PermissionError):
                process.kill()
            process.wait()

    @staticmethod
    def _record_owned_processes(
        snapshot: Mapping[int, Mapping[str, Any]],
        *,
        process_pid: int,
        runner_pid: int,
        preexisting: Mapping[int, str],
        owned: dict[int, dict[str, Any]],
        pid_reuse: list[dict[str, Any]],
    ) -> None:
        """Record descendants before and after a leader is reparented.

        While the direct child is alive, its normal descendant closure is the
        authoritative ownership boundary.  Once it exits, a Linux subreaper
        receives the orphaned descendants as children of ``runner_pid``; the
        second closure is therefore needed to retain ownership across the
        reparenting event.  Existing runner children are excluded by their
        PID/start-time snapshot, and a changed start time is treated as PID
        reuse rather than a process we may signal.
        """

        candidates: set[int] = set()
        rejected_descendants: set[int] = set()
        leader_record = owned.get(process_pid)
        leader_entry = snapshot.get(process_pid)
        if leader_record is None:
            # Without the leader's exact spawn identity, a child seen under
            # the runner could belong to a replacement process.  Refuse to
            # admit any part of that unproven subtree rather than treating an
            # empty/partial initial scan as ownership evidence.
            raise RunnerError(f"cannot establish owned leader identity for pid {process_pid}")
        # Never walk a PID whose exact start time is not the one captured at
        # spawn.  If the leader PID has been reused, its current descendants
        # belong to the replacement process and are foreign to this command.
        if leader_record is not None and leader_entry is not None:
            if _same_process_identity(leader_entry, leader_record):
                candidates.update(_descendants(snapshot, process_pid))
            else:
                reuse = {
                    "pid": process_pid,
                    "kind": "owned_pid_reused",
                    "expected_starttime": leader_record.get("starttime"),
                    "observed_starttime": str(leader_entry.get("starttime")),
                }
                if reuse not in pid_reuse:
                    pid_reuse.append(reuse)
                rejected_descendants.update(_descendants(snapshot, process_pid))
        for child_pid in _descendants(snapshot, runner_pid):
            if child_pid == runner_pid:
                continue
            entry = snapshot.get(child_pid)
            if entry is None:
                continue
            old_starttime = preexisting.get(child_pid)
            if old_starttime is not None:
                if child_pid in owned and child_pid != process_pid:
                    # A pre-existing identity must never remain in the owned
                    # set after it is discovered through a leader-tree race;
                    # otherwise cleanup could signal that unrelated child.
                    owned.pop(child_pid, None)
                    reuse = {
                        "pid": child_pid,
                        "kind": "preexisting_identity_admitted",
                        "expected_starttime": old_starttime,
                        "observed_starttime": str(entry.get("starttime")),
                    }
                    if reuse not in pid_reuse:
                        pid_reuse.append(reuse)
                if old_starttime != str(entry.get("starttime")):
                    reuse = {
                        "pid": child_pid,
                        "kind": "preexisting_pid_reused",
                        "expected_starttime": old_starttime,
                        "observed_starttime": str(entry.get("starttime")),
                    }
                    if reuse not in pid_reuse:
                        pid_reuse.append(reuse)
                # A pre-existing runner child is not admitted even when it
                # appears in the leader's tree; reject its entire current
                # subtree before processing candidates below.
                rejected_descendants.update(_descendants(snapshot, child_pid))
                continue
            if child_pid == process_pid:
                # The leader is admitted only through its exact identity
                # check above.  Do not re-admit a reused or unobserved PID via
                # the runner's reparented-child closure.
                continue
            existing = owned.get(child_pid)
            if existing is not None and not _same_process_identity(entry, existing):
                reuse = {
                    "pid": child_pid,
                    "kind": "owned_pid_reused",
                    "expected_starttime": existing.get("starttime"),
                    "observed_starttime": str(entry.get("starttime")),
                }
                if reuse not in pid_reuse:
                    pid_reuse.append(reuse)
                rejected_descendants.update(_descendants(snapshot, child_pid))
                continue
            candidates.add(child_pid)

        # The leader tree may have contained a reused/pre-existing PID before
        # the runner examined the reparented closure.  Remove those identities
        # and every descendant of them before admitting anything new.
        candidates.difference_update(rejected_descendants)
        for child_pid in sorted(candidates):
            if child_pid == runner_pid:
                continue
            entry = snapshot.get(child_pid)
            if entry is None:
                continue
            existing = owned.get(child_pid)
            if existing is not None:
                continue
            owned[child_pid] = _process_record(entry)

    @staticmethod
    def _reap_owned_processes(
        owned: Mapping[int, Mapping[str, Any]],
        *,
        leader_pid: int,
        snapshot: Mapping[int, Mapping[str, Any]] | None = None,
    ) -> None:
        """Reap only descendants that this runner has already identified."""

        snapshot = _proc_snapshot() if snapshot is None else snapshot
        for child_pid in sorted(owned):
            if child_pid == leader_pid:
                continue
            entry = snapshot.get(child_pid)
            if entry is None or not _same_process_identity(entry, owned[child_pid]):
                continue
            try:
                os.waitpid(child_pid, os.WNOHANG)
            except (ChildProcessError, OSError):
                # A descendant may have installed its own subreaper or may
                # already have been reaped.  In either case it is not ours to
                # wait for; liveness remains checked through /proc below.
                continue

    @staticmethod
    def _signal_owned_processes(
        owned: Mapping[int, Mapping[str, Any]],
        *,
        leader_pid: int,
        sig: signal.Signals,
    ) -> None:
        """Signal exact, start-time-matched owned PIDs; never a reused PID."""

        snapshot = _proc_snapshot()
        for child_pid in sorted(owned, reverse=True):
            if child_pid == leader_pid:
                continue
            entry = snapshot.get(child_pid)
            record = owned[child_pid]
            if entry is None or not _same_process_identity(entry, record):
                continue
            try:
                os.kill(child_pid, sig)
            except (ProcessLookupError, PermissionError):
                continue

    @staticmethod
    def _live_owned_processes(
        owned: Mapping[int, Mapping[str, Any]],
        *,
        leader_pid: int,
        snapshot: Mapping[int, Mapping[str, Any]],
    ) -> list[dict[str, Any]]:
        live: list[dict[str, Any]] = []
        for child_pid in sorted(owned):
            if child_pid == leader_pid:
                continue
            entry = snapshot.get(child_pid)
            if entry is not None and _same_process_identity(entry, owned[child_pid]):
                live.append(dict(owned[child_pid]))
        return live

    def _cleanup_owned_descendants(
        self,
        *,
        leader_pid: int,
        runner_pid: int,
        preexisting: Mapping[int, str],
        owned: dict[int, dict[str, Any]],
        pid_reuse: list[dict[str, Any]],
    ) -> dict[str, Any]:
        """Find, kill, and reap descendants that escaped the direct child.

        The returned ``escaped`` list is deliberately retained even when the
        children are successfully killed: an escaped process is a lifecycle
        violation and must fail the event rather than being silently hidden by
        cleanup.  ``survivors`` contains only identities still live after the
        bounded TERM/KILL cleanup window.
        """

        snapshot = _proc_snapshot()
        self._record_owned_processes(
            snapshot,
            process_pid=leader_pid,
            runner_pid=runner_pid,
            preexisting=preexisting,
            owned=owned,
            pid_reuse=pid_reuse,
        )
        escaped_by_pid: dict[int, dict[str, Any]] = {}

        def remember_escaped(current: Mapping[int, Mapping[str, Any]]) -> None:
            for pid, record in sorted(owned.items()):
                if (
                    pid != leader_pid
                    and current.get(pid) is not None
                    and _same_process_identity(current[pid], record)
                ):
                    escaped_by_pid.setdefault(pid, dict(record))

        remember_escaped(snapshot)
        if not escaped_by_pid:
            self._reap_owned_processes(owned, leader_pid=leader_pid, snapshot=snapshot)
            return {"escaped": [], "survivors": [], "pid_reuse": list(pid_reuse)}

        self._signal_owned_processes(owned, leader_pid=leader_pid, sig=signal.SIGTERM)
        deadline = time.monotonic() + 5.0
        survivors: list[dict[str, Any]] = list(escaped_by_pid.values())
        while time.monotonic() < deadline:
            snapshot = _proc_snapshot()
            self._record_owned_processes(
                snapshot,
                process_pid=leader_pid,
                runner_pid=runner_pid,
                preexisting=preexisting,
                owned=owned,
                pid_reuse=pid_reuse,
            )
            remember_escaped(snapshot)
            self._signal_owned_processes(owned, leader_pid=leader_pid, sig=signal.SIGTERM)
            self._reap_owned_processes(owned, leader_pid=leader_pid, snapshot=snapshot)
            survivors = self._live_owned_processes(
                owned,
                leader_pid=leader_pid,
                snapshot=snapshot,
            )
            if not survivors:
                return {
                    "escaped": list(escaped_by_pid.values()),
                    "survivors": [],
                    "pid_reuse": list(pid_reuse),
                }
            time.sleep(0.05)

        self._signal_owned_processes(owned, leader_pid=leader_pid, sig=signal.SIGKILL)
        deadline = time.monotonic() + 1.0
        while time.monotonic() < deadline:
            snapshot = _proc_snapshot()
            self._record_owned_processes(
                snapshot,
                process_pid=leader_pid,
                runner_pid=runner_pid,
                preexisting=preexisting,
                owned=owned,
                pid_reuse=pid_reuse,
            )
            remember_escaped(snapshot)
            # Re-capture before every final signal: a surviving descendant
            # may fork/reparent a new child after the first KILL snapshot.
            self._signal_owned_processes(owned, leader_pid=leader_pid, sig=signal.SIGKILL)
            self._reap_owned_processes(owned, leader_pid=leader_pid, snapshot=snapshot)
            survivors = self._live_owned_processes(
                owned,
                leader_pid=leader_pid,
                snapshot=snapshot,
            )
            if not survivors:
                break
            time.sleep(0.05)
        return {
            "escaped": list(escaped_by_pid.values()),
            "survivors": survivors,
            "pid_reuse": list(pid_reuse),
        }

    def _stabilize_target(self, target: Path) -> int:
        previous: int | None = None
        stable = 0
        for _ in range(10):
            size = self._check_target(target)
            if size == previous:
                stable += 1
                if stable >= 2:
                    return size
            else:
                stable = 0
            previous = size
            time.sleep(0.2)
        raise RunnerError(f"target size did not stabilize after child exit: {target}")

    def _prepare_log_path(self, path: Path) -> Path:
        path = _resolve_path(path)
        if not _under(path, self.output):
            raise RunnerError(f"log path escapes output_dir: {path}")
        if path.is_symlink() or path.exists():
            raise RunnerError(f"refusing to overwrite or follow log path: {path}")
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.parent.is_symlink():
            raise RunnerError(f"log parent may not be a symlink: {path.parent}")
        return path

    def _run_command(
        self,
        *,
        event_id: str,
        leg: str,
        argv: Sequence[str],
        cwd: Path,
        env: Mapping[str, str],
        identity: Mapping[str, Any],
        expected_outputs: Sequence[Path] = (),
        require_stdout: bool = False,
        stdout_path: Path | None = None,
        stderr_path: Path | None = None,
        preflight_target: bool = True,
        measured: bool = True,
        event_kind: str = "measurement",
    ) -> dict[str, Any]:
        overlap = self.overlap_lease if event_kind == "build" else None
        if overlap is not None:
            self.guard.require_compile_quiet(overlap)
        else:
            self.guard.require_quiet()
        target = _resolve_path(str(self.manifest["legs"][leg]["target_dir"]))
        target_before = self._check_target(target, preflight=preflight_target)
        stdout_path = self._prepare_log_path(stdout_path or (self.command_dir / f"{event_id}.stdout"))
        stderr_path = self._prepare_log_path(stderr_path or (self.command_dir / f"{event_id}.stderr"))
        start = time.monotonic_ns()
        started_at = _now()
        violation: list[dict[str, Any]] = []
        returncode: int | None = None
        process: subprocess.Popen[Any] | None = None
        process_observation: Any | None = None
        runner_error: str | None = None
        last_target_check = 0.0
        overlap_start: list[dict[str, Any]] = []
        overlap_end: list[dict[str, Any]] = []
        owned_job: dict[str, Any] | None = None
        runner_pid = os.getpid()
        _enable_child_subreaper()
        preexisting_snapshot = _proc_snapshot()
        preexisting = {
            pid: str(entry.get("starttime"))
            for pid in _descendants(preexisting_snapshot, runner_pid)
            if pid != runner_pid
            and (entry := preexisting_snapshot.get(pid)) is not None
        }
        owned: dict[int, dict[str, Any]] = {}
        pid_reuse: list[dict[str, Any]] = []
        lifecycle: dict[str, Any] = {"escaped": [], "survivors": [], "pid_reuse": []}
        if overlap is not None:
            overlap.prepare_phase(
                phase_id=event_id,
                cwd=cwd,
                argv=argv,
                target=target,
            )
            overlap_start = overlap.peer_snapshots(_proc_snapshot())
        with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
            try:
                process = subprocess.Popen(
                    list(argv),
                    cwd=str(cwd),
                    env=dict(env),
                    stdout=stdout,
                    stderr=stderr,
                    stdin=subprocess.DEVNULL,
                    start_new_session=True,
                    preexec_fn=self._preexec(),
                )
                process_entry = _read_proc(process.pid)
                if process_entry is None:
                    runner_error = f"cannot establish leader process identity for pid {process.pid}"
                    if process.poll() is None:
                        self._terminate(process)
                    returncode = process.returncode
                else:
                    owned[process.pid] = _process_record(process_entry)
                if runner_error is None and overlap is not None:
                    try:
                        overlap.bind_child(
                            process,
                            phase_id=event_id,
                            cwd=cwd,
                            argv=argv,
                            target=target,
                        )
                        owned_job = overlap.owned_jobs[-1]
                    except BaseException as error:
                        runner_error = f"overlap child binding failed at {event_id}: {error}"
                        self._terminate(process)
                        returncode = process.returncode
                if runner_error is None and self.process_observer is not None:
                    try:
                        process_observation = self.process_observer.process_started(
                            event_id=event_id,
                            leg=leg,
                            argv=list(argv),
                            cwd=cwd,
                            identity=identity,
                            process=process,
                            event_kind=event_kind,
                            measured=measured,
                        )
                    except BaseException as error:
                        runner_error = f"process observer failed at {event_id}: {error}"
                        self._terminate(process)
                        returncode = process.returncode
                if runner_error is None:
                    while True:
                        returncode = process.poll()
                        if returncode is not None:
                            snapshot = _proc_snapshot()
                            self._record_owned_processes(
                                snapshot,
                                process_pid=process.pid,
                                runner_pid=runner_pid,
                                preexisting=preexisting,
                                owned=owned,
                                pid_reuse=pid_reuse,
                            )
                            break
                        try:
                            snapshot = _proc_snapshot()
                            self._record_owned_processes(
                                snapshot,
                                process_pid=process.pid,
                                runner_pid=runner_pid,
                                preexisting=preexisting,
                                owned=owned,
                                pid_reuse=pid_reuse,
                            )
                            if overlap is not None:
                                offenders = self.guard.offenders(
                                    process.pid,
                                    overlap=overlap,
                                    expected_command_class="cargo-build",
                                )
                            else:
                                offenders = self.guard.offenders(process.pid)
                            if offenders:
                                violation.extend(offenders)
                                runner_error = "process guard rejected an unrelated Litchi process"
                                self._terminate(process)
                                returncode = process.returncode
                                break
                            now = time.monotonic()
                            if now - last_target_check >= 1.0:
                                self._check_target(target)
                                last_target_check = now
                        except ProcMonitorError as error:
                            runner_error = str(error)
                            self._terminate(process)
                            returncode = process.returncode
                            break
                        except RunnerError as error:
                            # The child can exit between poll() above and
                            # the lease/guard snapshot.  In that narrow
                            # window the bound PID is naturally absent, not
                            # reused; let the normal unbind/post-exit checks
                            # validate the completed child.
                            if (
                                overlap is not None
                                and process.poll() is not None
                                and overlap.bound_pid == process.pid
                            ):
                                returncode = process.returncode
                                break
                            runner_error = str(error)
                            self._terminate(process)
                            returncode = process.returncode
                            break
                        time.sleep(0.1)
            except BaseException as error:
                if process is not None and process.poll() is None:
                    self._terminate(process)
                runner_error = str(error)
                returncode = process.returncode if process is not None else None
        if self.process_observer is not None and process_observation is not None:
            try:
                process_observation = self.process_observer.process_finished(
                    process_observation,
                    returncode=returncode,
                )
            except BaseException as error:
                runner_error = runner_error or f"process observer failed at {event_id}: {error}"
                # Do not let an observer implementation detail become an
                # unserialisable event payload.  The failed event remains
                # journaled without a false observation pointer and the
                # runner still fails closed below.
                process_observation = None
        if overlap is not None:
            try:
                overlap.unbind_child()
            except RunnerError as error:
                runner_error = runner_error or f"overlap child unbinding failed at {event_id}: {error}"
        if process is not None:
            try:
                lifecycle = self._cleanup_owned_descendants(
                    leader_pid=process.pid,
                    runner_pid=runner_pid,
                    preexisting=preexisting,
                    owned=owned,
                    pid_reuse=pid_reuse,
                )
            except BaseException as error:
                runner_error = runner_error or f"owned child cleanup failed at {event_id}: {error}"
                lifecycle = {
                    "escaped": [],
                    "survivors": [],
                    "pid_reuse": list(pid_reuse),
                }
        ended_at = _now()
        post_descendants = [
            int(record["pid"])
            for record in lifecycle.get("survivors", [])
            if isinstance(record, Mapping) and isinstance(record.get("pid"), int)
        ]
        try:
            if overlap is not None:
                overlap_end = overlap.peer_snapshots(_proc_snapshot())
                post_offenders = self.guard.offenders(overlap=overlap)
            else:
                post_offenders = self.guard.offenders()
        except RunnerError as error:
            post_offenders = []
            runner_error = runner_error or str(error)
        if post_descendants:
            runner_error = runner_error or f"child descendants survived exit: {post_descendants}"
        if lifecycle.get("escaped"):
            runner_error = runner_error or (
                "child escaped its owned process tree: "
                f"{[record.get('pid') for record in lifecycle['escaped']]}"
            )
        if lifecycle.get("pid_reuse"):
            runner_error = runner_error or (
                "owned child PID was reused before cleanup: "
                f"{lifecycle['pid_reuse']}"
            )
        if post_offenders:
            runner_error = runner_error or "unrelated Litchi process remained after child exit"
        try:
            target_after = self._stabilize_target(target)
        except RunnerError as error:
            target_after = -1
            runner_error = runner_error or str(error)
        if target_after > self.target_limit:
            runner_error = runner_error or (
                f"target exceeds byte limit: {target} ({target_after} > {self.target_limit})"
            )
        try:
            stdout_hash, stdout_bytes = _json_hash(stdout_path)
            stderr_hash, stderr_bytes = _json_hash(stderr_path)
        except RunnerError as error:
            runner_error = runner_error or str(error)
            stdout_hash, stdout_bytes = _sha256_bytes(b""), 0
            stderr_hash, stderr_bytes = _sha256_bytes(b""), 0
        output_records: list[dict[str, Any]] = []
        for output in expected_outputs:
            try:
                if output.is_symlink():
                    raise RunnerError(f"child output is a symlink: {output}")
                if not output.exists() or not output.is_file():
                    output_records.append({"path": str(output), "exists": False})
                    continue
                output_hash, output_bytes = _sha256_file(output)
                output_records.append(
                    {"path": str(output), "exists": True, "bytes": output_bytes, "sha256": output_hash}
                )
            except RunnerError as error:
                output_records.append({"path": str(output), "exists": False, "error": str(error)})
                runner_error = runner_error or str(error)
        try:
            identity_after = self._identity(leg)
        except RunnerError as error:
            identity_after = None
            runner_error = runner_error or str(error)
        if identity_after is not None:
            for key in ("head", "tree", "source_tree_sha256", "clean", "status_sha256", "diff_sha256", "identity_sha256"):
                if identity_after.get(key) != identity.get(key):
                    runner_error = runner_error or f"source identity changed for {leg}: {key}"
        event = {
            "schema": SCHEMA,
            "event_id": event_id,
            "event_kind": event_kind,
            "measured": measured,
            "leg": leg,
            "argv": list(argv),
            "cwd": str(cwd),
            "environment": self._recorded_environment(env),
            "source_root": identity["source_root"],
            "target_dir": str(target),
            "output_dir": str(self.output),
            "start_utc": started_at,
            "end_utc": ended_at,
            "duration_ns": time.monotonic_ns() - start,
            "returncode": returncode,
            "stdout": {"path": str(stdout_path), "bytes": stdout_bytes, "sha256": stdout_hash},
            "stderr": {"path": str(stderr_path), "bytes": stderr_bytes, "sha256": stderr_hash},
            "outputs": output_records,
            "target": {"path": str(target), "bytes_before": target_before, "bytes_after": target_after},
            "loadavg": self.guard.loadavg(),
            "source_identity_sha256": identity["identity_sha256"],
            "source_head": identity["head"],
            "source_identity_after": identity_after,
            "process_guard_violations": violation,
            "post_exit_descendants": post_descendants,
            "post_exit_offenders": post_offenders,
            "owned_processes": [owned[pid] for pid in sorted(owned)],
            "escaped_descendants": lifecycle.get("escaped", []),
            "owned_pid_reuse": lifecycle.get("pid_reuse", []),
            "runner_error": runner_error,
            "execution_class": self._event_execution_class(event_kind),
        }
        if overlap is not None:
            event["overlap"] = {
                "owned_job": owned_job,
                "peer_jobs_start": overlap_start,
                "peer_jobs_end": overlap_end,
            }
        if process_observation is not None:
            event["process_observation"] = process_observation
        self._append_event(event)
        self.recorded_events.append(event)
        if runner_error is not None:
            raise RunnerError(f"{runner_error} during {event_id}")
        if returncode != 0:
            raise RunnerError(f"child failed for {event_id} with exit {returncode}; see {stderr_path}")
        if require_stdout and stdout_bytes == 0:
            raise RunnerError(f"child produced no stdout for {event_id}")
        missing = [record["path"] for record in output_records if not record.get("exists")]
        empty = [record["path"] for record in output_records if record.get("exists") and record.get("bytes", 0) == 0]
        if missing:
            raise RunnerError(f"child did not produce declared output(s) for {event_id}: {missing}")
        if empty:
            raise RunnerError(f"child produced empty declared output(s) for {event_id}: {empty}")
        return event

    @staticmethod
    def _control_value(report: Mapping[str, Any], path: str) -> Any:
        """Read a small, explicit JSON path used by a preflight handoff."""

        value: Any = report
        for component in path.split("."):
            if not component or not isinstance(value, Mapping) or component not in value:
                raise RunnerError(f"control report is missing comparison field {path!r}")
            value = value[component]
        return value

    def _run_controls(self, identities: Mapping[str, Mapping[str, Any]]) -> None:
        """Run untimed endpoint controls before any ABBA measurement child.

        Controls are real child processes and are retained in the event
        journal, but their explicit ``measured=false``/``event_kind=control``
        fields keep RSS and timing consumers from treating them as measured
        benchmark children.
        """

        controls = self.manifest.get("controls", [])
        if not controls:
            return
        seen_outputs: set[Path] = set()
        for control in controls:
            if not isinstance(control, Mapping):
                raise RunnerError("control entry must be an object")
            control_id = control.get("id")
            compare = control.get("compare")
            programs = control.get("programs")
            if (
                not isinstance(control_id, str)
                or not control_id
                or not isinstance(compare, list)
                or not compare
                or any(not isinstance(path, str) or not path for path in compare)
                or not isinstance(programs, Mapping)
            ):
                raise RunnerError("control entry has an invalid id, comparison list, or programs")
            reports: dict[str, Mapping[str, Any]] = {}
            for leg in ("before", "after"):
                program = programs.get(leg)
                if not isinstance(program, Mapping):
                    raise RunnerError(f"control {control_id} has no {leg} program")
                context = self._context(leg, 0)
                spec = self.manifest["legs"][leg]
                binary = _resolve_path(_render(str(program.get("binary", "")), context))
                if not binary.is_file() or not os.access(binary, os.X_OK):
                    raise RunnerError(f"control binary is missing or not executable: {binary}")
                output_values = [
                    _resolve_path(_render(str(value), context))
                    for value in program.get("outputs", [])
                ]
                if len(output_values) != 1:
                    raise RunnerError(f"control {control_id}/{leg} must declare exactly one output")
                output = output_values[0]
                if not _under(output, self.output):
                    raise RunnerError(f"control output escapes output_dir: {output}")
                if output in seen_outputs or output.is_symlink() or output.exists():
                    raise RunnerError(f"refusing to overwrite or follow control output: {output}")
                output.parent.mkdir(parents=True, exist_ok=True)
                if output.parent.is_symlink():
                    raise RunnerError(f"control output parent may not be a symlink: {output.parent}")
                seen_outputs.add(output)
                program_context = dict(context)
                program_context["output"] = output
                args = _render_argv([str(value) for value in program.get("argv", [])], program_context)
                argv = [str(binary), *args]
                stdout_path = (
                    _resolve_path(_render(str(program["stdout"]), program_context))
                    if "stdout" in program
                    else None
                )
                stderr_path = (
                    _resolve_path(_render(str(program["stderr"]), program_context))
                    if "stderr" in program
                    else None
                )
                source_root = Path(identities[leg]["source_root"])
                run_cwd = _resolve_path(
                    _render(str(spec.get("run_cwd", ".")), program_context),
                    base=source_root,
                )
                if not _under(run_cwd, source_root):
                    raise RunnerError(f"control run_cwd escapes source root: {run_cwd}")
                env = self._environment(spec, context)
                self._run_command(
                    event_id=f"control-{control_id}-{leg}",
                    leg=leg,
                    argv=argv,
                    cwd=run_cwd,
                    env=env,
                    identity=identities[leg],
                    expected_outputs=output_values,
                    require_stdout=bool(program.get("require_stdout", False)),
                    stdout_path=stdout_path,
                    stderr_path=stderr_path,
                    measured=False,
                    event_kind="control",
                )
                try:
                    value = json.loads(output.read_text(encoding="utf-8"))
                except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
                    raise RunnerError(f"control {control_id}/{leg} output is not valid JSON: {error}") from error
                if not isinstance(value, Mapping):
                    raise RunnerError(f"control {control_id}/{leg} output is not an object")
                if value.get("measured") is not False:
                    raise RunnerError(f"control {control_id}/{leg} output is marked measured")
                reports[leg] = value
            for path in compare:
                before = self._control_value(reports["before"], path)
                after = self._control_value(reports["after"], path)
                if before != after:
                    raise RunnerError(
                        f"control {control_id} comparison differs before/after for {path}: "
                        f"{before!r} != {after!r}"
                    )

    def _preexec(self):
        cpu = self.cpu
        memory = self.memory_limit

        def configure() -> None:
            if cpu is not None:
                os.sched_setaffinity(0, {cpu})
            resource.setrlimit(resource.RLIMIT_AS, (memory, memory))

        return configure

    def _write_receipt(self, status: str, *, identities: Mapping[str, Any], events: Sequence[Mapping[str, Any]], error: str | None = None) -> None:
        manifest_hash, _ = _sha256_file(self.manifest_path)
        runner_path = Path(__file__).resolve()
        runner_hash, _ = _sha256_file(runner_path)
        journal_hash, journal_bytes = (
            _sha256_file(self.events_path) if self.events_path.exists() else (_sha256_bytes(b""), 0)
        )
        receipt = {
            "schema": SCHEMA,
            "status": status,
            "phase": self.phase,
            "lane": self.manifest["lane"],
            "manifest": {"path": str(self.manifest_path), "sha256": manifest_hash},
            "runner": {
                "path": str(runner_path),
                "sha256": runner_hash,
            },
            "created_utc": _now(),
            "limits": {
                "target_preflight_bytes": self.target_preflight_limit,
                "target_hard_bytes": self.target_limit,
                "memory_bytes": self.memory_limit,
            },
            "cpu": self.cpu,
            "abba_order": list(ABBA_ORDER),
            "abba_blocks": int(self.manifest.get("abba_blocks", 1)),
            "build_execution": self._build_execution_record(),
            "identities": identities,
            "event_count": len(self.recorded_events),
            "measured_event_count": sum(
                event.get("measured") is True for event in self.recorded_events
            ),
            "control_event_count": sum(
                event.get("event_kind") == "control" for event in self.recorded_events
            ),
            "event_journal": {
                "path": str(self.events_path),
                "bytes": journal_bytes,
                "sha256": journal_hash,
                "count": len(self.recorded_events),
            },
        }
        if self.preparation_path is not None and self.preparation is not None:
            preparation_hash, preparation_bytes = _sha256_file(self.preparation_path)
            receipt["preparation"] = {
                "path": str(self.preparation_path),
                "bytes": preparation_bytes,
                "sha256": preparation_hash,
            }
        if error is not None:
            receipt["error"] = error
        if self.receipt_path.is_symlink() or self.receipt_path.exists():
            raise RunnerError(f"refusing to overwrite or follow receipt: {self.receipt_path}")
        with self.receipt_path.open("x", encoding="utf-8") as stream:
            stream.write(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
            stream.flush()
            os.fsync(stream.fileno())

    def build(self) -> None:
        identities: dict[str, Any] = {}
        try:
            self._acquire_overlap_lease()
            for leg in ("before", "after"):
                target = _resolve_path(str(self.manifest["legs"][leg]["target_dir"]))
                self._require_fresh_target(target)
                self._check_target(target, preflight=True)
            for leg in ("before", "after"):
                identity = self._identity(leg)
                self._validate_preparation_identity(leg, identity)
                target = _resolve_path(str(identity["target_dir"]))
                identity["target_initial_bytes"] = self._check_target(target, preflight=True)
                identities[leg] = identity
                source = Path(identity["source_root"])
                context = self._context(leg, 0)
                env = self._environment(self.manifest["legs"][leg], context)
                for command in self.manifest["legs"][leg]["build"]:
                    argv = normalize_cargo_argv(_render_argv(command["argv"], context))
                    if self.overlap_lease is not None and _cargo_command_class(argv) != "cargo-build":
                        raise RunnerError("compile_only build_execution encountered a non-build Cargo command")
                    cwd = _resolve_path(_render(str(command.get("cwd", ".")), context), base=source)
                    if not _under(cwd, source):
                        raise RunnerError(f"build cwd escapes source root: {cwd}")
                    _validate_cargo_paths(argv, cwd=cwd, source=source, target=Path(identity["target_dir"]))
                    outputs = self._build_artifact_paths(leg, command, context)
                    for output in outputs:
                        if output.is_symlink() or output.exists():
                            raise RunnerError(f"refusing to overwrite or follow build artifact: {output}")
                    self._run_command(
                        event_id=f"build-{leg}-{command['id']}",
                        leg=leg,
                        argv=argv,
                        cwd=cwd,
                        env=env,
                        identity=identity,
                        expected_outputs=outputs,
                        measured=False,
                        event_kind="build",
                    )
                artifacts = []
                for command in self.manifest["legs"][leg]["build"]:
                    for path in self._build_artifact_paths(leg, command, context):
                        if path.is_symlink() or not path.is_file() or not os.access(path, os.X_OK):
                            raise RunnerError(f"declared binary is missing or not executable: {path}")
                        digest, size = _sha256_file(path)
                        artifacts.append({"path": str(path), "bytes": size, "sha256": digest})
                identities[leg]["artifacts"] = artifacts
                identities[leg]["target_final_bytes"] = self._check_target(target)
            self._capture_overlap_end()
            if self.overlap_lease is not None:
                self.overlap_lease.release()
            self._write_receipt("passed", identities=identities, events=self.recorded_events)
        except (BaseException,) as error:
            try:
                self._capture_overlap_end()
            except RunnerError:
                pass
            if self.overlap_lease is not None:
                self.overlap_lease.release()
            self._write_receipt("failed", identities=identities, events=self.recorded_events, error=str(error))
            raise

    def _load_build_receipt(self) -> dict[str, Any]:
        path = self.output / "build-receipt.json"
        if path.is_symlink() or not path.is_file():
            raise RunnerError(f"ABBA requires a successful build receipt: {path}")
        try:
            receipt = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            raise RunnerError(f"cannot read build receipt: {error}") from error
        self._validate_build_receipt(receipt)
        return receipt

    def _validate_build_handoff(self, expected_command_logs: set[str]) -> None:
        """Reject anything a historical analyzer could mistake for a run."""

        allowed_top_level = {
            self.command_dir,
            self.output / "events-build.jsonl",
            self.output / "build-receipt.json",
        }
        actual_top_level = {entry for entry in self.output.iterdir()}
        if actual_top_level != allowed_top_level:
            raise RunnerError(
                "build-to-ABBA output handoff contains unexpected or stale entries: "
                f"{sorted(str(path) for path in actual_top_level ^ allowed_top_level)}"
            )
        if self.command_dir.is_symlink() or not self.command_dir.is_dir():
            raise RunnerError("build command-log directory is missing or symlinked")
        actual_command_logs = {entry for entry in self.command_dir.iterdir()}
        expected_command_log_paths = {Path(path) for path in expected_command_logs}
        if actual_command_logs != expected_command_log_paths:
            raise RunnerError(
                "build command-log handoff contains unexpected or stale entries: "
                f"{sorted(str(path) for path in actual_command_logs ^ expected_command_log_paths)}"
            )

    def _validate_build_receipt(self, receipt: Mapping[str, Any]) -> None:
        if (
            receipt.get("status") != "passed"
            or receipt.get("schema") != SCHEMA
            or receipt.get("phase") != "build"
            or receipt.get("lane") != self.manifest["lane"]
        ):
            raise RunnerError("build receipt is not a passed build-phase receipt")
        manifest = receipt.get("manifest")
        runner = receipt.get("runner")
        if not isinstance(manifest, dict) or manifest.get("path") != str(self.manifest_path):
            raise RunnerError("build receipt manifest path does not match the current manifest")
        manifest_hash, _ = _sha256_file(self.manifest_path)
        if manifest.get("sha256") != manifest_hash:
            raise RunnerError("build receipt manifest hash differs from the current manifest")
        runner_path = str(Path(__file__).resolve())
        runner_hash, _ = _sha256_file(Path(__file__).resolve())
        if not isinstance(runner, dict) or runner.get("path") != runner_path or runner.get("sha256") != runner_hash:
            raise RunnerError("build receipt was produced by a different runner")
        build_execution = receipt.get("build_execution")
        expected_execution = self.build_execution.get("class")
        if not isinstance(build_execution, dict) or build_execution.get("class") != expected_execution:
            raise RunnerError("build receipt execution class differs from the current manifest")
        if expected_execution == "compile_only":
            permit_value = self.build_execution.get("permit_path")
            if build_execution.get("permit_path") != str(_resolve_path(str(permit_value))):
                raise RunnerError("build receipt permit path differs from the current manifest")
        if self.preparation_path is not None:
            preparation = receipt.get("preparation")
            if not isinstance(preparation, dict) or preparation.get("path") != str(self.preparation_path):
                raise RunnerError("build receipt preparation path does not match the current preparation")
            preparation_hash, preparation_bytes = _sha256_file(self.preparation_path)
            if preparation.get("sha256") != preparation_hash or preparation.get("bytes") != preparation_bytes:
                raise RunnerError("build receipt preparation identity differs from the current receipt")
        limits = receipt.get("limits")
        expected_limits = {
            "target_preflight_bytes": TARGET_PREFLIGHT_BYTES,
            "target_hard_bytes": TARGET_HARD_LIMIT_BYTES,
            "memory_bytes": MEMORY_HARD_LIMIT_BYTES,
        }
        if limits != expected_limits:
            raise RunnerError("build receipt limits are not the runner's fixed limits")
        if receipt.get("abba_order") != list(ABBA_ORDER):
            raise RunnerError("build receipt ABBA order differs from the fixed order")
        if receipt.get("abba_blocks", 1) != int(self.manifest.get("abba_blocks", 1)):
            raise RunnerError("build receipt ABBA block count differs from the manifest")
        journal = receipt.get("event_journal")
        if not isinstance(journal, dict) or journal.get("path") != str(self.output / "events-build.jsonl"):
            raise RunnerError("build receipt event journal path does not match")
        if self.events_path != self.output / "events-abba.jsonl":
            raise RunnerError("ABBA runner has an unexpected current event journal")
        build_events_path = self.output / "events-build.jsonl"
        if build_events_path.is_symlink() or not build_events_path.is_file():
            raise RunnerError("build event journal is missing or symlinked")
        journal_hash, journal_bytes = _sha256_file(build_events_path)
        if journal.get("sha256") != journal_hash or journal.get("bytes") != journal_bytes:
            raise RunnerError("build event journal hash or size differs from the receipt")
        lines = build_events_path.read_text(encoding="utf-8").splitlines()
        if journal.get("count") != len(lines) or receipt.get("event_count") != len(lines):
            raise RunnerError("build event journal count differs from the receipt")
        try:
            events = [json.loads(line) for line in lines]
        except json.JSONDecodeError as error:
            raise RunnerError(f"build event journal contains invalid JSON: {error}") from error
        expected_events: list[tuple[str, str, dict[str, Any]]] = []
        for leg in ("before", "after"):
            source = Path(_resolve_path(str(self.manifest["legs"][leg]["source_root"])))
            context = self._context(leg, 0)
            for command in self.manifest["legs"][leg]["build"]:
                argv = normalize_cargo_argv(_render_argv(command["argv"], context))
                cwd = _resolve_path(_render(str(command.get("cwd", ".")), context), base=source)
                if not _under(cwd, source):
                    raise RunnerError(f"build cwd escapes source root: {cwd}")
                target = _resolve_path(str(self.manifest["legs"][leg]["target_dir"]))
                _validate_cargo_paths(argv, cwd=cwd, source=source, target=target)
                expected_events.append(
                    (
                        f"build-{leg}-{command['id']}",
                        leg,
                        {
                            "argv": argv,
                            "cwd": str(cwd),
                            "source": source,
                            "target": str(target),
                            "artifacts": [str(path) for path in self._build_artifact_paths(leg, command, context)],
                        },
                    )
                )
        if len(events) != len(expected_events):
            raise RunnerError("build event journal does not contain the complete build order")
        identities = receipt.get("identities")
        if not isinstance(identities, dict):
            raise RunnerError("build receipt identities are missing")
        seen_log_paths: set[str] = set()
        expected_command_logs: set[str] = set()
        for event, (event_id, leg, expectation) in zip(events, expected_events):
            if not isinstance(event, dict) or event.get("schema") != SCHEMA:
                raise RunnerError(f"invalid build event for {event_id}")
            if event.get("event_id") != event_id or event.get("leg") != leg or event.get("argv") != expectation["argv"]:
                raise RunnerError(f"build event order/argv mismatch for {event_id}")
            recorded = identities.get(leg)
            if not isinstance(recorded, dict):
                raise RunnerError(f"missing source identity for {leg}")
            if (
                event.get("source_root") != recorded.get("source_root")
                or event.get("target_dir") != recorded.get("target_dir")
                or event.get("cwd") != expectation["cwd"]
            ):
                raise RunnerError(f"resolved path identity mismatch for {event_id}")
            target_record = event.get("target")
            if (
                event.get("output_dir") != str(self.output)
                or not isinstance(target_record, dict)
                or target_record.get("path") != recorded.get("target_dir")
            ):
                raise RunnerError(f"output/target identity mismatch for {event_id}")
            if event.get("source_identity_sha256") != recorded.get("identity_sha256"):
                raise RunnerError(f"source identity hash mismatch for {event_id}")
            expected_event_class = "compile_only" if expected_execution == "compile_only" else "exclusive"
            if event.get("execution_class") != expected_event_class:
                raise RunnerError(f"build event execution class differs for {event_id}")
            if expected_event_class == "compile_only":
                overlap = event.get("overlap")
                owned_job = overlap.get("owned_job") if isinstance(overlap, dict) else None
                if (
                    not isinstance(overlap, dict)
                    or not isinstance(overlap.get("peer_jobs_start"), list)
                    or not isinstance(overlap.get("peer_jobs_end"), list)
                    or not isinstance(owned_job, dict)
                    or owned_job.get("root_pid", 0) <= 0
                    or not isinstance(owned_job.get("root_starttime"), str)
                    or owned_job.get("command_class") != "cargo-build"
                    or owned_job.get("phase") != "build"
                    or owned_job.get("phase_id") != event_id
                ):
                    raise RunnerError(f"build overlap receipt is missing for {event_id}")
            identity_after = event.get("source_identity_after")
            if not isinstance(identity_after, dict) or identity_after.get("identity_sha256") != recorded.get("identity_sha256"):
                raise RunnerError(f"source identity-after mismatch for {event_id}")
            if event.get("returncode") != 0 or event.get("runner_error") is not None:
                raise RunnerError(f"build event is not successful for {event_id}")
            if event.get("process_guard_violations") or event.get("post_exit_descendants") or event.get("post_exit_offenders"):
                raise RunnerError(f"build event was not quiet for {event_id}")
            if (
                not isinstance(target_record, dict)
                or target_record.get("path") != expectation["target"]
                or not isinstance(target_record.get("bytes_before"), int)
                or not isinstance(target_record.get("bytes_after"), int)
                or not 0 <= target_record["bytes_before"] <= TARGET_HARD_LIMIT_BYTES
                or not 0 <= target_record["bytes_after"] <= TARGET_HARD_LIMIT_BYTES
            ):
                raise RunnerError(f"build target limits/path are invalid for {event_id}")
            output_records = event.get("outputs")
            if (
                not isinstance(output_records, list)
                or not all(isinstance(item, dict) for item in output_records)
                or [item.get("path") for item in output_records] != expectation["artifacts"]
            ):
                raise RunnerError(f"build artifact path order differs for {event_id}")
            for output_record in output_records:
                output_path = _resolve_path(str(output_record.get("path", "")))
                if output_path in seen_log_paths:
                    raise RunnerError(f"build output path is reused for {event_id}: {output_path}")
                if output_path.is_symlink() or not output_path.is_file():
                    raise RunnerError(f"build output path is missing or symlinked for {event_id}: {output_path}")
                output_hash, output_bytes = _sha256_file(output_path)
                if output_hash != output_record.get("sha256") or output_bytes != output_record.get("bytes"):
                    raise RunnerError(f"build output hash mismatch for {event_id}: {output_path}")
                seen_log_paths.add(str(output_path))
            event_environment = event.get("environment")
            if not isinstance(event_environment, dict) or {
                "CARGO_TARGET_DIR": event_environment.get("CARGO_TARGET_DIR"),
                "CARGO_BUILD_JOBS": event_environment.get("CARGO_BUILD_JOBS"),
                "CARGO_INCREMENTAL": event_environment.get("CARGO_INCREMENTAL"),
                "LITCHI_MEM_MAX": event_environment.get("LITCHI_MEM_MAX"),
            } != {
                "CARGO_TARGET_DIR": expectation["target"],
                "CARGO_BUILD_JOBS": "1",
                "CARGO_INCREMENTAL": "0",
                "LITCHI_MEM_MAX": "20G",
            }:
                raise RunnerError(f"build environment limits/path differ for {event_id}")
            for stream_name in ("stdout", "stderr"):
                stream = event.get(stream_name)
                if not isinstance(stream, dict):
                    raise RunnerError(f"missing {stream_name} receipt for {event_id}")
                stream_path = _resolve_path(str(stream.get("path", "")))
                if (
                    not _under(stream_path, self.command_dir)
                    or stream_path.is_symlink()
                    or not stream_path.is_file()
                ):
                    raise RunnerError(f"invalid {stream_name} path for {event_id}")
                if str(stream_path) in seen_log_paths:
                    raise RunnerError(f"build log path is reused for {event_id}: {stream_path}")
                stream_hash, stream_bytes = _sha256_file(stream_path)
                if stream_hash != stream.get("sha256") or stream_bytes != stream.get("bytes"):
                    raise RunnerError(f"{stream_name} hash mismatch for {event_id}")
                seen_log_paths.add(str(stream_path))
                expected_command_logs.add(str(stream_path))
        self._validate_build_handoff(expected_command_logs)
        for leg in ("before", "after"):
            current = self._identity(leg)
            self._validate_preparation_identity(leg, current)
            recorded = identities.get(leg)
            if not isinstance(recorded, dict):
                raise RunnerError(f"missing identity for {leg}")
            for key in ("source_root", "target_dir", "output_dir", "head", "tree", "source_tree_sha256", "identity_sha256", "lock_paths", "probe_paths", "fixture_paths"):
                if current.get(key) != recorded.get(key):
                    raise RunnerError(f"build receipt identity changed for {leg}: {key}")
            if recorded.get("target_initial_bytes") != 0:
                raise RunnerError(f"build target was not fresh for {leg}")
            target = Path(recorded["target_dir"])
            target_now = self._check_target(target)
            if target_now != recorded.get("target_final_bytes"):
                raise RunnerError(f"build target changed before timing for {leg}")
            artifacts = recorded.get("artifacts")
            if not isinstance(artifacts, list) or not artifacts:
                raise RunnerError(f"build artifacts are missing for {leg}")
            for artifact in artifacts:
                if not isinstance(artifact, dict):
                    raise RunnerError(f"build artifact record is malformed for {leg}")
                artifact_path = _resolve_path(str(artifact.get("path", "")))
                if not _under(artifact_path, Path(recorded["target_dir"])) or artifact_path.is_symlink() or not artifact_path.is_file():
                    raise RunnerError(f"build artifact path is not a fresh target artifact: {artifact_path}")
                artifact_hash, artifact_bytes = _sha256_file(artifact_path)
                if artifact_hash != artifact.get("sha256") or artifact_bytes != artifact.get("bytes"):
                    raise RunnerError(f"build artifact hash changed before timing: {artifact_path}")

    def abba(self) -> None:
        identities: dict[str, Any] = {}
        try:
            build_receipt = self._load_build_receipt()
            for leg in ("before", "after"):
                identity = self._identity(leg)
                self._validate_preparation_identity(leg, identity)
                recorded = build_receipt.get("identities", {}).get(leg, {})
                if (
                    identity.get("head") != recorded.get("head")
                    or identity.get("tree") != recorded.get("tree")
                    or identity.get("source_tree_sha256") != recorded.get("source_tree_sha256")
                    or identity.get("identity_sha256") != recorded.get("identity_sha256")
                ):
                    raise RunnerError(f"{leg} source identity differs from the build receipt")
                recorded_artifacts = recorded.get("artifacts", [])
                if not isinstance(recorded_artifacts, list) or not recorded_artifacts:
                    raise RunnerError(f"build receipt has no artifacts for {leg}")
                recorded_paths: set[str] = set()
                for artifact in recorded_artifacts:
                    path = _resolve_path(str(artifact.get("path", "")))
                    if not path.is_file() or not os.access(path, os.X_OK):
                        raise RunnerError(f"built artifact is missing or not executable: {path}")
                    digest, size = _sha256_file(path)
                    if digest != artifact.get("sha256") or size != artifact.get("bytes"):
                        raise RunnerError(f"built artifact changed since build receipt: {path}")
                    recorded_paths.add(str(path))
                for program in self.manifest["legs"][leg]["programs"]:
                    context = self._context(leg, 0)
                    program_path = _resolve_path(_render(str(program["binary"]), context))
                    if str(program_path) not in recorded_paths:
                        raise RunnerError(f"program binary was not recorded by the build phase: {program_path}")
                identities[leg] = identity
            self._run_controls(identities)
            blocks = int(self.manifest.get("abba_blocks", 1))
            for block in range(1, blocks + 1):
                for index, leg in enumerate(ABBA_ORDER, start=1):
                    spec = self.manifest["legs"][leg]
                    context = self._context(leg, index, block=block)
                    env = self._environment(spec, context)
                    for program in spec["programs"]:
                        program_id = program["id"]
                        binary = _resolve_path(_render(str(program["binary"]), context))
                        if not binary.is_file() or not os.access(binary, os.X_OK):
                            raise RunnerError(f"measurement binary is missing or not executable: {binary}")
                        output_values = [
                            _resolve_path(_render(str(value), context))
                            for value in program.get("outputs", [])
                        ]
                        for output in output_values:
                            if not _under(output, self.output):
                                raise RunnerError(f"declared child output escapes output_dir: {output}")
                            if output.is_symlink() or output.exists():
                                raise RunnerError(f"refusing to overwrite or follow child output: {output}")
                            output.parent.mkdir(parents=True, exist_ok=True)
                            if output.parent.is_symlink():
                                raise RunnerError(f"child output parent may not be a symlink: {output.parent}")
                        program_context = dict(context)
                        if len(output_values) == 1:
                            program_context["output"] = output_values[0]
                        program_context["program"] = program_id
                        args = _render_argv([str(value) for value in program["argv"]], program_context)
                        argv = [str(binary), *args]
                        stdout_path = (
                            _resolve_path(_render(str(program["stdout"]), program_context))
                            if "stdout" in program
                            else None
                        )
                        stderr_path = (
                            _resolve_path(_render(str(program["stderr"]), program_context))
                            if "stderr" in program
                            else None
                        )
                        run_cwd = _resolve_path(str(spec.get("run_cwd", ".")), base=Path(identity["source_root"]))
                        if not _under(run_cwd, Path(identity["source_root"])):
                            raise RunnerError(f"run_cwd escapes source root: {run_cwd}")
                        event_id = (
                            f"abba-{index:02d}-{leg}-{program_id}"
                            if blocks == 1
                            else f"abba-block-{block:02d}-{index:02d}-{leg}-{program_id}"
                        )
                        self._run_command(
                            event_id=event_id,
                            leg=leg,
                            argv=argv,
                            cwd=run_cwd,
                            env=env,
                            identity=identities[leg],
                            expected_outputs=output_values,
                            require_stdout=bool(program.get("require_stdout", True)),
                            stdout_path=stdout_path,
                            stderr_path=stderr_path,
                        )
            validator_context = self._context("after", blocks * len(ABBA_ORDER) + 1, block=blocks)
            validator_spec = self.manifest["legs"]["after"]
            validator_env = self._environment(validator_spec, validator_context)
            for validator in self.manifest["validators"]:
                validator_id = validator["id"]
                argv = _render_argv([str(value) for value in validator["argv"]], validator_context)
                validator_script = _resolve_path(_render(str(validator["script"]), validator_context))
                if validator_script.is_symlink() or not validator_script.is_file():
                    raise RunnerError(f"validator script is missing or symlinked: {validator_script}")
                if len(argv) < 2 or _resolve_path(argv[1]) != validator_script:
                    raise RunnerError(f"validator argv does not invoke its declared script: {validator_id}")
                outputs = [
                    _resolve_path(_render(str(value), validator_context))
                    for value in validator.get("outputs", [])
                ]
                for output in outputs:
                    if not _under(output, self.output) or output.is_symlink() or output.exists():
                        raise RunnerError(f"validator output is invalid or already exists: {output}")
                    output.parent.mkdir(parents=True, exist_ok=True)
                    if output.parent.is_symlink():
                        raise RunnerError(f"validator output parent may not be a symlink: {output.parent}")
                validator_cwd = _resolve_path(
                    _render(str(validator.get("cwd", self.output)), validator_context)
                )
                self._run_command(
                    event_id=f"validator-{validator_id}",
                    leg="after",
                    argv=argv,
                    cwd=validator_cwd,
                    env=validator_env,
                    identity=identities["after"],
                    expected_outputs=outputs,
                    require_stdout=bool(validator.get("require_stdout", True)),
                    preflight_target=False,
                    measured=False,
                    event_kind="validator",
                )
            self._write_receipt("passed", identities=identities, events=self.recorded_events)
        except (BaseException,) as error:
            self._write_receipt("failed", identities=identities, events=self.recorded_events, error=str(error))
            raise


def _parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--phase", choices=("build", "abba"), required=True)
    parser.add_argument("--cpu", type=int, help="CPU affinity for child programs")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parse_args(argv or sys.argv[1:])
    try:
        manifest_path = _resolve_path(args.manifest, base=Path.cwd())
        manifest = load_manifest(manifest_path)
        runner = SafeRunner(manifest, manifest_path, args.phase, args.cpu)
        if args.phase == "build":
            runner.build()
        else:
            runner.abba()
    except RunnerError as error:
        print(f"native qualification refused: {error}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("native qualification interrupted; inspect retained logs", file=sys.stderr)
        return 130
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
