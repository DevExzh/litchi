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
from typing import Any, Iterable, Mapping, Sequence

from native_qualification.prepare_xlsx_raw_inputs import (
    BASELINE_REVISION,
    validate_preparation_receipt,
)


SCHEMA = "litchi.native-qualification.v1"
ABBA_ORDER = ("before", "after", "after", "before")
ROUNDS = 31
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_LIMIT_BYTES = 20_000_000_000
MEMORY_HARD_LIMIT_BYTES = 20_000_000_000
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


def _now() -> str:
    return _dt.datetime.now(_dt.timezone.utc).isoformat(timespec="milliseconds")


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
    if data.get("lane") == "w4-xlsx-raw":
        value = data["shared"].get("preparation_receipt")
        if not isinstance(value, str) or not value:
            raise RunnerError("XLSX raw lane needs a preparation receipt")
        preparation_receipt = _resolve_path(value)
        if _under(preparation_receipt, output) or preparation_receipt == output:
            raise RunnerError("XLSX preparation receipt must be outside the fresh output directory")
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


def _read_proc(pid: int) -> dict[str, Any] | None:
    proc = Path("/proc") / str(pid)
    try:
        comm = (proc / "comm").read_text(encoding="utf-8", errors="replace").strip()
        cmdline = (proc / "cmdline").read_bytes().split(b"\0")
        cmdline = [part.decode("utf-8", "replace") for part in cmdline if part]
        exe = os.readlink(proc / "exe")
        cwd = os.readlink(proc / "cwd")
        stat = (proc / "stat").read_text(encoding="utf-8", errors="replace")
    except (FileNotFoundError, PermissionError, OSError):
        return None
    right = stat.rsplit(") ", 1)
    if len(right) != 2:
        return None
    fields = right[1].split()
    try:
        parent = int(fields[1])
        process_group = int(fields[2])
    except (IndexError, ValueError):
        return None
    return {
        "pid": pid,
        "ppid": parent,
        "pgrp": process_group,
        "comm": comm,
        "exe": exe,
        "cwd": cwd,
        "cmdline": cmdline,
    }


def _proc_snapshot() -> dict[int, dict[str, Any]]:
    entries: dict[int, dict[str, Any]] = {}
    try:
        pids = [int(item.name) for item in Path("/proc").iterdir() if item.name.isdigit()]
    except OSError:
        return entries
    for pid in pids:
        entry = _read_proc(pid)
        if entry is not None:
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

    def offenders(self, allowed_root: int | None = None) -> list[dict[str, Any]]:
        snapshot = _proc_snapshot()
        allowed = _descendants(snapshot, allowed_root) if allowed_root is not None else set()
        return [
            entry
            for pid, entry in snapshot.items()
            if pid not in allowed and _in_litchi_scope(entry) and _is_heavy(entry)
        ]

    def require_quiet(self) -> None:
        offenders = self.offenders()
        if offenders:
            raise RunnerError(f"unrelated Litchi builder/profiler is running: {offenders!r}")

    @staticmethod
    def loadavg() -> list[float] | None:
        try:
            return [float(value) for value in os.getloadavg()]
        except (AttributeError, OSError):
            return None


def _json_hash(path: Path) -> tuple[str, int]:
    # ``Path.exists()`` is false for a dangling symlink.  Check the link first
    # so a replaced log cannot be silently recorded as an empty stream.
    if path.is_symlink():
        raise RunnerError(f"refusing to follow symlinked log: {path}")
    return _sha256_file(path) if path.exists() else (_sha256_bytes(b""), 0)


class SafeRunner:
    def __init__(self, manifest: Mapping[str, Any], manifest_path: Path, phase: str, cpu: int | None):
        self.manifest = manifest
        self.manifest_path = manifest_path.resolve()
        self.phase = phase
        self.cpu = cpu
        self.guard = ProcessGuard()
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

    def _wait_descendants_gone(self, pid: int) -> list[int]:
        deadline = time.monotonic() + 5.0
        remaining: set[int] = set()
        while time.monotonic() < deadline:
            snapshot = _proc_snapshot()
            remaining = {
                child_pid
                for child_pid, entry in snapshot.items()
                if child_pid != pid and (child_pid in _descendants(snapshot, pid) or entry.get("pgrp") == pid)
            }
            if not remaining:
                return []
            time.sleep(0.1)
        return sorted(remaining)

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
    ) -> dict[str, Any]:
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
        runner_error: str | None = None
        last_target_check = 0.0
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
                while True:
                    returncode = process.poll()
                    if returncode is not None:
                        break
                    try:
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
                    except RunnerError as error:
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
        ended_at = _now()
        post_descendants = self._wait_descendants_gone(process.pid) if process is not None else []
        post_offenders = self.guard.offenders()
        if post_descendants:
            runner_error = runner_error or f"child descendants survived exit: {post_descendants}"
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
            "runner_error": runner_error,
        }
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
            "identities": identities,
            "event_count": len(self.recorded_events),
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
            self._write_receipt("passed", identities=identities, events=self.recorded_events)
        except (BaseException,) as error:
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
        manifest_path = args.manifest.resolve()
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
