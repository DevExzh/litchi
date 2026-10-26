#!/usr/bin/env python3
"""Bounded source-build handoff for the current ODP qualification probe.

The ODP correctness runner intentionally consumes an executable and never
invokes Cargo.  This companion owns the preceding source/build boundary: it
checks the frozen source, every Cargo lock/config/probe identity, fresh
external target/output paths, and the fixed one-job/20G/process-guard policy;
then it records the observed executable identity.  ``--seal`` is the only
operation that may add that observed identity to a candidate input manifest.
No binary identity is accepted from the source binder.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
from typing import Any, Mapping

# When invoked as ``python tools/native_qualification/...py``, Python puts
# only this subdirectory on ``sys.path``.  Add the repository's tools
# directory explicitly so the reviewed ODG/ODP helpers are the same modules
# used by the test-suite and by package imports.
_REPO_ROOT = Path(__file__).resolve().parents[2]
_TOOLS_ROOT = _REPO_ROOT / "tools"
for _path in (str(_REPO_ROOT), str(_TOOLS_ROOT)):
    if _path not in sys.path:
        sys.path.insert(0, _path)

try:
    from tools.odg_joined_qualification_builder import (
        _environment_identity,
        _check_worktree_shape,
        _identity,
        _quiet_guard,
        _safe_path,
        _sha256_file,
        _stream_command,
        _tree_manifest,
    )
except ImportError:  # pragma: no cover
    from odg_joined_qualification_builder import (  # type: ignore[no-redef]
        _environment_identity,
        _check_worktree_shape,
        _identity,
        _quiet_guard,
        _safe_path,
        _sha256_file,
        _stream_command,
        _tree_manifest,
    )

try:
    from tools.odp_native_qualification import RunFailure, _digest, _load_json, _object, _positive_int
except ImportError:  # pragma: no cover
    from odp_native_qualification import (  # type: ignore[no-redef]
        RunFailure,
        _digest,
        _load_json,
        _object,
        _positive_int,
    )


PREPARATION_SCHEMA = "litchi.odp.native-qualification.source-build-preparation.v1"
BUILD_SCHEMA = "litchi.odp.native-qualification.source-build.v1"
BUILDER_REVISION = "odp-candidate-build-gate-v1-locked-source-receipt"
CANDIDATE_REF = "33a0a63d3553712c1276da592f7846a0c7dcabd6"
CANDIDATE_TREE = "656d66603b33aec273d529263aa10b504bd3b7f8"
PROBE_RELATIVE = "tools/perf-baseline/src/bin/odp_native_qualification_probe.rs"
PERF_MANIFEST_RELATIVE = "tools/perf-baseline/Cargo.toml"
PERF_LOCK_RELATIVE = "tools/perf-baseline/Cargo.lock"
ROOT_LOCK_RELATIVE = "Cargo.lock"
ROOT_MANIFEST_RELATIVE = "Cargo.toml"
TOOLCHAIN_RELATIVE = "rust-toolchain.toml"
CONFIG_RELATIVE = ".cargo/config.toml"
MAX_OUTPUT_BYTES = 4 * 1024 * 1024
MEMORY_LIMIT_BYTES = 20_000_000_000
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_LIMIT_BYTES = 20_000_000_000
BUILD_SPEC_KEYS = {"runner", "argv_template", "environment", "claim_scope"}
BUILD_RUNNER = "tools/native_qualification/odp_candidate_build_gate.py"
BUILD_CLAIM = "bounded locked source build only; no correctness, timing, or acceptance claim"


def _exact_keys(value: Mapping[str, Any], expected: set[str], label: str) -> None:
    if set(value) != expected:
        raise RunFailure(f"{label} fields differ from the bounded-build contract")


def _sha(path: Path) -> dict[str, Any]:
    digest, size = _sha256_file(path)
    return {"path": str(path), "bytes": size, "sha256": digest}


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=False,
        timeout=20,
    )
    if result.returncode:
        raise RunFailure(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout.strip()


def _fresh_external(path: Path, roots: tuple[Path, ...], label: str) -> Path:
    path = _safe_path(path.absolute(), label)
    for root in roots:
        try:
            path.relative_to(root)
        except ValueError:
            continue
        raise RunFailure(f"{label} must be outside source root: {path}")
    if path.exists() or path.is_symlink():
        raise RunFailure(f"{label} must be fresh and absent: {path}")
    return path


def _existing_external(path: Path, roots: tuple[Path, ...], label: str) -> Path:
    path = _safe_path(path.absolute(), label, must_exist=True)
    for root in roots:
        try:
            path.relative_to(root)
        except ValueError:
            continue
        raise RunFailure(f"{label} must be outside source root: {path}")
    if not path.is_dir() or path.is_symlink():
        raise RunFailure(f"{label} must be an existing regular directory: {path}")
    return path


def _under(path: Path, parent: Path) -> bool:
    return path == parent or parent in path.parents


def _identity(path: Path) -> dict[str, Any]:
    return _sha(path)


def _source_files(root: Path) -> dict[str, dict[str, Any]]:
    relative = {
        "root_lock": ROOT_LOCK_RELATIVE,
        "root_manifest": ROOT_MANIFEST_RELATIVE,
        "toolchain": TOOLCHAIN_RELATIVE,
        "cargo_config": CONFIG_RELATIVE,
        "perf_manifest": PERF_MANIFEST_RELATIVE,
        "perf_lock": PERF_LOCK_RELATIVE,
        "probe": PROBE_RELATIVE,
    }
    return {name: _identity(root / value) | {"relative": value} for name, value in relative.items()}


def _source_state(root: Path, expected_ref: str, expected_tree: str, expected_files: Mapping[str, Any] | None = None) -> dict[str, Any]:
    root = _safe_path(root.absolute(), "candidate source root", must_exist=True)
    shape = _check_worktree_shape(root, "candidate")
    head = _git(root, "rev-parse", "--verify", "HEAD")
    tree = _git(root, "rev-parse", "--verify", "HEAD^{tree}")
    status = _git(root, "status", "--porcelain", "--untracked-files=all")
    if head != expected_ref or tree != expected_tree or status:
        raise RunFailure("candidate source HEAD/tree/status differs from the frozen build input")
    files = _source_files(root)
    if expected_files is not None:
        for name, expected in expected_files.items():
            actual = files[name]
            if (actual["bytes"], actual["sha256"]) != (expected["bytes"], expected["sha256"]):
                raise RunFailure(f"candidate source file changed: {name}")
    return {
        "root": str(root), "ref": head, "tree": tree, "status": status,
        "source_shape": shape,
        "status_sha256": hashlib.sha256(status.encode()).hexdigest(),
        "source_manifest": _tree_manifest(root, head), "files": files,
    }


def _load_binder(path: Path) -> dict[str, Any]:
    path = _safe_path(path.absolute(), "ODP source binder", must_exist=True)
    value = _load_json(path)
    if value.get("schema") != "litchi.odp.native-qualification.c2a-source-binder.v1":
        raise RunFailure("ODP source binder schema differs")
    candidate = _object(value.get("candidate"), "source binder.candidate")
    if candidate.get("commit") != CANDIDATE_REF or candidate.get("tree") != CANDIDATE_TREE:
        raise RunFailure("ODP source binder is not bound to C2a")
    spec = _object(value.get("build"), "source binder.build")
    if set(spec) != BUILD_SPEC_KEYS:
        raise RunFailure("ODP source binder.build must be the structured bounded-build contract")
    if spec.get("runner") != BUILD_RUNNER or spec.get("claim_scope") != BUILD_CLAIM:
        raise RunFailure("ODP source binder build runner/claim scope differs")
    argv_template = spec.get("argv_template")
    expected_template = [
        "cargo", "build", "--offline", "--locked", "--jobs", "1", "--release",
        "--manifest-path", "{candidate_root}/tools/perf-baseline/Cargo.toml",
        "--target-dir", "{target_dir}", "--bin", "odp_native_qualification_probe",
    ]
    if argv_template != expected_template:
        raise RunFailure("ODP source binder build argv is not the approved structured composition")
    environment = _object(spec.get("environment"), "source binder.build.environment")
    expected_environment = {
        "CARGO_BUILD_JOBS": "1", "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_DEBUG": "0", "CARGO_TARGET_DIR": "{target_dir}",
        "LITCHI_MEM_MAX": "20G", "TMPDIR": "{tmpdir}",
    }
    if environment != expected_environment:
        raise RunFailure("ODP source binder build environment is not the approved capped composition")
    return value


def _candidate_path(binder: Mapping[str, Any]) -> Path:
    return _safe_path(Path(str(_object(binder.get("candidate"), "source binder.candidate")["root"])), "candidate source root", must_exist=True)


def _fixture_identity(binder: Mapping[str, Any]) -> dict[str, Any]:
    fixture = _object(binder.get("fixture"), "source binder.fixture")
    path = _safe_path(Path(str(fixture["path"])), "ODP fixture", must_exist=True)
    actual = _identity(path)
    if actual["bytes"] != fixture.get("bytes") or actual["sha256"] != fixture.get("sha256"):
        raise RunFailure("ODP fixture identity differs from source binder")
    return {**actual, "provenance": fixture.get("provenance", "")}


def prepare(
    binder_path: Path,
    *,
    target_dir: Path,
    build_output: Path,
    tmpdir: Path,
    cpu: int,
    output: Path,
) -> dict[str, Any]:
    binder_path = _safe_path(binder_path.absolute(), "ODP source binder", must_exist=True)
    output = _safe_path(output.absolute(), "ODP source preparation receipt")
    if output.exists() or output.is_symlink():
        raise RunFailure(f"refusing to overwrite ODP preparation receipt: {output}")
    binder = _load_binder(binder_path)
    candidate_root = _candidate_path(binder)
    roots = (candidate_root,)
    target_dir = _fresh_external(target_dir, roots, "ODP candidate target")
    build_output = _fresh_external(build_output, roots, "ODP build output")
    tmpdir = _existing_external(tmpdir, roots, "ODP build TMPDIR")
    if any(_under(left, right) or _under(right, left) for left, right in (
        (target_dir, build_output), (target_dir, tmpdir), (build_output, tmpdir)
    )):
        raise RunFailure("ODP target/output/TMPDIR paths must be pairwise disjoint")
    if _under(output, candidate_root) or any(_under(output, path) or _under(path, output)
                                             for path in (target_dir, build_output, tmpdir)):
        raise RunFailure("ODP preparation receipt must be outside source and build paths")
    candidate = _object(binder["candidate"], "source binder.candidate")
    expected_files = {
        "root_lock": {"bytes": candidate["cargo_lock"]["bytes"], "sha256": candidate["cargo_lock"]["sha256"]},
        "probe": {"bytes": candidate["probe"]["bytes"], "sha256": candidate["probe"]["sha256"]},
    }
    state = _source_state(candidate_root, CANDIDATE_REF, CANDIDATE_TREE, expected_files)
    fixture = _fixture_identity(binder)
    command = [
        "cargo", "build", "--offline", "--locked", "--jobs", "1", "--release",
        "--manifest-path", str(candidate_root / PERF_MANIFEST_RELATIVE),
        "--target-dir", str(target_dir),
        "--bin", "odp_native_qualification_probe",
    ]
    build_spec = _object(binder["build"], "source binder.build")
    expected_argv = [
        value.replace("{candidate_root}", str(candidate_root)).replace("{target_dir}", str(target_dir))
        for value in build_spec["argv_template"]
    ]
    if command != expected_argv:
        raise RunFailure("ODP generated build argv differs from the bound structured composition")
    overrides = {
        "CARGO_BUILD_JOBS": "1", "CARGO_INCREMENTAL": "0", "CARGO_PROFILE_RELEASE_DEBUG": "0",
        "CARGO_TARGET_DIR": str(target_dir), "LITCHI_MEM_MAX": "20G", "TMPDIR": str(tmpdir),
    }
    expected_environment = {
        key: value.replace("{target_dir}", str(target_dir)).replace("{tmpdir}", str(tmpdir))
        for key, value in _object(build_spec["environment"], "source binder.build.environment").items()
    }
    if overrides != expected_environment:
        raise RunFailure("ODP generated build environment differs from the bound composition")
    environment = _environment_identity(
        candidate_root, ref=CANDIDATE_REF, tree=CANDIDATE_TREE,
        source_closure={"entries": state["source_manifest"]["entries"], "bytes": state["source_manifest"]["bytes"], "sha256": state["source_manifest"]["sha256"]},
        overrides=overrides,
    )
    artifact = target_dir / "release" / "odp_native_qualification_probe"
    plan = {
        "schema": PREPARATION_SCHEMA,
        "builder_revision": BUILDER_REVISION,
        "builder_source": _sha(Path(__file__).resolve()),
        "runtime_source": _sha(candidate_root / "tools/odp_native_qualification.py"),
        "status": "source_prepared",
        "source_binder": {"path": str(binder_path), **_sha(binder_path)},
        "candidate": state,
        "fixture": fixture,
        "command": {"argv": command, "environment": environment, "environment_overrides": overrides},
        "artifact": {"path": str(artifact), "bytes": None, "sha256": None, "status": "pending bounded C2a build"},
        "target_dir": str(target_dir), "build_output": str(build_output), "tmpdir": str(tmpdir),
        "cpu": cpu, "max_load": 1.0, "timeout_s": 300.0,
        "limits": {"memory_bytes": MEMORY_LIMIT_BYTES, "target_preflight_bytes": TARGET_PREFLIGHT_BYTES, "target_hard_bytes": TARGET_HARD_LIMIT_BYTES, "max_output_bytes": MAX_OUTPUT_BYTES},
        "claim_scope": "source preparation only; no build, correctness, timing, or acceptance claim",
        "receipt": {"path": str(output)},
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(plan, indent=2, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())
    return plan


def _load_preparation(path: Path) -> dict[str, Any]:
    path = _safe_path(path.absolute(), "ODP source preparation receipt", must_exist=True)
    value = _load_json(path)
    if value.get("schema") != PREPARATION_SCHEMA or value.get("status") != "source_prepared":
        raise RunFailure("ODP preparation schema/status differs")
    binder = _object(value.get("source_binder"), "ODP preparation.source_binder")
    binder_path = _safe_path(Path(str(binder.get("path"))), "ODP source binder", must_exist=True)
    if _sha(binder_path) != {key: binder[key] for key in ("path", "bytes", "sha256")}:
        raise RunFailure("ODP source binder changed after preparation")
    if value.get("builder_source") != _sha(Path(__file__).resolve()):
        raise RunFailure("ODP build gate source changed after preparation")
    binder_value = _load_binder(binder_path)
    candidate = _object(value.get("candidate"), "ODP preparation.candidate")
    binder_candidate = _object(binder_value.get("candidate"), "ODP source binder.candidate")
    if candidate.get("root") != binder_candidate.get("root") or candidate.get("ref") != CANDIDATE_REF or candidate.get("tree") != CANDIDATE_TREE:
        raise RunFailure("ODP preparation candidate identity differs from the source binder")
    expected_files = candidate.get("files")
    if not isinstance(expected_files, dict):
        raise RunFailure("ODP preparation candidate source files are missing")
    for key, binder_key in (("root_lock", "cargo_lock"), ("probe", "probe")):
        expected = _object(expected_files.get(key), f"ODP preparation.candidate.files.{key}")
        bound = _object(binder_candidate.get(binder_key), f"ODP source binder.candidate.{binder_key}")
        if expected.get("bytes") != bound.get("bytes") or expected.get("sha256") != bound.get("sha256"):
            raise RunFailure(f"ODP preparation {key} identity differs from the source binder")
    target = _safe_path(Path(str(value.get("target_dir"))), "ODP candidate target")
    build_output = _safe_path(Path(str(value.get("build_output"))), "ODP build output")
    tmpdir = _safe_path(Path(str(value.get("tmpdir"))), "ODP build TMPDIR", must_exist=True)
    if _under(target, Path(candidate["root"])) or _under(build_output, Path(candidate["root"])) or _under(tmpdir, Path(candidate["root"])):
        raise RunFailure("ODP preparation build paths must be outside the candidate source")
    if any(_under(left, right) or _under(right, left) for left, right in (
        (target, build_output), (target, tmpdir), (build_output, tmpdir)
    )):
        raise RunFailure("ODP preparation build paths are not pairwise disjoint")
    spec = _object(binder_value.get("build"), "ODP source binder.build")
    expected_argv = [
        item.replace("{candidate_root}", str(candidate["root"])).replace("{target_dir}", str(target))
        for item in spec["argv_template"]
    ]
    command = _object(value.get("command"), "ODP preparation.command")
    if command.get("argv") != expected_argv:
        raise RunFailure("ODP preparation command differs from the source binder composition")
    expected_environment = {
        key: item.replace("{target_dir}", str(target)).replace("{tmpdir}", str(tmpdir))
        for key, item in spec["environment"].items()
    }
    if command.get("environment_overrides") != expected_environment:
        raise RunFailure("ODP preparation environment differs from the source binder composition")
    runtime_path = _safe_path(Path(str(candidate["root"])) / "tools/odp_native_qualification.py", "ODP runtime source", must_exist=True)
    if value.get("runtime_source") != _sha(runtime_path):
        raise RunFailure("ODP runtime source changed after preparation")
    return value


def build(preparation_path: Path) -> dict[str, Any]:
    preparation_path = _safe_path(preparation_path.absolute(), "ODP source preparation receipt", must_exist=True)
    preparation = _load_preparation(preparation_path)
    candidate = preparation["candidate"]
    root = Path(candidate["root"])
    expected_files = candidate["files"]
    before = _source_state(root, CANDIDATE_REF, CANDIDATE_TREE, expected_files)
    target = _safe_path(Path(preparation["target_dir"]), "ODP candidate target")
    build_output = _safe_path(Path(preparation["build_output"]), "ODP build output")
    if target.exists() or build_output.exists():
        raise RunFailure("ODP bounded build requires fresh target/output paths")
    target.mkdir(parents=True, exist_ok=False)
    build_output.mkdir(parents=True, exist_ok=False)
    logs = build_output / "logs"
    logs.mkdir()
    command = preparation["command"]
    guard_before = _quiet_guard(preparation["max_load"], "before-candidate-build")
    event = _stream_command(
        list(command["argv"]), root, logs, "candidate-normal", target, preparation["cpu"],
        timeout_s=preparation["timeout_s"], environment=command["environment"]["variables"], max_load=preparation["max_load"],
    )
    artifact_path = Path(preparation["artifact"]["path"])
    if artifact_path.is_symlink() or not artifact_path.is_file() or not os.access(artifact_path, os.X_OK):
        raise RunFailure("ODP bounded build did not produce an executable candidate")
    artifact = _identity(artifact_path)
    after = _source_state(root, CANDIDATE_REF, CANDIDATE_TREE, expected_files)
    event["environment"] = command["environment"]
    event["guard_before"] = guard_before
    event["guard_after"] = _quiet_guard(preparation["max_load"], "after-candidate-build")
    event["artifact"] = artifact
    event["source_before"] = before
    event["source_after"] = after
    receipt = {
        "schema": BUILD_SCHEMA, "builder_revision": BUILDER_REVISION, "status": "passed",
        "preparation": {"path": str(preparation_path), **_sha(preparation_path)},
        "candidate": after, "fixture": preparation["fixture"], "event": event,
        "artifact": artifact,
        "claim_scope": "bounded C2a candidate build and observed executable identity only; no correctness, timing, or acceptance claim",
    }
    receipt_path = build_output / "build-receipt.json"
    with receipt_path.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())
    receipt["receipt"] = {"path": str(receipt_path), **_sha(receipt_path)}
    return receipt


def _validate_log(value: Any, label: str, log_dir: Path) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"path", "bytes", "sha256"}, label)
    path = _safe_path(Path(str(item["path"])), label, must_exist=True)
    if path.parent != log_dir:
        raise RunFailure(f"{label} is outside the fresh ODP build log directory")
    actual = _sha(path)
    expected = {"path": str(path), "bytes": item["bytes"], "sha256": item["sha256"]}
    if actual != expected:
        raise RunFailure(f"{label} changed after the bounded build")
    return actual


def _validate_build_event(event: Mapping[str, Any], preparation: Mapping[str, Any]) -> None:
    required = {
        "label", "command", "cwd", "returncode", "wall_ns", "target", "stdout", "stderr",
        "monitor", "environment", "guard_before", "guard_after", "artifact", "source_before", "source_after",
    }
    _exact_keys(event, required, "ODP build receipt.event")
    if event["label"] != "candidate-normal" or event["returncode"] != 0:
        raise RunFailure("ODP build receipt process event did not pass")
    if type(event["wall_ns"]) is not int or event["wall_ns"] <= 0:
        raise RunFailure("ODP build receipt process duration is malformed")
    if event["command"] != preparation["command"]["argv"] or event["environment"] != preparation["command"]["environment"]:
        raise RunFailure("ODP build receipt process command/environment differs")
    if event["cwd"] != preparation["candidate"]["root"]:
        raise RunFailure("ODP build receipt process cwd differs")
    target = _object(event["target"], "ODP build receipt.event.target")
    _exact_keys(target, {"path", "bytes_before", "bytes_after", "hard_limit_bytes"}, "ODP build receipt.event.target")
    if (target["path"] != preparation["target_dir"]
            or type(target["bytes_before"]) is not int or target["bytes_before"] < 0
            or type(target["bytes_after"]) is not int or target["bytes_after"] < 0
            or target["hard_limit_bytes"] != TARGET_HARD_LIMIT_BYTES):
        raise RunFailure("ODP build receipt target accounting differs")
    log_dir = Path(preparation["build_output"]) / "logs"
    _validate_log(event["stdout"], "ODP build receipt.event.stdout", log_dir)
    _validate_log(event["stderr"], "ODP build receipt.event.stderr", log_dir)
    monitor = _object(event["monitor"], "ODP build receipt.event.monitor")
    _exact_keys(monitor, {"samples", "memory_limit_bytes", "max_output_bytes", "max_load", "owned_processes"}, "ODP build receipt.event.monitor")
    if (monitor["memory_limit_bytes"] != MEMORY_LIMIT_BYTES
            or monitor["max_output_bytes"] != MAX_OUTPUT_BYTES
            or monitor["max_load"] != preparation["max_load"]
            or monitor["owned_processes"] != []
            or not isinstance(monitor["samples"], list) or not monitor["samples"]):
        raise RunFailure("ODP build receipt live process monitor is incomplete")
    for sample in monitor["samples"]:
        row = _object(sample, "ODP build receipt.event.monitor.sample")
        _exact_keys(row, {"checked_unix_ns", "output_bytes", "rss_bytes", "load_1m", "outside_processes", "owned_processes"}, "ODP monitor sample")
        if (type(row["checked_unix_ns"]) is not int or row["checked_unix_ns"] <= 0
                or type(row["output_bytes"]) is not int or row["output_bytes"] < 0 or row["output_bytes"] > MAX_OUTPUT_BYTES
                or type(row["rss_bytes"]) is not int or row["rss_bytes"] < 0 or row["rss_bytes"] > MEMORY_LIMIT_BYTES
                or type(row["load_1m"]) not in (int, float) or row["load_1m"] < 0
                or row["load_1m"] > preparation["max_load"]
                or row["outside_processes"] != [] or not isinstance(row["owned_processes"], list)):
            raise RunFailure("ODP monitor sample exceeds the bounded process policy")
        for process in row["owned_processes"]:
            process_value = _object(process, "ODP monitor owned process")
            _exact_keys(process_value, {"pid", "start_time", "process_group"}, "ODP monitor owned process")
            if any(type(process_value[key]) is not int or process_value[key] <= 0
                   for key in ("pid", "start_time", "process_group")):
                raise RunFailure("ODP monitor owned process identity is malformed")
    for label in ("guard_before", "guard_after"):
        guard = _object(event[label], f"ODP build receipt.event.{label}")
        _exact_keys(guard, {"phase", "load_1m", "max_load", "heavy_processes", "checked_unix_ns"}, f"ODP {label}")
        if (type(guard["checked_unix_ns"]) is not int or guard["checked_unix_ns"] <= 0
                or guard["max_load"] != preparation["max_load"]
                or guard["heavy_processes"] != []):
            raise RunFailure("ODP quiet guard is not a valid clean receipt")


def seal(preparation_path: Path, build_receipt_path: Path, input_manifest: Path, output: Path) -> dict[str, Any]:
    preparation = _load_preparation(preparation_path)
    build_path = _safe_path(build_receipt_path.absolute(), "ODP build receipt", must_exist=True)
    build_receipt = _load_json(build_path)
    if build_receipt.get("schema") != BUILD_SCHEMA or build_receipt.get("status") != "passed":
        raise RunFailure("ODP build receipt is not a passed bounded-build receipt")
    preparation_identity = _object(build_receipt.get("preparation"), "ODP build receipt.preparation")
    if preparation_identity != _sha(Path(preparation_path).absolute()):
        raise RunFailure("ODP build receipt is bound to a different preparation")
    if build_receipt.get("candidate") != preparation.get("candidate"):
        raise RunFailure("ODP build receipt source identity differs from the prepared candidate")
    event = _object(build_receipt.get("event"), "ODP build receipt.event")
    _validate_build_event(event, preparation)
    if event.get("source_before") != preparation.get("candidate") or event.get("source_after") != preparation.get("candidate"):
        raise RunFailure("ODP build receipt process event is not source-bound")
    if event.get("command") != preparation["command"]["argv"] or event.get("environment") != preparation["command"]["environment"]:
        raise RunFailure("ODP build receipt process event command/environment differs")
    if event.get("cwd") != preparation["candidate"]["root"]:
        raise RunFailure("ODP build receipt process event cwd differs")
    artifact = _object(build_receipt.get("artifact"), "ODP build receipt.artifact")
    expected = Path(preparation["artifact"]["path"])
    if artifact.get("path") != str(expected) or _identity(expected) != artifact:
        raise RunFailure("ODP observed executable differs from the prepared artifact path/identity")
    base = _load_json(_safe_path(input_manifest.absolute(), "ODP input manifest", must_exist=True))
    if base.get("candidate", {}).get("ref") != CANDIDATE_REF or base.get("candidate", {}).get("root") != preparation["candidate"]["root"]:
        raise RunFailure("ODP input candidate differs from the bounded source build")
    base["candidate"]["binary"] = artifact["path"]
    base["candidate"]["target_dir"] = preparation["target_dir"]
    base["candidate_build_receipt"] = {"path": str(build_path), **_sha(build_path)}
    base["candidate_source_build"] = {
        "preparation": {"path": str(Path(preparation_path).absolute()), **_sha(Path(preparation_path))},
        "source": build_receipt["candidate"], "artifact": artifact,
    }
    output = _safe_path(output.absolute(), "sealed ODP input manifest")
    if _under(output, Path(preparation["candidate"]["root"])) or any(
        _under(output, Path(preparation[key])) or _under(Path(preparation[key]), output)
        for key in ("target_dir", "build_output", "tmpdir")
    ):
        raise RunFailure("sealed ODP input manifest must be outside source and build paths")
    if output.exists() or output.is_symlink():
        raise RunFailure(f"refusing to overwrite sealed ODP input manifest: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(base, indent=2, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())
    return base


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--binder", type=Path)
    parser.add_argument("--preparation", type=Path)
    parser.add_argument("--build-receipt", type=Path)
    parser.add_argument("--input", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--target-dir", type=Path)
    parser.add_argument("--build-output", type=Path)
    parser.add_argument("--tmpdir", type=Path)
    parser.add_argument("--cpu", type=int, default=3)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--prepare", action="store_true")
    mode.add_argument("--build", action="store_true")
    mode.add_argument("--seal", action="store_true")
    args = parser.parse_args(argv)
    try:
        if args.prepare:
            required = (args.binder, args.target_dir, args.build_output, args.tmpdir)
            if args.output is None or any(value is None for value in required):
                parser.error("--prepare requires --binder, --target-dir, --build-output, and --tmpdir")
            prepare(args.binder, target_dir=args.target_dir, build_output=args.build_output, tmpdir=args.tmpdir, cpu=args.cpu, output=args.output)
        elif args.build:
            if args.preparation is None:
                parser.error("--build requires --preparation")
            build(args.preparation)
        else:
            if args.preparation is None or args.build_receipt is None or args.input is None or args.output is None:
                parser.error("--seal requires --preparation, --build-receipt, and --input")
            seal(args.preparation, args.build_receipt, args.input, args.output)
    except (RunFailure, OSError, ValueError) as error:
        print(f"ODP candidate build gate refused: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
