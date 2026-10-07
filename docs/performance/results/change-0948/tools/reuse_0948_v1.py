#!/usr/bin/env python3
"""Shared 0948 provenance validators, including the retained 0945 reuse path."""

from __future__ import annotations

import hashlib
import json
import os
import re
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent.parent
LEGACY_SOURCE_REF = PACKET / "references" / "0945-source-equivalence"
SCHEMA = "litchi.performance.0945.native-baseline-reuse-v2"
COMMIT = "e037befa6ce52b3988c8ae2032ae652196f2a09d"
FIELDS = ("path", "bytes", "sha256", "git_oid", "git_mode")
MAX_JSON_BYTES = 64 * 1024 * 1024
SOURCE_BYTES = {"baseline": 259872266}
SOURCE_MANIFEST_FILES = {"baseline": 10788, "candidate": 10788}
# The default is the current 0948 build namespace.  A frozen plan may carry a
# different historical revision (for example, the retained 0945 v2 packet),
# and _build_binding below derives the effective revision from that plan.
BUILD_REVISIONS = {"baseline": "v1", "candidate": "v3"}
BUILD_BINDING_IDS = {"baseline": "baseline1", "candidate": "candidate2"}
UNSET_ENV = ["CARGO_BUILD_RUSTC_WRAPPER", "CARGO_ENCODED_RUSTFLAGS",
             "RUSTC_WRAPPER", "RUSTFLAGS"]


class ReuseReceiptError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReuseReceiptError(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON artifact {path}")
    require(path.stat().st_size <= MAX_JSON_BYTES, f"JSON artifact too large: {path}")

    def reject(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result

    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ReuseReceiptError(f"cannot read JSON {path}: {error}") from error


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and isinstance(value.get("path"), str),
            f"{label} is not a descriptor")
    raw = Path(value["path"])
    require(not raw.is_absolute() and raw.as_posix() == value["path"]
            and ".." not in raw.parts, f"{label} path is not packet-relative")
    unresolved = PACKET / raw
    require(not unresolved.is_symlink(), f"{label} target must not be a symlink: {unresolved}")
    path = unresolved.resolve()
    require(path.is_file() and not path.is_symlink(), f"{label} target is missing: {path}")
    require(path.is_relative_to(PACKET.resolve()), f"{label} target escaped packet: {path}")
    mode = path.stat().st_mode & 0o777
    actual = {"path": raw.as_posix(), "bytes": path.stat().st_size,
              "mode": mode, "executable": bool(mode & 0o111),
              "sha256": sha256(path)}
    for key in ("bytes", "mode", "executable", "sha256"):
        if key in value:
            require(value[key] == actual[key], f"{label}.{key} does not match {path}")
    return actual


def external_descriptor(value: Any, expected: Path, label: str) -> dict[str, Any]:
    """Check an explicitly allowed host artifact such as the Cargo shim."""
    require(isinstance(value, dict) and value.get("path") == str(expected),
            f"{label} path is not the frozen host path")
    require(expected.is_file(),
            f"{label} host artifact is missing: {expected}")
    resolved = expected.resolve()
    require(resolved == Path("/home/zhuhe/.cargo/bin/rustup")
            and resolved.is_file() and os.access(resolved, os.X_OK),
            f"{label} resolved executable is not the frozen rustup shim")
    mode = resolved.stat().st_mode & 0o777
    actual = {"path": str(expected), "resolved_path": str(resolved),
              "bytes": resolved.stat().st_size, "mode": mode,
              "executable": bool(mode & 0o111), "sha256": sha256(resolved)}
    require(value.get("resolved_path") == actual["resolved_path"],
            f"{label}.resolved_path changed")
    for key in ("bytes", "mode", "executable", "sha256"):
        if key in value:
            require(value[key] == actual[key], f"{label}.{key} does not match {expected}")
    return actual


def _plan_descriptor(plan: dict[str, Any], key: str, label: str) -> dict[str, Any]:
    source = plan.get("source")
    require(isinstance(source, dict), "frozen plan source binding is missing")
    return descriptor(source.get(key), label)


def _source_manifest_identity(path: Path, role: str) -> dict[str, Any]:
    raw, files = manifest_files(path, f"{role} build source manifest")
    require(raw.get("file_count") == SOURCE_MANIFEST_FILES[role]
            and isinstance(raw.get("bytes"), int) and raw["bytes"] > 0
            and len(files) == SOURCE_MANIFEST_FILES[role],
            f"{role} build source manifest census changed")
    if role == "baseline":
        require(raw.get("commit") == COMMIT, "baseline build source commit changed")
    else:
        require(raw.get("base_commit") == COMMIT
                and isinstance(raw.get("changes"), list)
                and len(raw["changes"]) == 2,
                "candidate build source delta changed")
    return raw


def _build_binding(plan: dict[str, Any] | None, role: str, kind: str) -> tuple[str, str]:
    """Return the frozen revision/id for a role and reject silent rebinding.

    The underlying baseline build remains the v1 build even though its native
    reuse receipt is the portable v2 receipt.  Candidate builders are the
    selected revision's wrappers and therefore have separate target/log
    namespaces.  The selected candidate revision is read from the frozen plan,
    rather than being embedded in the generic terminal validator.
    """
    require(role in BUILD_REVISIONS and kind in {"native", "allocator"},
            "invalid build role/kind")
    revisions = dict(BUILD_REVISIONS)
    binding_ids = dict(BUILD_BINDING_IDS)
    revision = revisions[role]
    binding_id = binding_ids[role]
    if plan is not None:
        require(isinstance(plan, dict), "frozen plan is not an object")
        plan_revisions = plan.get("build_revisions")
        plan_binding_ids = plan.get("build_binding_ids")
        roles = plan.get("roles")
        # Small retained-receipt controls intentionally pass only the source
        # subset of a plan.  Whenever a plan supplies build bindings, use those
        # values for the generic validator and reject malformed or path-unsafe
        # revision names.  This permits a new final build namespace (v3) while
        # preserving strict validation of an explicitly retained old namespace.
        if plan_revisions is not None or plan_binding_ids is not None:
            require(isinstance(plan_revisions, dict)
                    and set(plan_revisions) == set(BUILD_REVISIONS)
                    and isinstance(plan_binding_ids, dict)
                    and set(plan_binding_ids) == set(BUILD_BINDING_IDS),
                    "frozen plan build revision bindings are incomplete")
            for plan_role, plan_revision in plan_revisions.items():
                require(isinstance(plan_revision, str)
                        and re.fullmatch(r"v[0-9]+", plan_revision) is not None,
                        f"frozen plan {plan_role} build revision is unsafe")
            for plan_role, plan_binding_id in plan_binding_ids.items():
                require(isinstance(plan_binding_id, str)
                        and re.fullmatch(r"[A-Za-z0-9_.-]+", plan_binding_id) is not None,
                        f"frozen plan {plan_role} binding id is unsafe")
            revisions = dict(plan_revisions)
            binding_ids = dict(plan_binding_ids)
        revision = revisions[role]
        binding_id = binding_ids[role]
        # If role descriptors are present, bind each descriptor to the same
        # plan-derived revision/id.  A draft source-only plan may omit roles;
        # capture/reduce role checks enforce them before measurement.
        if roles is not None:
            require(isinstance(roles, dict) and isinstance(roles.get(role), dict),
                    f"frozen plan lacks {role} role binding")
            value = roles[role].get(kind)
            require(isinstance(value, dict)
                    and value.get("build_revision") == revision
                    and value.get("build_binding_id") == binding_id,
                    f"frozen plan {role}.{kind} build binding changed")
    return revision, binding_id


def validate_build_terminal(raw: dict[str, Any], role: str, kind: str,
                            plan: dict[str, Any],
                            binary: dict[str, Any] | None = None) -> None:
    """Validate a current 0945 builder terminal, including nested provenance.

    The capture and reducer both call this function.  It intentionally checks
    the builder's nested ``source`` object and exact Cargo contract instead of
    accepting a top-level exit code while silently dropping missing bindings.
    """
    require(role in {"baseline", "candidate"} and kind in {"native", "allocator"},
            "invalid build role/kind")
    revision, _ = _build_binding(plan, role, kind)
    prefix = f"{role}-{kind}"
    expected_schema = f"litchi.performance.0945.{prefix}-build-{revision}"
    require(raw.get("schema") == expected_schema and raw.get("status") == "terminal",
            f"{prefix} build terminal schema/status changed")
    require(raw.get("receipt_kind") == "terminal", f"{prefix} is not a terminal receipt")
    require(raw.get("returncode") == 0 and raw.get("timed_out") is False
            and raw.get("terminal_before_binary_parse") is True,
            f"{prefix} build did not terminate durably")
    require(raw.get("process_group") is True and raw.get("start_new_session") is True,
            f"{prefix} build process-group contract changed")
    cleanup = raw.get("cleanup")
    require(isinstance(cleanup, dict) and cleanup.get("process_group_clear") is True
            and cleanup.get("reaped") is True,
            f"{prefix} build process group was not reaped and cleared")

    expected_binary = ("litchi-perf-baseline-alloc" if kind == "allocator"
                       else "litchi-perf-baseline")
    expected_target = f"target-{prefix}-{revision}"
    source_plan = plan.get("source")
    require(isinstance(source_plan, dict), "frozen plan source binding is missing")
    expected_source = source_plan.get(f"{role}_root")
    require(isinstance(expected_source, str) and expected_source
            and not Path(expected_source).is_absolute()
            and Path(expected_source).as_posix() == expected_source
            and ".." not in Path(expected_source).parts,
            f"{prefix} source root is not a normalized frozen path")
    expected_manifest_descriptor = _plan_descriptor(
        plan, f"{role}_manifest", f"{role} plan source manifest")
    expected_manifest = expected_manifest_descriptor["path"]
    source = raw.get("source")
    require(isinstance(source, dict), f"{prefix} nested source binding is missing")
    require(source.get("source_root") == expected_source
            and source.get("source_commit") == COMMIT
            and source.get("file_count") == SOURCE_MANIFEST_FILES[role]
            and isinstance(source.get("bytes"), int) and source["bytes"] > 0,
            f"{prefix} nested source census/commit changed")
    manifest_desc = descriptor(source.get("manifest"), f"{prefix} source manifest")
    require(manifest_desc == expected_manifest_descriptor
            and manifest_desc["path"] == expected_manifest,
            f"{prefix} source manifest is not bound to the frozen role manifest")
    manifest_raw = _source_manifest_identity(PACKET / manifest_desc["path"], role)
    require(source["bytes"] == manifest_raw["bytes"],
            f"{prefix} source byte census does not match its manifest")
    require(source.get("manifest_size") == manifest_desc["bytes"],
            f"{prefix} source manifest size changed")

    expected_input_paths = {
        "cargo_manifest": f"{expected_source}/tools/perf-baseline/Cargo.toml",
        "cargo_lock": f"{expected_source}/tools/perf-baseline/Cargo.lock",
        "cargo_config": f"{expected_source}/.cargo/config.toml",
    }
    for field, expected_path in expected_input_paths.items():
        value = descriptor(source.get(field), f"{prefix} source {field}")
        require(value["path"] == expected_path,
                f"{prefix} source {field} path changed")
    if role == "baseline":
        export = descriptor(source.get("source_export"), f"{prefix} source export")
        expected_export = _plan_descriptor(plan, "source_export", "plan source export")
        require(export == expected_export and export["path"] == "source-export.json",
                f"{prefix} source export is not bound to the frozen export")
    else:
        require("source_export" not in source,
                "candidate build unexpectedly carries a baseline source export")
        freeze = descriptor(source.get("freeze_receipt"), f"{prefix} freeze receipt")
        require(freeze["path"].startswith("receipts/candidate-freeze-"),
                f"{prefix} freeze receipt path changed")
        expected_freeze = source_plan.get("candidate_freeze_receipt")
        if expected_freeze is not None:
            require(freeze == descriptor(expected_freeze, "plan candidate freeze receipt"),
                    f"{prefix} freeze receipt is not bound to the frozen plan")
        require(source.get("changes") == manifest_raw.get("changes"),
                f"{prefix} source delta is not bound to its manifest")

    expected_argv = [
        "/home/zhuhe/.cargo/bin/cargo", "build", "--release", "--locked", "--offline",
        "--manifest-path", f"/home/zhuhe/litchi-performance-0945/{expected_source}/tools/perf-baseline/Cargo.toml",
        "--bin", expected_binary,
    ]
    if kind == "allocator":
        expected_argv += ["--features", "allocator-metrics"]
    expected_argv += ["-j2"]
    require(raw.get("argv") == expected_argv, f"{prefix} Cargo argv/feature flags changed")
    require(raw.get("cwd") == f"/home/zhuhe/litchi-performance-0945/{expected_source}"
            and raw.get("cwd_relative") == expected_source,
            f"{prefix} build cwd changed")
    require(raw.get("deadline_seconds") == 1800,
            f"{prefix} build deadline changed")
    expected_env = {
        "PATH": "/home/zhuhe/.cargo/bin:/usr/bin:/bin", "HOME": "/home/zhuhe",
        "CARGO_HOME": "/home/zhuhe/.cargo", "RUSTUP_HOME": "/home/zhuhe/.rustup",
        "CARGO_TARGET_DIR": f"/home/zhuhe/litchi-performance-0945/{expected_target}",
        "CARGO_NET_OFFLINE": "true", "CARGO_INCREMENTAL": "0", "CARGO_BUILD_JOBS": "2",
        "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1", "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "CARGO_PROFILE_RELEASE_LTO": "fat", "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3",
        "RUSTUP_TOOLCHAIN": "1.95.0", "LANG": "C", "LC_ALL": "C", "TZ": "UTC",
        "RUST_BACKTRACE": "0",
        "TMPDIR": f"/home/zhuhe/litchi-performance-0945/tmp/{prefix}-{revision}",
        "TEMP": f"/home/zhuhe/litchi-performance-0945/tmp/{prefix}-{revision}",
        "TMP": f"/home/zhuhe/litchi-performance-0945/tmp/{prefix}-{revision}",
    }
    require(raw.get("environment") == dict(sorted(expected_env.items())),
            f"{prefix} build environment/profile flags changed")
    require(raw.get("unset_environment") == UNSET_ENV,
            f"{prefix} unset environment changed")
    require(raw.get("target_relative") == expected_target
            and raw.get("binary_relative") == f"{expected_target}/release/{expected_binary}",
            f"{prefix} target/binary binding changed")

    runner = descriptor(raw.get("runner"), f"{prefix} runner")
    expected_runner = f"build-tools/build_candidate_{kind}_v2.py"
    require(runner["path"] == expected_runner,
            f"{prefix} runner path changed")
    cargo = external_descriptor(raw.get("cargo"), Path("/home/zhuhe/.cargo/bin/cargo"),
                                f"{prefix} cargo")
    require(cargo["executable"], f"{prefix} cargo is not executable")
    for field, expected_path in (("stdout", f"builds/{prefix}-{revision}/build.stdout.log"),
                                 ("stderr", f"builds/{prefix}-{revision}/build.stderr.log")):
        value = descriptor(raw.get(field), f"{prefix} {field} log")
        require(value["path"] == expected_path, f"{prefix} {field} log path changed")

    binary_record = raw.get("binary")
    require(isinstance(binary_record, dict) and binary_record.get("exists") is True
            and binary_record.get("path") == f"{expected_target}/release/{expected_binary}",
            f"{prefix} binary terminal binding changed")
    terminal_binary = descriptor({"path": binary_record["path"]},
                                 f"{prefix} terminal binary")
    if binary is not None:
        require(binary.get("sha256") and binary.get("bytes")
                and terminal_binary["sha256"] == binary["sha256"]
                and terminal_binary["bytes"] == binary["bytes"],
                f"{prefix} role binary descriptor is incomplete")


def manifest_files(path: Path, label: str) -> tuple[dict[str, Any], dict[str, tuple[Any, ...]]]:
    raw = read_json(path)
    files = raw.get("files") if isinstance(raw, dict) else None
    require(isinstance(files, list) and files, f"{label} files are missing")
    output: dict[str, tuple[Any, ...]] = {}
    for item in files:
        require(isinstance(item, dict) and isinstance(item.get("path"), str),
                f"{label} has malformed file entry")
        name = item["path"]
        require(name not in output and not Path(name).is_absolute()
                and Path(name).as_posix() == name and ".." not in Path(name).parts,
                f"{label} has unsafe or duplicate path {name!r}")
        output[name] = tuple(item.get(key) for key in FIELDS)
    require(output, f"{label} is empty")
    return raw, output


def _build_contract(start: dict[str, Any]) -> dict[str, Any]:
    expected_argv = [
        "/home/zhuhe/.cargo/bin/cargo", "build", "--release", "--locked", "--offline",
        "--manifest-path", "/home/zhuhe/litchi-performance-0944/source-build/tools/perf-baseline/Cargo.toml",
        "--bin", "litchi-perf-baseline", "-j", "2",
    ]
    require(start.get("argv") == expected_argv, "original native build argv changed")
    env = start.get("env")
    require(isinstance(env, dict), "original native build environment is missing")
    exact = {
        "CARGO_BUILD_JOBS": "2", "CARGO_HOME": "/home/zhuhe/.cargo",
        "CARGO_INCREMENTAL": "0", "CARGO_NET_OFFLINE": "true",
        "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1", "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "CARGO_PROFILE_RELEASE_LTO": "fat", "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3",
        "HOME": "/home/zhuhe", "LANG": "C", "LC_ALL": "C",
        "PATH": "/home/zhuhe/.cargo/bin:/usr/bin:/bin", "RUSTUP_HOME": "/home/zhuhe/.rustup",
        "RUSTUP_TOOLCHAIN": "1.95.0", "RUST_BACKTRACE": "0", "TZ": "UTC",
        "CARGO_TARGET_DIR": "/home/zhuhe/litchi-performance-0944/target-native",
        "TEMP": "/home/zhuhe/litchi-performance-0944/tmp/build-native",
        "TMP": "/home/zhuhe/litchi-performance-0944/tmp/build-native",
        "TMPDIR": "/home/zhuhe/litchi-performance-0944/tmp/build-native",
    }
    require(env == exact, "original native build environment or flags changed")
    unset = ["RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_WRAPPER",
             "CARGO_BUILD_RUSTC_WRAPPER"]
    require(start.get("unset_env") == unset, "original native unset environment changed")
    return {"argv": expected_argv, "environment": exact, "unset_environment": unset,
            "cwd": "/home/zhuhe/litchi-performance-0944/source-build",
            "deadline_seconds": 900,
            "profile": {"release": True, "locked": True, "offline": True,
                         "jobs": 2, "allocator_metrics": False, "toolchain": "1.95.0"}}


def validate_receipt(path: Path, *, binary: dict[str, Any] | None = None,
                     plan: dict[str, Any] | None = None) -> dict[str, Any]:
    """Validate all nested descriptors and build/source contracts."""
    path = Path(path)
    if not path.is_absolute():
        path = PACKET / path
    path = path.resolve()
    top = descriptor({"path": path.relative_to(PACKET.resolve()).as_posix()}, "reuse receipt")
    raw = read_json(path)
    require(raw.get("schema") == SCHEMA and raw.get("status") == "passed"
            and raw.get("reuse_kind") == "native-baseline", "reuse receipt identity changed")
    bound_binary = descriptor(raw.get("binary"), "reuse binary")
    if binary is not None:
        require(bound_binary["sha256"] == binary.get("sha256")
                and bound_binary["bytes"] == binary.get("bytes"),
                "reuse binary differs from role binary")
    current = raw.get("current_source")
    require(isinstance(current, dict) and current.get("commit") == COMMIT
            and current.get("file_count") == 10788 and current.get("root") == "baseline-source",
            "current source binding changed")
    current_manifest_desc = descriptor(current.get("manifest"), "current source manifest")
    current_export_desc = descriptor(current.get("export"), "current source export")
    current_manifest_raw, current_files = manifest_files(
        PACKET / current_manifest_desc["path"], "current source manifest")
    require(current_manifest_raw.get("commit") == COMMIT
            and current_manifest_raw.get("file_count") == len(current_files),
            "current source manifest metadata changed")
    current_export = read_json(PACKET / current_export_desc["path"])
    require(current_export.get("schema") == "litchi.performance.0945.source-export-v1"
            and current_export.get("source_commit") == COMMIT
            and current_export.get("file_count") == 10788
            and current_export.get("source_bytes") == current_manifest_raw.get("bytes")
            and current_export.get("manifest", {}).get("sha256") == current_manifest_desc["sha256"],
            "current source export does not bind current manifest")
    equivalence = raw.get("source_equivalence")
    require(isinstance(equivalence, dict)
            and equivalence.get("file_count") == 10788
            and equivalence.get("matched_path_count") == 10788
            and equivalence.get("compared_fields") == list(FIELDS)
            and equivalence.get("mismatches") == {field: 0 for field in FIELDS},
            "source equivalence summary is not exact")
    require(descriptor(equivalence.get("current_manifest"), "equivalence current manifest") == current_manifest_desc,
            "equivalence current manifest binding changed")
    original_manifest_desc = descriptor(equivalence.get("original_manifest"), "original source manifest")
    original_manifest_raw, original_files = manifest_files(
        PACKET / original_manifest_desc["path"], "original source manifest")
    require(original_manifest_raw.get("origin_commit") == "933ee45cbb5d2ef9dc29583f7ec890ad234e04fb"
            and original_manifest_raw.get("file_count") == 10788,
            "original source manifest metadata changed")
    require(original_files == current_files, "original/current source closure differs")
    original_export_desc = descriptor(equivalence.get("original_export"), "original export receipt")
    original_export = read_json(PACKET / original_export_desc["path"])
    require(original_export.get("schema") == "litchi-xlsx-profile-source-export-v1"
            and original_export.get("manifest_sha256") == original_manifest_raw.get("full_export_manifest_sha256"),
            "original export is not bound to original closure")
    selection_desc = descriptor(equivalence.get("closure_selection"), "closure selection")
    selection = read_json(PACKET / selection_desc["path"])
    require(selection.get("schema") == "litchi-xlsx-profile-source-closure-selection-v1"
            and selection.get("tracked_3rdparty_count") == 0,
            "closure selection changed")
    original_build = raw.get("original_build")
    require(isinstance(original_build, dict), "original build binding is missing")
    start_desc = descriptor(original_build.get("start"), "original build start")
    terminal_desc = descriptor(original_build.get("terminal"), "original build terminal")
    start = read_json(PACKET / start_desc["path"])
    terminal = read_json(PACKET / terminal_desc["path"])
    require(start.get("schema") == "litchi-xlsx-profile-build-start-v1"
            and terminal.get("schema") == "litchi-xlsx-profile-build-terminal-v1",
            "original build receipt schemas changed")
    contract = _build_contract(start)
    require(original_build.get("contract") == contract, "portable build contract differs from original start")
    require(original_build.get("input_binding") == start.get("inputs"),
            "portable input binding differs from original start")
    require(original_build.get("terminal_status") == {
        "returncode": 0, "timed_out": False, "terminal_before_parse": True,
        "binary": terminal.get("binary")}, "portable terminal status differs")
    require(terminal.get("returncode") == 0 and terminal.get("timed_out") is False
            and terminal.get("terminal_before_parse") is True,
            "original build did not terminate successfully")
    require(terminal.get("binary", {}).get("sha256") == bound_binary["sha256"]
            and terminal.get("binary", {}).get("bytes") == bound_binary["bytes"],
            "original terminal binary differs")
    inputs = start.get("inputs")
    require(isinstance(inputs, dict)
            and inputs.get("source_manifest") == "source-build-manifest.json"
            and inputs.get("source_manifest_sha256") == original_manifest_desc["sha256"],
            "original source manifest input binding differs")
    if plan is not None:
        _build_binding(plan, "baseline", "native")
        source = plan.get("source", {})
        plan_manifest = descriptor(source.get("baseline_manifest"), "plan baseline manifest")
        plan_export = descriptor(source.get("source_export"), "plan source export")
        require(plan_manifest == current_manifest_desc and plan_export == current_export_desc,
                "reuse receipt is not bound to frozen plan source descriptors")
    portable = raw.get("portable_binding")
    require(portable == {
        "absolute_paths_preserved_only_inside_original_command_environment": True,
        "all_paths_relative_to_packet": True,
        "original_receipts_copied_from": "0944 build/native and source closure receipts",
    }, "portable binding policy changed")
    return top


def _strict_descriptor(value: Any, label: str) -> dict[str, Any]:
    """Require every digest field, rather than accepting a path-only binding."""
    actual = descriptor(value, label)
    require(all(key in value for key in ("bytes", "mode", "executable", "sha256")),
            f"{label} omits a digest or mode field")
    for key in ("bytes", "mode", "executable", "sha256"):
        require(value[key] == actual[key], f"{label}.{key} does not match its artifact")
    return actual


def _packet_file(value: Any, label: str, expected: str | None = None) -> Path:
    require(isinstance(value, str), f"{label} path is missing")
    raw = Path(value)
    require(not raw.is_absolute() and raw.as_posix() == value and ".." not in raw.parts,
            f"{label} path is not packet-relative")
    if expected is not None:
        require(value == expected, f"{label} path changed")
    path = PACKET / raw
    require(path.is_file() and not path.is_symlink(), f"{label} artifact is missing: {path}")
    return path


def _absolute_suffix(value: Any, suffix: str, label: str) -> Path:
    require(isinstance(value, str), f"{label} is missing")
    path = Path(value)
    require(path.is_absolute() and path.as_posix().endswith("/" + suffix),
            f"{label} does not have the frozen suffix /{suffix}")
    return path


def _common_receipt_fields(raw: dict[str, Any], prefix: str,
                           expected_root: str, expected_target: str,
                           expected_binary: str) -> None:
    """Check the fields emitted by the 0945 durable builder on every receipt."""
    require(raw.get("start_new_session") is True and raw.get("process_group") is True,
            f"{prefix} process-group launch contract changed")
    require(raw.get("terminal_before_binary_parse") is True,
            f"{prefix} terminal-before-parse custody changed")
    require(raw.get("cwd_relative") == expected_root,
            f"{prefix} cwd_relative changed")
    cwd = _absolute_suffix(raw.get("cwd"), expected_root, f"{prefix} cwd")
    # The recorded absolute cwd belongs to the original build host.  Replay
    # validates the packet-local root separately and never dereferences that
    # historical path.
    require((PACKET / expected_root).is_dir(),
            f"{prefix} packet-local cwd is not a directory: {PACKET / expected_root}")
    require(raw.get("target_relative") == expected_target
            and raw.get("binary_relative") == f"{expected_target}/release/{expected_binary}",
            f"{prefix} target/binary relative binding changed")
    expected_binary_record = raw.get("binary")
    require(isinstance(expected_binary_record, dict)
            and expected_binary_record.get("exists") is True
            and expected_binary_record.get("path") ==
            f"{expected_target}/release/{expected_binary}",
            f"{prefix} terminal binary binding changed")
    require(raw.get("launch_error") is None,
            f"{prefix} launch reported an error")
    require(isinstance(raw.get("pid"), int) and raw["pid"] > 0
            and isinstance(raw.get("process_group_id"), int)
            and raw["process_group_id"] > 0,
            f"{prefix} process identifiers are missing")


def validate_build_terminal_0948(raw: dict[str, Any], role: str, kind: str,
                                  plan: dict[str, Any],
                                  binary: dict[str, Any] | None = None) -> None:
    """Validate a 0948 terminal using the complete 0945 build custody contract.

    Absolute paths recorded by the builder are checked by stable packet-relative
    suffixes so a copied packet remains replayable.  Receipt paths, source
    descriptors, Cargo argv, environment, logs, and the start/launch/binary/result
    chain remain exact and create-only.
    """
    require(role in BUILD_REVISIONS and kind in {"native", "allocator"},
            "invalid 0948 build role/kind")
    revision, _ = _build_binding(plan, role, kind)
    prefix = f"{role}-{kind}"
    expected_schema = f"litchi.performance.0948.{prefix}-build-{revision}"
    require(raw.get("schema") == expected_schema and raw.get("status") == "terminal",
            f"{prefix} 0948 build terminal schema/status changed")
    require(raw.get("receipt_kind") == "terminal"
            and raw.get("returncode") == 0
            and raw.get("timed_out") is False
            and raw.get("terminal_before_binary_parse") is True,
            f"{prefix} build did not terminate durably")
    cleanup = raw.get("cleanup")
    require(isinstance(cleanup, dict)
            and cleanup.get("process_group_clear") is True
            and cleanup.get("reaped") is True,
            f"{prefix} build process group was not reaped and cleared")

    source_plan = plan.get("source")
    require(isinstance(source_plan, dict), "0948 plan source binding is missing")
    expected_root = source_plan.get(f"{role}_root")
    expected_commit = source_plan.get("source_commit")
    expected_count = source_plan.get("file_count")
    require(isinstance(expected_root, str) and expected_root
            and not Path(expected_root).is_absolute()
            and Path(expected_root).as_posix() == expected_root
            and ".." not in Path(expected_root).parts,
            f"{prefix} source root is not normalized")
    require(isinstance(expected_commit, str) and expected_commit
            and isinstance(expected_count, int) and expected_count > 0,
            "0948 source binding metadata is incomplete")

    source = raw.get("source")
    require(isinstance(source, dict)
            and source.get("source_root") == expected_root
            and source.get("source_commit") == expected_commit
            and source.get("file_count") == expected_count,
            f"{prefix} source census/commit changed")
    role_manifest_key = "baseline_manifest" if role == "baseline" else "candidate_manifest"
    expected_manifest = _strict_descriptor(source_plan.get(role_manifest_key),
                                           f"{prefix} plan source manifest")
    actual_manifest = _strict_descriptor(source.get("manifest"), f"{prefix} source manifest")
    require(actual_manifest == expected_manifest,
            f"{prefix} source manifest is not bound to the frozen plan")
    manifest_raw, manifest_entries = manifest_files(
        PACKET / actual_manifest["path"], f"{prefix} source manifest")
    manifest_file_count = manifest_raw.get("file_count", len(manifest_entries))
    manifest_bytes = manifest_raw.get("bytes", manifest_raw.get("source_bytes"))
    if manifest_bytes is None:
        manifest_bytes = sum(item[1] for item in manifest_entries.values()
                             if isinstance(item[1], int))
    require(manifest_file_count == expected_count
            and isinstance(manifest_bytes, int) and manifest_bytes > 0
            and all(isinstance(item[1], int) and item[1] >= 0
                    for item in manifest_entries.values()),
            f"{prefix} source manifest census changed")
    require(source.get("bytes") == manifest_bytes
            and source.get("manifest_size") == actual_manifest["bytes"],
            f"{prefix} source byte/manifest-size binding changed")
    if role == "baseline":
        require(manifest_raw.get("source_commit") == expected_commit
                and manifest_raw.get("source_root", expected_root) == expected_root,
                f"{prefix} baseline manifest identity changed")
    else:
        require(manifest_raw.get("base_commit") == expected_commit
                and manifest_raw.get("source_root", expected_root) == expected_root
                and source.get("changes") == manifest_raw.get("changes"),
                f"{prefix} candidate manifest delta changed")
    for field in ("cargo_manifest", "cargo_lock", "cargo_config"):
        item = _strict_descriptor(source.get(field), f"{prefix} source {field}")
        require(item["path"].startswith(expected_root + "/"),
                f"{prefix} source {field} escaped the role root")
    if role == "baseline":
        export = _strict_descriptor(source.get("source_export"), f"{prefix} source export")
        expected_export = _strict_descriptor(source_plan.get("source_export"),
                                             "0948 plan source export")
        require(export == expected_export, f"{prefix} source export changed")
    else:
        require("source_export" not in source,
                f"{prefix} unexpectedly carries baseline source export")
        expected_freeze = source_plan.get("candidate_freeze_receipt")
        require(expected_freeze is not None, f"{prefix} candidate freeze binding is missing")
        freeze = _strict_descriptor(source.get("freeze_receipt"),
                                    f"{prefix} candidate freeze receipt")
        require(freeze == _strict_descriptor(expected_freeze,
                                              "0948 plan candidate freeze receipt"),
                f"{prefix} candidate freeze receipt changed")

    expected_binary = ("litchi-perf-baseline-alloc" if kind == "allocator"
                       else "litchi-perf-baseline")
    expected_target = f"target-{prefix}-{revision}"
    _common_receipt_fields(raw, prefix, expected_root, expected_target, expected_binary)
    cwd = Path(raw["cwd"])
    packet_root = cwd.parent
    target_env_suffix = expected_target
    tmp_suffix = f"tmp/{prefix}-{revision}"
    argv = raw.get("argv")
    manifest_arg = f"{raw['cwd']}/tools/perf-baseline/Cargo.toml"
    expected_argv = [
        "/home/zhuhe/.cargo/bin/cargo", "build", "--release", "--locked", "--offline",
        "--manifest-path", manifest_arg, "--bin", expected_binary,
    ]
    if kind == "allocator":
        expected_argv += ["--features", "allocator-metrics"]
    expected_argv += ["-j2"]
    require(argv == expected_argv, f"{prefix} Cargo argv/feature flags changed")
    require(raw.get("deadline_seconds") == 1800,
            f"{prefix} build deadline changed")
    environment = raw.get("environment")
    require(isinstance(environment, dict), f"{prefix} build environment is missing")
    expected_env = {
        "PATH": "/home/zhuhe/.cargo/bin:/usr/bin:/bin", "HOME": "/home/zhuhe",
        "CARGO_HOME": "/home/zhuhe/.cargo", "RUSTUP_HOME": "/home/zhuhe/.rustup",
        "CARGO_NET_OFFLINE": "true", "CARGO_INCREMENTAL": "0", "CARGO_BUILD_JOBS": "2",
        "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1", "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "CARGO_PROFILE_RELEASE_LTO": "fat", "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3",
        "RUSTUP_TOOLCHAIN": "1.95.0", "LANG": "C", "LC_ALL": "C", "TZ": "UTC",
        "RUST_BACKTRACE": "0",
    }
    dynamic_env = {
        "CARGO_TARGET_DIR", "TMPDIR", "TEMP", "TMP",
    }
    require(set(environment) == set(expected_env) | dynamic_env
            and all(environment[key] == value for key, value in expected_env.items()),
            f"{prefix} build environment/profile flags changed")
    require(_absolute_suffix(environment.get("CARGO_TARGET_DIR"), target_env_suffix,
                             f"{prefix} CARGO_TARGET_DIR").parent == packet_root,
            f"{prefix} CARGO_TARGET_DIR packet root changed")
    for name in ("TMPDIR", "TEMP", "TMP"):
        tmp_path = _absolute_suffix(environment.get(name), tmp_suffix, f"{prefix} {name}")
        require(tmp_path.parent.parent == packet_root,
                f"{prefix} {name} packet root changed")
    require(raw.get("unset_environment") == UNSET_ENV,
            f"{prefix} unset environment changed")

    runner = _strict_descriptor(raw.get("runner"), f"{prefix} build runner")
    expected_runner = f"build-tools/build_{role}_{kind}_{revision}.py"
    require(runner["path"] == expected_runner,
            f"{prefix} build runner path changed")
    cargo = external_descriptor(raw.get("cargo"), Path("/home/zhuhe/.cargo/bin/cargo"),
                                f"{prefix} cargo")
    require(all(key in raw["cargo"] for key in ("bytes", "mode", "executable", "sha256",
                                                 "resolved_path"))
            and raw["cargo"].get("resolved_path") == cargo["resolved_path"],
            f"{prefix} Cargo executable custody is incomplete")
    for field, expected_path in (("stdout", f"builds/{prefix}-{revision}/build.stdout.log"),
                                 ("stderr", f"builds/{prefix}-{revision}/build.stderr.log")):
        value = _strict_descriptor(raw.get(field), f"{prefix} {field} log")
        require(value["path"] == expected_path, f"{prefix} {field} log path changed")

    terminal_binary = descriptor(raw.get("binary"), f"{prefix} terminal binary")
    if binary is not None:
        require(terminal_binary["bytes"] == binary.get("bytes")
                and terminal_binary["sha256"] == binary.get("sha256"),
                f"{prefix} terminal binary differs from role binding")

    build_dir = f"builds/{prefix}-{revision}"
    start_path = _packet_file(raw.get("start_receipt"), f"{prefix} start receipt",
                              f"{build_dir}/start.json")
    launch_path = _packet_file(raw.get("launch_receipt"), f"{prefix} launch receipt",
                               f"{build_dir}/launch.json")
    terminal_path = _packet_file(f"{build_dir}/terminal.json", f"{prefix} terminal receipt")
    require(read_json(terminal_path) == raw,
            f"{prefix} terminal receipt does not match its bound terminal")
    start = read_json(start_path)
    launch = read_json(launch_path)
    require(isinstance(start, dict) and start.get("schema") == expected_schema
            and start.get("receipt_kind") == "start",
            f"{prefix} start receipt schema/custody changed")
    require(isinstance(launch, dict) and launch.get("schema") == expected_schema
            and launch.get("receipt_kind") == "launch"
            and launch.get("status") == "launched",
            f"{prefix} launch receipt schema/custody changed")
    common_keys = ("argv", "cwd", "cwd_relative", "environment", "unset_environment",
                   "target_relative", "binary_relative", "deadline_seconds", "runner",
                   "cargo", "source", "start_new_session", "process_group",
                   "terminal_before_binary_parse")
    for key in common_keys:
        require(start.get(key) == raw.get(key), f"{prefix} start.{key} differs from terminal")
        require(launch.get(key) == raw.get(key), f"{prefix} launch.{key} differs from terminal")
    require(start.get("started_utc") == raw.get("started_utc")
            and launch.get("launched_utc") == raw.get("launched_utc")
            and launch.get("pid") == raw.get("pid")
            and launch.get("process_group_id") == raw.get("process_group_id"),
            f"{prefix} start/launch process custody changed")
    require(start.get("stdout_path") == f"{build_dir}/build.stdout.log"
            and start.get("stderr_path") == f"{build_dir}/build.stderr.log"
            and start.get("launch_receipt") == f"{build_dir}/launch.json"
            and start.get("terminal_receipt") == f"{build_dir}/terminal.json"
            and start.get("binary_receipt") == f"{build_dir}/binary.json"
            and start.get("result_receipt") == f"{build_dir}/result.json",
            f"{prefix} start receipt references changed")
    require(launch.get("start_receipt") == f"{build_dir}/start.json"
            and raw.get("start_receipt") == f"{build_dir}/start.json"
            and raw.get("launch_receipt") == f"{build_dir}/launch.json",
            f"{prefix} launch receipt references changed")
    _strict_descriptor(start.get("bindings"), f"{prefix} bindings receipt")
    bindings_path = _packet_file(start["bindings"]["path"], f"{prefix} bindings receipt")
    bindings = read_json(bindings_path)
    require(isinstance(bindings, dict) and bindings.get("schema") == expected_schema
            and bindings.get("receipt_kind") == "prelaunch-bindings"
            and bindings.get("bindings_path") == f"{build_dir}/bindings.json",
            f"{prefix} prelaunch binding receipt changed")
    for key in common_keys:
        require(bindings.get(key) == raw.get(key), f"{prefix} bindings.{key} differs")

    terminal_desc = _strict_descriptor({
        "path": f"{build_dir}/terminal.json",
        **{key: value for key, value in (
            ("bytes", terminal_path.stat().st_size),
            ("mode", terminal_path.stat().st_mode & 0o777),
            ("executable", bool(terminal_path.stat().st_mode & 0o111)),
            ("sha256", sha256(terminal_path)),
        )},
    }, f"{prefix} terminal receipt")
    binary_receipt_path = _packet_file(start["binary_receipt"], f"{prefix} binary receipt")
    result_receipt_path = _packet_file(start["result_receipt"], f"{prefix} result receipt")
    binary_receipt = read_json(binary_receipt_path)
    result_receipt = read_json(result_receipt_path)
    require(isinstance(binary_receipt, dict)
            and binary_receipt.get("schema") ==
            f"litchi.performance.0948.{prefix}-binary-receipt-{revision}"
            and binary_receipt.get("status") in {None, "passed"},
            f"{prefix} binary receipt schema/status changed")
    require(isinstance(result_receipt, dict)
            and result_receipt.get("schema") ==
            f"litchi.performance.0948.{prefix}-build-result-{revision}"
            and result_receipt.get("status") == "passed",
            f"{prefix} build result receipt schema/status changed")
    for receipt, label in ((binary_receipt, "binary"), (result_receipt, "result")):
        require(receipt.get("source_commit") == expected_commit,
                f"{prefix} {label} receipt source commit changed")
        require(receipt.get("build_terminal") == terminal_desc,
                f"{prefix} {label} receipt terminal binding changed")
    require(result_receipt.get("binary_receipt") == _strict_descriptor(
        {"path": start["binary_receipt"],
         "bytes": binary_receipt_path.stat().st_size,
         "mode": binary_receipt_path.stat().st_mode & 0o777,
         "executable": bool(binary_receipt_path.stat().st_mode & 0o111),
         "sha256": sha256(binary_receipt_path)}, f"{prefix} binary receipt descriptor"),
            f"{prefix} result binary receipt binding changed")
    bound = binary_receipt.get("binary")
    require(isinstance(bound, dict) and bound.get("status") == "present"
            and bound.get("path") == raw["binary"]["path"],
            f"{prefix} binary receipt identity changed")
    if binary is not None:
        require(bound.get("bytes") == binary.get("bytes")
                and bound.get("sha256") == binary.get("sha256"),
                f"{prefix} binary receipt differs from role binding")


def _reference_artifact(ref_dir: Path, value: Any, label: str,
                        expected_name: str | None = None) -> dict[str, Any]:
    """Validate a copied 0945 descriptor against its reference-directory file.

    The copied JSON preserves the original packet-relative path (for example
    ``builds/candidate-native-v2/terminal.json``), while the 0948 reference
    directory stores that artifact under its basename.  Digest and mode fields
    must still match exactly; only the containing packet path is adapted.
    """
    require(isinstance(value, dict) and isinstance(value.get("path"), str),
            f"{label} descriptor is missing")
    old_path = Path(value["path"])
    require(not old_path.is_absolute() and old_path.as_posix() == value["path"]
            and ".." not in old_path.parts,
            f"{label} original path is unsafe")
    if expected_name is not None:
        require(old_path.name == expected_name, f"{label} original basename changed")
    copied = ref_dir / old_path.name
    require(copied.is_file() and not copied.is_symlink(),
            f"{label} copied artifact is missing: {copied}")
    actual = descriptor({"path": copied.relative_to(PACKET).as_posix()}, label)
    for key in ("bytes", "mode", "executable", "sha256"):
        if key in value:
            require(value[key] == actual[key], f"{label}.{key} differs from copied artifact")
    return actual


def _legacy_reference_source(source: Any, prefix: str,
                             prior_manifest_sha: str) -> None:
    require(isinstance(source, dict), f"{prefix} 0945 source binding is missing")
    require(source.get("source_root") == "candidate-source-v2"
            and source.get("source_commit") == COMMIT
            and source.get("file_count") == 10788
            and source.get("bytes") == 259886223
            and source.get("manifest_size") == 3048241,
            f"{prefix} 0945 source census/commit changed")
    manifest = source.get("manifest")
    require(isinstance(manifest, dict)
            and manifest.get("path") == "candidate-source-manifest-v2.json"
            and manifest.get("bytes") == 3048241
            and manifest.get("sha256") == prior_manifest_sha,
            f"{prefix} 0945 source manifest binding changed")
    legacy_manifest = LEGACY_SOURCE_REF / "candidate-source-manifest-v2.json"
    legacy_manifest_desc = _strict_descriptor(
        {"path": legacy_manifest.relative_to(PACKET).as_posix(),
         "bytes": 3048241, "mode": 436, "executable": False,
         "sha256": prior_manifest_sha},
        f"{prefix} copied 0945 source manifest")
    require(legacy_manifest_desc["sha256"] == prior_manifest_sha,
            f"{prefix} copied 0945 source manifest digest changed")
    legacy_manifest_raw, legacy_files = manifest_files(
        legacy_manifest, f"{prefix} copied 0945 source manifest")
    require(legacy_manifest_raw.get("schema") == "litchi.performance.0945.source-manifest-v1"
            and legacy_manifest_raw.get("source_root") == "candidate-source-v2"
            and legacy_manifest_raw.get("base_commit") == COMMIT
            and legacy_manifest_raw.get("file_count") == 10788
            and legacy_manifest_raw.get("bytes") == 259886223,
            f"{prefix} copied 0945 source manifest metadata changed")
    freeze = source.get("freeze_receipt")
    require(isinstance(freeze, dict)
            and freeze.get("path") == "receipts/candidate-freeze-v2.json"
            and freeze.get("bytes") == 1269
            and freeze.get("sha256") ==
            "f624f161ad7e04499f6ae3a5f5079533ecc489c3690a2f46fdf947f3492dce32",
            f"{prefix} 0945 freeze binding changed")
    legacy_freeze = LEGACY_SOURCE_REF / "candidate-freeze-v2.json"
    legacy_freeze_desc = _strict_descriptor(
        {"path": legacy_freeze.relative_to(PACKET).as_posix(),
         "bytes": 1269, "mode": 436, "executable": False,
         "sha256": freeze["sha256"]},
        f"{prefix} copied 0945 freeze receipt")
    require(legacy_freeze_desc["sha256"] == freeze["sha256"],
            f"{prefix} copied 0945 freeze digest changed")
    legacy_freeze_raw = read_json(legacy_freeze)
    require(legacy_freeze_raw.get("status") == "frozen-before-corrected-quality-and-build"
            and legacy_freeze_raw.get("source_root") == "candidate-source-v2"
            and legacy_freeze_raw.get("file_count") == 10788
            and legacy_freeze_raw.get("source_bytes") == 259886223
            and legacy_freeze_raw.get("manifest") == {
                "bytes": 3048241, "path": "candidate-source-manifest-v2.json",
                "sha256": prior_manifest_sha},
            f"{prefix} copied 0945 freeze metadata changed")
    patch = legacy_freeze_raw.get("patch")
    require(isinstance(patch, dict) and patch.get("path") == "candidate-v2-portable.patch",
            f"{prefix} copied 0945 freeze patch binding changed")
    _reference_artifact(LEGACY_SOURCE_REF, patch,
                        f"{prefix} copied 0945 freeze patch", "candidate-v2-portable.patch")
    expected = {
        "cargo_config": ("candidate-source-v2/.cargo/config.toml", 470,
                          "50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d"),
        "cargo_lock": ("candidate-source-v2/tools/perf-baseline/Cargo.lock", 56289,
                        "db62bdd9174484613acae1f49f8187cf105d0b4976e2fea4de91df7372a893d5"),
        "cargo_manifest": ("candidate-source-v2/tools/perf-baseline/Cargo.toml", 3286,
                            "53fd47997ae969e709e2ee45575a6050f75de10df339f48dd427f35538768202"),
    }
    for field, (path, size, digest_value) in expected.items():
        item = source.get(field)
        require(isinstance(item, dict) and item.get("path") == path
                and item.get("bytes") == size and item.get("sha256") == digest_value,
                f"{prefix} 0945 {field} binding changed")
        manifest_name = path.removeprefix("candidate-source-v2/")
        manifest_item = legacy_files.get(manifest_name)
        require(isinstance(manifest_item, tuple)
                and manifest_item[1] == size and manifest_item[2] == digest_value,
                f"{prefix} 0945 {field} is not bound to the copied source manifest")
    changes = source.get("changes")
    require(isinstance(changes, list) and len(changes) == 2
            and {item.get("path") for item in changes if isinstance(item, dict)} == {
                "crates/litchi-opc/src/source_backed.rs",
                "crates/litchi-xlsx/src/cell_values/snapshot.rs",
            }, f"{prefix} 0945 source delta changed")
    require(source.get("changes") == legacy_freeze_raw.get("changes"),
            f"{prefix} 0945 source changes differ from the copied freeze")


def _validate_legacy_reference_build(ref_dir: Path, kind: str,
                                     entry: dict[str, Any],
                                     prior_manifest_sha: str) -> None:
    """Check the complete retained 0945 candidate build chain used as baseline."""
    prefix = f"candidate-{kind}"
    expected_binary = "litchi-perf-baseline-alloc" if kind == "allocator" else "litchi-perf-baseline"
    target = f"target-{prefix}-v2"
    expected_argv = [
        "/home/zhuhe/.cargo/bin/cargo", "build", "--release", "--locked", "--offline",
        "--manifest-path", f"/home/zhuhe/litchi-performance-0945/candidate-source-v2/tools/perf-baseline/Cargo.toml",
        "--bin", expected_binary,
    ]
    if kind == "allocator":
        expected_argv += ["--features", "allocator-metrics"]
    expected_argv += ["-j2"]
    expected_env = {
        "CARGO_BUILD_JOBS": "2", "CARGO_HOME": "/home/zhuhe/.cargo",
        "CARGO_INCREMENTAL": "0", "CARGO_NET_OFFLINE": "true",
        "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1", "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "CARGO_PROFILE_RELEASE_LTO": "fat", "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3",
        "HOME": "/home/zhuhe", "LANG": "C", "LC_ALL": "C",
        "PATH": "/home/zhuhe/.cargo/bin:/usr/bin:/bin", "RUSTUP_HOME": "/home/zhuhe/.rustup",
        "RUSTUP_TOOLCHAIN": "1.95.0", "RUST_BACKTRACE": "0", "TZ": "UTC",
        "CARGO_TARGET_DIR": f"/home/zhuhe/litchi-performance-0945/{target}",
        "TEMP": f"/home/zhuhe/litchi-performance-0945/tmp/{prefix}-v2",
        "TMP": f"/home/zhuhe/litchi-performance-0945/tmp/{prefix}-v2",
        "TMPDIR": f"/home/zhuhe/litchi-performance-0945/tmp/{prefix}-v2",
    }
    expected_runner = f"tools/build_candidate_{kind}_v2.py"
    terminal = read_json(ref_dir / "terminal.json")
    start = read_json(ref_dir / "start.json")
    launch = read_json(ref_dir / "launch.json")
    bindings = read_json(ref_dir / "bindings.json")
    binary_receipt = read_json(ref_dir / "binary.json")
    result = read_json(ref_dir / "result.json")
    expected_schema = f"litchi.performance.0945.{prefix}-build-v2"
    for value, receipt_kind in ((start, "start"), (launch, "launch"),
                                (terminal, "terminal"), (bindings, "prelaunch-bindings")):
        require(isinstance(value, dict) and value.get("schema") == expected_schema
                and value.get("receipt_kind") == receipt_kind,
                f"{prefix} 0945 {receipt_kind} schema changed")
    require(launch.get("status") == "launched"
            and terminal.get("status") == "terminal"
            and terminal.get("returncode") == 0
            and terminal.get("timed_out") is False
            and terminal.get("terminal_before_binary_parse") is True,
            f"{prefix} 0945 terminal status changed")
    cleanup = terminal.get("cleanup")
    require(isinstance(cleanup, dict) and cleanup.get("process_group_clear") is True
            and cleanup.get("reaped") is True,
            f"{prefix} 0945 process-group cleanup changed")
    require(terminal.get("start_new_session") is True and terminal.get("process_group") is True
            and terminal.get("launch_error") is None
            and isinstance(terminal.get("pid"), int) and terminal["pid"] > 0
            and terminal.get("process_group_id") == launch.get("process_group_id")
            and terminal.get("pid") == launch.get("pid"),
            f"{prefix} 0945 process custody changed")
    require(terminal.get("argv") == expected_argv
            and terminal.get("cwd") == "/home/zhuhe/litchi-performance-0945/candidate-source-v2"
            and terminal.get("cwd_relative") == "candidate-source-v2"
            and terminal.get("environment") == expected_env
            and terminal.get("unset_environment") == UNSET_ENV
            and terminal.get("target_relative") == target
            and terminal.get("binary_relative") == f"{target}/release/{expected_binary}"
            and terminal.get("deadline_seconds") == 1800,
            f"{prefix} 0945 command/environment binding changed")
    for value in (start, launch, bindings):
        for key in ("argv", "cwd", "cwd_relative", "environment", "unset_environment",
                    "target_relative", "binary_relative", "deadline_seconds", "runner",
                    "cargo", "source", "start_new_session", "process_group",
                    "terminal_before_binary_parse"):
            require(value.get(key) == terminal.get(key),
                    f"{prefix} {value.get('receipt_kind')} common field {key} changed")
    require(start.get("started_utc") == terminal.get("started_utc")
            and launch.get("launched_utc") == terminal.get("launched_utc"),
            f"{prefix} 0945 timestamp custody changed")
    _legacy_reference_source(terminal.get("source"), prefix, prior_manifest_sha)
    require(start.get("source") == launch.get("source") == bindings.get("source")
            == terminal.get("source"), f"{prefix} 0945 source receipt chain changed")
    runner = terminal.get("runner")
    require(isinstance(runner, dict) and runner.get("path") == expected_runner,
            f"{prefix} 0945 runner path changed")
    _reference_artifact(ref_dir, runner, f"{prefix} 0945 runner",
                        f"build_candidate_{kind}_v2.py")
    cargo = external_descriptor(terminal.get("cargo"), Path("/home/zhuhe/.cargo/bin/cargo"),
                                f"{prefix} 0945 cargo")
    require(all(key in terminal["cargo"] for key in ("bytes", "mode", "executable", "sha256",
                                                       "resolved_path"))
            and terminal["cargo"].get("resolved_path") == cargo["resolved_path"],
            f"{prefix} 0945 Cargo custody is incomplete")
    require(terminal.get("binary", {}).get("exists") is True
            and terminal.get("binary", {}).get("path") ==
            f"{target}/release/{expected_binary}",
            f"{prefix} 0945 terminal binary identity changed")
    for field, filename in (("stdout", "build.stdout.log"), ("stderr", "build.stderr.log")):
        value = terminal.get(field)
        require(isinstance(value, dict) and value.get("path") ==
                f"builds/{prefix}-v2/{filename}",
                f"{prefix} 0945 {field} descriptor changed")
        _reference_artifact(ref_dir, value, f"{prefix} 0945 {field}", filename)
    for key, filename in (("bindings", "bindings.json"), ("binary_receipt", "binary.json"),
                          ("result_receipt", "result.json")):
        value = start.get(key)
        require(isinstance(value, dict) if key == "bindings" else isinstance(value, str),
                f"{prefix} 0945 start {key} reference is malformed")
    _reference_artifact(ref_dir, start["bindings"], f"{prefix} 0945 bindings", "bindings.json")
    require(start.get("stdout_path") == f"builds/{prefix}-v2/build.stdout.log"
            and start.get("stderr_path") == f"builds/{prefix}-v2/build.stderr.log"
            and start.get("launch_receipt") == f"builds/{prefix}-v2/launch.json"
            and start.get("terminal_receipt") == f"builds/{prefix}-v2/terminal.json"
            and start.get("binary_receipt") == f"builds/{prefix}-v2/binary.json"
            and start.get("result_receipt") == f"builds/{prefix}-v2/result.json"
            and terminal.get("start_receipt") == f"builds/{prefix}-v2/start.json"
            and terminal.get("launch_receipt") == f"builds/{prefix}-v2/launch.json",
            f"{prefix} 0945 receipt references changed")
    require(bindings.get("bindings_path") == f"builds/{prefix}-v2/bindings.json",
            f"{prefix} 0945 bindings path changed")
    terminal_desc = descriptor({"path": f"references/0945-candidate-{kind}/terminal.json"},
                               f"{prefix} copied terminal")
    binary_desc = descriptor({"path": f"references/0945-candidate-{kind}/binary.json"},
                              f"{prefix} copied binary receipt")
    require(binary_receipt.get("schema") ==
            f"litchi.performance.0945.{prefix}-binary-receipt-v2"
            and binary_receipt.get("source_commit") == COMMIT
            and isinstance(binary_receipt.get("binary"), dict)
            and binary_receipt["binary"].get("status") == "present"
            and binary_receipt["binary"].get("path") == terminal["binary"]["path"],
            f"{prefix} 0945 binary receipt changed")
    _reference_artifact(ref_dir, binary_receipt.get("build_terminal"),
                        f"{prefix} 0945 binary terminal binding", "terminal.json")
    require(result.get("schema") ==
            f"litchi.performance.0945.{prefix}-build-result-v2"
            and result.get("status") == "passed"
            and result.get("source_commit") == COMMIT,
            f"{prefix} 0945 result receipt changed")
    _reference_artifact(ref_dir, result.get("terminal"),
                        f"{prefix} 0945 result terminal binding", "terminal.json")
    _reference_artifact(ref_dir, result.get("binary_receipt"),
                        f"{prefix} 0945 result binary binding", "binary.json")
    require(result.get("binary", {}).get("bytes") == entry.get("bytes")
            and result.get("binary", {}).get("sha256") == entry.get("sha256")
            and binary_receipt.get("binary", {}).get("bytes") == entry.get("bytes")
            and binary_receipt.get("binary", {}).get("sha256") == entry.get("sha256"),
            f"{prefix} 0945 binary digest differs from reuse receipt")
    require(terminal_desc["sha256"] == entry.get("build_receipt", {}).get("sha256")
            and binary_desc["sha256"] == _reference_artifact(
                ref_dir, result.get("binary_receipt"),
                f"{prefix} 0945 binary receipt descriptor", "binary.json")["sha256"],
            f"{prefix} 0945 copied receipt digest chain changed")


def _validate_source_equivalence_0948(raw: dict[str, Any], plan: dict[str, Any]) -> None:
    """Require packet-local per-file source equivalence for the reused baseline."""
    equivalence = raw.get("source_equivalence")
    require(isinstance(equivalence, dict),
            "0948 baseline reuse source_equivalence receipt is missing")
    require(equivalence.get("file_count") == 10788
            and equivalence.get("matched_path_count") == 10788
            and equivalence.get("compared_fields") == list(FIELDS)
            and equivalence.get("mismatches") == {field: 0 for field in FIELDS},
            "0948 baseline source equivalence is not exact")
    current_desc = _strict_descriptor(equivalence.get("current_manifest"),
                                      "0948 reuse current selected manifest")
    original_desc = _strict_descriptor(equivalence.get("original_manifest"),
                                       "0948 reuse prior selected manifest")
    original_export_desc = _strict_descriptor(equivalence.get("original_export"),
                                              "0948 reuse prior source export")
    selection_desc = _strict_descriptor(equivalence.get("closure_selection"),
                                        "0948 reuse closure selection")
    _, current_files = manifest_files(PACKET / current_desc["path"],
                                      "0948 reuse current selected manifest")
    current_manifest_raw = read_json(PACKET / current_desc["path"])
    original_manifest_raw, original_files = manifest_files(
        PACKET / original_desc["path"], "0948 reuse prior selected manifest")
    require(len(current_files) == len(original_files) == 10788
            and current_files == original_files,
            "0948 baseline selected source closure differs")
    plan_source = plan.get("source")
    require(isinstance(plan_source, dict), "0948 plan source binding is missing")
    plan_baseline_desc = _strict_descriptor(plan_source.get("baseline_manifest"),
                                            "0948 plan baseline source manifest")
    _, plan_baseline_files = manifest_files(PACKET / plan_baseline_desc["path"],
                                            "0948 plan baseline source manifest")
    require(len(plan_baseline_files) == 10789
            and set(plan_baseline_files) - set(current_files) == {"Cargo.lock"},
            "0948 selected source closure does not exclude only the ignored root lock")
    for name, values in current_files.items():
        wanted = plan_baseline_files[name]
        require(values[1:3] == wanted[1:3],
                f"0948 selected current source differs from the frozen baseline manifest: {name}")
    require(current_manifest_raw.get("schema") ==
            "litchi.performance.0948.selected-source-manifest.v1"
            and current_manifest_raw.get("source_root") == "baseline-source"
            and current_manifest_raw.get("source_commit") == plan.get("source", {}).get("source_commit")
            and current_manifest_raw.get("file_count") == 10788
            and current_manifest_raw.get("bytes") == 259886223
            and original_manifest_raw.get("schema") ==
            "litchi.performance.0948.prior-candidate-selected-source-manifest.v1"
            and original_manifest_raw.get("source_root") == "candidate-source-v2"
            and original_manifest_raw.get("source_commit") == COMMIT
            and original_manifest_raw.get("file_count") == 10788
            and original_manifest_raw.get("bytes") == 259886223
            and equivalence.get("current_commit") == plan.get("source", {}).get("source_commit")
            and equivalence.get("original_commit") == COMMIT,
            "0948 baseline source commit equivalence changed")
    original_export = read_json(PACKET / original_export_desc["path"])
    selection = read_json(PACKET / selection_desc["path"])
    require(isinstance(original_export, dict)
            and original_export.get("schema") ==
            "litchi.performance.0948.baseline-reuse-source-equivalence-export.v1"
            and original_export.get("file_count") == 10788
            and original_export.get("source_bytes") == 259886223
            and original_export.get("current_manifest") == current_desc
            and original_export.get("original_manifest") == original_desc
            and isinstance(selection, dict)
            and selection.get("schema") == "litchi-xlsx-profile-source-closure-selection-v1"
            and selection.get("origin_commit") ==
            "933ee45cbb5d2ef9dc29583f7ec890ad234e04fb"
            and selection.get("tracked_3rdparty_count") == 0,
            "0948 baseline source closure receipts are malformed")


def _validate_historical_source_closure(raw: dict[str, Any]) -> None:
    """Validate the retained 0944 closure receipts referenced by the v2 packet."""
    historical = raw.get("historical_source_closure")
    require(isinstance(historical, dict)
            and historical.get("origin_commit") ==
            "933ee45cbb5d2ef9dc29583f7ec890ad234e04fb",
            "0948 historical source closure binding is missing")
    manifest_desc = _strict_descriptor(
        historical.get("manifest"), "0948 historical source manifest")
    export_desc = _strict_descriptor(
        historical.get("export"), "0948 historical source export")
    selection_desc = _strict_descriptor(
        historical.get("closure_selection"), "0948 historical closure selection")
    require(manifest_desc["path"] ==
            "references/0945-source-equivalence/source-build-manifest.json"
            and manifest_desc["bytes"] == 3674260
            and manifest_desc["sha256"] ==
            "b63113ce5251218be817ba29c466b0ace535c546a416b97413ae8a73f0caea64"
            and export_desc["path"] ==
            "references/0945-source-equivalence/export-receipt.json"
            and export_desc["bytes"] == 734
            and export_desc["sha256"] ==
            "1bc0724ffb9b00910bcbf3453bd9bfd6e22133a4e1e534c1b49349ba3c1fda4b"
            and selection_desc["path"] ==
            "references/0945-source-equivalence/source-closure-selection.json"
            and selection_desc["bytes"] == 754
            and selection_desc["sha256"] ==
            "ffec8134b4d49b26bb68f57213b97b3c3a44439c735df1ec14ba4fbec42f4f21",
            "0948 historical source closure descriptor changed")
    manifest = read_json(PACKET / manifest_desc["path"])
    export = read_json(PACKET / export_desc["path"])
    selection = read_json(PACKET / selection_desc["path"])
    require(manifest.get("schema") == "litchi-xlsx-profile-source-closure-v1"
            and manifest.get("origin_commit") == historical["origin_commit"]
            and manifest.get("file_count") == 10788
            and export.get("schema") == "litchi-xlsx-profile-source-export-v1"
            and export.get("manifest_sha256") == manifest.get("full_export_manifest_sha256")
            and selection.get("schema") == "litchi-xlsx-profile-source-closure-selection-v1"
            and selection.get("origin_commit") == historical["origin_commit"]
            and selection.get("tracked_3rdparty_count") == 0,
            "0948 historical source closure contents changed")


def validate_baseline_reuse_0948(raw: dict[str, Any], plan: dict[str, Any],
                                 kind: str,
                                 binary: dict[str, Any] | None = None) -> None:
    """Validate the 0948 adapter around exact retained 0945 build receipts."""
    require(raw.get("schema") == "litchi.performance.0948.baseline-reuse.v2"
            and raw.get("status") == "pass"
            and raw.get("baseline_head") == plan.get("source", {}).get("source_commit")
            and raw.get("checked_files") == 10788
            and raw.get("all_selected_tracked_build_inputs_equal") is True,
            "0948 baseline reuse receipt identity/equivalence changed")
    prior = raw.get("prior_candidate_source_manifest")
    prior_desc = _strict_descriptor(prior, "0948 prior candidate source manifest")
    require(prior_desc["path"] ==
            "references/0945-source-equivalence/candidate-source-manifest-v2.json"
            and prior_desc["bytes"] == 3048241
            and prior_desc["sha256"] ==
            "da7128f699b4bf7a9a4745e99b99606830df1aa192ce389a0b7fe3cceb8a0994",
            "0948 prior candidate source manifest binding changed")
    _validate_source_equivalence_0948(raw, plan)
    _validate_historical_source_closure(raw)
    binaries = raw.get("binaries")
    require(isinstance(binaries, list) and len(binaries) == 2,
            "0948 baseline reuse binary list changed")
    by_kind = {item.get("kind"): item for item in binaries if isinstance(item, dict)}
    require(set(by_kind) == {"native", "allocator"},
            "0948 baseline reuse binary kinds changed")
    item = by_kind[kind]
    expected_name = f"baseline-{kind}"
    require(item.get("path") == f"binaries/{expected_name}"
            and item.get("original_path") ==
            f"/home/zhuhe/litchi-performance-0945/target-candidate-{kind}-v2/release/"
            f"{'litchi-perf-baseline-alloc' if kind == 'allocator' else 'litchi-perf-baseline'}"
            and item.get("prior_build_command_source") ==
            f"references/0945-candidate-{kind}/build_candidate_{kind}_v2.py",
            f"0948 {kind} reuse binary identity changed")
    copied_binary = _strict_descriptor({"path": item["path"],
                                        "bytes": item.get("bytes"),
                                        "mode": item.get("mode"),
                                        "executable": item.get("executable"),
                                        "sha256": item.get("sha256")},
                                       f"0948 {kind} baseline binary")
    if binary is not None:
        require(copied_binary["bytes"] == binary.get("bytes")
                and copied_binary["sha256"] == binary.get("sha256"),
                f"0948 {kind} reuse binary differs from plan binding")
    receipt_value = item.get("build_receipt")
    receipt_desc = _strict_descriptor(receipt_value,
                                      f"0948 {kind} prior build receipt")
    require(receipt_desc["path"] ==
            f"references/0945-candidate-{kind}/terminal.json"
            and receipt_desc["bytes"] == (5268 if kind == "native" else 5368)
            and receipt_desc["sha256"] == item["build_receipt"]["sha256"],
            f"0948 {kind} prior build receipt binding changed")
    command_source = item.get("prior_build_command_source")
    require(command_source ==
            f"references/0945-candidate-{kind}/build_candidate_{kind}_v2.py",
            f"0948 {kind} prior build command source changed")
    _reference_artifact(PACKET / f"references/0945-candidate-{kind}",
                        {"path": f"build_candidate_{kind}_v2.py"},
                        f"0948 {kind} prior build command source",
                        f"build_candidate_{kind}_v2.py")
    _validate_legacy_reference_build(PACKET / f"references/0945-candidate-{kind}",
                                     kind, item, prior["sha256"])


def main() -> None:
    receipt = PACKET / "receipts" / "baseline-reuse-v2.json"
    raw = read_json(receipt)
    require(raw.get("schema") == "litchi.performance.0948.baseline-reuse.v2"
            and raw.get("status") == "pass"
            and raw.get("checked_files") == 10788,
            "0948 baseline reuse receipt identity changed")
    prior = _strict_descriptor(raw.get("prior_candidate_source_manifest"),
                               "0948 prior candidate source manifest")
    require(prior["path"] ==
            "references/0945-source-equivalence/candidate-source-manifest-v2.json",
            "0948 prior candidate source manifest path changed")
    binaries = raw.get("binaries")
    require(isinstance(binaries, list) and {item.get("kind") for item in binaries
                                            if isinstance(item, dict)} == {"native", "allocator"},
            "0948 baseline reuse binary list changed")
    print(json.dumps({"status": "pass", "receipt": str(receipt.relative_to(PACKET))}, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, ReuseReceiptError) as error:
        print(f"baseline reuse validation failed: {error}")
        raise SystemExit(1)
