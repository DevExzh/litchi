#!/usr/bin/env python3
"""Recursive validator for the packet-local 0945 native reuse receipt."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent.parent
SCHEMA = "litchi.performance.0945.native-baseline-reuse-v2"
COMMIT = "e037befa6ce52b3988c8ae2032ae652196f2a09d"
FIELDS = ("path", "bytes", "sha256", "git_oid", "git_mode")
MAX_JSON_BYTES = 64 * 1024 * 1024
SOURCE_BYTES = {"baseline": 259872266}
SOURCE_MANIFEST_FILES = {"baseline": 10788, "candidate": 10788}
BUILD_REVISIONS = {"baseline": "v1", "candidate": "v2"}
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
    corrected v2 wrappers and therefore have separate target/log namespaces.
    """
    require(role in BUILD_REVISIONS and kind in {"native", "allocator"},
            "invalid build role/kind")
    revision = BUILD_REVISIONS[role]
    binding_id = BUILD_BINDING_IDS[role]
    if plan is not None:
        require(isinstance(plan, dict), "frozen plan is not an object")
        revisions = plan.get("build_revisions")
        binding_ids = plan.get("build_binding_ids")
        roles = plan.get("roles")
        # Small retained-receipt controls intentionally pass only the source
        # subset of a plan.  Enforce the explicit binding whenever a full
        # frozen plan (or any role binding fields) is supplied.
        if revisions is not None or binding_ids is not None or roles is not None:
            require(revisions == BUILD_REVISIONS and binding_ids == BUILD_BINDING_IDS,
                    "frozen plan build revision bindings changed")
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
    expected_runner = (f"tools/build_candidate_{kind}_v2.py" if role == "candidate"
                       else f"tools/build_baseline_{kind}.py")
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


def main() -> None:
    receipt = PACKET / "receipts" / "baseline-native-reuse-v2.json"
    actual = validate_receipt(receipt)
    print(json.dumps({"status": "passed", "receipt": actual}, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, ReuseReceiptError) as error:
        print(f"baseline reuse validation failed: {error}")
        raise SystemExit(1)
