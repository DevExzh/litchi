#!/usr/bin/env python3
"""Build and provenance admission for the paired ODG joined probes.

The runtime runner is intentionally not a build system.  This companion is
the only entry point allowed to create the two probe binaries: it prepares a
fully bound plan, applies the reviewed example overlay to one pristine
worktree at a time, streams a bounded Cargo child into fresh logs, and emits a
receipt which the runtime runner revalidates before starting a probe.

All paths in this module are treated as untrusted.  Existing output, receipt,
log, target, or overlay paths are never overwritten; symlink components,
`3rdparty` paths, sparse/skip-worktree source trees, stale refs, and changed
source/config identities fail closed.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import resource
import shutil
import signal
import subprocess
import sys
import time
from typing import Any, Mapping

try:
    from odp_native_qualification import (
        MAX_CHILD_OUTPUT,
        RunFailure,
        _digest,
        _load_json,
        _nonnegative_int,
        _object,
        _positive_int,
        _string,
        _run_git,
        binary_identity,
        load_average,
        process_scan,
    )
except ImportError:  # pragma: no cover - direct tools import.
    from tools.odp_native_qualification import (  # type: ignore[no-redef]
        MAX_CHILD_OUTPUT,
        RunFailure,
        _digest,
        _load_json,
        _nonnegative_int,
        _object,
        _positive_int,
        _string,
        _run_git,
        binary_identity,
        load_average,
        process_scan,
    )


PREPARE_SCHEMA = "litchi.odg.joined-native-qualification.prepare.v2"
BUILD_SCHEMA = "litchi.odg.joined-native-qualification.build.v2"
BUILDER_REVISION = "odg-joined-builder-v2-streamed-receipt"
PRISTINE_BASELINE_REF = "ad52437f5e5e2c71ea476eedd5ddcf61aee2021d"
EXPECTED_EXAMPLE = "litchi-perf-odg-joined-qualification"
EXPECTED_OVERLAY_DESTINATION = "crates/litchi-odg/examples/litchi-perf-odg-joined-qualification.rs"
EXPECTED_COMMON_DESTINATION = "crates/litchi-odg/examples/odg_joined_qualification_common.rs"
EXPECTED_SOURCE_NAMES = {
    "common": "odg_joined_qualification_common.rs",
    "allocator": "odg_joined_qualification_allocator.rs",
    "baseline": "odg_joined_qualification_ad524_probe.rs",
    "candidate": "odg_joined_qualification_probe.rs",
}
MAX_BUILD_OUTPUT = 64 * 1024 * 1024
MEMORY_LIMIT_BYTES = 20_000_000_000
DESCENDANT_GRACE_S = 2.0


def _exact_keys(value: Mapping[str, Any], expected: set[str], label: str) -> None:
    if set(value) != expected:
        raise RunFailure(f"{label} fields differ from the build contract")


def _raw_path(value: Any, location: str, base: Path) -> Path:
    raw = _string(value, location)
    path = Path(raw)
    return (base / path).absolute() if not path.is_absolute() else path.absolute()


def _sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunFailure(f"cannot hash a non-regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
                size += len(chunk)
    except OSError as error:
        raise RunFailure(f"cannot hash {path}: {error}") from error
    return digest.hexdigest(), size


def _safe_path(path: Path, label: str, *, must_exist: bool = False) -> Path:
    """Reject symlink components and untracked upstream trees without follow."""

    path = path.absolute()
    if "3rdparty" in path.parts:
        raise RunFailure(f"{label} must not reference 3rdparty: {path}")
    current = Path(path.anchor)
    for component in path.parts[1:]:
        current /= component
        try:
            current.lstat()
        except FileNotFoundError:
            if must_exist:
                raise RunFailure(f"{label} does not exist: {path}")
            break
        except OSError as error:
            raise RunFailure(f"cannot inspect {label} {current}: {error}") from error
        if current.is_symlink():
            raise RunFailure(f"{label} contains a symlink component: {current}")
        if current != path and not current.is_dir():
            raise RunFailure(f"{label} has a non-directory component: {current}")
    if must_exist and not path.exists():
        raise RunFailure(f"{label} does not exist: {path}")
    return path


def _fresh_dir(path: Path, label: str) -> None:
    path = _safe_path(path, label)
    if path.exists():
        raise RunFailure(f"{label} must be a new directory: {path}")
    path.mkdir(parents=True, exist_ok=False)


def _fresh_file(path: Path, value: bytes, label: str) -> None:
    path = _safe_path(path, label)
    if path.exists():
        raise RunFailure(f"{label} must not overwrite an existing file: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(value)
            stream.flush()
            os.fsync(stream.fileno())
    except OSError as error:
        raise RunFailure(f"cannot create fresh {label} {path}: {error}") from error


def _write_fresh_json(path: Path, value: Any, label: str) -> None:
    _fresh_file(path, (json.dumps(value, indent=2, sort_keys=True) + "\n").encode(), label)


def _manifest_sha(path: Path) -> str:
    return _sha256_file(path)[0]


def _raw_manifest(path: Path) -> dict[str, Any]:
    path = _safe_path(path.absolute(), "manifest", must_exist=True)
    value = _load_json(path)
    return value


def _endpoint_value(raw: Mapping[str, Any], name: str) -> dict[str, Any]:
    value = _object(raw.get(name), f"manifest.{name}")
    expected = {
        "root", "ref", "tree", "profile", "binary", "target_dir", "lock",
        "toolchain", "cargo_config", "source_manifest", "allocator_binary",
        "allocator_target_dir", "features",
    }
    _exact_keys(value, expected, f"manifest.{name}")
    return value


def _build_value(raw: Mapping[str, Any]) -> dict[str, Any]:
    value = _object(raw.get("build"), "manifest.build")
    _exact_keys(value, {"prepare_receipt", "receipt", "final_candidate_ref", "cpu"}, "manifest.build")
    return value


def _protocol_max_load(raw: Mapping[str, Any]) -> float:
    protocol = _object(raw.get("protocol"), "manifest.protocol")
    value = protocol.get("max_load")
    if type(value) not in (int, float) or value <= 0:
        raise RunFailure("manifest.protocol.max_load must be positive for live guards")
    return float(value)


def _identity_value(value: Any, label: str, base: Path) -> tuple[Path, int, str]:
    item = _object(value, label)
    _exact_keys(item, {"path", "bytes", "sha256"}, label)
    path = _safe_path(_raw_path(item.get("path"), f"{label}.path", base), f"{label}.path", must_exist=True)
    return path, _positive_int(item.get("bytes"), f"{label}.bytes"), _digest(item.get("sha256"), f"{label}.sha256")


def _identity(path: Path) -> dict[str, Any]:
    sha, size = _sha256_file(path)
    return {"path": str(path), "bytes": size, "sha256": sha}


def _check_identity(value: Any, label: str, base: Path, root: Path | None = None) -> dict[str, Any]:
    path, expected_bytes, expected_sha = _identity_value(value, label, base)
    if root is not None:
        try:
            path.relative_to(root)
        except ValueError as error:
            raise RunFailure(f"{label}.path must be below its source root") from error
    actual = _identity(path)
    if actual["bytes"] != expected_bytes or actual["sha256"] != expected_sha:
        raise RunFailure(f"{label} identity differs from the manifest")
    return actual


def _outside(path: Path, root: Path, label: str) -> None:
    if path == root or root in path.parents:
        raise RunFailure(f"{label} must be outside source root {root}")


def _git_output(root: Path, *args: str) -> str:
    return _run_git(root, *args)


def _check_worktree_shape(root: Path, name: str) -> dict[str, Any]:
    root = _safe_path(root, f"{name}.root", must_exist=True)
    if not root.is_dir():
        raise RunFailure(f"{name}.root is not a directory")
    top = Path(_git_output(root, "rev-parse", "--show-toplevel")).resolve()
    if top != root.resolve():
        raise RunFailure(f"{name}.root must be the Git worktree root")
    sparse = ""
    try:
        sparse = _git_output(root, "config", "--get", "core.sparseCheckout")
    except RunFailure:
        sparse = ""
    if sparse.lower() in {"true", "1", "yes", "on"}:
        raise RunFailure(f"{name}.root uses sparse checkout")
    indexed = _run_git_bytes(root, "ls-files", "-t", "-z")
    entries: list[str] = []
    for item in indexed.split(b"\0"):
        if not item:
            continue
        if len(item) < 3 or item[1:2] != b" ":
            raise RunFailure(f"{name}.root returned malformed ls-files status")
        status, raw_path = item[0:1], item[2:]
        path = root / os.fsdecode(raw_path)
        _safe_path(path, f"{name} tracked path", must_exist=False)
        if status == b"S":
            raise RunFailure(f"{name}.root contains skip-worktree entry {raw_path!r}")
        if not path.exists():
            raise RunFailure(f"{name}.root omits tracked file {raw_path!r}")
        entries.append(os.fsdecode(raw_path))
    status = _git_output(root, "status", "--porcelain", "--untracked-files=all")
    if status:
        raise RunFailure(f"{name}.root is dirty")
    return {"root": str(root), "tracked_entries": len(entries), "sparse": False}


def _run_git_bytes(root: Path, *args: str) -> bytes:
    try:
        completed = subprocess.run(
            ["git", "-C", str(root), *args],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            timeout=30,
        )
    except (OSError, subprocess.SubprocessError) as error:
        raise RunFailure(f"cannot invoke git for {root}: {error}") from error
    if completed.returncode != 0:
        detail = completed.stderr.decode(errors="replace").strip()
        raise RunFailure(f"git {' '.join(args)} failed for {root}: {detail}")
    return completed.stdout


def _tree_manifest(root: Path, revision: str) -> dict[str, Any]:
    raw = _run_git_bytes(root, "ls-tree", "-r", "--full-tree", "--long", revision)
    entries = [item for item in raw.splitlines() if item]
    return {
        "entries": len(entries),
        "bytes": len(raw),
        "sha256": hashlib.sha256(raw).hexdigest(),
    }


def _check_tree_manifest(value: Any, actual: Mapping[str, Any], label: str) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"entries", "bytes", "sha256"}, label)
    expected = {
        "entries": _positive_int(item.get("entries"), f"{label}.entries"),
        "bytes": _positive_int(item.get("bytes"), f"{label}.bytes"),
        "sha256": _digest(item.get("sha256"), f"{label}.sha256"),
    }
    if dict(actual) != expected:
        raise RunFailure(f"{label} differs from the complete Git tree")
    return expected


def _endpoint_identity(raw: Mapping[str, Any], name: str, base: Path, build: Mapping[str, Any], *, allow_built_artifacts: bool = False) -> dict[str, Any]:
    value = _endpoint_value(raw, name)
    root = _safe_path(_raw_path(value["root"], f"manifest.{name}.root", base), f"{name}.root", must_exist=True)
    shape = _check_worktree_shape(root, name)
    ref = _string(value["ref"], f"manifest.{name}.ref")
    tree = _string(value["tree"], f"manifest.{name}.tree")
    head = _git_output(root, "rev-parse", "--verify", "HEAD")
    actual_tree = _git_output(root, "rev-parse", "--verify", "HEAD^{tree}")
    if head != ref or actual_tree != tree:
        raise RunFailure(f"{name}: HEAD/tree do not match the manifest")
    if name == "baseline" and ref != PRISTINE_BASELINE_REF:
        raise RunFailure(f"baseline must be pristine pinned ref {PRISTINE_BASELINE_REF}")
    if name == "candidate" and ref != _string(build["final_candidate_ref"], "manifest.build.final_candidate_ref"):
        raise RunFailure("candidate ref is not bound to the final candidate ref")
    expected_profile = "legacy-apply-joined" if name == "baseline" else "current-edit-join-commit"
    if value["profile"] != expected_profile:
        raise RunFailure(f"{name}.profile is not the reviewed paired profile")
    features = value["features"]
    if (not isinstance(features, list)
            or any(not isinstance(feature, str) or not feature for feature in features)
            or features != sorted(set(features))):
        raise RunFailure(f"{name}.features must be a sorted unique list")
    lock = _check_identity(value["lock"], f"{name}.lock", base, root)
    toolchain = _check_identity(value["toolchain"], f"{name}.toolchain", base, root)
    cargo_config = _check_identity(value["cargo_config"], f"{name}.cargo_config", base, root)
    for field, expected_relative in (
        ("lock", Path("Cargo.lock")),
        ("toolchain", Path("rust-toolchain.toml")),
        ("cargo_config", Path(".cargo/config.toml")),
    ):
        declared, _, _ = _identity_value(value[field], f"{name}.{field}", base)
        if declared != root / expected_relative:
            raise RunFailure(f"{name}.{field} must use the reviewed {expected_relative} path")
    tree_identity = _tree_manifest(root, head)
    source_manifest = _check_tree_manifest(value["source_manifest"], tree_identity, f"{name}.source_manifest")
    binary = _safe_path(_raw_path(value["binary"], f"manifest.{name}.binary", base), f"{name}.binary")
    allocator_binary = _safe_path(
        _raw_path(value["allocator_binary"], f"manifest.{name}.allocator_binary", base),
        f"{name}.allocator_binary",
    )
    target_dir = _safe_path(_raw_path(value["target_dir"], f"manifest.{name}.target_dir", base), f"{name}.target_dir")
    allocator_target_dir = _safe_path(
        _raw_path(value["allocator_target_dir"], f"manifest.{name}.allocator_target_dir", base),
        f"{name}.allocator_target_dir",
    )
    if target_dir == allocator_target_dir:
        raise RunFailure(f"{name}: normal and allocator target directories must differ")
    _outside(target_dir, root, f"{name}.target_dir")
    _outside(allocator_target_dir, root, f"{name}.allocator_target_dir")
    for artifact, label, target in (
        (binary, "binary", target_dir),
        (allocator_binary, "allocator_binary", allocator_target_dir),
    ):
        if artifact.is_symlink():
            raise RunFailure(f"{name}.{label} must not be a symlink")
        if artifact.exists() and not allow_built_artifacts:
            raise RunFailure(f"{name}.{label} must not exist before the fresh build")
        try:
            artifact.relative_to(target)
        except ValueError as error:
            raise RunFailure(f"{name}.{label} must be below its target directory") from error
    return {
        "name": name,
        "root": str(root),
        "ref": ref,
        "tree": tree,
        "profile": value["profile"],
        "features": list(features),
        "source_shape": shape,
        "lock": lock,
        "toolchain": toolchain,
        "cargo_config": cargo_config,
        "source_manifest": source_manifest,
        "binary": str(binary),
        "allocator_binary": str(allocator_binary),
        "target_dir": str(target_dir),
        "allocator_target_dir": str(allocator_target_dir),
    }


def _probe_plan(raw: Mapping[str, Any], base: Path, roots: tuple[Path, Path]) -> list[dict[str, Any]]:
    sources = _object(raw.get("probe_sources"), "manifest.probe_sources")
    expected_destinations = {
        "common": EXPECTED_COMMON_DESTINATION,
        "allocator": "crates/litchi-odg/examples/odg_joined_qualification_allocator.rs",
        "baseline": EXPECTED_OVERLAY_DESTINATION,
        "candidate": EXPECTED_OVERLAY_DESTINATION,
    }
    result: list[dict[str, Any]] = []
    for name in ("common", "allocator", "baseline", "candidate"):
        item = _object(sources.get(name), f"manifest.probe_sources.{name}")
        _exact_keys(item, {"path", "bytes", "sha256", "destination"}, f"manifest.probe_sources.{name}")
        path = _safe_path(_raw_path(item["path"], f"manifest.probe_sources.{name}.path", base), f"probe {name}", must_exist=True)
        if (path.name != EXPECTED_SOURCE_NAMES[name] or path.parent.name != "perf-baseline"
                or path.parent.parent.name != "tools"):
            raise RunFailure(f"probe {name} is not the reviewed source adapter")
        if item["destination"] != expected_destinations[name]:
            raise RunFailure(f"probe {name} has an unreviewed destination")
        for root in roots:
            if path == root or root in path.parents:
                raise RunFailure(f"probe {name} must not come from an endpoint worktree")
        actual = _identity(path)
        if actual["bytes"] != _positive_int(item["bytes"], f"probe {name}.bytes") or actual["sha256"] != _digest(item["sha256"], f"probe {name}.sha256"):
            raise RunFailure(f"probe {name} identity differs from the manifest")
        result.append({"name": name, **actual, "destination": item["destination"]})
    return result


def _fixture_plan(raw: Mapping[str, Any], base: Path) -> dict[str, Any]:
    item = _object(raw.get("fixture"), "manifest.fixture")
    _exact_keys(item, {"path", "bytes", "sha256", "provenance"}, "manifest.fixture")
    path = _safe_path(_raw_path(item["path"], "manifest.fixture.path", base), "fixture", must_exist=True)
    if (path.name != "joined-content.xml" or path.parent.name != "odg"
            or path.parent.parent.name != "test-data"):
        raise RunFailure("fixture is not the reviewed test-data/odg/joined-content.xml")
    actual = _identity(path)
    if actual["bytes"] != _positive_int(item["bytes"], "manifest.fixture.bytes") or actual["sha256"] != _digest(item["sha256"], "manifest.fixture.sha256"):
        raise RunFailure("fixture identity differs from the manifest")
    content = path.read_bytes()
    if b"<office:document-content" not in content or b"<draw:page" not in content:
        raise RunFailure("fixture is not the reviewed ODG content XML")
    return {**actual, "path": str(path), "provenance": _string(item["provenance"], "manifest.fixture.provenance")}


def _command(endpoint: Mapping[str, Any], *, allocator: bool, cpu: int) -> dict[str, Any]:
    target = endpoint["allocator_target_dir"] if allocator else endpoint["target_dir"]
    features = endpoint["features"]
    command = [
        "taskset", "--cpu-list", str(cpu), "cargo", "build", "--offline", "--locked",
        "--jobs", "1", "--manifest-path", str(Path(endpoint["root"]) / "Cargo.toml"),
        "--target-dir", target, "--example", EXPECTED_EXAMPLE,
    ]
    if features:
        command.extend(("--features", ",".join(features)))
    return {
        "argv": command,
        "environment": {
            "CARGO_INCREMENTAL": "0",
            "CARGO_TERM_COLOR": "never",
            "RUSTFLAGS": "--cfg odg_joined_allocator" if allocator else "",
        },
    }


def prepare_plan(manifest_path: Path, *, allow_built_artifacts: bool = False) -> dict[str, Any]:
    manifest_path = _safe_path(manifest_path.absolute(), "manifest", must_exist=True)
    raw = _raw_manifest(manifest_path)
    base = manifest_path.parent
    build = _build_value(raw)
    cpu = _nonnegative_int(build["cpu"], "manifest.build.cpu")
    baseline = _endpoint_identity(raw, "baseline", base, build, allow_built_artifacts=allow_built_artifacts)
    candidate = _endpoint_identity(raw, "candidate", base, build, allow_built_artifacts=allow_built_artifacts)
    if baseline["root"] == candidate["root"] or baseline["ref"] == candidate["ref"]:
        raise RunFailure("baseline and candidate source identities must differ")
    fixture = _fixture_plan(raw, base)
    for root in (Path(baseline["root"]), Path(candidate["root"])):
        _outside(Path(fixture["path"]), root, "fixture")
    probes = _probe_plan(raw, base, (Path(baseline["root"]), Path(candidate["root"])))
    prepare_receipt = _safe_path(_raw_path(build["prepare_receipt"], "manifest.build.prepare_receipt", base), "prepare receipt")
    receipt = _safe_path(_raw_path(build["receipt"], "manifest.build.receipt", base), "build receipt")
    if prepare_receipt == receipt:
        raise RunFailure("prepare and build receipts must be distinct")
    for root in (Path(baseline["root"]), Path(candidate["root"])):
        _outside(prepare_receipt, root, "prepare receipt")
        _outside(receipt, root, "build receipt")
    plan = {
        "schema": PREPARE_SCHEMA,
        "builder_revision": BUILDER_REVISION,
        "manifest": str(manifest_path),
        "manifest_sha256": _manifest_sha(manifest_path),
        "baseline": baseline,
        "candidate": candidate,
        "fixture": fixture,
        "probe_sources": probes,
        "cpu": cpu,
        "max_load": _protocol_max_load(raw),
        "overlays": [
            {"source": next(item for item in probes if item["name"] == "common")["path"], "destination": EXPECTED_COMMON_DESTINATION},
            {"source": next(item for item in probes if item["name"] == "allocator")["path"], "destination": "crates/litchi-odg/examples/odg_joined_qualification_allocator.rs"},
            {"source": next(item for item in probes if item["name"] == "baseline")["path"], "destination": EXPECTED_OVERLAY_DESTINATION, "endpoint": "baseline"},
            {"source": next(item for item in probes if item["name"] == "candidate")["path"], "destination": EXPECTED_OVERLAY_DESTINATION, "endpoint": "candidate"},
        ],
        "commands": {
            "baseline": {"normal": _command(baseline, allocator=False, cpu=cpu), "allocator": _command(baseline, allocator=True, cpu=cpu)},
            "candidate": {"normal": _command(candidate, allocator=False, cpu=cpu), "allocator": _command(candidate, allocator=True, cpu=cpu)},
        },
        "receipts": {"prepare": str(prepare_receipt), "build": str(receipt)},
    }
    return plan


def prepare(manifest_path: Path) -> dict[str, Any]:
    plan = prepare_plan(manifest_path)
    build = _build_value(_raw_manifest(manifest_path))
    path = _safe_path(_raw_path(build["prepare_receipt"], "manifest.build.prepare_receipt", manifest_path.parent), "prepare receipt")
    _write_fresh_json(path, plan, "prepare receipt")
    return plan


def _load_receipt(path: Path, schema: str) -> dict[str, Any]:
    path = _safe_path(path.absolute(), "receipt", must_exist=True)
    value = _load_json(path)
    if value.get("schema") != schema:
        raise RunFailure(f"receipt schema must be {schema}")
    return value


def validate_prepare_receipt(manifest_path: Path) -> dict[str, Any]:
    raw = _raw_manifest(manifest_path)
    build = _build_value(raw)
    path = _safe_path(_raw_path(build["prepare_receipt"], "manifest.build.prepare_receipt", manifest_path.parent), "prepare receipt", must_exist=True)
    receipt = _load_receipt(path, PREPARE_SCHEMA)
    current = prepare_plan(manifest_path, allow_built_artifacts=True)
    if receipt != current:
        raise RunFailure("prepare receipt does not exactly match current source/config identities")
    return receipt


def _overlay_copy(source: Path, destination: Path, root: Path) -> dict[str, Any]:
    source = _safe_path(source, "overlay source", must_exist=True)
    destination = _safe_path(root / destination, "overlay destination")
    if destination.exists():
        raise RunFailure(f"overlay destination already exists; refusing overwrite: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        with source.open("rb") as input_stream, destination.open("xb") as output_stream:
            shutil.copyfileobj(input_stream, output_stream, 1024 * 1024)
            output_stream.flush()
            os.fsync(output_stream.fileno())
    except OSError as error:
        raise RunFailure(f"cannot create overlay {destination}: {error}") from error
    return _identity(destination)


def _proc_rss(pid: int) -> int:
    try:
        for line in (Path("/proc") / str(pid) / "status").read_text(encoding="ascii").splitlines():
            if line.startswith("VmRSS:"):
                return int(line.split()[1]) * 1024
    except FileNotFoundError:
        return 0
    except (OSError, ValueError, IndexError) as error:
        raise RunFailure(f"cannot inspect RSS for child {pid}: {error}") from error
    return 0


def _descendants(pid: int) -> set[int]:
    result: set[int] = set()
    pending = [pid]
    while pending:
        current = pending.pop()
        children_path = Path("/proc") / str(current) / "task" / str(current) / "children"
        try:
            children = [int(value) for value in children_path.read_text(encoding="ascii").split()]
        except FileNotFoundError:
            children = []
        except (OSError, ValueError) as error:
            raise RunFailure(f"cannot inspect descendants of {current}: {error}") from error
        for child in children:
            if child not in result:
                result.add(child)
                pending.append(child)
    return result


def _quiet_guard(max_load: float, phase: str) -> dict[str, Any]:
    load = load_average()
    processes = process_scan()
    result = {
        "phase": phase,
        "load_1m": load,
        "max_load": max_load,
        "heavy_processes": processes,
        "checked_unix_ns": time.time_ns(),
    }
    if load > max_load:
        raise RunFailure(f"{phase}: one-minute load {load} exceeds {max_load}")
    if processes:
        raise RunFailure(f"{phase}: unrelated heavy processes are running: {processes}")
    return result


def _kill_group(process: subprocess.Popen[bytes]) -> None:
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return
    except OSError as error:
        raise RunFailure(f"cannot terminate child process group: {error}") from error
    deadline = time.monotonic() + 1.0
    while time.monotonic() < deadline and process.poll() is None:
        time.sleep(0.02)
    if process.poll() is None:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass


def _stream_command(
    command: list[str],
    cwd: Path,
    log_dir: Path,
    label: str,
    target_dir: Path,
    cpu: int,
    timeout_s: float = 3600.0,
    environment: Mapping[str, str] | None = None,
    max_output_bytes: int = MAX_BUILD_OUTPUT,
    max_load: float = 1000.0,
) -> dict[str, Any]:
    stdout_path = log_dir / f"{label}.stdout"
    stderr_path = log_dir / f"{label}.stderr"
    if stdout_path.exists() or stderr_path.exists():
        raise RunFailure(f"{label}: refusing stale child logs")
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    env.update({"CARGO_INCREMENTAL": "0", "CARGO_TERM_COLOR": "never"})
    if environment is not None:
        env.update(environment)
    if "allocator" in label:
        env["RUSTFLAGS"] = "--cfg odg_joined_allocator"
    def child_setup() -> None:
        os.setsid()
        if hasattr(os, "sched_setaffinity"):
            os.sched_setaffinity(0, {cpu})
        resource.setrlimit(resource.RLIMIT_AS, (MEMORY_LIMIT_BYTES, MEMORY_LIMIT_BYTES))
    try:
        stdout = stdout_path.open("xb")
        stderr = stderr_path.open("xb")
        process = subprocess.Popen(
            command,
            cwd=str(cwd),
            env=env,
            stdin=subprocess.DEVNULL,
            stdout=stdout,
            stderr=stderr,
            start_new_session=False,
            preexec_fn=child_setup,
        )
    except (OSError, ValueError) as error:
        raise RunFailure(f"{label}: cannot start streamed child: {error}") from error
    started = time.monotonic_ns()
    started_monotonic = time.monotonic()
    violation: str | None = None
    samples: list[dict[str, Any]] = []
    target_before = _directory_bytes(target_dir)
    try:
        while process.poll() is None:
            stdout.flush()
            stderr.flush()
            output_bytes = stdout_path.stat().st_size + stderr_path.stat().st_size
            load = load_average()
            descendants = _descendants(process.pid)
            rss = sum(_proc_rss(pid) for pid in {process.pid, *descendants})
            outside = [item for item in process_scan() if item["pid"] not in {process.pid, *descendants}]
            if output_bytes > max_output_bytes:
                violation = f"child output exceeded {max_output_bytes} bytes"
            elif rss > MEMORY_LIMIT_BYTES:
                violation = f"child RSS exceeded {MEMORY_LIMIT_BYTES} bytes"
            elif load > max_load:
                violation = f"host load exceeded build ceiling: {load} > {max_load}"
            elif outside:
                violation = f"unrelated heavy processes appeared during child: {outside}"
            elif _directory_bytes(target_dir) > 20_000_000_000:
                violation = "target directory exceeded 20 GB"
            elif time.monotonic() - started_monotonic > timeout_s:
                violation = f"child exceeded timeout {timeout_s}s"
            if violation:
                _kill_group(process)
                break
            samples.append({"checked_unix_ns": time.time_ns(), "output_bytes": output_bytes, "rss_bytes": rss,
                            "load_1m": load, "outside_processes": outside})
            time.sleep(0.05)
        return_code = process.wait(timeout=5)
    except (OSError, subprocess.SubprocessError) as error:
        _kill_group(process)
        raise RunFailure(f"{label}: streamed child wait failed: {error}") from error
    finally:
        stdout.flush()
        stderr.flush()
        stdout.close()
        stderr.close()
    if violation:
        raise RunFailure(f"{label}: {violation}")
    deadline = time.monotonic() + DESCENDANT_GRACE_S
    descendants = _descendants(process.pid)
    while descendants and time.monotonic() < deadline:
        time.sleep(0.05)
        descendants = _descendants(process.pid)
    if descendants:
        raise RunFailure(f"{label}: child left descendants running after recheck: {sorted(descendants)}")
    target_after = _directory_bytes(target_dir)
    final_output_bytes = stdout_path.stat().st_size + stderr_path.stat().st_size
    samples.append({
        "checked_unix_ns": time.time_ns(),
        "output_bytes": final_output_bytes,
        "rss_bytes": 0,
        "load_1m": load_average(),
        "outside_processes": process_scan(),
    })
    stdout_sha, stdout_bytes = _sha256_file(stdout_path)
    stderr_sha, stderr_bytes = _sha256_file(stderr_path)
    if return_code != 0:
        raise RunFailure(f"{label}: build child exited {return_code}")
    return {
        "label": label,
        "command": command,
        "cwd": str(cwd),
        "returncode": return_code,
        "wall_ns": time.monotonic_ns() - started,
        "target": {"path": str(target_dir), "bytes_before": target_before, "bytes_after": target_after,
                    "hard_limit_bytes": 20_000_000_000},
        "stdout": {"path": str(stdout_path), "bytes": stdout_bytes, "sha256": stdout_sha},
        "stderr": {"path": str(stderr_path), "bytes": stderr_bytes, "sha256": stderr_sha},
        "monitor": {"samples": samples, "memory_limit_bytes": MEMORY_LIMIT_BYTES,
                     "max_output_bytes": max_output_bytes, "max_load": max_load},
    }


def _directory_bytes(path: Path) -> int:
    if not path.exists() or path.is_symlink():
        return 0
    if path.is_file():
        return path.stat().st_size
    return sum(item.stat().st_size for item in path.rglob("*") if item.is_file() and not item.is_symlink())


def _remove_overlay(root: Path, relative: str) -> None:
    path = _safe_path(root / relative, "overlay cleanup", must_exist=True)
    if not path.is_file() or path.is_symlink():
        raise RunFailure(f"overlay cleanup target is not a regular file: {path}")
    path.unlink()


def build(manifest_path: Path) -> dict[str, Any]:
    plan = validate_prepare_receipt(manifest_path)
    raw = _raw_manifest(manifest_path)
    build_value = _build_value(raw)
    build_receipt = _safe_path(_raw_path(build_value["receipt"], "manifest.build.receipt", manifest_path.parent), "build receipt")
    if build_receipt.exists():
        raise RunFailure(f"build receipt already exists: {build_receipt}")
    logs_dir = build_receipt.parent / "logs"
    _fresh_dir(logs_dir, "build log directory")
    events: list[dict[str, Any]] = []
    endpoint_plans = (("baseline", plan["baseline"]), ("candidate", plan["candidate"]))
    for name, endpoint in endpoint_plans:
        root = Path(endpoint["root"])
        target = Path(endpoint["target_dir"])
        allocator_target = Path(endpoint["allocator_target_dir"])
        _fresh_dir(target, f"{name} target directory")
        _fresh_dir(allocator_target, f"{name} allocator target directory")
        for kind, target_dir in (("normal", target), ("allocator", allocator_target)):
            overlay_paths: list[str] = []
            event: dict[str, Any] | None = None
            try:
                for overlay in plan["overlays"]:
                    if overlay.get("endpoint") not in (None, name):
                        continue
                    relative = overlay["destination"]
                    _overlay_copy(Path(overlay["source"]), root / relative, root)
                    overlay_paths.append(relative)
                command_spec = plan["commands"][name][kind]
                command = command_spec["argv"]
                guard_before = _quiet_guard(plan["max_load"], f"before-{name}-{kind}")
                event = _stream_command(
                    command, root, logs_dir, f"{name}-{kind}", target_dir, plan["cpu"],
                    environment=command_spec["environment"],
                    max_load=plan["max_load"],
                )
                event["environment"] = command_spec["environment"]
                event["guard_before"] = guard_before
                event["guard_after"] = _quiet_guard(plan["max_load"], f"after-{name}-{kind}")
                binary_key = "allocator_binary" if kind == "allocator" else "binary"
                binary = Path(endpoint[binary_key])
                if not binary.is_file() or not os.access(binary, os.X_OK):
                    raise RunFailure(f"{name}/{kind}: expected executable was not produced: {binary}")
                event["artifact"] = binary_identity(binary)
                event["source_before"] = {"ref": endpoint["ref"], "tree": endpoint["tree"]}
            finally:
                for relative in reversed(overlay_paths):
                    _remove_overlay(root, relative)
                shape = _check_worktree_shape(root, name)
                if shape["root"] != endpoint["root"]:
                    raise RunFailure(f"{name}: source root changed during build")
                if event is not None:
                    event["source_after"] = {
                        "ref": _git_output(root, "rev-parse", "--verify", "HEAD"),
                        "tree": _git_output(root, "rev-parse", "--verify", "HEAD^{tree}"),
                        "source_manifest": _tree_manifest(root, endpoint["ref"]),
                        "lock": _identity(Path(endpoint["lock"]["path"])),
                        "toolchain": _identity(Path(endpoint["toolchain"]["path"])),
                        "cargo_config": _identity(Path(endpoint["cargo_config"]["path"])),
                    }
                    events.append(event)
    receipt = {
        "schema": BUILD_SCHEMA,
        "builder_revision": BUILDER_REVISION,
        "status": "passed",
        "manifest": plan["manifest"],
        "manifest_sha256": plan["manifest_sha256"],
        "prepare_receipt": plan["receipts"]["prepare"],
        "prepare_sha256": _manifest_sha(Path(plan["receipts"]["prepare"])),
        "plan": plan,
        "events": events,
        "completed_unix_ns": time.time_ns(),
    }
    _write_fresh_json(build_receipt, receipt, "build receipt")
    return receipt


def _validate_log_identity(value: Any, label: str, expected_dir: Path) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"path", "bytes", "sha256"}, label)
    path = _safe_path(Path(_string(item["path"], f"{label}.path")).absolute(), f"{label}.path", must_exist=True)
    if path.parent != expected_dir:
        raise RunFailure(f"{label} is outside the fresh build log directory")
    actual = _identity(path)
    expected = {"path": str(path), "bytes": _nonnegative_int(item["bytes"], f"{label}.bytes"), "sha256": _digest(item["sha256"], f"{label}.sha256")}
    if actual != expected:
        raise RunFailure(f"{label} was changed after the build")
    return actual


def validate_build_receipt(manifest_path: Path) -> dict[str, Any]:
    manifest_path = _safe_path(manifest_path.absolute(), "manifest", must_exist=True)
    raw = _raw_manifest(manifest_path)
    build = _build_value(raw)
    receipt_path = _safe_path(_raw_path(build["receipt"], "manifest.build.receipt", manifest_path.parent), "build receipt", must_exist=True)
    receipt = _load_receipt(receipt_path, BUILD_SCHEMA)
    plan = validate_prepare_receipt(manifest_path)
    if receipt.get("status") != "passed" or receipt.get("builder_revision") != BUILDER_REVISION:
        raise RunFailure("build receipt is not a passed receipt from this builder")
    if receipt.get("manifest") != plan["manifest"] or receipt.get("manifest_sha256") != plan["manifest_sha256"]:
        raise RunFailure("build receipt is bound to a different manifest")
    if receipt.get("prepare_receipt") != plan["receipts"]["prepare"]:
        raise RunFailure("build receipt is bound to a different prepare receipt")
    prepare_path = Path(plan["receipts"]["prepare"])
    if receipt.get("prepare_sha256") != _manifest_sha(prepare_path):
        raise RunFailure("prepare receipt bytes changed after the build")
    if receipt.get("plan") != plan:
        raise RunFailure("build receipt plan differs from the current prepared plan")
    events = receipt.get("events")
    if not isinstance(events, list) or len(events) != 4:
        raise RunFailure("build receipt must contain exactly four endpoint/build-kind events")
    expected = [(name, kind) for name in ("baseline", "candidate") for kind in ("normal", "allocator")]
    observed: set[tuple[str, str]] = set()
    observed_order: list[tuple[str, str]] = []
    for event_index, event in enumerate(events):
        item = _object(event, "build event")
        _exact_keys(item, {"label", "command", "cwd", "environment", "returncode", "wall_ns", "stdout", "stderr", "monitor", "target", "guard_before", "guard_after", "artifact", "source_before", "source_after"}, "build event")
        label = _string(item["label"], "build event.label")
        parts = label.split("-", 1)
        if len(parts) != 2 or parts[0] not in {"baseline", "candidate"} or parts[1] not in {"normal", "allocator"}:
            raise RunFailure(f"unknown build event label {label!r}")
        identity = (parts[0], parts[1])
        if identity in observed:
            raise RunFailure(f"duplicate build event {label}")
        if identity != tuple(expected[event_index]):
            raise RunFailure("build events are not in deterministic endpoint/kind order")
        observed.add(identity)
        observed_order.append(identity)
        if item["returncode"] != 0 or _positive_int(item["wall_ns"], "build event.wall_ns") <= 0:
            raise RunFailure(f"build event {label} did not pass")
        _validate_log_identity(item["stdout"], f"{label}.stdout", receipt_path.parent / "logs")
        _validate_log_identity(item["stderr"], f"{label}.stderr", receipt_path.parent / "logs")
        monitor = _object(item["monitor"], f"{label}.monitor")
        _exact_keys(monitor, {"samples", "memory_limit_bytes", "max_output_bytes", "max_load"}, f"{label}.monitor")
        if not isinstance(monitor["samples"], list) or not monitor["samples"]:
            raise RunFailure(f"{label}: live monitor samples are missing")
        for sample in monitor["samples"]:
            row = _object(sample, f"{label}.monitor.sample")
            _exact_keys(row, {"checked_unix_ns", "output_bytes", "rss_bytes", "load_1m", "outside_processes"}, f"{label}.monitor.sample")
            _positive_int(row["checked_unix_ns"], "monitor.checked_unix_ns")
            if type(row["output_bytes"]) is not int or row["output_bytes"] < 0 or row["output_bytes"] > MAX_BUILD_OUTPUT:
                raise RunFailure(f"{label}: output monitor exceeded its bound")
            if type(row["rss_bytes"]) is not int or row["rss_bytes"] < 0 or row["rss_bytes"] > MEMORY_LIMIT_BYTES:
                raise RunFailure(f"{label}: RSS monitor exceeded its bound")
            if row["outside_processes"] != []:
                raise RunFailure(f"{label}: unrelated process appeared in monitor")
        if (monitor["memory_limit_bytes"] != MEMORY_LIMIT_BYTES
                or monitor["max_output_bytes"] != MAX_BUILD_OUTPUT
                or monitor["max_load"] != plan["max_load"]):
            raise RunFailure(f"{label}: monitor limits differ from the fixed builder limits")
        for guard_label in ("guard_before", "guard_after"):
            guard = _object(item[guard_label], f"{label}.{guard_label}")
            _exact_keys(guard, {"phase", "load_1m", "max_load", "heavy_processes", "checked_unix_ns"},
                        f"{label}.{guard_label}")
            if guard["max_load"] != plan["max_load"] or guard["heavy_processes"] != []:
                raise RunFailure(f"{label}: quiet guard identity or process list changed")
        target = _object(item["target"], f"{label}.target")
        _exact_keys(target, {"path", "bytes_before", "bytes_after", "hard_limit_bytes"}, f"{label}.target")
        endpoint_target = plan[identity[0]]["allocator_target_dir" if identity[1] == "allocator" else "target_dir"]
        if target["path"] != endpoint_target or target["hard_limit_bytes"] != 20_000_000_000:
            raise RunFailure(f"{label}: target identity differs from its endpoint")
        artifact = _object(item["artifact"], f"{label}.artifact")
        _exact_keys(artifact, {"path", "bytes", "sha256"}, f"{label}.artifact")
        artifact_path = _safe_path(Path(_string(artifact["path"], f"{label}.artifact.path")).absolute(), f"{label}.artifact.path", must_exist=True)
        expected_artifact = plan[identity[0]]["allocator_binary" if identity[1] == "allocator" else "binary"]
        if str(artifact_path) != expected_artifact:
            raise RunFailure(f"{label}: artifact path differs from the prepared command")
        if _identity(artifact_path) != artifact:
            raise RunFailure(f"{label}: built artifact bytes changed")
        expected_command = plan[identity[0]]["commands"][identity[0]][identity[1]]
        if item["command"] != expected_command["argv"]:
            raise RunFailure(f"{label}: command differs from the prepared command")
        if item["environment"] != expected_command["environment"]:
            raise RunFailure(f"{label}: build environment differs from the prepared command")
        if item["cwd"] != plan[identity[0]]["root"]:
            raise RunFailure(f"{label}: command cwd differs from the endpoint root")
        source_before = _object(item["source_before"], f"{label}.source_before")
        _exact_keys(source_before, {"ref", "tree"}, f"{label}.source_before")
        if source_before != {"ref": plan[identity[0]]["ref"], "tree": plan[identity[0]]["tree"]}:
            raise RunFailure(f"{label}: source identity is not bound to its endpoint")
        source_after = _object(item["source_after"], f"{label}.source_after")
        _exact_keys(source_after, {"ref", "tree", "source_manifest", "lock", "toolchain", "cargo_config"},
                    f"{label}.source_after")
        endpoint = plan[identity[0]]
        if source_after["ref"] != endpoint["ref"] or source_after["tree"] != endpoint["tree"]:
            raise RunFailure(f"{label}: source changed during the build")
        if source_after["source_manifest"] != endpoint["source_manifest"]:
            raise RunFailure(f"{label}: complete source manifest changed during the build")
        for field in ("lock", "toolchain", "cargo_config"):
            if source_after[field] != endpoint[field]:
                raise RunFailure(f"{label}: {field} changed during the build")
    if observed_order != expected:
        raise RunFailure("build receipt event set is incomplete")
    return receipt


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True, type=Path)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--prepare", action="store_true")
    group.add_argument("--build", action="store_true")
    group.add_argument("--validate", action="store_true")
    args = parser.parse_args(argv)
    try:
        if args.prepare:
            prepare(args.manifest)
        elif args.build:
            build(args.manifest)
        else:
            validate_build_receipt(args.manifest)
    except (RunFailure, OSError, ValueError) as error:
        print(str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
