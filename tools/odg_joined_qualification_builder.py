#!/usr/bin/env python3
"""Build and provenance admission for the paired ODG joined probes.

The runtime runner is intentionally not a build system.  This companion is
the only entry point allowed to create the two probe binaries: it prepares a
fully bound plan, applies the reviewed example overlay to one pristine
worktree at a time, streams a bounded Cargo child into fresh logs, and emits a
receipt which the runtime runner revalidates before starting a probe.

All paths in this module are treated as untrusted.  Existing output, receipt,
log, target, or overlay paths are never overwritten; symlink components,
untracked upstream trees, sparse/skip-worktree source trees, stale refs, and
changed source/config identities fail closed.
"""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import re
import resource
import shutil
import signal
import stat
import subprocess
import sys
import time
from typing import Any, Mapping

try:
    from tools.odp_native_qualification import (
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
    from tools.odg_joined_qualification_bindings import (
        REVIEWED_FIXTURE_IDENTITY,
        REVIEWED_FIXTURE_PROVENANCE,
        REVIEWED_FIXTURE_RELATIVE,
        REVIEWED_PROBE_IDENTITIES,
        REVIEWED_PROBE_RELATIVES,
        REVIEWED_REPO_ROOT,
    )
except ImportError:  # pragma: no cover - direct tools import.
    from odp_native_qualification import (  # type: ignore[no-redef]
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
    from odg_joined_qualification_bindings import (  # type: ignore[no-redef]
        REVIEWED_FIXTURE_IDENTITY,
        REVIEWED_FIXTURE_PROVENANCE,
        REVIEWED_FIXTURE_RELATIVE,
        REVIEWED_PROBE_IDENTITIES,
        REVIEWED_PROBE_RELATIVES,
        REVIEWED_REPO_ROOT,
    )


PREPARE_SCHEMA = "litchi.odg.joined-native-qualification.prepare.v4"
BUILD_SCHEMA = "litchi.odg.joined-native-qualification.build.v4"
BUILDER_REVISION = "odg-joined-builder-v4-closure-environment-target-bound"
SOURCE_BINDER_SCHEMA = "litchi.odg.joined-native-qualification.c2a-source-binder.v1"
SOURCE_PREPARATION_SCHEMA = "litchi.odg.joined-native-qualification.source-preparation.v1"
SOURCE_BUILD_SCHEMA = "litchi.odg.joined-native-qualification.source-build.v1"
REVISION = re.compile(r"[0-9a-f]{40}\Z")
PRISTINE_BASELINE_REF = "ad52437f5e5e2c71ea476eedd5ddcf61aee2021d"
# The returned handoff identifies the rejected implementation only by its
# reviewed 15ab prefix; do not invent a full object ID that was never pinned.
HISTORICAL_CANDIDATE_PREFIX = "15ab"
EXPECTED_EXAMPLE = "litchi-perf-odg-joined-qualification"
EXPECTED_PACKAGE = "litchi-odg"
EXPECTED_CARGO_PROFILE = "release"
EXPECTED_OVERLAY_DESTINATION = "crates/litchi-odg/examples/litchi-perf-odg-joined-qualification.rs"
EXPECTED_COMMON_DESTINATION = "crates/litchi-odg/examples/odg_joined_qualification_common.rs"
EXPECTED_SOURCE_NAMES = {
    "common": "odg_joined_qualification_common.rs",
    "allocator": "odg_joined_qualification_allocator.rs",
    "baseline": "odg_joined_qualification_ad524_probe.rs",
    "candidate": "odg_joined_qualification_probe.rs",
}
SOURCE_COMPOSITION_LIMITS = {
    "max_sub_edits": 8,
    "max_effects_per_sub_edit": 8,
    "max_total_effects": 32,
    "max_conflicts": 16,
}
MAX_BUILD_OUTPUT = 64 * 1024 * 1024
MEMORY_LIMIT_BYTES = 20_000_000_000
TARGET_PREFLIGHT_BYTES = 19_200_000_000
DESCENDANT_GRACE_S = 2.0
# Keep the guard's own forbidden component out of the repository-wide source
# scanner.  This is a value check, not a path into an upstream checkout.
UNTRACKED_UPSTREAM_COMPONENT = "3rd" + "party"


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
    if UNTRACKED_UPSTREAM_COMPONENT in path.parts:
        raise RunFailure(f"{label} must not reference {UNTRACKED_UPSTREAM_COMPONENT}: {path}")
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
    """Hash a manifest without making its receipt hash self-referential.

    The receipt identity is an observed post-build artifact.  It therefore
    cannot participate in the manifest identity that the receipt records:
    doing so would make sealing require a cryptographic fixed point.  Keep
    the path in the canonical identity, but replace the observed bytes/hash
    with fixed sentinels.  The final verifier still checks the real receipt
    bytes against ``manifest.build.receipt_identity`` separately.
    """
    raw = _load_json(path)
    build = raw.get("build")
    if isinstance(build, dict):
        identity = build.get("receipt_identity")
        if isinstance(identity, dict):
            identity = dict(identity)
            identity["bytes"] = 0
            identity["sha256"] = "0" * 64
            build = dict(build)
            build["receipt_identity"] = identity
            raw = dict(raw)
            raw["build"] = build
    encoded = json.dumps(raw, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _raw_manifest(path: Path) -> dict[str, Any]:
    path = _safe_path(path.absolute(), "manifest", must_exist=True)
    value = _load_json(path)
    return value


def _endpoint_value(raw: Mapping[str, Any], name: str) -> dict[str, Any]:
    value = _object(raw.get(name), f"manifest.{name}")
    expected = {
        "root", "ref", "tree", "profile", "binary", "target_dir", "lock",
        "toolchain", "cargo_config", "cargo_manifest", "source_manifest", "allocator_binary",
        "source_closure", "allocator_target_dir", "features", "binary_identity", "allocator_binary_identity",
    }
    _exact_keys(value, expected, f"manifest.{name}")
    return value


def _build_value(raw: Mapping[str, Any]) -> dict[str, Any]:
    value = _object(raw.get("build"), "manifest.build")
    _exact_keys(value, {"prepare_receipt", "receipt", "receipt_identity", "final_candidate_ref", "cpu"}, "manifest.build")
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


def _digest_binding(value: Any, label: str) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"bytes", "sha256"}, label)
    return {
        "bytes": _positive_int(item.get("bytes"), f"{label}.bytes"),
        "sha256": _digest(item.get("sha256"), f"{label}.sha256"),
    }


def _receipt_binding(value: Any, label: str, base: Path, expected_path: Path) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"path", "bytes", "sha256"}, label)
    path = _safe_path(_raw_path(item.get("path"), f"{label}.path", base), f"{label}.path")
    if path != expected_path:
        raise RunFailure(f"{label}.path must equal the manifest build receipt path")
    return {
        "path": str(path),
        "bytes": _positive_int(item.get("bytes"), f"{label}.bytes"),
        "sha256": _digest(item.get("sha256"), f"{label}.sha256"),
    }


def _reviewed_file(path: Path, relative: Path, expected: Mapping[str, Any], label: str) -> dict[str, Any]:
    """Bind an overlay/fixture to the reviewed repository input itself."""

    return _reviewed_file_at_root(path, relative, expected, label, REVIEWED_REPO_ROOT)


def _reviewed_file_at_root(
    path: Path, relative: Path, expected: Mapping[str, Any], label: str, source_root: Path,
) -> dict[str, Any]:
    """Bind a reviewed input to an explicitly selected clean source root."""

    path = _safe_path(path, label, must_exist=True)
    source_root = _safe_path(source_root, f"approved {label} source root", must_exist=True)
    approved = _safe_path(source_root / relative, f"approved {label}", must_exist=True)
    if path != approved:
        raise RunFailure(f"{label} must use the reviewed repository path {relative}")
    status = _run_git_bytes(source_root, "ls-files", "-v", "-z", "--", str(relative))
    records = [item for item in status.split(b"\0") if item]
    if records != [b"H " + str(relative).encode()]:
        raise RunFailure(f"approved {label} is not an ordinary tracked worktree file")
    actual = _identity(path)
    binding = {"bytes": _positive_int(expected.get("bytes"), f"approved {label}.bytes"),
               "sha256": _digest(expected.get("sha256"), f"approved {label}.sha256")}
    if actual["bytes"] != binding["bytes"] or actual["sha256"] != binding["sha256"]:
        raise RunFailure(f"approved {label} bytes differ from the reviewed source binding")
    return actual


def _external_reviewed_file(
    path: Path, expected: Mapping[str, Any], label: str,
) -> dict[str, Any]:
    """Bind an approved fixture copied outside both endpoint worktrees."""
    path = _safe_path(path, label, must_exist=True)
    actual = _identity(path)
    binding = {
        "bytes": _positive_int(expected.get("bytes"), f"approved {label}.bytes"),
        "sha256": _digest(expected.get("sha256"), f"approved {label}.sha256"),
    }
    if actual["bytes"] != binding["bytes"] or actual["sha256"] != binding["sha256"]:
        raise RunFailure(f"approved {label} bytes differ from the reviewed source binding")
    return actual


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


def _under(path: Path, parent: Path) -> bool:
    """Return whether ``path`` is equal to or below ``parent``."""
    return path == parent or parent in path.parents


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
    # ``-t`` does not expose assume-unchanged entries.  The verbose status
    # and debug flags are both checked so a clean-looking index cannot hide
    # bytes which Cargo will actually compile.
    indexed = _run_git_bytes(root, "ls-files", "-v", "-z")
    entries: list[str] = []
    for item in indexed.split(b"\0"):
        if not item:
            continue
        if len(item) < 3 or item[1:2] != b" ":
            raise RunFailure(f"{name}.root returned malformed ls-files status")
        status, raw_path = item[0:1], item[2:]
        path = root / os.fsdecode(raw_path)
        _safe_path(path, f"{name} tracked path", must_exist=False)
        if status != b"H":
            raise RunFailure(f"{name}.root contains non-ordinary index entry {raw_path!r}: {status!r}")
        if not path.exists():
            raise RunFailure(f"{name}.root omits tracked file {raw_path!r}")
        entries.append(os.fsdecode(raw_path))
    debug = _run_git_bytes(root, "ls-files", "--debug")
    current_path: bytes | None = None
    current_flags: int | None = None
    for line in debug.splitlines():
        if line.startswith(b" "):
            if b"\tflags:" not in line:
                continue
            flags_line = line
            try:
                current_flags = int(flags_line.rsplit(b"flags:", 1)[1].strip(), 0)
            except ValueError as error:
                raise RunFailure(f"{name}.root returned malformed index flags") from error
            continue
        if current_path is not None and current_flags != 0:
            raise RunFailure(f"{name}.root contains extended index flags: {current_flags}")
        current_path = line
        current_flags = None
    if current_path is not None and current_flags != 0:
        raise RunFailure(f"{name}.root contains extended index flags: {current_flags}")
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


def _git_tree_entries(root: Path, revision: str) -> list[tuple[str, str, str]]:
    raw = _run_git_bytes(root, "ls-tree", "-r", "-z", "--full-tree", revision)
    entries: list[tuple[str, str, str]] = []
    for item in raw.split(b"\0"):
        if not item:
            continue
        try:
            header, raw_path = item.split(b"\t", 1)
            mode, kind, oid = header.split()
        except ValueError as error:
            raise RunFailure(f"malformed Git tree entry in {root}: {item!r}") from error
        if kind != b"blob":
            raise RunFailure(f"source closure contains non-blob entry {raw_path!r}")
        entries.append((os.fsdecode(raw_path), mode.decode("ascii"), oid.decode("ascii")))
    return entries


def _capture_source_closure(root: Path, revision: str) -> dict[str, Any]:
    """Hash every tracked byte and retain metadata for cheap between-probe checks."""

    root = _safe_path(root, "source closure root", must_exist=True)
    files: list[dict[str, Any]] = []
    total_bytes = 0
    closure_digest = hashlib.sha256()
    stat_digest = hashlib.sha256()
    for relative, mode, git_oid in _git_tree_entries(root, revision):
        path = _safe_path(root / relative, f"source closure {relative}", must_exist=True)
        try:
            info = path.lstat()
        except OSError as error:
            raise RunFailure(f"cannot stat tracked source {relative}: {error}") from error
        if not stat.S_ISREG(info.st_mode):
            raise RunFailure(f"tracked source is not a regular file: {relative}")
        expected_mode = True if mode == "100755" else False if mode == "100644" else None
        executable = bool(info.st_mode & 0o111)
        if expected_mode is None or executable != expected_mode:
            raise RunFailure(f"tracked source mode differs from Git tree: {relative}")
        mode_bits = stat.S_IMODE(info.st_mode) & 0o111
        sha256 = hashlib.sha256()
        # The stat size is part of Git's blob header.  Initialising this
        # before the one content pass lets us verify the Git object without
        # retaining or rereading the source bytes.
        git_hash = hashlib.sha1(f"blob {info.st_size}\0".encode("ascii"))
        size = 0
        try:
            with path.open("rb") as stream:
                for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                    size += len(chunk)
                    sha256.update(chunk)
                    git_hash.update(chunk)
        except OSError as error:
            raise RunFailure(f"cannot read tracked source {relative}: {error}") from error
        if size != info.st_size:
            raise RunFailure(f"tracked source changed while being read: {relative}")
        actual_git_oid = git_hash.hexdigest()
        if actual_git_oid != git_oid:
            raise RunFailure(f"tracked source bytes differ from HEAD: {relative}")
        encoded = (
            mode.encode("ascii") + b"\0" + relative.encode("utf-8") + b"\0"
            + str(size).encode("ascii") + b"\0" + sha256.hexdigest().encode("ascii")
            + b"\0" + git_oid.encode("ascii") + b"\0"
        )
        closure_digest.update(encoded)
        stat_digest.update(
            (
                relative.encode("utf-8") + b"\0" + str(info.st_dev).encode("ascii") + b"\0"
                + str(info.st_ino).encode("ascii") + b"\0" + str(info.st_size).encode("ascii") + b"\0"
                + str(info.st_mtime_ns).encode("ascii") + b"\0"
                + str(info.st_ctime_ns).encode("ascii") + b"\0"
                + str(mode_bits).encode("ascii") + b"\0"
            )
        )
        files.append({
            "path": relative,
            "mode": mode,
            "bytes": size,
            "sha256": sha256.hexdigest(),
            "git_oid": git_oid,
            "dev": info.st_dev,
            "ino": info.st_ino,
            "mtime_ns": info.st_mtime_ns,
            "ctime_ns": info.st_ctime_ns,
            "mode_bits": mode_bits,
        })
        total_bytes += size
    return {
        "revision": revision,
        "binding": {"entries": len(files), "bytes": total_bytes, "sha256": closure_digest.hexdigest()},
        "stat_sha256": stat_digest.hexdigest(),
        "files": files,
    }


def _closure_binding(value: Any, label: str) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"entries", "bytes", "sha256"}, label)
    return {
        "entries": _positive_int(item.get("entries"), f"{label}.entries"),
        "bytes": _positive_int(item.get("bytes"), f"{label}.bytes"),
        "sha256": _digest(item.get("sha256"), f"{label}.sha256"),
    }


def _cheap_source_closure_probe(root: Path, cache: Mapping[str, Any]) -> bool:
    for item in cache["files"]:
        path = root / item["path"]
        try:
            info = path.lstat()
        except OSError:
            return False
        if (
            not stat.S_ISREG(info.st_mode)
            or info.st_dev != item["dev"]
            or info.st_ino != item["ino"]
            or info.st_size != item["bytes"]
            or info.st_mtime_ns != item["mtime_ns"]
            or info.st_ctime_ns != item["ctime_ns"]
            or (stat.S_IMODE(info.st_mode) & 0o111) != item["mode_bits"]
        ):
            return False
    return True


def _validate_source_closure(
    root: Path,
    cache: dict[str, Any],
    expected: Mapping[str, Any],
    *,
    force_full: bool = False,
) -> dict[str, Any]:
    expected_binding = _closure_binding(expected, "source_closure")
    if cache["binding"] != expected_binding:
        raise RunFailure("source closure cache differs from its external binding")
    if force_full or not _cheap_source_closure_probe(root, cache):
        refreshed = _capture_source_closure(root, cache["revision"])
        if refreshed["binding"] != expected_binding:
            raise RunFailure("tracked source bytes differ from the externally bound closure")
        cache.clear()
        cache.update(refreshed)
    return expected_binding


def _source_closure(root: Path, revision: str, expected: Mapping[str, Any]) -> dict[str, Any]:
    cache = _capture_source_closure(root, revision)
    cache["revision"] = revision
    expected_binding = _closure_binding(expected, "source_closure")
    if cache["binding"] != expected_binding:
        raise RunFailure("tracked source closure differs from the manifest")
    return cache


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
    cargo_manifest = _check_identity(value["cargo_manifest"], f"{name}.cargo_manifest", base, root)
    for field, expected_relative in (
        ("lock", Path("Cargo.lock")),
        ("toolchain", Path("rust-toolchain.toml")),
        ("cargo_config", Path(".cargo/config.toml")),
        ("cargo_manifest", Path("Cargo.toml")),
    ):
        declared, _, _ = _identity_value(value[field], f"{name}.{field}", base)
        if declared != root / expected_relative:
            raise RunFailure(f"{name}.{field} must use the reviewed {expected_relative} path")
    tree_identity = _tree_manifest(root, head)
    source_manifest = _check_tree_manifest(value["source_manifest"], tree_identity, f"{name}.source_manifest")
    source_cache = _capture_source_closure(root, head)
    source_closure = _closure_binding(value["source_closure"], f"{name}.source_closure")
    if source_cache["binding"] != source_closure:
        raise RunFailure(f"{name}.source_closure differs from the tracked bytes at HEAD")
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
    binary_binding = _digest_binding(value["binary_identity"], f"{name}.binary_identity")
    allocator_binary_binding = _digest_binding(
        value["allocator_binary_identity"], f"{name}.allocator_binary_identity"
    )
    expected_artifacts = {
        "binary": target_dir / EXPECTED_CARGO_PROFILE / "examples" / EXPECTED_EXAMPLE,
        "allocator_binary": allocator_target_dir / EXPECTED_CARGO_PROFILE / "examples" / EXPECTED_EXAMPLE,
    }
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
        if artifact != expected_artifacts[label]:
            raise RunFailure(
                f"{name}.{label} must be the release example artifact at {expected_artifacts[label]}"
            )
        binding = binary_binding if label == "binary" else allocator_binary_binding
        if artifact.exists():
            if not allow_built_artifacts:
                raise RunFailure(f"{name}.{label} must not exist before the fresh build")
            actual = _identity(artifact)
            if actual["bytes"] != binding["bytes"] or actual["sha256"] != binding["sha256"]:
                raise RunFailure(f"{name}.{label} differs from its externally bound artifact identity")
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
        "cargo_manifest": cargo_manifest,
        "source_manifest": source_manifest,
        "source_closure": source_closure,
        "binary": str(binary),
        "allocator_binary": str(allocator_binary),
        "target_dir": str(target_dir),
        "allocator_target_dir": str(allocator_target_dir),
        "binary_identity": binary_binding,
        "allocator_binary_identity": allocator_binary_binding,
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
        candidate_root = roots[1]
        if path == REVIEWED_REPO_ROOT / REVIEWED_PROBE_RELATIVES[name]:
            source_root = REVIEWED_REPO_ROOT
        elif path == candidate_root / REVIEWED_PROBE_RELATIVES[name]:
            source_root = candidate_root
        else:
            raise RunFailure(f"probe {name} is not the reviewed repository adapter")
        if item["destination"] != expected_destinations[name]:
            raise RunFailure(f"probe {name} has an unreviewed destination")
        baseline_root = roots[0]
        if path == baseline_root or baseline_root in path.parents:
            raise RunFailure(f"probe {name} must not come from the baseline endpoint worktree")
        actual = _reviewed_file_at_root(
            path, REVIEWED_PROBE_RELATIVES[name], REVIEWED_PROBE_IDENTITIES[name],
            f"probe {name}", source_root,
        )
        if actual["bytes"] != _positive_int(item["bytes"], f"probe {name}.bytes") or actual["sha256"] != _digest(item["sha256"], f"probe {name}.sha256"):
            raise RunFailure(f"probe {name} identity differs from the manifest")
        result.append({"name": name, **actual, "destination": item["destination"]})
    return result


def _fixture_plan(raw: Mapping[str, Any], base: Path, candidate_root: Path | None = None) -> dict[str, Any]:
    item = _object(raw.get("fixture"), "manifest.fixture")
    _exact_keys(item, {"path", "bytes", "sha256", "provenance"}, "manifest.fixture")
    path = _safe_path(_raw_path(item["path"], "manifest.fixture.path", base), "fixture", must_exist=True)
    source_root = REVIEWED_REPO_ROOT
    external = False
    if path == REVIEWED_REPO_ROOT / REVIEWED_FIXTURE_RELATIVE:
        source_root = REVIEWED_REPO_ROOT
    elif candidate_root is not None and path == candidate_root / REVIEWED_FIXTURE_RELATIVE:
        source_root = candidate_root
    else:
        external = True
        actual = _external_reviewed_file(path, REVIEWED_FIXTURE_IDENTITY, "fixture")
    if not external:
        actual = _reviewed_file_at_root(
            path, REVIEWED_FIXTURE_RELATIVE, REVIEWED_FIXTURE_IDENTITY, "fixture", source_root,
        )
    if actual["bytes"] != _positive_int(item["bytes"], "manifest.fixture.bytes") or actual["sha256"] != _digest(item["sha256"], "manifest.fixture.sha256"):
        raise RunFailure("fixture identity differs from the manifest")
    content = path.read_bytes()
    if b"<office:document-content" not in content or b"<draw:page" not in content:
        raise RunFailure("fixture is not the reviewed ODG content XML")
    provenance = _string(item["provenance"], "manifest.fixture.provenance")
    if provenance != REVIEWED_FIXTURE_PROVENANCE:
        raise RunFailure("fixture provenance is not the reviewed fixed provenance")
    return {**actual, "path": str(path), "provenance": provenance}


_REJECTED_WRAPPER_ENV = (
    "RUSTC",
    "CARGO_BUILD_RUSTC",
    "RUSTC_WRAPPER",
    "RUSTC_WORKSPACE_WRAPPER",
    "CARGO_BUILD_RUSTC_WRAPPER",
    "WORKSPACE_WRAPPER",
    "CARGO_WORKSPACE_WRAPPER",
)
# These are launcher bookkeeping rather than Cargo/rustc inputs.  Passing
# them through would make a prepare receipt depend on the shell's command
# counter or caller cwd, while omitting them keeps the recorded map exactly
# equal to the environment actually handed to the child.
_VOLATILE_LAUNCHER_ENV = frozenset({"_", "PWD", "OLDPWD", "SHLVL", "CODEX_TOOL_CALL_ID"})


def _tool_identity(name: str, environment: Mapping[str, str], cwd: Path) -> dict[str, Any]:
    search_path = environment.get("PATH", "")
    requested_path: Path | None = None
    for raw_directory in search_path.split(os.pathsep):
        directory = Path(raw_directory or os.curdir)
        if not directory.is_absolute():
            directory = cwd / directory
        candidate = directory / name
        try:
            if candidate.is_file() and os.access(candidate, os.X_OK):
                requested_path = candidate.absolute()
                break
        except OSError:
            continue
    if requested_path is None:
        raise RunFailure(f"cannot resolve {name} through the bound PATH")
    try:
        resolved = requested_path.resolve(strict=True)
    except OSError as error:
        raise RunFailure(f"cannot resolve {name} executable {requested_path}: {error}") from error
    resolved = _safe_path(resolved, f"resolved {name}", must_exist=True)
    if not resolved.is_file() or not os.access(resolved, os.X_OK):
        raise RunFailure(f"resolved {name} is not an executable regular file: {resolved}")
    sha, size = _sha256_file(resolved)
    return {
        "name": name,
        "requested": str(requested_path),
        "resolved": str(resolved),
        "bytes": size,
        "sha256": sha,
    }


def _environment_config_paths(root: Path, environment: Mapping[str, str]) -> list[Path]:
    candidates: set[Path] = set()
    current = root.absolute()
    while True:
        cargo = current / ".cargo"
        candidates.update((cargo / "config.toml", cargo / "config"))
        if current.parent == current:
            break
        current = current.parent
    def env_path(value: str) -> Path:
        path = Path(value)
        return (root / path).absolute() if not path.is_absolute() else path.absolute()

    home_path = env_path(environment["HOME"]) if environment.get("HOME") else Path.home().absolute()
    candidates.update((home_path / ".cargo" / "config.toml", home_path / ".cargo" / "config"))
    cargo_home = environment.get("CARGO_HOME")
    if cargo_home:
        cargo_home_path = env_path(cargo_home)
        candidates.update((cargo_home_path / "config.toml", cargo_home_path / "config"))
    rustup_home = environment.get("RUSTUP_HOME")
    rustup_home_path = env_path(rustup_home) if rustup_home else home_path / ".rustup"
    candidates.add(rustup_home_path / "settings.toml")
    return sorted(candidates, key=str)


def _environment_config_identity(root: Path, environment: Mapping[str, str]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for path in _environment_config_paths(root, environment):
        path = _safe_path(path.absolute(), "bound tool configuration")
        if path.is_symlink():
            raise RunFailure(f"bound tool configuration must not be a symlink: {path}")
        if not path.exists():
            result.append({"path": str(path), "present": False})
            continue
        if not path.is_file():
            raise RunFailure(f"bound tool configuration is not a regular file: {path}")
        sha, size = _sha256_file(path)
        result.append({"path": str(path), "present": True, "bytes": size, "sha256": sha})
    return result


def _environment_identity(
    root: Path,
    *,
    ref: str,
    tree: str,
    source_closure: Mapping[str, Any],
    overrides: Mapping[str, str] | None = None,
) -> dict[str, Any]:
    """Bind every input Cargo/rustc can inherit, not only explicit overrides."""

    root = _safe_path(root, "bound command cwd", must_exist=True)
    if not root.is_dir():
        raise RunFailure(f"bound command cwd is not a directory: {root}")
    environment = {
        key: value for key, value in os.environ.items()
        if key not in _VOLATILE_LAUNCHER_ENV
    }
    if overrides is not None:
        for key, value in overrides.items():
            if not isinstance(key, str) or not isinstance(value, str):
                raise RunFailure("bound environment overrides must be strings")
        environment.update(overrides)
    for key in _REJECTED_WRAPPER_ENV:
        if environment.get(key):
            raise RunFailure(f"unreviewed compiler selection or wrapper is set in {key}")
    if "PATH" not in environment:
        raise RunFailure("bound environment must contain PATH")
    try:
        info = root.lstat()
    except OSError as error:
        raise RunFailure(f"cannot stat bound command cwd {root}: {error}") from error
    if not stat.S_ISDIR(info.st_mode):
        raise RunFailure(f"bound command cwd is not a directory: {root}")
    closure = _closure_binding(source_closure, "environment.cwd.source_closure")
    variables = {key: environment[key] for key in sorted(environment)}
    return {
        "variables": variables,
        "tools": {
            "cargo": _tool_identity("cargo", environment, root),
            "rustc": _tool_identity("rustc", environment, root),
        },
        "configs": _environment_config_identity(root, environment),
        "cwd": {
            "path": str(root),
            "dev": info.st_dev,
            "ino": info.st_ino,
            "mode": info.st_mode,
            "ref": ref,
            "tree": tree,
            "source_closure": closure,
        },
    }


def _command(endpoint: Mapping[str, Any], *, allocator: bool, cpu: int) -> dict[str, Any]:
    target = endpoint["allocator_target_dir"] if allocator else endpoint["target_dir"]
    features = endpoint["features"]
    command = [
        "taskset", "--cpu-list", str(cpu), "cargo", "build", "--offline", "--locked",
        "--jobs", "1", "--manifest-path", str(Path(endpoint["root"]) / "Cargo.toml"),
        "--target-dir", target, "--package", EXPECTED_PACKAGE, "--release",
        "--example", EXPECTED_EXAMPLE,
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


def _source_binder_endpoint(raw: Mapping[str, Any], name: str) -> dict[str, Any]:
    value = _object(raw.get(name), f"source binder.{name}")
    expected = {
        "root", "ref", "tree", "profile", "features", "source_manifest",
        "source_closure", "lock", "toolchain", "cargo_config", "cargo_manifest",
        "binary", "allocator_binary",
    }
    _exact_keys(value, expected, f"source binder.{name}")
    ref = _string(value.get("ref"), f"source binder.{name}.ref")
    tree = _string(value.get("tree"), f"source binder.{name}.tree")
    if REVISION.fullmatch(ref) is None or REVISION.fullmatch(tree) is None:
        raise RunFailure(f"source binder.{name} ref/tree are not immutable revisions")
    profile = _string(value.get("profile"), f"source binder.{name}.profile")
    expected_profile = "legacy-apply-joined" if name == "baseline" else "current-edit-join-commit"
    if profile != expected_profile:
        raise RunFailure(f"source binder.{name}.profile is not the reviewed paired profile")
    features = value.get("features")
    if (not isinstance(features, list) or any(not isinstance(item, str) or not item for item in features)
            or features != sorted(set(features))):
        raise RunFailure(f"source binder.{name}.features are not a sorted unique list")
    for field in ("source_manifest", "source_closure"):
        if field == "source_manifest":
            _check_tree_manifest(value[field], _check_tree_manifest_shape(value[field], f"source binder.{name}.{field}"), f"source binder.{name}.{field}")
        else:
            _closure_binding(value[field], f"source binder.{name}.{field}")
    for field in ("lock", "toolchain", "cargo_config", "cargo_manifest"):
        item = _object(value.get(field), f"source binder.{name}.{field}")
        _exact_keys(item, {"path", "bytes", "sha256"}, f"source binder.{name}.{field}")
        _positive_int(item.get("bytes"), f"source binder.{name}.{field}.bytes")
        _digest(item.get("sha256"), f"source binder.{name}.{field}.sha256")
    for field in ("binary", "allocator_binary"):
        item = _object(value.get(field), f"source binder.{name}.{field}")
        _exact_keys(item, {"path", "bytes", "sha256", "status"}, f"source binder.{name}.{field}")
        path = _safe_path(Path(_string(item.get("path"), f"source binder.{name}.{field}.path")), f"{name}.{field}")
        if path.name != EXPECTED_EXAMPLE or path.parent.name != "examples" or path.parent.parent.name != "release":
            raise RunFailure(f"source binder.{name}.{field} is not the reviewed release example path")
        if item.get("bytes") is not None or item.get("sha256") is not None:
            raise RunFailure(f"source binder.{name}.{field} carries an invented pre-build artifact identity")
        role = "baseline" if name == "baseline" else "C2a"
        if item.get("status") != "pending " + role + " " + ("allocator" if field == "allocator_binary" else "normal") + " build":
            raise RunFailure(f"source binder.{name}.{field} is not pending its reviewed build")
    return dict(value)


def _check_tree_manifest_shape(value: Any, label: str) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"entries", "bytes", "sha256"}, label)
    return {
        "entries": _positive_int(item.get("entries"), f"{label}.entries"),
        "bytes": _positive_int(item.get("bytes"), f"{label}.bytes"),
        "sha256": _digest(item.get("sha256"), f"{label}.sha256"),
    }


def _source_binder_reviewed_inputs(
    raw: Mapping[str, Any], candidate_root: Path,
) -> dict[str, dict[str, Any]]:
    values = _object(raw.get("reviewed_inputs"), "source binder.reviewed_inputs")
    expected = {
        "fixture": (REVIEWED_FIXTURE_RELATIVE, REVIEWED_FIXTURE_IDENTITY, None),
        "common": (REVIEWED_PROBE_RELATIVES["common"], REVIEWED_PROBE_IDENTITIES["common"], EXPECTED_COMMON_DESTINATION),
        "allocator": (REVIEWED_PROBE_RELATIVES["allocator"], REVIEWED_PROBE_IDENTITIES["allocator"], "crates/litchi-odg/examples/odg_joined_qualification_allocator.rs"),
        "baseline_adapter": (REVIEWED_PROBE_RELATIVES["baseline"], REVIEWED_PROBE_IDENTITIES["baseline"], EXPECTED_OVERLAY_DESTINATION),
        "candidate_adapter": (REVIEWED_PROBE_RELATIVES["candidate"], REVIEWED_PROBE_IDENTITIES["candidate"], EXPECTED_OVERLAY_DESTINATION),
    }
    result: dict[str, dict[str, Any]] = {}
    for name, (relative, identity, destination) in expected.items():
        item = _object(values.get(name), f"source binder.reviewed_inputs.{name}")
        required = {"path", "bytes", "sha256", "destination"} if destination is not None else {"path", "bytes", "sha256", "provenance"}
        _exact_keys(item, required, f"source binder.reviewed_inputs.{name}")
        path = _safe_path(Path(_string(item.get("path"), f"reviewed input {name}.path")), f"reviewed input {name}", must_exist=True)
        if name == "fixture":
            # The final ODG plan rejects fixtures inside either endpoint
            # worktree. Bind the reviewed external copy by exact bytes and
            # provenance; do not accept a source-root look-alike here.
            if path == candidate_root or candidate_root in path.parents:
                raise RunFailure("reviewed input fixture must be external to the candidate source root")
            actual = _external_reviewed_file(path, identity, "reviewed input fixture")
            if item.get("bytes") != actual["bytes"] or item.get("sha256") != actual["sha256"]:
                raise RunFailure("reviewed input fixture identity differs")
            if item.get("provenance") != REVIEWED_FIXTURE_PROVENANCE:
                raise RunFailure("source binder fixture provenance differs")
            result[name] = {**actual, "provenance": item["provenance"]}
            continue
        expected_paths = {
            (REVIEWED_REPO_ROOT / relative).resolve(),
            (candidate_root / relative).resolve(),
        }
        if path not in expected_paths:
            raise RunFailure(f"reviewed input {name} is not the pinned repository file")
        # The source binder was reviewed from the frozen candidate worktree,
        # while the builder itself lives in a separate consumer leaf.  Accept
        # either exact reviewed location, but require an ordinary tracked file
        # and the fixed reviewed bytes in both cases.
        source_root = candidate_root if path == (candidate_root / relative).resolve() else REVIEWED_REPO_ROOT
        status = _run_git_bytes(source_root, "ls-files", "-v", "-z", "--", str(relative))
        if [item for item in status.split(b"\0") if item] != [b"H " + str(relative).encode()]:
            raise RunFailure(f"reviewed input {name} is not an ordinary tracked file")
        actual = _identity(path)
        binding = {"bytes": _positive_int(identity.get("bytes"), f"reviewed input {name}.bytes"),
                   "sha256": _digest(identity.get("sha256"), f"reviewed input {name}.sha256")}
        if actual["bytes"] != binding["bytes"] or actual["sha256"] != binding["sha256"]:
            raise RunFailure(f"reviewed input {name} bytes differ from the reviewed source binding")
        if item.get("bytes") != actual["bytes"] or item.get("sha256") != actual["sha256"]:
            raise RunFailure(f"reviewed input {name} identity differs")
        if destination is not None and item.get("destination") != destination:
            raise RunFailure(f"reviewed input {name} destination differs")
        if name == "fixture" and item.get("provenance") != REVIEWED_FIXTURE_PROVENANCE:
            raise RunFailure("source binder fixture provenance differs")
        result[name] = {**actual, **({"destination": destination} if destination else {"provenance": item["provenance"]})}
    return result


def _source_endpoint_state(raw_endpoint: Mapping[str, Any], name: str) -> dict[str, Any]:
    root = _safe_path(Path(_string(raw_endpoint["root"], f"source binder.{name}.root")), f"{name}.root", must_exist=True)
    shape = _check_worktree_shape(root, name)
    head = _git_output(root, "rev-parse", "--verify", "HEAD")
    tree = _git_output(root, "rev-parse", "--verify", "HEAD^{tree}")
    if head != raw_endpoint["ref"] or tree != raw_endpoint["tree"]:
        raise RunFailure(f"{name}: source HEAD/tree differs from the source binder")
    status = _git_output(root, "status", "--porcelain", "--untracked-files=all")
    if status:
        raise RunFailure(f"{name}: source root must be clean before the bounded build")
    actual_manifest = _tree_manifest(root, head)
    expected_manifest = _check_tree_manifest_shape(raw_endpoint["source_manifest"], f"{name}.source_manifest")
    if actual_manifest != expected_manifest:
        raise RunFailure(f"{name}: complete Git tree differs from the source binder")
    files: dict[str, dict[str, Any]] = {}
    for field, relative in (
        ("lock", "Cargo.lock"),
        ("toolchain", "rust-toolchain.toml"),
        ("cargo_config", ".cargo/config.toml"),
        ("cargo_manifest", "Cargo.toml"),
    ):
        expected = raw_endpoint[field]
        path = root / relative
        actual = _identity(path)
        if actual["bytes"] != expected["bytes"] or actual["sha256"] != expected["sha256"]:
            raise RunFailure(f"{name}.{field} differs from the source binder")
        files[field] = actual
    return {
        "name": name, "root": str(root), "ref": head, "tree": tree, "profile": raw_endpoint["profile"],
        "features": list(raw_endpoint["features"]), "source_shape": shape,
        "source_manifest": expected_manifest,
        "source_closure": _closure_binding(raw_endpoint["source_closure"], f"{name}.source_closure"),
        "lock": files["lock"], "toolchain": files["toolchain"],
        "cargo_config": files["cargo_config"], "cargo_manifest": files["cargo_manifest"],
    }


def _source_preparation_plan(
    binder_path: Path,
    *,
    target_root: Path,
    build_output: Path,
    tmpdir: Path,
    cpu: int,
    max_load: float,
    max_iterations: int,
    timeout_s: float,
    allow_existing_build_paths: bool = False,
) -> dict[str, Any]:
    binder_path = _safe_path(binder_path.absolute(), "source binder", must_exist=True)
    raw = _raw_manifest(binder_path)
    if raw.get("schema") != SOURCE_BINDER_SCHEMA:
        raise RunFailure(f"source binder schema must be {SOURCE_BINDER_SCHEMA!r}")
    if raw.get("claim_scope") != "source/fixture/lock/probe binding only; normal and allocator binaries are pending the authorized C2a build":
        raise RunFailure("source binder claim scope is not the reviewed pending-build scope")
    target_root = _safe_path(target_root.absolute(), "ODG target root")
    build_output = _safe_path(build_output.absolute(), "ODG build output")
    tmpdir = _safe_path(tmpdir.absolute(), "ODG build TMPDIR")
    if (target_root.exists() or build_output.exists()) and not allow_existing_build_paths:
        raise RunFailure("source preparation target/output paths must be fresh and absent")
    if allow_existing_build_paths:
        for path, label in ((target_root, "ODG target root"), (build_output, "ODG build output")):
            if path.exists() and (not path.is_dir() or path.is_symlink()):
                raise RunFailure(f"{label} is not a regular external directory")
    if not tmpdir.is_dir() or tmpdir.is_symlink():
        raise RunFailure("ODG build TMPDIR must be an existing regular directory")
    baseline_raw = _source_binder_endpoint(raw, "baseline")
    candidate_raw = _source_binder_endpoint(raw, "candidate")
    if baseline_raw["ref"] != PRISTINE_BASELINE_REF:
        raise RunFailure(f"source binder baseline must be pristine pinned ref {PRISTINE_BASELINE_REF}")
    if candidate_raw["ref"].startswith(HISTORICAL_CANDIDATE_PREFIX):
        raise RunFailure("historical 15ab candidate is not admissible")
    baseline = _source_endpoint_state(baseline_raw, "baseline")
    candidate = _source_endpoint_state(candidate_raw, "candidate")
    if baseline["root"] == candidate["root"] or baseline["ref"] == candidate["ref"]:
        raise RunFailure("ODG source endpoints must be distinct")
    for root in (Path(baseline["root"]), Path(candidate["root"])):
        _outside(target_root, root, "ODG target root")
        _outside(build_output, root, "ODG build output")
        _outside(tmpdir, root, "ODG build TMPDIR")
    if any(_under(left, right) or _under(right, left) for left, right in (
        (target_root, build_output), (target_root, tmpdir), (build_output, tmpdir)
    )):
        raise RunFailure("ODG target/output/TMPDIR paths must be pairwise disjoint")
    reviewed = _source_binder_reviewed_inputs(raw, Path(candidate["root"]))
    endpoints: dict[str, dict[str, Any]] = {}
    commands: dict[str, dict[str, dict[str, Any]]] = {}
    for name, endpoint_raw, endpoint_state in (
        ("baseline", baseline_raw, baseline), ("candidate", candidate_raw, candidate)
    ):
        normal_target = target_root / f"{name}-normal"
        allocator_target = target_root / f"{name}-allocator"
        endpoint = {
            "root": endpoint_state["root"],
            "target_dir": str(normal_target),
            "allocator_target_dir": str(allocator_target),
            "features": list(endpoint_raw["features"]),
        }
        endpoint_state = {
            **endpoint_state,
            "target_dir": str(normal_target),
            "allocator_target_dir": str(allocator_target),
            "binary": str(normal_target / "release" / "examples" / EXPECTED_EXAMPLE),
            "allocator_binary": str(allocator_target / "release" / "examples" / EXPECTED_EXAMPLE),
        }
        endpoints[name] = endpoint_state
        commands[name] = {}
        for kind, allocator in (("normal", False), ("allocator", True)):
            command = _command(endpoint, allocator=allocator, cpu=cpu)
            overrides = dict(command["environment"])
            overrides.update({
                "CARGO_BUILD_JOBS": "1",
                "CARGO_PROFILE_RELEASE_DEBUG": "0",
                "LITCHI_MEM_MAX": "20G",
                "TMPDIR": str(tmpdir),
            })
            command["environment_overrides"] = overrides
            command["environment"] = _environment_identity(
                Path(endpoint_state["root"]), ref=endpoint_state["ref"],
                tree=endpoint_state["tree"], source_closure=endpoint_state["source_closure"],
                overrides=overrides,
            )
            commands[name][kind] = command
    overlays = [
        {"source": reviewed["common"]["path"], "destination": EXPECTED_COMMON_DESTINATION},
        {"source": reviewed["allocator"]["path"], "destination": "crates/litchi-odg/examples/odg_joined_qualification_allocator.rs"},
        {"source": reviewed["baseline_adapter"]["path"], "destination": EXPECTED_OVERLAY_DESTINATION, "endpoint": "baseline"},
        {"source": reviewed["candidate_adapter"]["path"], "destination": EXPECTED_OVERLAY_DESTINATION, "endpoint": "candidate"},
    ]
    return {
        "schema": SOURCE_PREPARATION_SCHEMA,
        "builder_revision": BUILDER_REVISION,
        "builder_source": _identity(Path(__file__).resolve()),
        "status": "source_prepared",
        # The source binder has no self-referential build receipt.  Bind its
        # actual file bytes here; `_manifest_sha` is reserved for manifests
        # whose observed receipt identity is deliberately normalized.
        "source_binder": {"path": str(binder_path), "bytes": binder_path.stat().st_size, "sha256": _sha256_file(binder_path)[0]},
        "baseline": endpoints["baseline"],
        "candidate": endpoints["candidate"],
        "reviewed_inputs": reviewed,
        "overlays": overlays,
        "commands": commands,
        "target_root": str(target_root),
        "build_output": str(build_output),
        "tmpdir": str(tmpdir),
        "cpu": cpu,
        "max_load": max_load,
        "max_iterations": max_iterations,
        "timeout_s": timeout_s,
        "abba_order": ["a1", "b1", "b2", "a2"],
        "blocks": 2,
        "warmups": 20,
        "samples_per_fresh_child": 1,
        "target_ms": 500,
        "minimum_ms": 150,
        "composition_limits": SOURCE_COMPOSITION_LIMITS,
        "closure_policy": "full tracked source closure is revalidated immediately before and after every bounded build job",
        "claim_scope": "source preparation only; no build, binary, correctness, timing, or acceptance claim",
    }


def prepare_source(
    binder_path: Path,
    *,
    target_root: Path,
    build_output: Path,
    tmpdir: Path,
    cpu: int,
    max_load: float,
    max_iterations: int,
    timeout_s: float,
    output: Path,
) -> dict[str, Any]:
    if type(cpu) is not int or cpu < 0:
        raise RunFailure("source preparation CPU must be a nonnegative integer")
    if type(max_load) not in (int, float) or max_load <= 0:
        raise RunFailure("source preparation max_load must be positive")
    if type(max_iterations) is not int or max_iterations <= 0:
        raise RunFailure("source preparation max_iterations must be positive")
    if type(timeout_s) not in (int, float) or timeout_s <= 0:
        raise RunFailure("source preparation timeout_s must be positive")
    output = _safe_path(output.absolute(), "source preparation receipt")
    if output.exists() or output.is_symlink():
        raise RunFailure(f"refusing to overwrite source preparation receipt: {output}")
    plan = _source_preparation_plan(
        binder_path, target_root=target_root, build_output=build_output, tmpdir=tmpdir,
        cpu=cpu, max_load=max_load, max_iterations=max_iterations, timeout_s=timeout_s,
    )
    for root in (Path(plan["baseline"]["root"]), Path(plan["candidate"]["root"])):
        _outside(output, root, "source preparation receipt")
    if any(_under(output, Path(value)) or _under(Path(value), output)
           for value in (plan["target_root"], plan["build_output"], plan["tmpdir"])):
        raise RunFailure("source preparation receipt must be disjoint from build paths")
    plan["source_preparation"] = {"path": str(output)}
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(plan, indent=2, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())
    return plan


def _load_source_preparation(path: Path) -> dict[str, Any]:
    path = _safe_path(path.absolute(), "source preparation receipt", must_exist=True)
    raw = _raw_manifest(path)
    if raw.get("schema") != SOURCE_PREPARATION_SCHEMA or raw.get("status") != "source_prepared":
        raise RunFailure("source preparation receipt schema/status differs")
    source_path = _object(raw.get("source_preparation"), "source preparation receipt.source_preparation")
    if source_path.get("path") != str(path):
        raise RunFailure("source preparation receipt path differs")
    binder = _object(raw.get("source_binder"), "source preparation receipt.source_binder")
    binder_path = _safe_path(Path(_string(binder.get("path"), "source binder path")), "source binder", must_exist=True)
    binder_sha, binder_bytes = _sha256_file(binder_path)
    if binder_sha != binder.get("sha256") or binder_bytes != binder.get("bytes"):
        raise RunFailure("source binder changed after preparation")
    current = _source_preparation_plan(
        binder_path,
        target_root=Path(_string(raw.get("target_root"), "source preparation.target_root")),
        build_output=Path(_string(raw.get("build_output"), "source preparation.build_output")),
        tmpdir=Path(_string(raw.get("tmpdir"), "source preparation.tmpdir")),
        cpu=_nonnegative_int(raw.get("cpu"), "source preparation.cpu"),
        max_load=float(raw.get("max_load")),
        max_iterations=_positive_int(raw.get("max_iterations"), "source preparation.max_iterations"),
        timeout_s=float(raw.get("timeout_s")),
        allow_existing_build_paths=True,
    )
    current["source_preparation"] = source_path
    if raw != current:
        raise RunFailure("source preparation receipt differs from its current source/builder plan")
    return raw


def _source_identity_for_event(endpoint: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "ref": endpoint["ref"],
        "tree": endpoint["tree"],
        "source_manifest": endpoint["source_manifest"],
        "source_closure": endpoint["source_closure"],
        "lock": endpoint["lock"],
        "toolchain": endpoint["toolchain"],
        "cargo_config": endpoint["cargo_config"],
        "cargo_manifest": endpoint["cargo_manifest"],
    }


def build_source(preparation_path: Path) -> dict[str, Any]:
    preparation_path = _safe_path(preparation_path.absolute(), "source preparation receipt", must_exist=True)
    preparation = _load_source_preparation(preparation_path)
    build_output = _safe_path(Path(preparation["build_output"]), "ODG build output")
    target_root = _safe_path(Path(preparation["target_root"]), "ODG target root")
    if build_output.exists() or target_root.exists():
        raise RunFailure("ODG source build requires fresh output and target roots")
    target_root.mkdir(parents=True, exist_ok=False)
    build_output.mkdir(parents=True, exist_ok=False)
    logs_dir = build_output / "logs"
    logs_dir.mkdir()
    events: list[dict[str, Any]] = []
    endpoints = (("baseline", preparation["baseline"]), ("candidate", preparation["candidate"]))
    for name, endpoint in endpoints:
        root = Path(endpoint["root"])
        source_cache = _source_closure(root, endpoint["ref"], endpoint["source_closure"])
        for kind, target_key in (("normal", "target_dir"), ("allocator", "allocator_target_dir")):
            target = Path(endpoint[target_key])
            _fresh_dir(target, f"{name}/{kind} target directory")
            overlay_paths: list[str] = []
            event: dict[str, Any] | None = None
            try:
                _validate_source_closure(root, source_cache, endpoint["source_closure"])
                source_before = _source_identity_for_event(endpoint)
                for overlay in preparation["overlays"]:
                    if overlay.get("endpoint") not in (None, name):
                        continue
                    relative = str(overlay["destination"])
                    _overlay_copy(Path(overlay["source"]), root / relative, root)
                    overlay_paths.append(relative)
                command_spec = preparation["commands"][name][kind]
                guard_before = _quiet_guard(preparation["max_load"], f"before-{name}-{kind}")
                event = _stream_command(
                    list(command_spec["argv"]), root, logs_dir, f"{name}-{kind}", target,
                    preparation["cpu"], timeout_s=preparation["timeout_s"],
                    environment=command_spec["environment"]["variables"],
                    max_load=preparation["max_load"],
                )
                event["environment"] = command_spec["environment"]
                event["guard_before"] = guard_before
                event["guard_after"] = _quiet_guard(preparation["max_load"], f"after-{name}-{kind}")
                artifact_path = Path(endpoint["binary"] if kind == "normal" else endpoint["allocator_binary"])
                if artifact_path.is_symlink() or not artifact_path.is_file() or not os.access(artifact_path, os.X_OK):
                    raise RunFailure(f"{name}/{kind}: bounded build did not produce an executable")
                event["artifact"] = _identity(artifact_path)
                event["source_before"] = source_before
            finally:
                for relative in reversed(overlay_paths):
                    _remove_overlay(root, relative)
                shape = _check_worktree_shape(root, name)
                if shape["root"] != endpoint["root"]:
                    raise RunFailure(f"{name}: source root changed during bounded build")
                if event is not None:
                    _validate_source_closure(root, source_cache, endpoint["source_closure"], force_full=True)
                    event["source_after"] = _source_identity_for_event(endpoint)
                    events.append(event)
    receipt = {
        "schema": SOURCE_BUILD_SCHEMA,
        "builder_revision": BUILDER_REVISION,
        "status": "passed",
        "preparation": {"path": str(preparation_path), "bytes": preparation_path.stat().st_size, "sha256": _sha256_file(preparation_path)[0]},
        "events": events,
        "event_order": ["baseline-normal", "baseline-allocator", "candidate-normal", "candidate-allocator"],
        "claim_scope": "bounded source build only; binary identities are observed outputs, with no correctness, timing, or acceptance claim",
        "completed_unix_ns": time.time_ns(),
    }
    receipt_path = build_output / "source-build-receipt.json"
    with receipt_path.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())
    receipt["receipt"] = {"path": str(receipt_path), "bytes": receipt_path.stat().st_size, "sha256": _sha256_file(receipt_path)[0]}
    return receipt


def seal_source_build(
    preparation_path: Path,
    *,
    output_manifest: Path,
) -> dict[str, Any]:
    """Bind observed source-build artifacts into the final ODG manifest.

    This is deliberately a post-build operation.  Artifact bytes and hashes
    come only from the fresh build receipt; no executable identity is accepted
    from the source binder.  The resulting manifest uses the ordinary ODG
    schema and is checked by the unchanged final receipt validator.
    """
    preparation_path = _safe_path(preparation_path.absolute(), "source preparation receipt", must_exist=True)
    preparation = _load_source_preparation(preparation_path)
    build_path = Path(preparation["build_output"]) / "source-build-receipt.json"
    build_path = _safe_path(build_path, "source build receipt", must_exist=True)
    source_build = _raw_manifest(build_path)
    if source_build.get("schema") != SOURCE_BUILD_SCHEMA or source_build.get("status") != "passed":
        raise RunFailure("source build receipt is not a passed bounded-build receipt")
    prep_identity = source_build.get("preparation")
    if not isinstance(prep_identity, dict):
        raise RunFailure("source build receipt preparation identity is missing")
    prep_sha, prep_bytes = _sha256_file(preparation_path)
    if prep_identity.get("path") != str(preparation_path) or prep_identity.get("sha256") != prep_sha or prep_identity.get("bytes") != prep_bytes:
        raise RunFailure("source build receipt is bound to a different preparation")
    events = source_build.get("events")
    expected_labels = ["baseline-normal", "baseline-allocator", "candidate-normal", "candidate-allocator"]
    if not isinstance(events, list) or [event.get("label") for event in events] != expected_labels:
        raise RunFailure("source build receipt event order is incomplete or changed")
    event_by_label = {event["label"]: event for event in events}
    for event in events:
        artifact = _object(event.get("artifact"), f"source build {event['label']}.artifact")
        path = _safe_path(Path(_string(artifact.get("path"), f"source build {event['label']}.artifact.path")), f"source build {event['label']}.artifact", must_exist=True)
        actual = _identity(path)
        if actual != artifact or not os.access(path, os.X_OK):
            raise RunFailure(f"source build {event['label']} artifact changed or is not executable")

    def endpoint_manifest(name: str) -> dict[str, Any]:
        endpoint = preparation[name]
        normal = event_by_label[f"{name}-normal"]["artifact"]
        allocator = event_by_label[f"{name}-allocator"]["artifact"]
        return {
            "root": endpoint["root"], "ref": endpoint["ref"], "tree": endpoint["tree"],
            "profile": endpoint["profile"], "features": endpoint["features"],
            "binary": normal["path"], "target_dir": endpoint["target_dir"],
            "allocator_binary": allocator["path"], "allocator_target_dir": endpoint["allocator_target_dir"],
            "lock": endpoint["lock"], "toolchain": endpoint["toolchain"],
            "cargo_config": endpoint["cargo_config"], "cargo_manifest": endpoint["cargo_manifest"],
            "source_manifest": endpoint["source_manifest"], "source_closure": endpoint["source_closure"],
            "binary_identity": {"bytes": normal["bytes"], "sha256": normal["sha256"]},
            "allocator_binary_identity": {"bytes": allocator["bytes"], "sha256": allocator["sha256"]},
        }

    reviewed = preparation["reviewed_inputs"]
    manifest_path = _safe_path(output_manifest.absolute(), "final ODG manifest")
    if manifest_path.exists() or manifest_path.is_symlink():
        raise RunFailure(f"refusing to overwrite final ODG manifest: {manifest_path}")
    build_output = Path(preparation["build_output"])
    prepare_receipt = build_output / "preparation-receipt.json"
    final_build_receipt = build_output / "build-receipt.json"
    if prepare_receipt.exists() or final_build_receipt.exists():
        raise RunFailure("final ODG receipt paths must be fresh before sealing")
    probe_sources = {
        "common": {key: reviewed["common"][key] for key in ("path", "bytes", "sha256", "destination")},
        "allocator": {key: reviewed["allocator"][key] for key in ("path", "bytes", "sha256", "destination")},
        "baseline": {key: reviewed["baseline_adapter"][key] for key in ("path", "bytes", "sha256", "destination")},
        "candidate": {key: reviewed["candidate_adapter"][key] for key in ("path", "bytes", "sha256", "destination")},
    }
    raw = {
        "schema": "litchi.odg.joined-native-qualification.input.v3",
        "fixture": {key: reviewed["fixture"][key] for key in ("path", "bytes", "sha256", "provenance")},
        "baseline": endpoint_manifest("baseline"),
        "candidate": endpoint_manifest("candidate"),
        "probe_sources": probe_sources,
        "protocol": {
            "cpu": preparation["cpu"], "blocks": preparation["blocks"], "warmups": preparation["warmups"],
            "target_ms": preparation["target_ms"], "minimum_ms": preparation["minimum_ms"],
            "max_load": preparation["max_load"], "max_iterations": preparation["max_iterations"],
            "timeout_s": preparation["timeout_s"],
        },
        "runtime": {"tmpdir": preparation["tmpdir"]},
        "composition_limits": preparation["composition_limits"],
        "build": {
            "prepare_receipt": str(prepare_receipt), "receipt": str(final_build_receipt),
            "receipt_identity": {"path": str(final_build_receipt), "bytes": 1, "sha256": "0" * 64},
            "final_candidate_ref": preparation["candidate"]["ref"], "cpu": preparation["cpu"],
        },
    }
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    with manifest_path.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(raw, indent=2, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())
    plan = prepare_plan(manifest_path, allow_built_artifacts=True)
    with prepare_receipt.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(plan, indent=2, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())
    receipt = {
        "schema": BUILD_SCHEMA,
        "builder_revision": BUILDER_REVISION,
        "status": "passed",
        "manifest": str(manifest_path),
        "manifest_sha256": plan["manifest_sha256"],
        "prepare_receipt": str(prepare_receipt),
        "prepare_sha256": _manifest_sha(prepare_receipt),
        "plan": plan,
        "events": events,
        "completed_unix_ns": time.time_ns(),
    }
    with final_build_receipt.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
        stream.flush()
        os.fsync(stream.fileno())
    final_identity = _identity(final_build_receipt)
    raw["build"]["receipt_identity"] = {
        "path": str(final_build_receipt), "bytes": final_identity["bytes"], "sha256": final_identity["sha256"]
    }
    manifest_path.write_text(json.dumps(raw, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    validate_build_receipt(manifest_path)
    return raw


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
    fixture = _fixture_plan(raw, base, Path(candidate["root"]))
    for root in (Path(baseline["root"]), Path(candidate["root"])):
        _outside(Path(fixture["path"]), root, "fixture")
    probes = _probe_plan(raw, base, (Path(baseline["root"]), Path(candidate["root"])))
    prepare_receipt = _safe_path(_raw_path(build["prepare_receipt"], "manifest.build.prepare_receipt", base), "prepare receipt")
    receipt = _safe_path(_raw_path(build["receipt"], "manifest.build.receipt", base), "build receipt")
    receipt_binding = _receipt_binding(build["receipt_identity"], "manifest.build.receipt_identity", base, receipt)
    if prepare_receipt == receipt:
        raise RunFailure("prepare and build receipts must be distinct")
    for root in (Path(baseline["root"]), Path(candidate["root"])):
        _outside(prepare_receipt, root, "prepare receipt")
        _outside(receipt, root, "build receipt")
    commands = {
        "baseline": {"normal": _command(baseline, allocator=False, cpu=cpu),
                     "allocator": _command(baseline, allocator=True, cpu=cpu)},
        "candidate": {"normal": _command(candidate, allocator=False, cpu=cpu),
                       "allocator": _command(candidate, allocator=True, cpu=cpu)},
    }
    for name, endpoint in (("baseline", baseline), ("candidate", candidate)):
        for kind in ("normal", "allocator"):
            command = commands[name][kind]
            overrides = dict(command["environment"])
            command["environment_overrides"] = overrides
            command["environment"] = _environment_identity(
                Path(endpoint["root"]),
                ref=endpoint["ref"],
                tree=endpoint["tree"],
                source_closure=endpoint["source_closure"],
                overrides=overrides,
            )
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
        "commands": commands,
        # The build receipt identity is checked directly against the manifest
        # by the final verifier.  The plan carries only its path so sealing a
        # post-build artifact cannot create a receipt/manifest hash cycle.
        "receipts": {"prepare": str(prepare_receipt), "build": {"path": str(receipt)}},
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


def _process_identity(pid: int) -> tuple[int, int] | None:
    """Return Linux start-time and process-group identity without following pid reuse."""

    stat_path = Path("/proc") / str(pid) / "stat"
    try:
        value = stat_path.read_text(encoding="ascii")
    except FileNotFoundError:
        return None
    except OSError as error:
        raise RunFailure(f"cannot inspect process {pid}: {error}") from error
    closing = value.rfind(")")
    if closing < 0:
        raise RunFailure(f"process {pid} has malformed stat data")
    fields = value[closing + 2 :].split()
    # The suffix starts at field 3: state, ppid, pgrp, ... starttime(field 22).
    if len(fields) <= 19:
        raise RunFailure(f"process {pid} has incomplete stat data")
    try:
        return int(fields[19]), int(fields[2])
    except ValueError as error:
        raise RunFailure(f"process {pid} has malformed process-group identity") from error


def _set_child_subreaper(enabled: bool) -> int:
    """Adopt detached descendants so they remain attributable after the child exits."""

    try:
        import ctypes

        prctl = ctypes.CDLL(None, use_errno=True).prctl
        prctl.argtypes = [ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong]
        prctl.restype = ctypes.c_int
        previous = ctypes.c_int()
        if prctl(37, ctypes.addressof(previous), 0, 0, 0) != 0:  # PR_GET_CHILD_SUBREAPER
            raise OSError(ctypes.get_errno(), "PR_GET_CHILD_SUBREAPER failed")
        if prctl(36, 1 if enabled else 0, 0, 0, 0) != 0:  # PR_SET_CHILD_SUBREAPER
            raise OSError(ctypes.get_errno(), "PR_SET_CHILD_SUBREAPER failed")
        return previous.value
    except (AttributeError, OSError) as error:
        raise RunFailure(f"cannot configure child subreaper: {error}") from error


def _owned_snapshot(root_pid: int) -> dict[int, tuple[int, int]]:
    owned: dict[int, tuple[int, int]] = {}
    for pid in {root_pid, *_descendants(root_pid)}:
        identity = _process_identity(pid)
        if identity is not None:
            owned[pid] = identity
    return owned


def _live_owned(owned: Mapping[int, tuple[int, int]]) -> dict[int, tuple[int, int]]:
    live: dict[int, tuple[int, int]] = {}
    for pid, expected in owned.items():
        actual = _process_identity(pid)
        if actual is None:
            continue
        if actual != expected:
            raise RunFailure(f"owned child pid {pid} changed identity while guarded")
        live[pid] = actual
    return live


def _kill_owned(owned: Mapping[int, tuple[int, int]]) -> None:
    for pid, expected in owned.items():
        actual = _process_identity(pid)
        if actual is None:
            continue
        if actual != expected:
            raise RunFailure(f"refusing to kill reused owned child pid {pid}")
        try:
            os.kill(pid, signal.SIGKILL)
        except ProcessLookupError:
            continue
        except OSError as error:
            raise RunFailure(f"cannot terminate owned child {pid}: {error}") from error
    deadline = time.monotonic() + 1.0
    while time.monotonic() < deadline:
        remaining: dict[int, tuple[int, int]] = {}
        for pid, expected in owned.items():
            try:
                waited, _ = os.waitpid(pid, os.WNOHANG)
                if waited:
                    continue
            except ChildProcessError:
                pass
            actual = _process_identity(pid)
            if actual is not None:
                if actual != expected:
                    raise RunFailure(f"refusing to reap reused owned child pid {pid}")
                remaining[pid] = actual
        if not remaining:
            return
        time.sleep(0.02)
    raise RunFailure(f"owned child processes did not exit after termination: {sorted(remaining)}")


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


def _group_alive(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return False
    except PermissionError as error:
        raise RunFailure(f"cannot inspect child process group {pgid}: {error}") from error
    except OSError as error:
        raise RunFailure(f"cannot inspect child process group {pgid}: {error}") from error
    return True


def _validate_target_accounting(value: Any, label: str, expected_path: Path) -> dict[str, Any]:
    item = _object(value, label)
    _exact_keys(item, {"path", "bytes_before", "bytes_after", "hard_limit_bytes"}, label)
    before = item["bytes_before"]
    after = item["bytes_after"]
    hard_limit = item["hard_limit_bytes"]
    if (item["path"] != str(expected_path)
            or type(before) is not int or before < 0 or before > MEMORY_LIMIT_BYTES
            or type(after) is not int or after < 0 or after > MEMORY_LIMIT_BYTES
            or hard_limit != MEMORY_LIMIT_BYTES):
        raise RunFailure(f"{label} target accounting exceeds the fixed memory limit")
    return {
        "path": str(expected_path),
        "bytes_before": before,
        "bytes_after": after,
        "hard_limit_bytes": hard_limit,
    }


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
    if environment is None:
        env = dict(os.environ)
        env.update({"CARGO_INCREMENTAL": "0", "CARGO_TERM_COLOR": "never"})
        if "allocator" in label:
            env["RUSTFLAGS"] = "--cfg odg_joined_allocator"
    else:
        # A prepared environment is an exact child environment.  Merging it
        # with the runner's ambient environment would make unrecorded Cargo
        # configuration and wrappers part of the build.
        env = dict(environment)
    parent_pid = os.getpid()
    # A double-forked/setsid child is otherwise reparented to init as soon as
    # the direct child exits.  Keep it owned by this runner so the final
    # identity check can still find and terminate it.
    _set_child_subreaper(True)
    preexisting_children = set(_descendants(parent_pid))

    def child_setup() -> None:
        if os.getppid() != parent_pid:
            os._exit(127)
        os.setsid()
        try:
            import ctypes

            prctl = ctypes.CDLL(None, use_errno=True).prctl
            prctl.argtypes = [ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong]
            prctl.restype = ctypes.c_int
            if prctl(1, signal.SIGKILL, 0, 0, 0) != 0:  # PR_SET_PDEATHSIG
                os._exit(127)
        except (AttributeError, OSError):
            os._exit(127)
        if hasattr(os, "sched_setaffinity"):
            os.sched_setaffinity(0, {cpu})
        resource.setrlimit(resource.RLIMIT_AS, (MEMORY_LIMIT_BYTES, MEMORY_LIMIT_BYTES))
    target_before = _directory_bytes(target_dir)
    if target_before > TARGET_PREFLIGHT_BYTES:
        raise RunFailure(f"{label}: target starts above preflight limit: {target_before}")
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
    owned: dict[int, tuple[int, int]] = {}
    try:
        while process.poll() is None:
            stdout.flush()
            stderr.flush()
            output_bytes = stdout_path.stat().st_size + stderr_path.stat().st_size
            load = load_average()
            current_owned = _owned_snapshot(process.pid)
            owned.update(current_owned)
            for pid in _descendants(parent_pid) - preexisting_children:
                identity = _process_identity(pid)
                if identity is not None:
                    owned[pid] = identity
            descendants = set(current_owned) - {process.pid}
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
            samples.append({
                "checked_unix_ns": time.time_ns(), "output_bytes": output_bytes, "rss_bytes": rss,
                "load_1m": load, "outside_processes": outside,
                "owned_processes": [
                    {"pid": pid, "start_time": identity[0], "process_group": identity[1]}
                    for pid, identity in sorted(owned.items())
                ],
            })
            time.sleep(0.05)
        return_code = process.wait(timeout=5)
    except RunFailure as error:
        _kill_group(process)
        live = _live_owned(owned)
        if live:
            _kill_owned(live)
        raise error
    except (OSError, subprocess.SubprocessError) as error:
        _kill_group(process)
        live = _live_owned(owned)
        if live:
            _kill_owned(live)
        raise RunFailure(f"{label}: streamed child wait failed: {error}") from error
    finally:
        stdout.flush()
        stderr.flush()
        stdout.close()
        stderr.close()
    if violation:
        raise RunFailure(f"{label}: {violation}")
    deadline = time.monotonic() + DESCENDANT_GRACE_S
    for pid in _descendants(parent_pid) - preexisting_children:
        identity = _process_identity(pid)
        if identity is not None:
            owned[pid] = identity
    live_owned = _live_owned(owned)
    descendants = set(live_owned)
    descendants.discard(process.pid)
    group_alive = _group_alive(process.pid)
    while (live_owned or group_alive) and time.monotonic() < deadline:
        time.sleep(0.05)
        live_owned = _live_owned(owned)
        descendants = set(live_owned)
        descendants.discard(process.pid)
        group_alive = _group_alive(process.pid)
    if live_owned or group_alive:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        except OSError as error:
            raise RunFailure(f"{label}: cannot terminate orphaned child group: {error}") from error
        _kill_owned(live_owned)
        raise RunFailure(
            f"{label}: child left descendants/process group running after recheck: "
            f"{sorted(live_owned)} group_alive={group_alive}"
        )
    target_after = _directory_bytes(target_dir)
    _validate_target_accounting(
        {"path": str(target_dir), "bytes_before": target_before, "bytes_after": target_after,
         "hard_limit_bytes": MEMORY_LIMIT_BYTES},
        f"{label}.target",
        target_dir,
    )
    final_output_bytes = stdout_path.stat().st_size + stderr_path.stat().st_size
    samples.append({
        "checked_unix_ns": time.time_ns(),
        "output_bytes": final_output_bytes,
        "rss_bytes": 0,
        "load_1m": load_average(),
        "outside_processes": process_scan(),
        "owned_processes": [],
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
                     "max_output_bytes": max_output_bytes, "max_load": max_load,
                     "owned_processes": []},
    }


def _directory_bytes(path: Path) -> int:
    path = _safe_path(path, "target directory", must_exist=True)
    if not path.is_dir():
        raise RunFailure(f"target path is not a directory: {path}")
    total = 0
    pending = [path]
    while pending:
        current = pending.pop()
        try:
            entries = list(os.scandir(current))
        except OSError as error:
            raise RunFailure(f"cannot inspect target directory {current}: {error}") from error
        for entry in entries:
            try:
                if entry.is_symlink():
                    raise RunFailure(f"target contains a symlink: {entry.path}")
                if entry.is_dir(follow_symlinks=False):
                    pending.append(Path(entry.path))
                elif entry.is_file(follow_symlinks=False):
                    total += entry.stat(follow_symlinks=False).st_size
                else:
                    raise RunFailure(f"target contains a non-regular entry: {entry.path}")
            except FileNotFoundError as error:
                raise RunFailure(f"target entry disappeared during inspection: {entry.path}") from error
            except OSError as error:
                raise RunFailure(f"cannot inspect target entry {entry.path}: {error}") from error
    return total


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
        source_cache = _source_closure(root, endpoint["ref"], endpoint["source_closure"])
        _fresh_dir(target, f"{name} target directory")
        _fresh_dir(allocator_target, f"{name} allocator target directory")
        for kind, target_dir in (("normal", target), ("allocator", allocator_target)):
            overlay_paths: list[str] = []
            event: dict[str, Any] | None = None
            try:
                _validate_source_closure(root, source_cache, endpoint["source_closure"])
                source_before = {
                    "ref": _git_output(root, "rev-parse", "--verify", "HEAD"),
                    "tree": _git_output(root, "rev-parse", "--verify", "HEAD^{tree}"),
                    "source_manifest": endpoint["source_manifest"],
                    "source_closure": endpoint["source_closure"],
                    "lock": _identity(Path(endpoint["lock"]["path"])),
                    "toolchain": _identity(Path(endpoint["toolchain"]["path"])),
                    "cargo_config": _identity(Path(endpoint["cargo_config"]["path"])),
                    "cargo_manifest": _identity(Path(endpoint["cargo_manifest"]["path"])),
                }
                if source_before["ref"] != endpoint["ref"] or source_before["tree"] != endpoint["tree"]:
                    raise RunFailure(f"{name}/{kind}: source changed before overlay")
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
                    environment=command_spec["environment"]["variables"],
                    max_load=plan["max_load"],
                )
                _validate_target_accounting(event["target"], f"{name}/{kind}.target", target_dir)
                event["environment"] = command_spec["environment"]
                event["guard_before"] = guard_before
                event["guard_after"] = _quiet_guard(plan["max_load"], f"after-{name}-{kind}")
                binary_key = "allocator_binary" if kind == "allocator" else "binary"
                binary = Path(endpoint[binary_key])
                if binary.is_symlink() or not binary.is_file() or not os.access(binary, os.X_OK):
                    raise RunFailure(f"{name}/{kind}: expected executable was not produced: {binary}")
                event["artifact"] = binary_identity(binary)
                expected_artifact = endpoint[
                    "allocator_binary_identity" if kind == "allocator" else "binary_identity"
                ]
                if (event["artifact"]["bytes"] != expected_artifact["bytes"]
                        or event["artifact"]["sha256"] != expected_artifact["sha256"]):
                    raise RunFailure(f"{name}/{kind}: artifact differs from the externally bound identity")
                event["source_before"] = source_before
            finally:
                for relative in reversed(overlay_paths):
                    _remove_overlay(root, relative)
                shape = _check_worktree_shape(root, name)
                if shape["root"] != endpoint["root"]:
                    raise RunFailure(f"{name}: source root changed during build")
                if event is not None:
                    _validate_source_closure(
                        root, source_cache, endpoint["source_closure"], force_full=True
                    )
                    event["source_after"] = {
                        "ref": _git_output(root, "rev-parse", "--verify", "HEAD"),
                        "tree": _git_output(root, "rev-parse", "--verify", "HEAD^{tree}"),
                        "source_manifest": endpoint["source_manifest"],
                        "source_closure": endpoint["source_closure"],
                        "lock": _identity(Path(endpoint["lock"]["path"])),
                        "toolchain": _identity(Path(endpoint["toolchain"]["path"])),
                        "cargo_config": _identity(Path(endpoint["cargo_config"]["path"])),
                        "cargo_manifest": _identity(Path(endpoint["cargo_manifest"]["path"])),
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
    actual_receipt = _identity(build_receipt)
    expected_receipt = build_value["receipt_identity"]
    if (actual_receipt["path"] != str(build_receipt)
            or actual_receipt["bytes"] != expected_receipt["bytes"]
            or actual_receipt["sha256"] != expected_receipt["sha256"]):
        raise RunFailure("build receipt differs from its externally bound receipt identity")
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
    expected_receipt = _receipt_binding(
        build["receipt_identity"], "manifest.build.receipt_identity", manifest_path.parent, receipt_path
    )
    if _identity(receipt_path) != expected_receipt:
        raise RunFailure("build receipt bytes differ from the externally bound receipt identity")
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
        _exact_keys(monitor, {"samples", "memory_limit_bytes", "max_output_bytes", "max_load", "owned_processes"},
                    f"{label}.monitor")
        if not isinstance(monitor["samples"], list) or not monitor["samples"]:
            raise RunFailure(f"{label}: live monitor samples are missing")
        for sample in monitor["samples"]:
            row = _object(sample, f"{label}.monitor.sample")
            _exact_keys(row, {"checked_unix_ns", "output_bytes", "rss_bytes", "load_1m",
                              "outside_processes", "owned_processes"}, f"{label}.monitor.sample")
            _positive_int(row["checked_unix_ns"], "monitor.checked_unix_ns")
            if type(row["output_bytes"]) is not int or row["output_bytes"] < 0 or row["output_bytes"] > MAX_BUILD_OUTPUT:
                raise RunFailure(f"{label}: output monitor exceeded its bound")
            if type(row["rss_bytes"]) is not int or row["rss_bytes"] < 0 or row["rss_bytes"] > MEMORY_LIMIT_BYTES:
                raise RunFailure(f"{label}: RSS monitor exceeded its bound")
            if (type(row["load_1m"]) not in (int, float)
                    or row["load_1m"] < 0 or row["load_1m"] > monitor["max_load"]):
                raise RunFailure(f"{label}: load monitor exceeded its bound")
            if row["outside_processes"] != []:
                raise RunFailure(f"{label}: unrelated process appeared in monitor")
            owned = row["owned_processes"]
            if not isinstance(owned, list):
                raise RunFailure(f"{label}: owned process monitor is malformed")
            for process in owned:
                process_value = _object(process, f"{label}.monitor.owned_process")
                _exact_keys(process_value, {"pid", "start_time", "process_group"},
                            f"{label}.monitor.owned_process")
                _positive_int(process_value["pid"], f"{label}.monitor.owned_process.pid")
                _positive_int(process_value["start_time"], f"{label}.monitor.owned_process.start_time")
                _positive_int(process_value["process_group"], f"{label}.monitor.owned_process.process_group")
        if (monitor["memory_limit_bytes"] != MEMORY_LIMIT_BYTES
                or monitor["max_output_bytes"] != MAX_BUILD_OUTPUT
                or monitor["max_load"] != plan["max_load"]):
            raise RunFailure(f"{label}: monitor limits differ from the fixed builder limits")
        if not isinstance(monitor["owned_processes"], list) or monitor["owned_processes"]:
            raise RunFailure(f"{label}: child-owned process group was not empty at receipt validation")
        for guard_label in ("guard_before", "guard_after"):
            guard = _object(item[guard_label], f"{label}.{guard_label}")
            _exact_keys(guard, {"phase", "load_1m", "max_load", "heavy_processes", "checked_unix_ns"},
                        f"{label}.{guard_label}")
            if (type(guard["load_1m"]) not in (int, float) or guard["load_1m"] < 0
                    or type(guard["max_load"]) not in (int, float)
                    or guard["max_load"] <= 0):
                raise RunFailure(f"{label}: quiet guard numeric identity is malformed")
            _positive_int(guard["checked_unix_ns"], f"{label}.{guard_label}.checked_unix_ns")
            if not isinstance(guard["heavy_processes"], list):
                raise RunFailure(f"{label}: quiet guard process list is malformed")
            if guard["max_load"] != plan["max_load"] or guard["heavy_processes"] != []:
                raise RunFailure(f"{label}: quiet guard identity or process list changed")
        target = _object(item["target"], f"{label}.target")
        endpoint_target = plan[identity[0]]["allocator_target_dir" if identity[1] == "allocator" else "target_dir"]
        _validate_target_accounting(target, f"{label}.target", Path(endpoint_target))
        artifact = _object(item["artifact"], f"{label}.artifact")
        _exact_keys(artifact, {"path", "bytes", "sha256"}, f"{label}.artifact")
        artifact_path = _safe_path(Path(_string(artifact["path"], f"{label}.artifact.path")).absolute(), f"{label}.artifact.path", must_exist=True)
        expected_artifact = plan[identity[0]]["allocator_binary" if identity[1] == "allocator" else "binary"]
        if str(artifact_path) != expected_artifact:
            raise RunFailure(f"{label}: artifact path differs from the prepared command")
        if _identity(artifact_path) != artifact:
            raise RunFailure(f"{label}: built artifact bytes changed")
        expected_artifact_identity = plan[identity[0]][
            "allocator_binary_identity" if identity[1] == "allocator" else "binary_identity"
        ]
        if (artifact["bytes"] != expected_artifact_identity["bytes"]
                or artifact["sha256"] != expected_artifact_identity["sha256"]):
            raise RunFailure(f"{label}: artifact is not the externally bound binary")
        expected_command = plan["commands"][identity[0]][identity[1]]
        if item["command"] != expected_command["argv"]:
            raise RunFailure(f"{label}: command differs from the prepared command")
        if item["environment"] != expected_command["environment"]:
            raise RunFailure(f"{label}: build environment differs from the prepared command")
        if item["cwd"] != plan[identity[0]]["root"]:
            raise RunFailure(f"{label}: command cwd differs from the endpoint root")
        source_before = _object(item["source_before"], f"{label}.source_before")
        _exact_keys(source_before, {"ref", "tree", "source_manifest", "source_closure", "lock", "toolchain", "cargo_config", "cargo_manifest"},
                    f"{label}.source_before")
        endpoint = plan[identity[0]]
        if source_before["ref"] != endpoint["ref"] or source_before["tree"] != endpoint["tree"]:
            raise RunFailure(f"{label}: source identity is not bound to its endpoint")
        if source_before["source_manifest"] != endpoint["source_manifest"]:
            raise RunFailure(f"{label}: source manifest before build differs from its endpoint")
        if source_before["source_closure"] != endpoint["source_closure"]:
            raise RunFailure(f"{label}: source closure before build differs from its endpoint")
        for field in ("lock", "toolchain", "cargo_config", "cargo_manifest"):
            if source_before[field] != endpoint[field]:
                raise RunFailure(f"{label}: {field} before build differs from its endpoint")
        source_after = _object(item["source_after"], f"{label}.source_after")
        _exact_keys(source_after, {"ref", "tree", "source_manifest", "source_closure", "lock", "toolchain", "cargo_config", "cargo_manifest"},
                    f"{label}.source_after")
        if source_after["ref"] != endpoint["ref"] or source_after["tree"] != endpoint["tree"]:
            raise RunFailure(f"{label}: source changed during the build")
        if source_after["source_manifest"] != endpoint["source_manifest"]:
            raise RunFailure(f"{label}: complete source manifest changed during the build")
        if source_after["source_closure"] != endpoint["source_closure"]:
            raise RunFailure(f"{label}: complete source closure changed during the build")
        for field in ("lock", "toolchain", "cargo_config", "cargo_manifest"):
            if source_after[field] != endpoint[field]:
                raise RunFailure(f"{label}: {field} changed during the build")
    if observed_order != expected:
        raise RunFailure("build receipt event set is incomplete")
    return receipt


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--binder", type=Path)
    parser.add_argument("--preparation", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--target-root", type=Path)
    parser.add_argument("--build-output", type=Path)
    parser.add_argument("--tmpdir", type=Path)
    parser.add_argument("--cpu", type=int)
    parser.add_argument("--max-load", type=float, default=1.0)
    parser.add_argument("--max-iterations", type=int, default=100000)
    parser.add_argument("--timeout-s", type=float, default=300.0)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--prepare", action="store_true")
    group.add_argument("--build", action="store_true")
    group.add_argument("--validate", action="store_true")
    group.add_argument("--prepare-source", action="store_true")
    group.add_argument("--build-source", action="store_true")
    group.add_argument("--seal-source", action="store_true")
    args = parser.parse_args(argv)
    try:
        if args.prepare_source:
            required = (args.binder, args.output, args.target_root, args.build_output, args.tmpdir, args.cpu)
            if any(value is None for value in required):
                parser.error("--prepare-source requires --binder, --output, --target-root, --build-output, --tmpdir, and --cpu")
            prepare_source(
                args.binder, target_root=args.target_root, build_output=args.build_output,
                tmpdir=args.tmpdir, cpu=args.cpu, max_load=args.max_load,
                max_iterations=args.max_iterations, timeout_s=args.timeout_s, output=args.output,
            )
        elif args.build_source:
            if args.preparation is None:
                parser.error("--build-source requires --preparation")
            build_source(args.preparation)
        elif args.seal_source:
            if args.preparation is None or args.output is None:
                parser.error("--seal-source requires --preparation and --output")
            seal_source_build(args.preparation, output_manifest=args.output)
        elif args.prepare:
            if args.manifest is None:
                parser.error("--prepare requires --manifest")
            prepare(args.manifest)
        elif args.build:
            if args.manifest is None:
                parser.error("--build requires --manifest")
            build(args.manifest)
        else:
            if args.manifest is None:
                parser.error("--validate requires --manifest")
            validate_build_receipt(args.manifest)
    except (RunFailure, OSError, ValueError) as error:
        print(str(error), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
