#!/usr/bin/env python3
"""Prepare the version-two DrawingML source pair.

Version one prepared a historical worktree by copying the current
``tools/perf-baseline/src/lib.rs`` and reversing a growing list of consumer
spellings.  Version two has no source rewrite.  It materializes the one
reviewed historical library directly from a fixed Git ref/blob and copies only
the shared probe and route inputs from the candidate.

The final candidate source is deliberately not bound in this change.  Until
the coordinator records its exact commit, manifests, locks, and vendor-tree
identity, preparation fails before it mutates either endpoint.  This keeps a
partial or user-supplied hash from becoming qualification evidence.

This module is source preparation only.  It does not invoke Cargo, compile a
probe, run a native process, or measure a workload.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any, Mapping


SCHEMA = "litchi.drawingml.raw-preparation.v2"

# The historical library is a trusted object from the final reviewed source
# capture.  These values are protocol constants, not CLI inputs.  The Git
# object id is SHA-1 as reported by ``git ls-tree``; the SHA-256 and byte count
# bind the materialized payload independently of Git's object database.
HISTORICAL_SOURCE_REF = "472b6705475b073e7fe08f6e11ee17fb1e29f838"
HISTORICAL_SOURCE_PATH = "tools/perf-baseline/src/lib.rs"
HISTORICAL_SOURCE_BLOB_SHA1 = "fbc55c5dcdb7cf86e5377e4f8012954a7fb4e572"
HISTORICAL_SOURCE_SHA256 = "280d61ae1f714af78be6cecf21a4432a9ea2a27478e89c237b5d358bc02db02a"
HISTORICAL_SOURCE_BYTES = 3_088_216

# These are the reviewed historical parents used by the five raw routes, the
# theme-family route, and the four true owner transitions.  A baseline hash
# supplied by a manifest is accepted only if it is one of these source fences.
APPROVED_BASELINE_REVISIONS = frozenset(
    {
        "2f7db7ee7fa01893e9099217fc4243687d196f63",
        "996651b6b19717a0192dfd4144bccd367fb2ea44",
        "11728db894181ffbcdbcea3ae0008d51fdd2d8af",
        "945eb1d6b9ef6c56fa26c82e12758d38c1b9015d",
    }
)

# All three role locks are separate identities.  The root lock is ignored by
# the repository and is therefore checked as an external materialized input;
# the two tool locks are tracked files.  Candidate values are known from the
# reviewed font bridge, but the candidate source commit and manifests remain
# intentionally unbound below.
ROLE_LOCKS: dict[str, dict[str, dict[str, Any]]] = {
    "workspace": {
        "path": "Cargo.lock",
        "baseline": {
            "sha256": "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1",
            "bytes": 139_606,
        },
        "candidate": {
            "sha256": "efff87af6c339a44e267987c5c79ba3fb1910cef734f0a9f212f08f865dac811",
            "bytes": 139_189,
        },
    },
    "perf": {
        "path": "tools/perf-baseline/Cargo.lock",
        "baseline": {
            "sha256": "eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c",
            "bytes": 57_227,
        },
        "candidate": {
            "sha256": "de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea",
            "bytes": 56_941,
        },
    },
    "native": {
        "path": "tools/native-resave/Cargo.lock",
        "baseline": {
            "sha256": "51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378",
            "bytes": 56_975,
        },
        "candidate": {
            "sha256": "989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae",
            "bytes": 56_689,
        },
    },
}
ROLE_LOCK_PATHS = tuple(spec["path"] for spec in ROLE_LOCKS.values())

# These are checked separately from the lock files because sparse checkouts
# can retain a commit/tree identity while omitting the materialized files.
CANDIDATE_MANIFEST_PATHS = (
    "Cargo.toml",
    "tools/perf-baseline/Cargo.toml",
    "tools/native-resave/Cargo.toml",
)
# Cargo's source graph is not determined by the manifests alone.  Bind the
# selected toolchain and runner-owned Cargo configuration in the same receipt
# so a compile consumer cannot silently switch compiler/configuration inputs.
TOOLCHAIN_PATHS = (
    "rust-toolchain.toml",
    ".cargo/config.toml",
)
VENDOR_ROOTS = (
    "vendor/allsorts",
    "vendor/flate2",
    "vendor/zlib-rs",
)

# Shared probes/routes are copied byte-for-byte from the candidate.  The
# historical library is the only source path that comes from the trusted Git
# object.  Lock files are role-specific and are never copied across roles.
PROBE_COPY_PATHS = (
    "tools/perf-baseline/Cargo.toml",
    HISTORICAL_SOURCE_PATH,
    "tools/perf-baseline/src/drawingml_raw_probe.rs",
    "tools/perf-baseline/src/bin/drawingml_raw_native.rs",
    "tools/perf-baseline/src/bin/drawingml_raw_allocator.rs",
    "tools/perf-baseline/src/bin/support/drawingml_counting_allocator.rs",
    "tools/native_qualification/drawingml_raw_validator.py",
    "tools/native_qualification/fixtures/drawingml_raw_fixtures.json",
)
ALLOWED_BASELINE_DIRTY_PATHS = frozenset(PROBE_COPY_PATHS)

# This is deliberately None until the final candidate source capture is
# frozen.  It is not populated from command-line values or an external JSON
# file.  The coordinator must commit the exact revision, manifest hashes, and
# vendor-tree digest here before any v2 preparation can execute.
CANDIDATE_BINDING: dict[str, Any] | None = None

# The dependency policy is explicit even while source binding is pending.  It
# describes the three root registry -> path transitions; the role lock hashes
# above bind the concrete lock bytes once the candidate is frozen.
DEPENDENCY_PATH_TRANSITIONS = (
    {
        "package": "allsorts",
        "version": "0.17.0",
        "baseline_source": "registry",
        "candidate_source": "path",
        "candidate_path": "vendor/allsorts",
    },
    {
        "package": "flate2",
        "version": "1.1.10",
        "baseline_source": "registry",
        "candidate_source": "path",
        "candidate_path": "vendor/flate2",
    },
    {
        "package": "zlib-rs",
        "version": "0.6.8",
        "baseline_source": "registry",
        "candidate_source": "path",
        "candidate_path": "vendor/zlib-rs",
    },
)


class PreparationError(ValueError):
    """A source, policy, or materialization invariant failed."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise PreparationError(message)


def _canonical(value: Any) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n"
    ).encode("utf-8")


def _sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _git_blob_sha1(value: bytes) -> str:
    return hashlib.sha1(f"blob {len(value)}\0".encode("ascii") + value).hexdigest()


def _reject_symlink_components(path: Path, label: str) -> None:
    raw = Path(path)
    cursor = Path(raw.anchor) if raw.is_absolute() else Path()
    pieces = raw.parts[1:] if raw.is_absolute() else raw.parts
    for piece in pieces:
        cursor /= piece
        _expect(not cursor.is_symlink(), f"{label} path contains a symlink: {cursor}")


def _git(root: Path, *arguments: str, binary: bool = False) -> bytes | str:
    try:
        result = subprocess.run(
            ["git", "-C", str(root), *arguments],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        detail = (
            error.stderr.decode("utf-8", errors="replace").strip()
            if isinstance(error, subprocess.CalledProcessError)
            else str(error)
        )
        raise PreparationError(f"git {' '.join(arguments)} failed: {detail}") from error
    if binary:
        return bytes(result.stdout)
    return result.stdout.decode("utf-8")


def _root(value: Path) -> Path:
    root = value.resolve(strict=False)
    _expect(root.is_dir(), f"source root is not a directory: {root}")
    top = Path(str(_git(root, "rev-parse", "--show-toplevel").strip())).resolve(strict=False)
    _expect(top == root, f"source root is not a worktree top-level: {root}")
    return root


def _head(root: Path) -> str:
    value = str(_git(root, "rev-parse", "HEAD")).strip()
    _expect(len(value) == 40 and value == value.lower(), f"invalid Git HEAD in {root}")
    return value


def _tree(root: Path) -> str:
    value = str(_git(root, "rev-parse", "HEAD^{tree}")).strip()
    _expect(len(value) == 40 and value == value.lower(), f"invalid Git tree in {root}")
    return value


def _status_paths(root: Path) -> tuple[str, ...]:
    raw = str(_git(root, "status", "--porcelain=v1", "--untracked-files=all"))
    paths: list[str] = []
    for line in raw.splitlines():
        _expect(len(line) >= 4, f"malformed Git status line: {line!r}")
        value = line[3:]
        _expect(" -> " not in value, f"renames are not allowed: {line!r}")
        paths.append(value)
    return tuple(sorted(paths))


def _file_identity(root: Path, relative: str) -> dict[str, Any]:
    path = root / relative
    _reject_symlink_components(path, f"{relative} identity")
    _expect(path.is_file(), f"required file is missing: {path}")
    payload = path.read_bytes()
    return {
        "path": relative,
        "bytes": len(payload),
        "sha256": _sha256_bytes(payload),
    }


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _assert_disjoint_paths(paths: Mapping[str, Path]) -> None:
    resolved = {label: path.resolve(strict=False) for label, path in paths.items()}
    items = list(resolved.items())
    for index, (left_label, left) in enumerate(items):
        for right_label, right in items[index + 1 :]:
            _expect(
                not (left == right or _under(left, right) or _under(right, left)),
                f"paths must be disjoint: {left_label}={left}, {right_label}={right}",
            )


def _historical_library_identity(root: Path) -> dict[str, Any]:
    """Require the materialized historical file to match the fixed blob."""

    identity = _file_identity(root, HISTORICAL_SOURCE_PATH)
    payload = (root / HISTORICAL_SOURCE_PATH).read_bytes()
    _expect(
        identity["sha256"] == HISTORICAL_SOURCE_SHA256,
        "materialized historical library SHA-256 differs",
    )
    _expect(
        identity["bytes"] == HISTORICAL_SOURCE_BYTES,
        "materialized historical library byte count differs",
    )
    _expect(
        _git_blob_sha1(payload) == HISTORICAL_SOURCE_BLOB_SHA1,
        "materialized historical library Git blob differs",
    )
    return identity


def _lock_identities(root: Path, role: str) -> dict[str, dict[str, Any]]:
    _expect(role in {"baseline", "candidate"}, f"unknown lock role: {role}")
    result: dict[str, dict[str, Any]] = {}
    for spec in ROLE_LOCKS.values():
        relative = str(spec["path"])
        actual = _file_identity(root, relative)
        expected = dict(spec[role])
        _expect(actual["sha256"] == expected["sha256"], f"{role} lock hash differs: {relative}")
        _expect(actual["bytes"] == expected["bytes"], f"{role} lock byte count differs: {relative}")
        result[relative] = actual
    return result


def _parse_tree_entries(root: Path, revision: str = "HEAD") -> tuple[dict[str, str], ...]:
    raw = bytes(_git(root, "ls-tree", "-r", "--full-tree", "-z", revision, binary=True))
    entries: list[dict[str, str]] = []
    for record in raw.split(b"\0"):
        if not record:
            continue
        try:
            header, path_bytes = record.split(b"\t", 1)
            mode, kind, object_id = header.decode("ascii").split(" ", 2)
            path = path_bytes.decode("utf-8")
        except (ValueError, UnicodeDecodeError) as error:
            raise PreparationError("Git tree contains an unreadable path") from error
        _expect(kind == "blob", f"Git tree contains non-blob entry: {path}")
        entries.append({"mode": mode, "object": object_id, "path": path})
    _expect(bool(entries), f"Git tree is empty: {revision}")
    return tuple(sorted(entries, key=lambda item: item["path"]))


def _vendor_tree_identity(root: Path) -> dict[str, Any]:
    """Bind every tracked vendor blob and fail on sparse materialization."""

    entries = [
        entry
        for entry in _parse_tree_entries(root)
        if entry["path"] == "vendor" or entry["path"].startswith("vendor/")
    ]
    _expect(entries, "candidate has no tracked vendor tree")
    allowed = tuple(f"{name}/" for name in VENDOR_ROOTS)
    unexpected = [
        entry["path"]
        for entry in entries
        if not any(entry["path"].startswith(prefix) for prefix in allowed)
    ]
    _expect(not unexpected, f"candidate has an unreviewed vendor tree: {unexpected[:3]}")
    tracked_paths = {entry["path"] for entry in entries}
    # Git status intentionally omits ignored files. Walk the materialized
    # vendor roots as well, so a sparse checkout or ignored overlay cannot
    # hide an input from the bound tree digest.
    for vendor_root in VENDOR_ROOTS:
        directory = root / vendor_root
        _reject_symlink_components(directory, f"vendor root {vendor_root}")
        _expect(directory.is_dir() and not directory.is_symlink(), f"vendor root is missing: {directory}")
        for current, directories, files in os.walk(directory, topdown=True, followlinks=False):
            current_path = Path(current)
            for name in directories:
                child = current_path / name
                _expect(not child.is_symlink(), f"vendor tree contains a symlink: {child}")
            for name in files:
                child = current_path / name
                _expect(not child.is_symlink(), f"vendor tree contains a symlink: {child}")
                relative = child.relative_to(root).as_posix()
                _expect(
                    relative in tracked_paths,
                    f"vendor tree contains an untracked or ignored overlay: {relative}",
                )
    observed: list[dict[str, Any]] = []
    root_counts = {name: 0 for name in VENDOR_ROOTS}
    root_bytes = {name: 0 for name in VENDOR_ROOTS}
    for entry in entries:
        relative = entry["path"]
        if relative == "vendor":
            continue
        matched = next((name for name in VENDOR_ROOTS if relative.startswith(f"{name}/")), None)
        _expect(matched is not None, f"vendor path is outside reviewed roots: {relative}")
        _expect(entry["mode"] not in {"120000", "160000"}, f"vendor path is not a regular blob: {relative}")
        identity = _file_identity(root, relative)
        payload = (root / relative).read_bytes()
        _expect(
            _git_blob_sha1(payload) == entry["object"],
            f"tracked vendor blob differs on disk: {relative}",
        )
        item = {
            "path": relative,
            "mode": entry["mode"],
            "git_blob_sha1": entry["object"],
            "bytes": identity["bytes"],
            "sha256": identity["sha256"],
        }
        observed.append(item)
        root_counts[matched] += 1
        root_bytes[matched] += identity["bytes"]
    for name in VENDOR_ROOTS:
        _expect(root_counts[name] > 0, f"tracked vendor tree is missing: {name}")
    return {
        "roots": {
            name: {"files": root_counts[name], "bytes": root_bytes[name]}
            for name in VENDOR_ROOTS
        },
        "files": len(observed),
        "bytes": sum(item["bytes"] for item in observed),
        "sha256": _sha256_bytes(_canonical(observed)),
        "entries": observed,
    }


def _trusted_historical_payload(trusted_root: Path) -> tuple[bytes, dict[str, Any]]:
    trusted_root = _root(trusted_root)
    actual_commit = str(
        _git(trusted_root, "rev-parse", "--verify", f"{HISTORICAL_SOURCE_REF}^{{commit}}")
    ).strip()
    _expect(actual_commit == HISTORICAL_SOURCE_REF, "trusted historical ref is not the fixed reviewed commit")
    tree = _parse_tree_entries(trusted_root, HISTORICAL_SOURCE_REF)
    matched = [entry for entry in tree if entry["path"] == HISTORICAL_SOURCE_PATH]
    _expect(len(matched) == 1, "trusted historical library path is not unique")
    _expect(matched[0]["object"] == HISTORICAL_SOURCE_BLOB_SHA1, "trusted historical blob changed")
    payload = bytes(_git(trusted_root, "show", f"{HISTORICAL_SOURCE_REF}:{HISTORICAL_SOURCE_PATH}", binary=True))
    _expect(len(payload) == HISTORICAL_SOURCE_BYTES, "trusted historical library byte count changed")
    _expect(_sha256_bytes(payload) == HISTORICAL_SOURCE_SHA256, "trusted historical library hash changed")
    _expect(_git_blob_sha1(payload) == HISTORICAL_SOURCE_BLOB_SHA1, "trusted historical Git blob changed")
    return payload, {
        "ref": HISTORICAL_SOURCE_REF,
        "path": HISTORICAL_SOURCE_PATH,
        "blob_sha1": HISTORICAL_SOURCE_BLOB_SHA1,
        "sha256": HISTORICAL_SOURCE_SHA256,
        "bytes": HISTORICAL_SOURCE_BYTES,
    }


def materialize_historical_library(
    trusted_root: Path,
    baseline_root: Path,
    *,
    payload: bytes | None = None,
    source: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Write the fixed historical blob and return its source-bound receipt."""

    baseline_root = _root(baseline_root)
    if payload is None:
        payload, trusted_source = _trusted_historical_payload(trusted_root)
    else:
        trusted_source = dict(source or {})
        _expect(
            trusted_source
            == {
                "ref": HISTORICAL_SOURCE_REF,
                "path": HISTORICAL_SOURCE_PATH,
                "blob_sha1": HISTORICAL_SOURCE_BLOB_SHA1,
                "sha256": HISTORICAL_SOURCE_SHA256,
                "bytes": HISTORICAL_SOURCE_BYTES,
            },
            "preflight historical source identity differs",
        )
        _expect(len(payload) == HISTORICAL_SOURCE_BYTES, "preflight historical library byte count differs")
        _expect(_sha256_bytes(payload) == HISTORICAL_SOURCE_SHA256, "preflight historical library SHA-256 differs")
        _expect(_git_blob_sha1(payload) == HISTORICAL_SOURCE_BLOB_SHA1, "preflight historical library Git blob differs")
    target = baseline_root / HISTORICAL_SOURCE_PATH
    _reject_symlink_components(target, "historical library destination")
    _expect(not target.is_symlink(), f"historical library destination is a symlink: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(payload)
    observed = _historical_library_identity(baseline_root)
    return {
        "trusted": trusted_source,
        "destination": observed,
    }


def _require_candidate_binding() -> dict[str, Any]:
    if CANDIDATE_BINDING is None:
        raise PreparationError(
            "DrawingML v2 candidate policy is incomplete: final source commit, "
            "manifests, and vendor-tree identity are not frozen"
        )
    binding = CANDIDATE_BINDING
    _expect(binding.get("status") == "bound", "DrawingML v2 candidate policy is not bound")
    revision = binding.get("revision")
    _expect(isinstance(revision, str) and len(revision) == 40, "bound candidate revision is missing")
    manifests = binding.get("manifests")
    _expect(isinstance(manifests, dict) and set(manifests) == set(CANDIDATE_MANIFEST_PATHS), "bound candidate manifests are incomplete")
    vendor = binding.get("vendor_tree")
    _expect(isinstance(vendor, dict) and vendor.get("sha256"), "bound candidate vendor tree is missing")
    return binding


def _candidate_identity(root: Path, binding: Mapping[str, Any]) -> dict[str, Any]:
    root = _root(root)
    expected_revision = str(binding["revision"])
    actual_revision = _head(root)
    _expect(actual_revision == expected_revision, "candidate HEAD differs from the bound source policy")
    _expect(not _status_paths(root), "candidate source must be clean")
    manifests = {path: _file_identity(root, path) for path in CANDIDATE_MANIFEST_PATHS}
    expected_manifests = binding["manifests"]
    for path, expected in expected_manifests.items():
        _expect(manifests[path] == {"path": path, "bytes": expected["bytes"], "sha256": expected["sha256"]}, f"candidate manifest differs: {path}")
    locks = _lock_identities(root, "candidate")
    toolchain = {path: _file_identity(root, path) for path in TOOLCHAIN_PATHS}
    probe_identities = {
        relative: _file_identity(root, relative) for relative in PROBE_COPY_PATHS
    }
    vendor = _vendor_tree_identity(root)
    _expect(vendor["sha256"] == binding["vendor_tree"]["sha256"], "candidate vendor tree differs from the bound policy")
    _expect(vendor["files"] == binding["vendor_tree"]["files"], "candidate vendor file count differs")
    _expect(vendor["bytes"] == binding["vendor_tree"]["bytes"], "candidate vendor byte count differs")
    return {
        "root": str(root),
        "head": actual_revision,
        "tree": _tree(root),
        "status_paths": [],
        "manifests": manifests,
        "toolchain": toolchain,
        "locks": locks,
        "vendor_tree": {key: value for key, value in vendor.items() if key != "entries"},
        "probe_paths": list(PROBE_COPY_PATHS),
        # Keep the eight copied inputs independently bound.  The path list is
        # retained for compatibility with the source policy, while this map
        # prevents a receipt from proving only that names existed.
        "probe_identities": probe_identities,
    }


def _baseline_identity(
    root: Path,
    baseline_revision: str,
    *,
    require_historical_library: bool = True,
) -> dict[str, Any]:
    root = _root(root)
    _expect(baseline_revision in APPROVED_BASELINE_REVISIONS, "baseline revision is not an approved DrawingML source fence")
    _expect(_head(root) == baseline_revision, "baseline HEAD differs from the selected source fence")
    status = _status_paths(root)
    _expect(set(status).issubset(ALLOWED_BASELINE_DIRTY_PATHS), f"baseline changed outside v2 adapter paths: {status!r}")
    locks = _lock_identities(root, "baseline")
    toolchain = {path: _file_identity(root, path) for path in TOOLCHAIN_PATHS}
    if require_historical_library:
        historical_library = _historical_library_identity(root)
    else:
        historical_path = root / HISTORICAL_SOURCE_PATH
        historical_library = (
            _file_identity(root, HISTORICAL_SOURCE_PATH)
            if historical_path.is_file() and not historical_path.is_symlink()
            else {"path": HISTORICAL_SOURCE_PATH, "missing": True}
        )
    return {
        "root": str(root),
        "head": baseline_revision,
        "tree": _tree(root),
        "status_paths": list(status),
        "locks": locks,
        "toolchain": toolchain,
        "historical_library": historical_library,
        "probe_paths": {path: _file_identity(root, path) for path in PROBE_COPY_PATHS},
    }


def _copy_probe_paths(candidate_root: Path, baseline_root: Path) -> None:
    for relative in PROBE_COPY_PATHS:
        source = candidate_root / relative
        target = baseline_root / relative
        _reject_symlink_components(source, f"candidate probe {relative}")
        _reject_symlink_components(target, f"baseline probe {relative}")
        _expect(source.is_file() and not source.is_symlink(), f"candidate probe is missing: {source}")
        _expect(not target.is_symlink(), f"baseline probe destination is a symlink: {target}")
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)


def _receipt_policy(binding: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "status": "bound",
        "revision": binding["revision"],
        "manifests": binding["manifests"],
        "vendor_tree": binding["vendor_tree"],
    }


def prepare(
    baseline_root: Path,
    candidate_root: Path,
    baseline_revision: str,
    candidate_revision: str,
    trusted_root: Path,
    receipt_path: Path,
) -> None:
    """Materialize a v2 source pair after the checked-in candidate policy binds."""

    binding = _require_candidate_binding()
    _expect(candidate_revision == binding["revision"], "candidate revision is not the bound policy revision")
    _expect(baseline_revision in APPROVED_BASELINE_REVISIONS, "baseline revision is not approved")
    candidate_root = _root(candidate_root)
    baseline_root = _root(baseline_root)
    receipt_path = receipt_path.resolve(strict=False)
    _assert_disjoint_paths(
        {
            "candidate source": candidate_root,
            "baseline source": baseline_root,
            "preparation receipt": receipt_path,
        }
    )
    _candidate_identity(candidate_root, binding)
    _baseline_identity(
        baseline_root,
        baseline_revision,
        require_historical_library=False,
    )
    # Verify the trusted object before copying even one candidate route file.
    # A missing ref/blob must leave the baseline worktree untouched.
    historical_payload, historical_source = _trusted_historical_payload(trusted_root)
    _copy_probe_paths(candidate_root, baseline_root)
    historical = materialize_historical_library(
        trusted_root,
        baseline_root,
        payload=historical_payload,
        source=historical_source,
    )
    baseline = _baseline_identity(baseline_root, baseline_revision)
    _expect(
        set(baseline["status_paths"]).issubset(ALLOWED_BASELINE_DIRTY_PATHS),
        "baseline contains an unapproved post-materialization change",
    )
    candidate = _candidate_identity(candidate_root, binding)
    _reject_symlink_components(receipt_path, "preparation receipt")
    _expect(not receipt_path.exists() and not receipt_path.is_symlink(), f"refusing to overwrite receipt: {receipt_path}")
    receipt_path.parent.mkdir(parents=True, exist_ok=True)
    receipt = {
        "schema": SCHEMA,
        "protocol_version": 2,
        "candidate_policy": _receipt_policy(binding),
        "baseline_revision": baseline_revision,
        "candidate_revision": candidate_revision,
        "historical_source": historical,
        "dependency_path_transitions": list(DEPENDENCY_PATH_TRANSITIONS),
        "role_locks": ROLE_LOCKS,
        "candidate_manifests": list(CANDIDATE_MANIFEST_PATHS),
        "toolchain_paths": list(TOOLCHAIN_PATHS),
        "vendor_roots": list(VENDOR_ROOTS),
        "probe_copy_paths": list(PROBE_COPY_PATHS),
        "baseline": baseline,
        "candidate": candidate,
    }
    with receipt_path.open("x", encoding="utf-8") as stream:
        json.dump(receipt, stream, indent=2, sort_keys=True)
        stream.write("\n")


def validate_preparation_receipt(
    receipt_path: Path,
    *,
    baseline_root: Path,
    candidate_root: Path,
    baseline_revision: str,
    candidate_revision: str,
    trusted_root: Path,
) -> dict[str, Any]:
    """Validate a v2 receipt against the still-bound source pair."""

    binding = _require_candidate_binding()
    receipt_path = receipt_path.resolve(strict=False)
    baseline_root = _root(baseline_root)
    candidate_root = _root(candidate_root)
    _assert_disjoint_paths(
        {
            "candidate source": candidate_root,
            "baseline source": baseline_root,
            "preparation receipt": receipt_path,
        }
    )
    _reject_symlink_components(receipt_path, "preparation receipt")
    _expect(receipt_path.is_file() and not receipt_path.is_symlink(), f"preparation receipt is missing: {receipt_path}")
    try:
        receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read preparation receipt: {error}") from error
    _expect(isinstance(receipt, dict), "preparation receipt root is not an object")
    _expect(receipt.get("schema") == SCHEMA, "preparation receipt schema differs")
    _expect(receipt.get("protocol_version") == 2, "preparation receipt protocol differs")
    _expect(receipt.get("candidate_policy") == _receipt_policy(binding), "candidate policy receipt differs")
    _expect(receipt.get("baseline_revision") == baseline_revision, "baseline revision receipt differs")
    _expect(receipt.get("candidate_revision") == candidate_revision, "candidate revision receipt differs")
    _, trusted_source = _trusted_historical_payload(trusted_root)
    expected_trusted = {
        "ref": HISTORICAL_SOURCE_REF,
        "path": HISTORICAL_SOURCE_PATH,
        "blob_sha1": HISTORICAL_SOURCE_BLOB_SHA1,
        "sha256": HISTORICAL_SOURCE_SHA256,
        "bytes": HISTORICAL_SOURCE_BYTES,
    }
    _expect(
        receipt.get("historical_source", {}).get("trusted") == expected_trusted
        and trusted_source == expected_trusted,
        "historical trusted source receipt differs",
    )
    destination = receipt.get("historical_source", {}).get("destination")
    _expect(
        isinstance(destination, dict)
        and destination.get("sha256") == HISTORICAL_SOURCE_SHA256
        and destination.get("bytes") == HISTORICAL_SOURCE_BYTES,
        "historical destination receipt differs",
    )
    _expect(receipt.get("role_locks") == ROLE_LOCKS, "role lock policy receipt differs")
    _expect(receipt.get("candidate_manifests") == list(CANDIDATE_MANIFEST_PATHS), "candidate manifest policy receipt differs")
    _expect(receipt.get("toolchain_paths") == list(TOOLCHAIN_PATHS), "toolchain path policy receipt differs")
    _expect(receipt.get("vendor_roots") == list(VENDOR_ROOTS), "vendor root policy receipt differs")
    _expect(receipt.get("probe_copy_paths") == list(PROBE_COPY_PATHS), "probe path policy receipt differs")
    current_candidate = _candidate_identity(candidate_root, binding)
    current_baseline = _baseline_identity(baseline_root, baseline_revision)
    _expect(receipt.get("candidate") == current_candidate, "candidate source identity changed")
    _expect(receipt.get("baseline") == current_baseline, "baseline source identity changed")
    return receipt


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-root", type=Path, required=True)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--baseline-revision", required=True)
    parser.add_argument("--candidate-revision", required=True)
    parser.add_argument("--trusted-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    arguments = parser.parse_args(argv)
    try:
        prepare(
            arguments.baseline_root,
            arguments.candidate_root,
            arguments.baseline_revision,
            arguments.candidate_revision,
            arguments.trusted_root,
            arguments.output,
        )
    except (OSError, PreparationError) as error:
        parser.error(str(error))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
