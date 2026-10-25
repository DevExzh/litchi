#!/usr/bin/env python3
"""Prepare the version-two signed-native qualification source pair.

This module is deliberately a source-only boundary.  It records the inputs a
future signed-native run must use and verifies those inputs before a consumer
or a native process may be started.  The final candidate source was not
frozen when this policy was recorded, so the checked-in binding is ``None``.
Consequently :func:`prepare` and :func:`execute` currently fail closed before
creating an output directory, receipt, build, or measurement.  Once the
coordinator supplies the immutable binding, they delegate to the tracked
consumer in :mod:`native_qualification.signed_native_v2_consumer`.

The old 0362ul24 receipts are historical v1 evidence.  They are never
relabelled as v2, and this module does not accept a v1 receipt as a v2 source
or qualification result.

No Cargo, compiler, native probe, rustdoc, or timing command is invoked by the
pending source-only path.  All digests used for a bound candidate are
recomputed from the selected source tree; a caller cannot supply a digest or
commit on the command line to make an unbound policy executable.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tarfile
import tomllib
from typing import Any, Mapping


HERE = Path(__file__).resolve().parent
REPOSITORY = HERE.parents[1]
POLICY_PATH = HERE / "manifests" / "signed-native-v2.json"
FIXTURE_MANIFEST = REPOSITORY / "docs/unified-ops/results/0362ul24/fixtures.json"

SCHEMA = "litchi.signed-native.policy.v2"
PREPARATION_SCHEMA = "litchi.signed-native.preparation.v2"
POLICY_VERSION = 2
ADAPTER_POLICY = {
    "schema": "litchi.signed-native.adapter.v2",
    "source": "consumer_only_external_project",
    "legacy_feature": "legacy_api",
    "api_branches": [
        "legacy_commit_fields",
        "current_commit_accessors",
        "legacy_xlsb_option",
        "current_xlsb_resolve",
    ],
    "route_count": 16,
}
EXECUTION_POLICY = {
    "capabilities_schema": "litchi.signed-native.capabilities.v2",
    "execution_schema": "litchi.signed-native.execution.v2",
    "child_schema": "litchi.signed-native.child.v2",
    "probe_schema": "litchi.signed-native.probe.v2",
    "batch_proof_schema": "litchi.signed-native.batch-proof.v2",
    "route_evidence_required": True,
    "api_branches": ["current_api", "legacy_api"],
    "build_receipt_required": True,
    "child_timeout_seconds": 300.0,
    "child_output_limit_bytes": 4 * 1024 * 1024,
    "target_cap_bytes": 20_000_000_000,
    "tmp_cap_bytes": 20_000_000_000,
    "signed_batch": {
        "selection": "immediate_refusal",
        "iterator_reads": 0,
        "publication_attempts": 0,
    },
}
SIGNED_BATCH_POLICY = {
    "status": "owner_confirmed",
    "selection": "immediate_refusal",
}

FULL_SHA = re.compile(r"^[0-9a-f]{40}$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")

# The current signature policy test matrix.  Keep this set/order in lockstep
# with crates/litchi-ops/src/tests/signature.rs and the historical route
# validator.  In particular, no route is silently dropped because its result
# is correctness-only.
NATIVE_FORMATS = (
    "DOCX",
    "PPTX",
    "XLSX",
    "XLSB",
    "DOC",
    "PPT",
    "XLS",
)
ROUTES = (
    "signed_leaf/noop_commit",
    "signed_facade/noop_commit",
    "signed_leaf/noop_write",
    "signed_facade/noop_write",
    "signed_leaf/noop_replay",
    "signed_facade/noop_replay",
    "signed_leaf/refusal",
    "signed_facade/refusal",
    "signed_leaf/unsign_write",
    "signed_facade/unsign_write",
    "signed_leaf/unsigned_negative",
    "signed_facade/unsigned_negative",
    "signed_leaf/source_freshness",
    "signed_facade/source_freshness",
    "signed_leaf/budget",
    "signed_facade/budget",
)
CORRECTNESS_ONLY = frozenset(
    {
        "signed_leaf/refusal",
        "signed_facade/refusal",
        "signed_leaf/unsigned_negative",
        "signed_facade/unsigned_negative",
        "signed_leaf/source_freshness",
        "signed_facade/source_freshness",
        "signed_leaf/budget",
        "signed_facade/budget",
    }
)
TIMED_ROUTES = tuple(route for route in ROUTES if route not in CORRECTNESS_ONLY)
BASELINE_NOT_APPLICABLE = frozenset(
    {
        "signed_leaf/unsigned_negative",
        "signed_facade/unsigned_negative",
        "signed_leaf/source_freshness",
        "signed_facade/source_freshness",
        "signed_leaf/budget",
        "signed_facade/budget",
    }
)

# These are fixed source identities from the old preparation record.  They
# describe the historical input boundary only; no old receipt is upgraded.
HISTORICAL_SOURCE = {
    "record": "0362ul24",
    "schema": "litchi_signed_native_qualification.v1",
    "baseline_commit": "f2bbbc8d07ab5da26c180f1ff74813d39f632d54",
    "old_candidate_commit": "472b6705475b073e7fe08f6e11ee17fb1e29f838",
    "receipts_are_historical": True,
    "receipts_accepted_by_v2": False,
}

# The three role locks are known inputs from the current source preparation.
# The root lock is intentionally listed even though the sparse candidate leaf
# does not materialise it yet; its absence is a source-freeze failure.
ROLE_LOCKS: dict[str, dict[str, Any]] = {
    "workspace": {
        "path": "Cargo.lock",
        "baseline": {
            "sha256": "1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c",
            "bytes": 139441,
        },
        "candidate": {
            # The workspace lock is the exact C-780 source lock.  The source
            # commit/archive itself remains pending until the coordinator
            # supplies the immutable post-134a production capture.
            "sha256": "02fda2543226f28c2ded9bbf7e18618f8b0aab1c4d02c7b2c4fdd73511e4ff7a",
            "bytes": 139207,
        },
    },
    "perf": {
        "path": "tools/perf-baseline/Cargo.lock",
        "baseline": {
            "sha256": "eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c",
            "bytes": 57227,
        },
        "candidate": {
            "sha256": "de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea",
            "bytes": 56941,
        },
    },
    "native": {
        "path": "tools/native-resave/Cargo.lock",
        "baseline": {
            "sha256": "51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378",
            "bytes": 56975,
        },
        "candidate": {
            "sha256": "989da70f10deab4460a77e4afa3a1b840b4697bbea2c07e134e22d9b4c5084ae",
            "bytes": 56689,
        },
    },
}

SOURCE_MANIFEST_PATHS = (
    "Cargo.toml",
    "tools/perf-baseline/Cargo.toml",
    "tools/native-resave/Cargo.toml",
)
VENDOR_ROOTS = (
    "vendor/allsorts",
    "vendor/flate2",
    "vendor/zlib-rs",
)
TOOLCHAIN_PATHS = ("rust-toolchain.toml", ".cargo/config.toml")
TOOLCHAIN_IDENTITIES = {
    "rust-toolchain.toml": {
        "path": "rust-toolchain.toml",
        "bytes": 86,
        "sha256": "e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310",
    },
    ".cargo/config.toml": {
        "path": ".cargo/config.toml",
        "bytes": 470,
        "sha256": "50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d",
    },
}

# These are source/fixture inputs, not result receipts.  Their fixed hashes
# prevent a future run from silently changing the route or fixture contract.
PROBE_INPUTS = {
    "probe": {
        "path": "docs/unified-ops/results/0362ul24/signed_native_probe.rs",
        "bytes": 46851,
        "sha256": "cabc4fed22648df915a667f7be661adcf2124eacb7fc5efb6f784300c1d36e2f",
    },
    "qualification_driver": {
        "path": "docs/unified-ops/results/0362ul24/signed_native_qualification.py",
        "bytes": 24171,
        "sha256": "3ec79eb432c1af1f8d88699916e8607beb62d63c98d0ce3d1360e4950ad013d5",
    },
    "validator": {
        "path": "docs/unified-ops/results/0362ul24/validate_signed_native.py",
        "bytes": 91077,
        "sha256": "f5a64930577c98a29cf2da7e288de799d3d13a93521e539c1f0aa5111ba71fa3",
    },
    "memory_collector": {
        "path": "docs/unified-ops/results/0362ul24/signed_native_memory.py",
        "bytes": 16981,
        "sha256": "f1b3f432ae04ea57d4a7797d645d7fa2a8e13d170635a1681815f33adc6d157e",
    },
    "final_collector": {
        "path": "docs/unified-ops/results/0362ul24/signed_native_finalize.py",
        "bytes": 5330,
        "sha256": "febc82bbd4082d9f03839a68bb9e7731c99a9692abfca65c2aa8d0a8faf57823",
    },
}
FIXTURE_INPUT = {
    "path": "docs/unified-ops/results/0362ul24/fixtures.json",
    "bytes": 4988,
    "sha256": "729d326b9aab089a9b360852b7a29abe0ca7337ca4d93c0a21398ccc8f823af7",
}

# This is deliberately ``None``.  It may only be replaced by a later source
# freeze that records all fields in the checked-in manifest and updates this
# code constant in the same review.  There is no CLI or JSON override.
CANDIDATE_BINDING: dict[str, Any] | None = None


class PolicyError(ValueError):
    """The policy or a source identity violates the qualification boundary."""


class IncompletePolicyError(PolicyError):
    """The source-freeze fields are intentionally not complete yet."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise PolicyError(message)


def _canonical(value: Any) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        + "\n"
    ).encode("utf-8")


def _sha256_bytes(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def _git_blob_sha1(payload: bytes) -> str:
    return hashlib.sha1(f"blob {len(payload)}\0".encode("ascii") + payload).hexdigest()


def _reject_symlink_components(path: Path, label: str) -> None:
    lexical = path.absolute()
    cursor = Path(lexical.anchor)
    for component in lexical.parts[1:]:
        cursor /= component
        _expect(not cursor.is_symlink(), f"{label} path contains a symlink: {cursor}")


def _safe_relative(value: Any, label: str) -> str:
    _expect(isinstance(value, str) and value, f"{label} is missing")
    path = Path(value)
    _expect(not path.is_absolute() and "\\" not in value, f"{label} must be relative")
    _expect(
        all(part not in {"", ".", ".."} for part in path.parts),
        f"{label} contains traversal or an unnormalised component",
    )
    _expect("3rdparty" not in path.parts, f"{label} may not use an upstream checkout")
    return value


def _git(root: Path, *arguments: str, binary: bool = False) -> str | bytes:
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
        raise PolicyError(f"git {' '.join(arguments)} failed: {detail}") from error
    return bytes(result.stdout) if binary else result.stdout.decode("utf-8")


def _root(value: Path) -> Path:
    _reject_symlink_components(value, "source root")
    root = value.resolve(strict=False)
    _expect(root.is_dir(), f"source root is not a directory: {root}")
    top = Path(str(_git(root, "rev-parse", "--show-toplevel")).strip()).resolve(strict=False)
    _expect(top == root, f"source root is not a worktree top-level: {root}")
    return root


def _head(root: Path, revision: str = "HEAD") -> str:
    value = str(_git(root, "rev-parse", "--verify", f"{revision}^{{commit}}")).strip()
    _expect(FULL_SHA.fullmatch(value) is not None, f"Git revision is not a full commit: {revision}")
    return value


def _tree_id(root: Path, revision: str) -> str:
    value = str(_git(root, "rev-parse", "--verify", f"{revision}^{{tree}}")).strip()
    _expect(FULL_SHA.fullmatch(value) is not None, f"Git tree is not a full id: {revision}")
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


def _tree_entries(root: Path, revision: str = "HEAD") -> tuple[dict[str, str], ...]:
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
            raise PolicyError("Git tree contains an unreadable path") from error
        _expect(kind == "blob", f"Git tree contains a non-blob entry: {path}")
        entries.append({"mode": mode, "object": object_id, "path": path})
    _expect(bool(entries), f"Git tree is empty: {revision}")
    return tuple(sorted(entries, key=lambda item: item["path"]))


def _file_identity(root: Path, relative: str) -> dict[str, Any]:
    relative = _safe_relative(relative, "tracked path")
    path = root / relative
    _reject_symlink_components(path, f"{relative} identity")
    _expect(path.is_file() and not path.is_symlink(), f"required regular file is missing: {path}")
    payload = path.read_bytes()
    return {"path": relative, "bytes": len(payload), "sha256": _sha256_bytes(payload)}


def _artifact_identity(path: Path, label: str) -> dict[str, Any]:
    _reject_symlink_components(path, label)
    _expect(path.is_file() and not path.is_symlink(), f"{label} is not a regular file: {path}")
    absolute = path.absolute()
    _expect("3rdparty" not in absolute.parts, f"{label} may not use an upstream checkout")
    forbidden = (
        REPOSITORY / "docs" / "performance" / "results",
        REPOSITORY / "docs" / "report",
        REPOSITORY / "binaries",
    )
    resolved = absolute.resolve(strict=False)
    for root in forbidden:
        try:
            resolved.relative_to(root.resolve())
        except ValueError:
            continue
        raise PolicyError(f"{label} must remain outside {root}")
    payload = resolved.read_bytes()
    return {"path": str(resolved), "bytes": len(payload), "sha256": _sha256_bytes(payload)}


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _assert_disjoint_paths(paths: Mapping[str, Path]) -> None:
    """Reject equal and ancestor/descendant source, archive, and output paths."""

    resolved: dict[str, Path] = {}
    for label, path in paths.items():
        _reject_symlink_components(path, label)
        resolved[label] = path.resolve(strict=False)
    items = list(resolved.items())
    for index, (left_label, left) in enumerate(items):
        for right_label, right in items[index + 1 :]:
            _expect(
                not (left == right or _under(left, right) or _under(right, left)),
                f"paths must be disjoint: {left_label}={left}, {right_label}={right}",
            )


def _checked_artifact(path: Path, expected: Mapping[str, Any], label: str) -> dict[str, Any]:
    _expect(isinstance(expected.get("bytes"), int), f"{label} byte count is missing")
    expected_sha = expected.get("sha256")
    _expect(isinstance(expected_sha, str) and SHA256.fullmatch(expected_sha) is not None, f"{label} hash is malformed")
    actual = _artifact_identity(path, label)
    _expect(actual["bytes"] == expected["bytes"], f"{label} byte count changed")
    _expect(actual["sha256"] == expected_sha, f"{label} digest changed")
    return actual


def _vendor_tree_identity(root: Path, revision: str = "HEAD") -> dict[str, Any]:
    """Hash every tracked vendor blob and reject sparse or modified files."""

    entries = [
        entry
        for entry in _tree_entries(root, revision)
        if entry["path"].startswith("vendor/")
    ]
    _expect(entries, "candidate has no tracked vendor files")
    prefixes = tuple(f"{name}/" for name in VENDOR_ROOTS)
    unexpected = [
        entry["path"]
        for entry in entries
        if not any(entry["path"].startswith(prefix) for prefix in prefixes)
    ]
    _expect(not unexpected, f"candidate has an unreviewed vendor path: {unexpected[:3]}")
    tracked_paths = {entry["path"] for entry in entries}
    # Git status omits ignored files. Walk each materialised root as well so
    # a sparse checkout or ignored/untracked overlay cannot hide a source
    # input from the bound tree digest.
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
                _expect(child.is_file(), f"vendor tree contains a non-regular file: {child}")
                relative = child.relative_to(root).as_posix()
                _expect(
                    relative in tracked_paths,
                    f"vendor tree contains an untracked or ignored overlay: {relative}",
                )
    observed: list[dict[str, Any]] = []
    counts = {name: 0 for name in VENDOR_ROOTS}
    sizes = {name: 0 for name in VENDOR_ROOTS}
    for entry in entries:
        relative = entry["path"]
        _expect(entry["mode"] not in {"120000", "160000"}, f"vendor path is not a regular blob: {relative}")
        root_name = next(name for name in VENDOR_ROOTS if relative.startswith(f"{name}/"))
        identity = _file_identity(root, relative)
        payload = (root / relative).read_bytes()
        _expect(_git_blob_sha1(payload) == entry["object"], f"tracked vendor blob differs on disk: {relative}")
        observed.append(
            {
                "path": relative,
                "mode": entry["mode"],
                "git_blob_sha1": entry["object"],
                "bytes": identity["bytes"],
                "sha256": identity["sha256"],
            }
        )
        counts[root_name] += 1
        sizes[root_name] += identity["bytes"]
    for name in VENDOR_ROOTS:
        _expect(counts[name] > 0, f"tracked vendor tree is missing: {name}")
    return {
        "roots": {name: {"files": counts[name], "bytes": sizes[name]} for name in VENDOR_ROOTS},
        "files": len(observed),
        "bytes": sum(item["bytes"] for item in observed),
        "sha256": _sha256_bytes(_canonical(observed)),
        "entries": observed,
    }


def _source_tree_digest(entries: tuple[dict[str, str], ...]) -> str:
    return _sha256_bytes(_canonical(list(entries)))


def _archive_tree_identity(
    root: Path,
    revision: str,
    archive: Mapping[str, Any],
) -> dict[str, Any]:
    """Prove that a bound tar archive contains exactly the selected Git tree.

    A digest of an archive only proves the bytes of that archive.  This
    second check proves the archive's regular-file set, modes, and Git blob
    ids against ``revision``.  Ignored lock files are allowed only through the
    separately bound ``ignored_files`` list; no other extra file is admitted.
    """

    root = _root(root)
    _expect(archive.get("format") == "git-tree-tar-v1", "source archive format is not the reviewed Git-tree format")
    prefix = archive.get("prefix")
    _expect(isinstance(prefix, str), "source archive prefix is missing")
    _expect(not prefix.startswith("/") and "\\" not in prefix, "source archive prefix is unsafe")
    if prefix:
        _expect(prefix.endswith("/"), "source archive prefix must end with '/'")
    expected_tree = _tree_entries(root, revision)
    expected_by_path = {entry["path"]: entry for entry in expected_tree}
    expected_tree_id = _tree_id(root, revision)
    expected_tree_sha = _source_tree_digest(expected_tree)
    _expect(archive.get("git_tree") == expected_tree_id, "source archive Git tree binding differs")
    _expect(archive.get("tree_sha256") == expected_tree_sha, "source archive tracked-tree digest binding differs")

    ignored_rows = archive.get("ignored_files")
    _expect(isinstance(ignored_rows, list), "source archive ignored-file policy is missing")
    ignored: dict[str, dict[str, Any]] = {}
    for row in ignored_rows:
        _expect(isinstance(row, dict), "source archive ignored-file row is malformed")
        relative = _safe_relative(row.get("path"), "source archive ignored-file path")
        _expect(relative not in expected_by_path and relative not in ignored, f"source archive ignored-file path is duplicated or tracked: {relative}")
        _expect(isinstance(row.get("bytes"), int) and row["bytes"] >= 0, f"source archive ignored-file byte count is malformed: {relative}")
        _expect(SHA256.fullmatch(str(row.get("sha256"))) is not None, f"source archive ignored-file digest is malformed: {relative}")
        _expect(row.get("mode") in {"100644", "100755"}, f"source archive ignored-file mode is not regular: {relative}")
        ignored[relative] = dict(row)

    archive_path = Path(str(archive.get("path", "")))
    _assert_disjoint_paths({"source root": root, "source archive": archive_path})
    actual_archive = _checked_artifact(archive_path, archive, "candidate source archive")
    seen: set[str] = set()
    try:
        stream = tarfile.open(archive_path, mode="r:*")
    except (OSError, tarfile.TarError) as error:
        raise PolicyError(f"cannot open source archive: {error}") from error
    try:
        for member in stream:
            _expect(not member.issym() and not member.islnk(), f"source archive contains a link: {member.name}")
            if member.isdir():
                continue
            _expect(member.isfile(), f"source archive contains a non-regular member: {member.name}")
            name = member.name
            if prefix:
                _expect(name.startswith(prefix), f"source archive member is outside its prefix: {name}")
                relative = name[len(prefix) :]
            else:
                relative = name
            _safe_relative(relative, "source archive member")
            _expect(relative not in seen, f"source archive contains a duplicate member: {relative}")
            seen.add(relative)
            file_stream = stream.extractfile(member)
            _expect(file_stream is not None, f"source archive member cannot be read: {relative}")
            payload = file_stream.read()
            expected = expected_by_path.get(relative)
            if expected is not None:
                _expect(expected["mode"] in {"100644", "100755"}, f"selected Git tree contains a non-regular file: {relative}")
                actual_mode = "100755" if member.mode & 0o111 else "100644"
                _expect(actual_mode == expected["mode"], f"source archive mode differs from Git tree: {relative}")
                _expect(_git_blob_sha1(payload) == expected["object"], f"source archive blob differs from Git tree: {relative}")
            else:
                expected_ignored = ignored.get(relative)
                _expect(expected_ignored is not None, f"source archive contains an unbound file: {relative}")
                actual_mode = "100755" if member.mode & 0o111 else "100644"
                _expect(actual_mode == expected_ignored["mode"], f"source archive ignored-file mode differs: {relative}")
                _expect(len(payload) == expected_ignored["bytes"], f"source archive ignored-file size differs: {relative}")
                _expect(_sha256_bytes(payload) == expected_ignored["sha256"], f"source archive ignored-file digest differs: {relative}")
    except (OSError, tarfile.TarError) as error:
        raise PolicyError(f"cannot read source archive: {error}") from error
    finally:
        stream.close()
    _expect(set(expected_by_path).issubset(seen), "source archive is missing a selected Git tree file")
    return {
        "archive": actual_archive,
        "git_tree": expected_tree_id,
        "tree_sha256": expected_tree_sha,
        "tracked_files": len(expected_by_path),
        "ignored_files": sorted(ignored),
    }


def _role_lock_identities(root: Path, role: str) -> dict[str, dict[str, Any]]:
    _expect(role in {"baseline", "candidate"}, f"unknown lock role: {role}")
    result: dict[str, dict[str, Any]] = {}
    for spec in ROLE_LOCKS.values():
        relative = str(spec["path"])
        actual = _file_identity(root, relative)
        expected = spec[role]
        _expect(actual["bytes"] == expected["bytes"], f"{role} lock byte count differs: {relative}")
        _expect(actual["sha256"] == expected["sha256"], f"{role} lock hash differs: {relative}")
        result[relative] = actual
    return result


def _fixture_policy(root: Path = REPOSITORY) -> dict[str, Any]:
    root = _root(root)
    manifest_path = root / FIXTURE_INPUT["path"]
    _checked_artifact(manifest_path, FIXTURE_INPUT, "signed fixture manifest")
    try:
        value = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PolicyError(f"cannot read signed fixture manifest: {error}") from error
    _expect(isinstance(value, dict) and value.get("schema") == "litchi_signed_fixture_manifest.v1", "signed fixture schema changed")
    fixtures = value.get("fixtures")
    _expect(isinstance(fixtures, list) and fixtures, "signed fixture list is empty")
    names: set[str] = set()
    formats: set[str] = set()
    for index, fixture in enumerate(fixtures):
        _expect(isinstance(fixture, dict), f"fixture {index} is not an object")
        name = fixture.get("name")
        _expect(isinstance(name, str) and name and name not in names, f"fixture {index} name is missing or duplicated")
        relative = _safe_relative(fixture.get("relative_path"), f"fixture {name} path")
        _expect(relative.startswith("test-data/"), f"fixture {name} is outside tracked test-data")
        payload_identity = _file_identity(root, relative)
        _expect(
            payload_identity["sha256"] == fixture.get("input_sha256"),
            f"fixture {name} payload differs from its manifest digest",
        )
        format_name = fixture.get("format")
        _expect(format_name in NATIVE_FORMATS, f"fixture {name} has an unknown format")
        _expect(fixture.get("signature_presence") == "detected", f"fixture {name} lacks signature presence")
        _expect(fixture.get("cryptographic_verification") == "not_performed", f"fixture {name} makes a crypto claim")
        for key in ("input_sha256",):
            _expect(SHA256.fullmatch(str(fixture.get(key))) is not None, f"fixture {name} has no input digest")
        if fixture.get("fixture_mode") == "producer_signed":
            _expect(SHA256.fullmatch(str(fixture.get("signed_sha256"))) is not None, f"fixture {name} has no signed digest")
        else:
            _expect(fixture.get("signed_sha256") is None and fixture.get("signed_hash_mode") == "probe_generated", f"fixture {name} synthetic hash policy changed")
        _expect(isinstance(fixture.get("target"), str) and fixture["target"], f"fixture {name} target is missing")
        _expect(isinstance(fixture.get("property"), str) and fixture["property"], f"fixture {name} property is missing")
        names.add(name)
        formats.add(format_name)
    _expect(formats == set(NATIVE_FORMATS), f"signed fixture formats differ: {sorted(formats)}")
    return {"artifact": FIXTURE_INPUT, "schema": value["schema"], "fixtures": len(fixtures), "formats": list(NATIVE_FORMATS)}


def _validate_route_contract(value: Mapping[str, Any]) -> None:
    _expect(tuple(value.get("formats", ())) == NATIVE_FORMATS, "policy format order differs")
    routes = value.get("routes")
    _expect(isinstance(routes, list) and tuple(routes) == ROUTES, "policy route list differs")
    _expect(set(value.get("correctness_only", ())) == set(CORRECTNESS_ONLY), "correctness-only route set differs")
    _expect(set(value.get("baseline_not_applicable", ())) == set(BASELINE_NOT_APPLICABLE), "baseline disposition set differs")
    semantics = value.get("semantics")
    _expect(isinstance(semantics, dict), "signed semantic contract is missing")
    expected_semantics = {
        "signature_presence": "detected_metadata_only",
        "cryptographic_verification": "not_performed",
        "noop": "diagnostics_unchanged_patch_noop_signed_bytes_equal",
        "cancellation": "cancelled_intermediate_edit_remains_signed_and_does_not_publish",
        "replay": "content_bound_base_match_signed_bytes_equal",
        "freshness": "SourceChanged_io_error_without_publication",
        "refusal": "ProtectedDocument_document_protected_without_publication",
        "unsigned_negative": "PatchConflict_invalid_input_without_signature_remarking",
        "budget": "LimitExceeded_operations_without_publication",
        "unsign": "new_unsigned_output_signature_parts_absent_stream_bounded",
    }
    _expect(semantics == expected_semantics, "signed semantic contract changed")


def _validate_policy_document(value: Mapping[str, Any], *, execution: bool) -> dict[str, Any]:
    _expect(value.get("schema") == SCHEMA, "signed-native policy schema differs")
    _expect(value.get("policy_version") == POLICY_VERSION, "signed-native policy version differs")
    _expect(value.get("historical") == HISTORICAL_SOURCE, "historical v1 boundary changed")
    _validate_route_contract(value)
    candidate = value.get("candidate")
    _expect(isinstance(candidate, dict), "candidate policy is missing")
    _expect(candidate.get("source_manifests") == list(SOURCE_MANIFEST_PATHS), "candidate manifest paths changed")
    _expect(candidate.get("vendor_roots") == list(VENDOR_ROOTS), "candidate vendor roots changed")
    _expect(candidate.get("role_locks") == ROLE_LOCKS, "candidate role lock policy changed")
    _expect(candidate.get("fixture") == {"path": FIXTURE_INPUT["path"], "sha256": FIXTURE_INPUT["sha256"], "bytes": FIXTURE_INPUT["bytes"]}, "candidate fixture binding changed")
    _expect(candidate.get("probe_inputs") == PROBE_INPUTS, "candidate probe input policy changed")
    _expect(candidate.get("toolchain_paths") == list(TOOLCHAIN_PATHS), "toolchain paths changed")
    _expect(candidate.get("toolchain") == TOOLCHAIN_IDENTITIES, "toolchain identity policy changed")
    _expect(candidate.get("dependency_transitions") == [
        {"package": "allsorts", "version": "0.17.0", "baseline_source": "registry", "candidate_source": "path", "candidate_path": "vendor/allsorts"},
        {"package": "flate2", "version": "1.1.10", "baseline_source": "registry", "candidate_source": "path", "candidate_path": "vendor/flate2"},
        {"package": "zlib-rs", "version": "0.6.8", "baseline_source": "registry", "candidate_source": "path", "candidate_path": "vendor/zlib-rs"},
    ], "dependency path policy changed")
    _expect(value.get("signed_batch") == SIGNED_BATCH_POLICY, "signed-batch owner choice changed")
    _expect(value.get("adapter") == ADAPTER_POLICY, "signed-native consumer adapter policy changed")
    _expect(value.get("execution") == EXECUTION_POLICY, "signed-native execution receipt policy changed")
    status = candidate.get("status")
    _expect(status in {"pending_final_source_freeze", "bound"}, "candidate policy status is invalid")
    source = candidate.get("source")
    _expect(isinstance(source, dict), "candidate source identity is missing")
    if status == "pending_final_source_freeze":
        for key in ("commit", "git_tree", "tree_sha256", "archive"):
            _expect(source.get(key) is None, f"pending candidate supplied a {key}")
        _expect(candidate.get("manifests") is None, "pending candidate supplied manifests")
        _expect(candidate.get("vendor_tree") is None, "pending candidate supplied vendor identity")
        _expect(candidate.get("lock_graph") is None, "pending candidate supplied lock graph")
        _expect(candidate.get("features") is None, "pending candidate supplied feature selection")
        _expect(status == "pending_final_source_freeze", "candidate pending status changed")
        if execution:
            raise IncompletePolicyError(
                "signed-native v2 candidate is incomplete: final source commit, tree, archive, "
                "manifests, lock graph, vendor files, and features are not frozen"
            )
        return {"status": status, "execution": "refused_pending_source_freeze"}

    _expect(CANDIDATE_BINDING is not None, "manifest claims a bound candidate but code binding is still incomplete")
    _validate_bound_binding(CANDIDATE_BINDING)
    _expect(candidate == CANDIDATE_BINDING, "manifest candidate differs from the trusted code binding")
    return {"status": status, "execution": "bound"}


def _validate_bound_binding(binding: Mapping[str, Any]) -> None:
    _expect(binding.get("status") == "bound", "candidate binding is not marked bound")
    _expect(binding.get("source_manifests") == list(SOURCE_MANIFEST_PATHS), "bound source manifest paths are incomplete")
    _expect(binding.get("role_locks") == ROLE_LOCKS, "bound role lock policy differs")
    _expect(binding.get("vendor_roots") == list(VENDOR_ROOTS), "bound vendor roots differ")
    _expect(binding.get("fixture") == FIXTURE_INPUT, "bound fixture identity differs")
    _expect(binding.get("probe_inputs") == PROBE_INPUTS, "bound probe input identities differ")
    _expect(binding.get("toolchain_paths") == list(TOOLCHAIN_PATHS), "bound toolchain paths differ")
    _expect(binding.get("toolchain") == TOOLCHAIN_IDENTITIES, "bound toolchain identities differ")
    source = binding.get("source")
    _expect(isinstance(source, dict), "bound candidate source is missing")
    commit = source.get("commit")
    tree = source.get("git_tree")
    tree_sha = source.get("tree_sha256")
    _expect(isinstance(commit, str) and FULL_SHA.fullmatch(commit) is not None, "bound candidate commit is not a full id")
    _expect(isinstance(tree, str) and FULL_SHA.fullmatch(tree) is not None, "bound candidate Git tree is not a full id")
    _expect(isinstance(tree_sha, str) and SHA256.fullmatch(tree_sha) is not None, "bound candidate tree digest is malformed")
    archive = source.get("archive")
    _expect(isinstance(archive, dict), "bound candidate source archive is missing")
    archive_path = Path(str(archive.get("path", "")))
    _expect(isinstance(archive.get("path"), str) and archive_path.is_absolute(), "bound source archive path is not absolute")
    _expect("3rdparty" not in archive_path.parts, "bound source archive may not use an upstream checkout")
    try:
        archive_path.resolve(strict=False).relative_to(REPOSITORY.resolve())
    except ValueError:
        pass
    else:
        raise PolicyError("bound source archive must remain outside this checkout")
    _expect(isinstance(archive.get("bytes"), int) and archive["bytes"] > 0, "bound source archive size is missing")
    _expect(isinstance(archive.get("sha256"), str) and SHA256.fullmatch(archive["sha256"]) is not None, "bound source archive hash is malformed")
    _expect(archive.get("format") == "git-tree-tar-v1", "bound source archive format is not the reviewed Git-tree format")
    _expect(isinstance(archive.get("prefix"), str) and (not archive["prefix"] or archive["prefix"].endswith("/")), "bound source archive prefix is malformed")
    _expect(archive.get("git_tree") == tree, "bound source archive is not tied to the selected Git tree")
    _expect(archive.get("tree_sha256") == tree_sha, "bound source archive is not tied to the selected tracked-tree digest")
    expected_ignored = [
        {
            "path": ROLE_LOCKS["workspace"]["path"],
            "bytes": ROLE_LOCKS["workspace"]["candidate"]["bytes"],
            "sha256": ROLE_LOCKS["workspace"]["candidate"]["sha256"],
            "mode": "100644",
        }
    ]
    _expect(archive.get("ignored_files") == expected_ignored, "bound source archive ignored-lock policy differs")
    manifests = binding.get("manifests")
    _expect(isinstance(manifests, dict) and set(manifests) == set(SOURCE_MANIFEST_PATHS), "bound source manifests are incomplete")
    for relative, identity in manifests.items():
        _safe_relative(relative, "bound source manifest")
        _expect(isinstance(identity, dict) and isinstance(identity.get("bytes"), int) and SHA256.fullmatch(str(identity.get("sha256"))) is not None, f"bound source manifest is malformed: {relative}")
    vendor = binding.get("vendor_tree")
    _expect(isinstance(vendor, dict), "bound vendor tree is missing")
    _expect(isinstance(vendor.get("files"), int) and vendor["files"] > 0, "bound vendor file count is missing")
    _expect(isinstance(vendor.get("bytes"), int) and vendor["bytes"] > 0, "bound vendor byte count is missing")
    _expect(SHA256.fullmatch(str(vendor.get("sha256"))) is not None, "bound vendor digest is malformed")
    entries = vendor.get("entries")
    _expect(isinstance(entries, list) and len(entries) == vendor["files"], "bound vendor file list is incomplete")
    paths: set[str] = set()
    for entry in entries:
        _expect(isinstance(entry, dict), "bound vendor entry is malformed")
        path = _safe_relative(entry.get("path"), "bound vendor path")
        _expect(path not in paths and any(path.startswith(f"{root}/") for root in VENDOR_ROOTS), "bound vendor path is outside the reviewed roots or duplicated")
        _expect(entry.get("mode") in {"100644", "100755"}, f"bound vendor mode is not regular: {path}")
        _expect(FULL_SHA.fullmatch(str(entry.get("git_blob_sha1"))) is not None, f"bound vendor Git blob is malformed: {path}")
        _expect(isinstance(entry.get("bytes"), int) and entry["bytes"] >= 0 and SHA256.fullmatch(str(entry.get("sha256"))) is not None, f"bound vendor digest is malformed: {path}")
        paths.add(path)
    graph = binding.get("lock_graph")
    _expect(isinstance(graph, dict) and set(graph) == set(ROLE_LOCKS), "bound lock graph is incomplete")
    for role, identity in graph.items():
        _expect(isinstance(identity, dict) and identity.get("path") == ROLE_LOCKS[role]["path"], f"bound lock graph path differs: {role}")
        _expect(isinstance(identity.get("bytes"), int) and SHA256.fullmatch(str(identity.get("sha256"))) is not None, f"bound lock graph identity is malformed: {role}")
        _expect(SHA256.fullmatch(str(identity.get("graph_sha256"))) is not None, f"bound lock graph digest is malformed: {role}")
    features = binding.get("features")
    _expect(isinstance(features, dict) and features.get("package") == "litchi-ops" and features.get("all_features") is True and isinstance(features.get("selected"), list) and features["selected"], "bound feature selection is incomplete")


def load_policy() -> dict[str, Any]:
    """Load only the checked-in policy; arbitrary policy paths are unsupported."""

    _reject_symlink_components(POLICY_PATH, "signed-native policy")
    _expect(POLICY_PATH.is_file() and not POLICY_PATH.is_symlink(), f"signed-native policy is missing: {POLICY_PATH}")
    try:
        value = json.loads(POLICY_PATH.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PolicyError(f"cannot read signed-native policy: {error}") from error
    _expect(isinstance(value, dict), "signed-native policy root is not an object")
    return value


def validate_policy(*, execution: bool = False) -> dict[str, Any]:
    """Validate the trusted checked-in policy and return its disposition."""

    value = load_policy()
    result = _validate_policy_document(value, execution=execution)
    # Validate the fixture bytes and schema even while candidate preparation is
    # pending.  This ensures a pending policy cannot hide a changed fixture.
    _expect(value["candidate"]["fixture"] == FIXTURE_INPUT, "fixture identity is not fixed")
    _fixture_policy()
    return result


def _lock_graph_identity(root: Path, role: str) -> str:
    relative = str(ROLE_LOCKS[role]["path"])
    path = root / relative
    _reject_symlink_components(path, f"{role} lock graph")
    try:
        lock = tomllib.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, tomllib.TOMLDecodeError) as error:
        raise PolicyError(f"{role} lock graph is not valid TOML: {error}") from error
    packages = lock.get("package")
    _expect(isinstance(packages, list), f"{role} lock graph has no package list")
    normalized: list[dict[str, Any]] = []
    for package in packages:
        _expect(isinstance(package, dict), f"{role} lock graph has a malformed package")
        name = package.get("name")
        version = package.get("version")
        _expect(isinstance(name, str) and name and isinstance(version, str) and version, f"{role} lock graph package identity is incomplete")
        dependencies = package.get("dependencies", [])
        _expect(isinstance(dependencies, list) and all(isinstance(item, str) for item in dependencies), f"{role} lock graph dependencies are malformed")
        normalized.append(
            {
                "name": name,
                "version": version,
                "source": package.get("source"),
                "checksum": package.get("checksum"),
                "dependencies": sorted(dependencies),
            }
        )
    normalized.sort(key=lambda item: (item["name"], item["version"], item["source"] or ""))
    return _sha256_bytes(_canonical(normalized))


def candidate_identity(root: Path) -> dict[str, Any]:
    """Verify a bound candidate tree against the code-trusted binding.

    The public preparation path uses only the checked-in
    ``CANDIDATE_BINDING`` constant and therefore cannot trust an arbitrary
    digest supplied in a manifest or command line.
    """

    trusted = CANDIDATE_BINDING
    _expect(trusted is not None, "candidate source identity is not frozen")
    _validate_bound_binding(trusted)
    root = _root(root)
    expected_source = trusted["source"]
    archive_path = Path(str(expected_source["archive"]["path"]))
    _assert_disjoint_paths({"candidate source": root, "candidate archive": archive_path})
    expected_commit = str(expected_source["commit"])
    _expect(_head(root) == expected_commit, "candidate HEAD differs from the bound source commit")
    _expect(not _status_paths(root), "candidate source must be clean")
    entries = _tree_entries(root)
    _expect(_tree_id(root, "HEAD") == expected_source["git_tree"], "candidate Git tree differs from the bound source")
    _expect(_source_tree_digest(entries) == expected_source["tree_sha256"], "candidate tracked-tree digest differs")
    manifests = {path: _file_identity(root, path) for path in SOURCE_MANIFEST_PATHS}
    _expect(manifests == trusted["manifests"], "candidate source manifests differ")
    locks = _role_lock_identities(root, "candidate")
    expected_graph = trusted["lock_graph"]
    for role, identity in locks.items():
        _expect(expected_graph[role]["path"] == identity["path"], f"candidate lock path differs: {role}")
        _expect(expected_graph[role]["bytes"] == identity["bytes"], f"candidate lock bytes differ: {role}")
        _expect(expected_graph[role]["sha256"] == identity["sha256"], f"candidate lock hash differs: {role}")
        _expect(_lock_graph_identity(root, role) == expected_graph[role]["graph_sha256"], f"candidate lock graph differs: {role}")
    toolchain = {
        path: _file_identity(root, path) for path in TOOLCHAIN_PATHS
    }
    _expect(toolchain == TOOLCHAIN_IDENTITIES, "candidate toolchain inputs differ")
    probe_inputs = {
        name: _file_identity(root, str(item["path"]))
        for name, item in PROBE_INPUTS.items()
    }
    _expect(probe_inputs == PROBE_INPUTS, "candidate probe inputs differ")
    fixture = _file_identity(root, FIXTURE_INPUT["path"])
    _expect(fixture == FIXTURE_INPUT, "candidate fixture input differs")
    _fixture_policy(root)
    vendor = _vendor_tree_identity(root)
    expected_vendor = trusted["vendor_tree"]
    _expect(vendor == expected_vendor, "candidate vendor tree differs from the bound file list")
    archive = _archive_tree_identity(root, "HEAD", expected_source["archive"])
    archive_identity = dict(expected_source["archive"])
    archive_identity.update(archive["archive"])
    lock_graph = {
        role: {
            "path": identity["path"],
            "bytes": identity["bytes"],
            "sha256": identity["sha256"],
            "graph_sha256": expected_graph[role]["graph_sha256"],
        }
        for role, identity in locks.items()
    }
    _expect(
        all(_lock_graph_identity(root, role) == expected_graph[role]["graph_sha256"] for role in ROLE_LOCKS),
        "candidate lock graph differs",
    )
    return {
        "source_root": str(root),
        "commit": expected_commit,
        "git_tree": expected_source["git_tree"],
        "tree_sha256": expected_source["tree_sha256"],
        "manifests": manifests,
        "locks": locks,
        "lock_graph": lock_graph,
        "vendor_tree": vendor,
        "toolchain": toolchain,
        "probe_inputs": probe_inputs,
        "fixture": fixture,
        "archive": archive_identity,
        "features": trusted["features"],
    }


def prepare(source_root: Path, output_root: Path) -> dict[str, Any]:
    """Prepare the external v2 consumer after the checked-in source gate passes.

    The current checked-in policy remains pending and therefore returns before
    this import.  Once a reviewed binding is committed, all source checks are
    completed before the consumer creates its fresh output root.
    """

    validate_policy(execution=True)
    source_root = _root(source_root)
    _expect(CANDIDATE_BINDING is not None, "candidate source identity is not frozen")
    archive_path = Path(str(CANDIDATE_BINDING["source"]["archive"]["path"]))
    _assert_disjoint_paths(
        {
            "candidate source": source_root,
            "candidate archive": archive_path,
            "preparation output": output_root,
        }
    )
    identity = candidate_identity(source_root)
    # Import lazily so the pending source-only policy has no consumer side
    # effects and so this module remains usable by the policy tests.
    from native_qualification import signed_native_v2_consumer as consumer

    return consumer.prepare(source_root, output_root, identity=identity)


def execute(
    prepared_root: Path | None = None,
    capabilities: Path | None = None,
    *,
    output: Path | None = None,
) -> dict[str, Any]:
    """Consume a built v2 receipt and typed protocol capabilities.

    The execution gate is checked before the consumer is imported.  Thus a
    pending policy still fails closed and cannot create an output directory or
    relabel a historical v1 receipt.
    """

    validate_policy(execution=True)
    _expect(CANDIDATE_BINDING is not None, "candidate source identity is not frozen")
    from native_qualification import signed_native_v2_consumer as consumer

    return consumer.execute(prepared_root, capabilities, output=output)


def reject_v1_receipt(receipt: Mapping[str, Any]) -> None:
    """Prevent a historical ul24 receipt from being relabelled as v2."""

    _expect(isinstance(receipt, Mapping), "receipt is not an object")
    _expect(receipt.get("schema") != "litchi_signed_native_qualification.v1", "historical ul24 receipt cannot be relabelled as v2")
    _expect(receipt.get("schema") != "litchi_signed_native_probe.v1", "historical ul24 probe receipt cannot be relabelled as v2")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    arguments = parser.parse_args(argv)
    try:
        prepare(arguments.source_root, arguments.output_root)
    except (OSError, ValueError) as error:
        parser.error(str(error))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
