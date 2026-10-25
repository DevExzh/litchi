#!/usr/bin/env python3
"""Prepare the version-two signed-native qualification source pair.

This module is deliberately a source-only boundary.  It records the two
immutable signed-native endpoints and verifies each source closure before a
consumer or a native process may be started.  The historical candidate
binding constant remains for compatibility with the earlier pending policy,
but endpoint preparation uses the separately pinned f2 assembly for `before`
and the complete checked-in candidate binding for `after`; the latter fails
closed while it is absent.  No feature flag selects an API spelling.
Preparation delegates to the
tracked consumer in :mod:`native_qualification.signed_native_v2_consumer`.

The old 0362ul24 receipts are historical v1 evidence.  They are never
relabelled as v2, and this module does not accept a v1 receipt as a v2 source
or qualification result.

No Cargo, compiler, native probe, rustdoc, or timing command is invoked by
this source-only preparation path.  All digests used for a bound endpoint are
recomputed from the selected source tree; a caller cannot supply a digest or
commit on the command line to make an unbound policy executable.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
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
    "legacy_feature": None,
    "api_branches": {
        "before": ["legacy_commit_fields", "legacy_xlsb_option"],
        "after": ["current_commit_accessors", "current_xlsb_resolve"],
    },
    "route_count": 16,
}
EXECUTION_POLICY = {
    "runner_schema": "litchi.signed-native.runner.v2",
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

# Endpoint bindings are deliberately separate.  The historical f2 checkout
# cannot compile the current API and the 472b checkout cannot compile the
# removed API; a feature flag in one generated project is therefore not an
# endpoint selector.  These are provenance pins, not result receipts.  The
# actual source/archive/lock graph is revalidated by ``endpoint_identity``
# before a consumer directory is created.
ENDPOINTS = ("before", "after")
ENDPOINT_SOURCE_KINDS = {"before": "f2", "after": "candidate"}
ENDPOINT_API_BRANCH = {"before": "legacy_api", "after": "current_api"}
ENDPOINT_ROLE = {"before": "baseline", "after": "candidate"}
ENDPOINT_ROUTE_STATUS = {
    "before": {
        **{route: "capability_missing" for route in ROUTES if route.endswith("/noop_commit") or route.endswith("/noop_write") or route.endswith("/noop_replay")},
        **{route: "expected_refusal" for route in ROUTES if route.endswith("/refusal")},
        **{route: "available" for route in ROUTES if route.endswith("/unsign_write")},
        **{route: "not_applicable" for route in BASELINE_NOT_APPLICABLE},
    },
    "after": {
        **{route: "available" for route in TIMED_ROUTES},
        **{route: "expected_refusal" for route in CORRECTNESS_ONLY},
    },
}

HISTORICAL_ENDPOINT_HARNESS = {
    "before": {
        "role": "baseline",
        "source_kind": "f2",
        "commit": "f2bbbc8d07ab5da26c180f1ff74813d39f632d54",
        "source_root": "/home/zhuhe/code/litchi-worktrees/.coord/w8-f2-baseline-sAirXf/src",
        "source_archive": {
            "path": "/home/zhuhe/code/litchi-worktrees/.coord/w8-f2-baseline-sAirXf/f2-source.tar",
            "bytes": 6614528000,
            "sha256": "de625729ec8e6f8055e619443ad31bf1a91719cd31465975f951d54dbc1c2280",
        },
        "source_assembly": {
            "path": "/home/zhuhe/code/litchi-worktrees/.coord/w8-f2-baseline-sAirXf/out/provenance/original-baseline-source-assembly.json",
            "bytes": 787,
            "sha256": "b9de004ed2a74a509ebc2e3879b8d8812fcf543b92cb8bccf740ec82486ea44d",
        },
        "lock": {
            "path": "Cargo.lock",
            "bytes": 139441,
            "sha256": "1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c",
        },
    },
    "after": {
        "role": "candidate",
        "source_kind": "472b",
        "commit": "472b6705475b073e7fe08f6e11ee17fb1e29f838",
        "source_root": "/home/zhuhe/code/litchi-worktrees/.coord/final-api-capture-472b670547/source",
        "source_archive": {
            "path": "/home/zhuhe/code/litchi-worktrees/.coord/w8-472b-native.V1lU2E/candidate/472b-source.tar",
            "bytes": 6634147840,
            "sha256": "44c11e933974c8fa3d6a67b46739bedeef866e60fbd4360ffe7f3bb8a4dc2274",
        },
        "source_assembly": {
            "path": "/home/zhuhe/code/litchi-worktrees/.coord/w8-472b-native.V1lU2E/candidate/source-assembly.json",
            "bytes": 483,
            "sha256": "cc9d9f67e8d4474685fa0daf4ba9054adcb508ca6dcc047f49fda39aa6b1289f",
        },
        "lock": {
            "path": "Cargo.lock",
            "bytes": 139606,
            "sha256": "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1",
        },
    },
}

# The historical 472b capture remains useful for the reviewed harness and
# probe provenance, but it is not an executable after endpoint.  The after
# endpoint is selected only by the complete checked-in CANDIDATE_BINDING.
BEFORE_SOURCE_BINDING = HISTORICAL_ENDPOINT_HARNESS["before"]
HISTORICAL_AFTER_HARNESS = HISTORICAL_ENDPOINT_HARNESS["after"]
ENDPOINT_SOURCE_BINDINGS = {"before": BEFORE_SOURCE_BINDING}
ENDPOINT_POLICY_BINDINGS = {
    "before": BEFORE_SOURCE_BINDING,
    "after": {
        "role": "candidate",
        "source_kind": "candidate",
        "binding": "CANDIDATE_BINDING",
        "status": "pending_final_source_freeze",
    },
}

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
            "sha256": "57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f",
            "bytes": 57028,
        },
        "candidate": {
            "sha256": "de1f8ae53b3455a4246da0b9340896881cd72ca0295b8e52444e0d55e6c9c4ea",
            "bytes": 56941,
        },
    },
    "native": {
        "path": "tools/native-resave/Cargo.lock",
        "baseline": {
            "sha256": "6aaa0e3800a4db72afbe8783aab382accdb4a6355d06bb97fb1c12dd78839408",
            "bytes": 56929,
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

# The legacy candidate slot remains ``None`` because it represented the
# pre-endpoint pending policy.  It is not consulted by the endpoint path when
# the reviewed 472b source is selected; there is no CLI or JSON override.
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


def _harness_file_path(relative: str) -> Path:
    """Return a reviewed harness input from this immutable source checkout."""

    _safe_relative(relative, "harness input")
    path = REPOSITORY / relative
    _reject_symlink_components(path, f"harness input {relative}")
    _expect(path.is_file() and not path.is_symlink(), f"reviewed harness input is missing: {path}")
    return path


def _probe_input_path(root: Path, endpoint: str, relative: str) -> Path:
    _expect(endpoint in ENDPOINTS, f"unknown signed-native endpoint: {endpoint}")
    # The immutable endpoint trees do not carry the 0362ul24 docs inputs.  A
    # reviewed endpoint-specific input is accepted when it is actually
    # present (which keeps the source-only fixtures honest); otherwise the
    # explicitly pinned shared protocol/harness input is staged.  Missing
    # production paths must never become a fabricated source identity.
    candidate = root / relative
    _reject_symlink_components(candidate, f"{endpoint} probe input")
    if candidate.exists() or candidate.is_symlink():
        _expect(candidate.is_file() and not candidate.is_symlink(), f"{endpoint} probe input is not a regular file: {candidate}")
        return candidate
    return _harness_file_path(relative)


def _fixture_manifest_path(root: Path, endpoint: str) -> Path:
    _expect(endpoint in ENDPOINTS, f"unknown signed-native endpoint: {endpoint}")
    candidate = root / FIXTURE_INPUT["path"]
    _reject_symlink_components(candidate, f"{endpoint} fixture manifest")
    if candidate.exists() or candidate.is_symlink():
        _expect(candidate.is_file() and not candidate.is_symlink(), f"{endpoint} fixture manifest is not a regular file: {candidate}")
        return candidate
    return _harness_file_path(FIXTURE_INPUT["path"])


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


def _fixture_policy(
    root: Path = REPOSITORY,
    *,
    manifest_root: Path | None = None,
    allow_extracted: bool = False,
) -> dict[str, Any]:
    if allow_extracted:
        try:
            root = _root(root)
        except PolicyError:
            _reject_symlink_components(root, "fixture source root")
            root = root.resolve(strict=False)
            _expect(root.is_dir() and (root / "Cargo.toml").is_file(), f"fixture source root is not a usable workspace: {root}")
    else:
        root = _root(root)
    manifest_path = (manifest_root or root) / FIXTURE_INPUT["path"]
    _reject_symlink_components(manifest_path, "signed fixture manifest")
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
    _expect(value.get("endpoints") == ENDPOINT_POLICY_BINDINGS, "endpoint source bindings changed")
    _expect(value.get("historical_harness") == HISTORICAL_AFTER_HARNESS, "historical harness provenance changed")
    _validate_route_contract(value)
    endpoint_routes = value.get("endpoint_routes")
    _expect(endpoint_routes == ENDPOINT_ROUTE_STATUS, "endpoint route disposition matrix changed")
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
    assembly = binding.get("source_assembly")
    _expect(isinstance(assembly, dict), "bound candidate source assembly is missing")
    assembly_path = Path(str(assembly.get("path", "")))
    _expect(isinstance(assembly.get("path"), str) and assembly_path.is_absolute(), "bound source assembly path is not absolute")
    _expect("3rd" + "party" not in assembly_path.parts, "bound source assembly may not use an upstream checkout")
    try:
        assembly_path.resolve(strict=False).relative_to(REPOSITORY.resolve())
    except ValueError:
        pass
    else:
        raise PolicyError("bound source assembly must remain outside this checkout")
    _expect(isinstance(assembly.get("bytes"), int) and assembly["bytes"] > 0, "bound source assembly size is missing")
    _expect(isinstance(assembly.get("sha256"), str) and SHA256.fullmatch(assembly["sha256"]) is not None, "bound source assembly hash is malformed")
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
    probe_inputs = {}
    probe_provenance = {}
    for name, item in PROBE_INPUTS.items():
        probe_path = _probe_input_path(root, "after", str(item["path"]))
        actual = _artifact_identity(probe_path, f"candidate {name} probe input")
        _expect(
            {"path": str(item["path"]), "bytes": actual["bytes"], "sha256": actual["sha256"]} == item,
            f"candidate probe input differs: {name}",
        )
        probe_inputs[name] = dict(item)
        probe_provenance[name] = actual
    _expect(probe_inputs == PROBE_INPUTS, "candidate probe inputs differ")
    fixture_path = _fixture_manifest_path(root, "after")
    fixture_actual = _artifact_identity(fixture_path, "candidate fixture manifest")
    fixture = dict(FIXTURE_INPUT)
    _expect(
        {"path": FIXTURE_INPUT["path"], "bytes": fixture_actual["bytes"], "sha256": fixture_actual["sha256"]} == FIXTURE_INPUT,
        "candidate fixture input differs",
    )
    _expect(fixture == FIXTURE_INPUT, "candidate fixture input differs")
    fixture_provenance = fixture_actual
    manifest_root = REPOSITORY if fixture_path == REPOSITORY / FIXTURE_INPUT["path"] else root
    _fixture_policy(root, manifest_root=manifest_root)
    vendor = _vendor_tree_identity(root)
    expected_vendor = trusted["vendor_tree"]
    _expect(vendor == expected_vendor, "candidate vendor tree differs from the bound file list")
    archive = _archive_tree_identity(root, "HEAD", expected_source["archive"])
    archive_identity = dict(expected_source["archive"])
    archive_identity.update(archive["archive"])
    assembly_spec = trusted["source_assembly"]
    assembly_identity = _checked_artifact(
        Path(str(assembly_spec["path"])),
        assembly_spec,
        "candidate source assembly",
    )
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
        "probe_provenance": probe_provenance,
        "fixture": fixture,
        "fixture_provenance": fixture_provenance,
        "archive": archive_identity,
        "source_assembly": assembly_identity,
        "features": trusted["features"],
    }


def _reviewed_source_helper() -> Any:
    """Load the reviewed W8 source-closure verifier without executing it."""

    path = REPOSITORY / "docs/unified-ops/results/0362ul24/prepare_signed_native_consumer.py"
    _expect(path.is_file() and not path.is_symlink(), f"reviewed source verifier is missing: {path}")
    spec = importlib.util.spec_from_file_location("signed_native_v1_source_fence", path)
    _expect(spec is not None and spec.loader is not None, "cannot load reviewed source verifier")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        raise PolicyError(f"reviewed source verifier failed to load: {error}") from error
    return module


def _endpoint_root(value: Path, endpoint: str) -> Path:
    """Resolve a pinned endpoint, allowing the extracted f2 tree without .git."""

    _expect(endpoint in ENDPOINTS, f"unknown signed-native endpoint: {endpoint}")
    try:
        return _root(value)
    except PolicyError:
        # The reviewed f2 source is an extracted, immutable workspace rather
        # than a Git checkout.  The source verifier below supplies its commit
        # and archive fence; do not silently accept this fallback for 472b.
        _expect(endpoint == "before", "candidate source must be a Git worktree")
        _reject_symlink_components(value, "baseline source root")
        root = value.resolve(strict=False)
        _expect(root.is_dir(), f"baseline source root is not a directory: {root}")
        _expect((root / "Cargo.toml").is_file(), f"baseline source lacks Cargo.toml: {root}")
        return root


def _endpoint_source_binding(endpoint: str) -> Mapping[str, Any]:
    _expect(endpoint in ENDPOINTS, f"unknown signed-native endpoint: {endpoint}")
    if endpoint == "before":
        return ENDPOINT_SOURCE_BINDINGS[endpoint]
    return _candidate_endpoint_binding()


def _candidate_endpoint_binding() -> Mapping[str, Any]:
    """Project the complete checked-in candidate binding into endpoint pins."""

    trusted = CANDIDATE_BINDING
    if trusted is None:
        raise IncompletePolicyError(
            "signed-native v2 after endpoint is unavailable: CANDIDATE_BINDING is not frozen"
        )
    _expect(isinstance(trusted, Mapping), "candidate binding is malformed")
    source = trusted.get("source")
    locks = trusted.get("role_locks")
    _expect(isinstance(source, Mapping), "candidate source binding is missing")
    _expect(isinstance(locks, Mapping), "candidate role-lock binding is missing")
    _expect(set(locks) == set(ROLE_LOCKS), "candidate role-lock set is incomplete")
    for role, expected in ROLE_LOCKS.items():
        lock = locks.get(role)
        _expect(isinstance(lock, Mapping), f"candidate {role} lock binding is missing")
        _expect(lock.get("path") == expected["path"], f"candidate {role} lock path differs")
        candidate_lock = lock.get("candidate")
        _expect(isinstance(candidate_lock, Mapping), f"candidate {role} lock identity is missing")
        _expect(isinstance(candidate_lock.get("bytes"), int) and candidate_lock["bytes"] > 0, f"candidate {role} lock size is missing")
        _expect(SHA256.fullmatch(str(candidate_lock.get("sha256"))) is not None, f"candidate {role} lock digest is malformed")
    workspace = locks["workspace"]
    candidate_lock = workspace["candidate"]
    archive = source.get("archive")
    _expect(isinstance(archive, Mapping), "candidate source archive binding is missing")
    _expect(isinstance(source.get("commit"), str) and FULL_SHA.fullmatch(source["commit"]) is not None, "candidate source commit is missing")
    _expect(isinstance(source.get("git_tree"), str) and FULL_SHA.fullmatch(source["git_tree"]) is not None, "candidate source tree is missing")
    _expect(isinstance(source.get("tree_sha256"), str) and SHA256.fullmatch(source["tree_sha256"]) is not None, "candidate source digest is missing")
    assembly = trusted.get("source_assembly")
    _expect(isinstance(assembly, Mapping), "candidate source assembly binding is missing")
    _expect(isinstance(assembly.get("path"), str) and Path(assembly["path"]).is_absolute(), "candidate source assembly path is missing")
    _expect(isinstance(assembly.get("bytes"), int) and assembly["bytes"] > 0, "candidate source assembly size is missing")
    _expect(SHA256.fullmatch(str(assembly.get("sha256"))) is not None, "candidate source assembly digest is malformed")
    return {
        "role": "candidate",
        "source_kind": "candidate",
        "commit": source.get("commit"),
        "source_archive": archive,
        "source_assembly": trusted.get("source_assembly"),
        "role_locks": locks,
        "lock": {
            "path": workspace.get("path"),
            "bytes": candidate_lock.get("bytes"),
            "sha256": candidate_lock.get("sha256"),
        },
    }


def _endpoint_lock_spec(endpoint: str) -> dict[str, dict[str, Any]]:
    binding = _endpoint_source_binding(endpoint)
    if endpoint == "after" and isinstance(binding.get("role_locks"), Mapping):
        role_locks = binding["role_locks"]
        role = "candidate"
    else:
        role_locks = ROLE_LOCKS
        role = "baseline"
    result: dict[str, dict[str, Any]] = {}
    for name, spec in role_locks.items():
        _expect(isinstance(spec, Mapping), f"{endpoint} {name} lock policy is malformed")
        selected = spec.get(role)
        _expect(isinstance(selected, Mapping), f"{endpoint} {name} lock identity is missing")
        result[name] = {
            "path": spec.get("path"),
            "bytes": selected.get("bytes"),
            "sha256": selected.get("sha256"),
        }
    return result


def _endpoint_identity_from_reviewed(root: Path, endpoint: str) -> dict[str, Any]:
    """Build the v2 closure from the retained W8 source verifier.

    This verifies the extracted f2 before assembly independently.  The after
    endpoint never uses this historical verifier; it is checked by
    candidate_identity against the complete candidate binding.  No API
    feature is used to switch between source graphs.
    """

    _expect(endpoint == "before", "historical harness provenance is not an after endpoint")
    binding = _endpoint_source_binding(endpoint)
    kind = str(binding["source_kind"])
    helper = _reviewed_source_helper()
    archive_path = Path(str(binding["source_archive"]["path"]))
    # The reviewed helper performs the fixed archive hash and, for both
    # endpoints, the member/tree closure proof.  Avoid a second full archive
    # hash here; the returned artifact is compared with this static pin below.
    assembly_spec = binding.get("source_assembly")
    assembly_path: Path | None = None
    assembly_artifact: dict[str, Any] | None = None
    if isinstance(assembly_spec, Mapping):
        assembly_path = Path(str(assembly_spec["path"]))
        _reject_symlink_components(assembly_path, f"{endpoint} source assembly")
        _expect(assembly_path.is_file() and not assembly_path.is_symlink(), f"{endpoint} source assembly is missing")
        assembly_artifact = _artifact_identity(assembly_path, f"{endpoint} source assembly")
        if isinstance(assembly_spec.get("bytes"), int):
            _expect(assembly_artifact["bytes"] == assembly_spec["bytes"], f"{endpoint} source assembly size changed")
        _expect(assembly_artifact["sha256"] == assembly_spec["sha256"], f"{endpoint} source assembly digest changed")
    try:
        reviewed = helper.source_identity(
            root,
            kind,
            source_archive=archive_path,
            source_assembly=assembly_path,
            verify_closure=True,
        )
    except Exception as error:
        raise PolicyError(f"{endpoint} source closure failed reviewed verification: {error}") from error
    source_tree = reviewed.get("source_tree") or {}
    tree_sha = source_tree.get("sha256")
    _expect(isinstance(tree_sha, str) and SHA256.fullmatch(tree_sha), f"{endpoint} source tree digest is missing")
    git_tree: str | None = None
    try:
        git_tree = _tree_id(root, "HEAD")
    except PolicyError:
        _expect(endpoint == "before", "after endpoint source has no Git tree")
    manifests = {relative: _file_identity(root, relative) for relative in SOURCE_MANIFEST_PATHS}
    lock_specs = _endpoint_lock_spec(endpoint)
    locks: dict[str, dict[str, Any]] = {}
    lock_graph: dict[str, dict[str, Any]] = {}
    for role, spec in lock_specs.items():
        relative = str(spec["path"])
        lock_artifact = _checked_artifact(root / relative, spec, f"{endpoint} {role} source lock")
        locks[relative] = lock_artifact
        lock_graph[role] = {
            **lock_artifact,
            "graph_sha256": _lock_graph_identity_for_path(root / relative, f"{endpoint} {role}"),
        }
    toolchain = {relative: _file_identity(root, relative) for relative in TOOLCHAIN_PATHS}
    _expect(toolchain == TOOLCHAIN_IDENTITIES, f"{endpoint} toolchain inputs differ")
    probe_inputs = {}
    probe_provenance = {}
    for name, item in PROBE_INPUTS.items():
        probe_path = _probe_input_path(root, endpoint, str(item["path"]))
        actual = _artifact_identity(probe_path, f"{endpoint} {name} probe input")
        _expect(
            {"path": str(item["path"]), "bytes": actual["bytes"], "sha256": actual["sha256"]} == item,
            f"{endpoint} reviewed probe input differs: {name}",
        )
        probe_inputs[name] = dict(item)
        probe_provenance[name] = actual
    _expect(probe_inputs == PROBE_INPUTS, f"{endpoint} reviewed probe inputs differ")
    fixture_path = _fixture_manifest_path(root, endpoint)
    fixture_actual = _artifact_identity(fixture_path, f"{endpoint} fixture manifest")
    fixture = dict(FIXTURE_INPUT)
    _expect(
        {"path": FIXTURE_INPUT["path"], "bytes": fixture_actual["bytes"], "sha256": fixture_actual["sha256"]} == FIXTURE_INPUT,
        f"{endpoint} fixture manifest differs",
    )
    _expect(fixture == FIXTURE_INPUT, f"{endpoint} fixture input differs")
    _fixture_policy(root, manifest_root=REPOSITORY, allow_extracted=endpoint == "before")
    try:
        features = list(helper._reviewed_harness_features(kind))
    except Exception as error:
        raise PolicyError(f"{endpoint} feature projection could not be loaded: {error}") from error
    _expect(features and all(isinstance(item, str) and item for item in features), f"{endpoint} feature projection is empty")
    source_archive = dict(reviewed.get("source_archive") or binding["source_archive"])
    _expect(source_archive.get("bytes") == binding["source_archive"]["bytes"], f"{endpoint} source archive size differs from the trusted pin")
    _expect(source_archive.get("sha256") == binding["source_archive"]["sha256"], f"{endpoint} source archive digest differs from the trusted pin")
    source_assembly = reviewed.get("source_assembly") or assembly_artifact
    return {
        "source_root": str(root),
        "endpoint": endpoint,
        "source_role": binding["role"],
        "source_kind": kind,
        "commit": str(reviewed["commit"]),
        "git_tree": git_tree,
        "tree_sha256": tree_sha,
        "source_tree": source_tree,
        "reviewed_policy": reviewed.get("reviewed_policy"),
        "manifests": manifests,
        "locks": locks,
        "lock_graph": lock_graph,
        "role_locks": _endpoint_lock_spec(endpoint),
        "vendor_tree": {"source_role": source_tree.get("source_role"), "verified_closure": reviewed.get("verified_closure")},
        "toolchain": toolchain,
        "probe_inputs": probe_inputs,
        "probe_provenance": probe_provenance,
        "fixture": fixture,
        "fixture_provenance": fixture_actual,
        "archive": source_archive,
        "source_assembly": source_assembly,
        "features": {"package": "litchi-ops", "all_features": True, "selected": features},
    }


def _lock_graph_identity_for_path(path: Path, label: str) -> str:
    try:
        lock = tomllib.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, tomllib.TOMLDecodeError) as error:
        raise PolicyError(f"{label} lock graph is not valid TOML: {error}") from error
    packages = lock.get("package")
    _expect(isinstance(packages, list), f"{label} lock graph has no package list")
    normalized: list[dict[str, Any]] = []
    for package in packages:
        _expect(isinstance(package, dict), f"{label} lock graph has a malformed package")
        _expect(isinstance(package.get("name"), str) and isinstance(package.get("version"), str), f"{label} lock graph package identity is incomplete")
        dependencies = package.get("dependencies", [])
        _expect(isinstance(dependencies, list) and all(isinstance(item, str) for item in dependencies), f"{label} lock graph dependencies are malformed")
        normalized.append({"name": package["name"], "version": package["version"], "source": package.get("source"), "checksum": package.get("checksum"), "dependencies": sorted(dependencies)})
    normalized.sort(key=lambda item: (item["name"], item["version"], item["source"] or ""))
    return _sha256_bytes(_canonical(normalized))


def endpoint_identity(root: Path, endpoint: str) -> dict[str, Any]:
    """Return the independently verified identity for one real endpoint."""

    _expect(endpoint in ENDPOINTS, f"unknown signed-native endpoint: {endpoint}")
    if endpoint == "after":
        # Resolve the trusted binding before touching the caller's source or
        # an output path.  A historical 472b harness is never an after leg.
        _candidate_endpoint_binding()
    source = _endpoint_root(root, endpoint)
    if endpoint == "after":
        identity = candidate_identity(source)
        identity.setdefault("source_kind", "candidate")
        identity.setdefault("source_role", "candidate")
    else:
        identity = _endpoint_identity_from_reviewed(source, endpoint)
    identity = dict(identity)
    identity["endpoint"] = endpoint
    identity["source_kind"] = ENDPOINT_SOURCE_KINDS[endpoint]
    identity["source_role"] = ENDPOINT_ROLE[endpoint]
    identity["api_branch"] = ENDPOINT_API_BRANCH[endpoint]
    return identity


def validate_endpoint_policy(endpoint: str) -> dict[str, Any]:
    """Validate endpoint pins without treating the other leg as its source."""

    _expect(endpoint in ENDPOINTS, f"unknown signed-native endpoint: {endpoint}")
    document = load_policy()
    _validate_policy_document(document, execution=False)
    trusted = _endpoint_source_binding(endpoint)
    _expect(ENDPOINT_ROUTE_STATUS[endpoint] and set(ENDPOINT_ROUTE_STATUS[endpoint]) == set(ROUTES), f"{endpoint} route disposition matrix is incomplete")
    _fixture_policy()
    return {"endpoint": endpoint, "source": trusted, "api_branch": ENDPOINT_API_BRANCH[endpoint], "routes": ENDPOINT_ROUTE_STATUS[endpoint]}


def prepare(source_root: Path, output_root: Path, *, endpoint: str = "after") -> dict[str, Any]:
    """Prepare one endpoint after its checked-in source gate passes."""

    validate_endpoint_policy(endpoint)
    source_root = _endpoint_root(source_root, endpoint)
    archive_path = Path(str(_endpoint_source_binding(endpoint)["source_archive"]["path"]))
    _assert_disjoint_paths(
        {
            "candidate source": source_root,
            "candidate archive": archive_path,
            "preparation output": output_root,
        }
    )
    identity = endpoint_identity(source_root, endpoint)
    # Import lazily so source validation remains independent of the consumer
    # and no output is created until the endpoint fence has passed.
    from native_qualification import signed_native_v2_consumer as consumer

    return consumer.prepare(source_root, output_root, endpoint=endpoint, identity=identity)


def prepare_pair(
    before_source: Path,
    after_source: Path,
    output_root: Path,
) -> dict[str, Any]:
    """Prepare both immutable endpoint projects under disjoint output roots."""

    validate_endpoint_policy("before")
    validate_endpoint_policy("after")
    before = _endpoint_root(before_source, "before")
    after = _endpoint_root(after_source, "after")
    after_binding = _endpoint_source_binding("after")
    output = output_root.absolute()
    _assert_disjoint_paths(
        {
            "before source": before,
            "after source": after,
            "before archive": Path(str(ENDPOINT_SOURCE_BINDINGS["before"]["source_archive"]["path"])),
            "after archive": Path(str(after_binding["source_archive"]["path"])),
            "pair output": output,
        }
    )
    _expect(not output.exists() and not output.is_symlink(), f"refusing to overwrite pair output: {output}")
    from native_qualification import signed_native_v2_consumer as consumer

    output.mkdir(parents=True)
    before_result = consumer.prepare(before, output / "before", endpoint="before")
    after_result = consumer.prepare(after, output / "after", endpoint="after")
    pair = {
        "schema": "litchi.signed-native.preparation-pair.v2",
        "status": "prepared",
        "endpoints": {
            "before": before_result["receipt"],
            "after": after_result["receipt"],
        },
    }
    pair_path = output / "pair-receipt.json"
    _expect(not pair_path.exists(), f"refusing to overwrite pair receipt: {pair_path}")
    pair_path.write_bytes(_canonical(pair))
    return {"status": "prepared", "output": str(output), "receipt": pair, "pair_receipt": str(pair_path)}


def execute(
    prepared_root: Path | None = None,
    capabilities: Path | None = None,
    *,
    output: Path | None = None,
    endpoint: str | None = None,
) -> dict[str, Any]:
    """Consume a built v2 receipt and typed protocol capabilities.

    A no-input invocation still checks the historical pending policy and fails
    closed; an explicit prepared endpoint delegates only after its own source
    receipt has been revalidated.
    """

    if prepared_root is None:
        # Keep the no-input invocation fail-closed rather than guessing an
        # endpoint or consuming a historical receipt.
        validate_policy(execution=True)
    if endpoint is not None:
        validate_endpoint_policy(endpoint)
    from native_qualification import signed_native_v2_consumer as consumer

    return consumer.execute(prepared_root, capabilities, output=output, endpoint=endpoint)


def reject_v1_receipt(receipt: Mapping[str, Any]) -> None:
    """Prevent a historical ul24 receipt from being relabelled as v2."""

    _expect(isinstance(receipt, Mapping), "receipt is not an object")
    _expect(receipt.get("schema") != "litchi_signed_native_qualification.v1", "historical ul24 receipt cannot be relabelled as v2")
    _expect(receipt.get("schema") != "litchi_signed_native_probe.v1", "historical ul24 probe receipt cannot be relabelled as v2")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--endpoint", choices=ENDPOINTS, default="after")
    parser.add_argument("--source-root", type=Path)
    parser.add_argument("--before-source", type=Path)
    parser.add_argument("--after-source", type=Path)
    parser.add_argument("--output-root", type=Path, required=True)
    arguments = parser.parse_args(argv)
    try:
        if arguments.before_source is not None or arguments.after_source is not None:
            _expect(arguments.before_source is not None and arguments.after_source is not None, "pair preparation requires both endpoint source roots")
            prepare_pair(arguments.before_source, arguments.after_source, arguments.output_root)
        else:
            _expect(arguments.source_root is not None, "single endpoint preparation requires --source-root")
            prepare(arguments.source_root, arguments.output_root, endpoint=arguments.endpoint)
    except (OSError, ValueError) as error:
        parser.error(str(error))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
