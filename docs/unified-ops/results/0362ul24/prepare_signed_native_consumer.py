#!/usr/bin/env python3
"""Prepare and build the signed-native consumer outside a source checkout.

The signed qualification binary is an external consumer.  This driver keeps
the two endpoint source trees immutable, makes the f2 API differences explicit
in a tiny compile-time adapter, and writes endpoint manifests only after a
real release binary exists.  ``plan`` and ``prepare`` are source/provenance
operations; ``build`` is deliberately a separate, owner-granted operation.

No generated project, target directory, binary, archive, lockfile, or raw log
may be placed in the repository.  The caller supplies an external output
directory for all of them.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tarfile
from typing import Any


HERE = Path(__file__).resolve().parent
PROBE = HERE / "signed_native_probe.rs"
FIXTURES = HERE / "fixtures.json"
F2_COMMIT = "f2bbbc8d07ab5da26c180f1ff74813d39f632d54"
FINAL_COMMIT = "472b6705475b073e7fe08f6e11ee17fb1e29f838"
SOURCE_KINDS = {"f2": F2_COMMIT, "472b": FINAL_COMMIT}
FULL_SHA = re.compile(r"^[0-9a-f]{40}$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")
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
TARGET_CAP_BYTES = 20_000_000_000
CHILD_TIMEOUT_SECONDS = 300.0
REVIEWED_TOOLCHAIN = {
    "bytes": 86,
    "sha256": "e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310",
}
REVIEWED_CARGO_CONFIG = {
    "bytes": 470,
    "sha256": "50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d",
}
REVIEWED_SOURCE_LOCKS = {
    "f2": {
        "bytes": 139_441,
        "sha256": "1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c",
    },
    "472b": {
        "bytes": 139_606,
        "sha256": "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1",
    },
}
REVIEWED_SOURCE_ARCHIVES = {
    "f2": {
        "bytes": 6_614_528_000,
        "sha256": "de625729ec8e6f8055e619443ad31bf1a91719cd31465975f951d54dbc1c2280",
    },
    "472b": {
        "bytes": 6_634_147_840,
        "sha256": "44c11e933974c8fa3d6a67b46739bedeef866e60fbd4360ffe7f3bb8a4dc2274",
    },
}
REVIEWED_POLICY_RELATIVE = Path("docs/unified-ops/results/0362yh/original-to-final-policy.json")
REVIEWED_PROJECTION_RELATIVE = {
    "f2": Path("docs/unified-ops/results/0362yh/images/before/projection.json"),
    "472b": Path("docs/unified-ops/results/0362yh/images/after/projection.json"),
}


class PreparationError(RuntimeError):
    pass


def fail(message: str) -> PreparationError:
    return PreparationError(message)


def _repository() -> Path:
    return HERE.parents[3]


def _reviewed_module(relative: str, name: str) -> Any:
    path = _repository() / relative
    if not path.is_file() or path.is_symlink():
        raise fail(f"reviewed helper is missing: {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise fail(f"cannot load reviewed helper: {path}")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        raise fail(f"reviewed helper failed to load: {path}: {error}") from error
    return module


def _w8_baseline() -> Any:
    return _reviewed_module("tools/w8_original_baseline.py", "w8_original_baseline_signed_consumer")


def _w8_policy() -> Any:
    return _reviewed_module("tools/w8_projection_policy.py", "w8_projection_policy_signed_consumer")


def _selected_probe(path: Path | None) -> Path:
    selected = (path or PROBE).absolute()
    if selected.is_symlink() or not selected.is_file():
        raise fail(f"signed probe is not a regular file: {selected}")
    safe_external(selected, "signed probe")
    return selected.resolve()


def canonical(value: object) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        + "\n"
    ).encode("utf-8")


def digest(path: Path) -> str:
    h = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                h.update(block)
    except OSError as error:
        raise fail(f"cannot read {path}: {error}") from error
    return h.hexdigest()


def artifact(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if not path.is_file() or path.is_symlink():
        raise fail(f"artifact is not a regular non-symlink file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": digest(path)}


def safe_external(path: Path, label: str) -> Path:
    lexical = path.absolute()
    cursor = lexical
    while cursor != cursor.parent:
        if cursor.is_symlink():
            raise fail(f"{label} may not traverse a symlink: {cursor}")
        cursor = cursor.parent
    path = lexical.resolve()
    if "3rdparty" in path.parts:
        raise fail(f"{label} may not use 3rdparty: {path}")
    # ``HERE`` is ``<checkout>/docs/unified-ops/results/0362ul24``.
    # Worktrees use a `.git` file, so this script must not infer checkout
    # validity from that path; the repository root is the fourth parent.
    repository = HERE.parents[3]
    forbidden = (
        repository / "docs" / "performance" / "results",
        repository / "docs" / "report",
        repository / "binaries",
    )
    for root in forbidden:
        try:
            path.relative_to(root.resolve())
        except ValueError:
            continue
        raise fail(f"{label} must be outside {root}")
    return path


def require_commit(value: str, label: str) -> str:
    if FULL_SHA.fullmatch(value) is None:
        raise fail(f"{label} is not a full commit id")
    return value


def run(
    command: list[str],
    cwd: Path,
    *,
    capture: bool = True,
    environment: dict[str, str] | None = None,
) -> str:
    try:
        result = subprocess.run(
            command,
            cwd=cwd,
            check=True,
            text=True,
            env=environment,
            stdout=subprocess.PIPE if capture else None,
            stderr=subprocess.STDOUT if capture else None,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        raise fail(f"command failed in {cwd}: {' '.join(command)}: {error}") from error
    return result.stdout or ""


def _verified_receipt_artifact(value: object, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise fail(f"{label} artifact is missing")
    path = value.get("path")
    expected_bytes = value.get("bytes")
    expected_sha = value.get("sha256")
    if not isinstance(path, str) or not isinstance(expected_bytes, int) or not isinstance(expected_sha, str):
        raise fail(f"{label} artifact is malformed")
    if SHA256.fullmatch(expected_sha) is None:
        raise fail(f"{label} artifact hash is malformed")
    actual = artifact(Path(path))
    if actual != {"path": str(Path(path).resolve()), "bytes": expected_bytes, "sha256": expected_sha}:
        raise fail(f"{label} artifact changed since its receipt")
    return actual


def _checked_fixed_artifact(path: Path, expected: dict[str, Any], label: str) -> dict[str, Any]:
    actual = artifact(path)
    if actual["bytes"] != expected["bytes"] or actual["sha256"] != expected["sha256"]:
        raise fail(f"{label} differs from the reviewed fixed identity")
    return actual


def _git_tree_records(root: Path, commit: str) -> dict[str, tuple[str, str]]:
    try:
        raw = subprocess.run(
            ["git", "ls-tree", "-r", "--full-tree", "-z", commit],
            cwd=root,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        ).stdout
    except (OSError, subprocess.CalledProcessError) as error:
        raise fail(f"cannot read reviewed source tree {commit}: {error}") from error
    records: dict[str, tuple[str, str]] = {}
    for entry in raw.split(b"\0"):
        if not entry:
            continue
        try:
            header, path_bytes = entry.split(b"\t", 1)
            mode, kind, object_id = header.decode("ascii").split(" ", 2)
            path = path_bytes.decode("utf-8")
        except (ValueError, UnicodeDecodeError) as error:
            raise fail("reviewed source tree contains an unreadable path record") from error
        if kind != "blob" or path in records:
            raise fail(f"reviewed source tree contains an unsupported or duplicate path: {path}")
        records[path] = (mode, object_id)
    if not records:
        raise fail(f"reviewed source tree is empty: {commit}")
    return records


def _reviewed_source_policy(kind: str) -> dict[str, Any]:
    policy = _w8_policy()
    try:
        value = policy.load_reviewed_policy()
        side = "before" if kind == "f2" else "after"
        projection_path = _repository() / REVIEWED_PROJECTION_RELATIVE[kind]
        projection = json.loads(projection_path.read_bytes())
        if not isinstance(projection, dict):
            raise ValueError("projection is not an object")
        # The policy loader verifies the canonical fixed pre/post images and
        # the exact registry/local-edge delta.  Keep that check in every plan
        # so a caller cannot replace the reviewed graph with a convenient
        # regenerated lock.
        before_path = _repository() / REVIEWED_PROJECTION_RELATIVE["f2"]
        after_path = _repository() / REVIEWED_PROJECTION_RELATIVE["472b"]
        before = json.loads(before_path.read_bytes())
        after = json.loads(after_path.read_bytes())
        policy.validate_projection_pair(before, after)
    except Exception as error:
        raise fail(f"reviewed W8 source policy is not valid: {error}") from error
    return {
        "policy": artifact(_repository() / REVIEWED_POLICY_RELATIVE),
        "projection": artifact(projection_path),
        "projection_side": side,
        "projection_value": projection,
        "policy_version": value["policy_version"],
    }


def _reviewed_harness_features(kind: str) -> list[str]:
    """Return the endpoint's exact reviewed all-features selection.

    The source legs intentionally have different feature sets (the historical
    f2 leg predates ``sign``).  Deriving this list from the fixed projection
    keeps the generated consumer graph aligned with the reviewed lock images;
    a hand-maintained subset would make the dependency receipt unverifiable.
    """
    projection = _reviewed_source_policy(kind)["projection_value"]
    features = projection.get("features")
    if (
        not isinstance(features, dict)
        or features.get("package") != "litchi-ops"
        or features.get("all_features") is not True
        or not isinstance(features.get("selected"), list)
        or not features["selected"]
        or any(not isinstance(item, str) or not item for item in features["selected"])
    ):
        raise fail("reviewed harness feature selection is incomplete")
    return list(features["selected"])


def _package_key(package: dict[str, Any]) -> tuple[str, str, str | None, str | None]:
    return (
        str(package.get("name")),
        str(package.get("version")),
        package.get("source") if isinstance(package.get("source"), str) else None,
        package.get("checksum") if isinstance(package.get("checksum"), str) else None,
    )


def _fixed_dependency_keys(
    dependency: str,
    fixed_by_name: dict[str, list[tuple[str, str, str | None, str | None]]],
) -> set[tuple[str, str, str | None, str | None]]:
    pieces = dependency.rsplit(" ", 1)
    name = pieces[0]
    candidates = fixed_by_name.get(name, [])
    if len(pieces) == 2 and FULL_SHA.fullmatch(pieces[1]) is None:
        version = pieces[1]
        candidates = [candidate for candidate in candidates if candidate[1] == version]
    if not candidates:
        raise fail(f"reviewed projection dependency has no package identity: {dependency}")
    if len(candidates) > 1:
        raise fail(f"reviewed projection dependency is ambiguous: {dependency}")
    return set(candidates)


def _validate_harness_metadata(metadata: dict[str, Any], kind: str) -> dict[str, Any]:
    reviewed = _reviewed_source_policy(kind)
    fixed_packages = reviewed["projection_value"].get("lock", {}).get("packages")
    packages = metadata.get("packages")
    resolve = metadata.get("resolve")
    if not isinstance(fixed_packages, list) or not isinstance(packages, list) or not isinstance(resolve, dict):
        raise fail("cargo metadata lacks the reviewed package/resolve projection")
    fixed_by_key = {_package_key(package): package for package in fixed_packages if isinstance(package, dict)}
    fixed_by_name: dict[str, list[tuple[str, str, str | None, str | None]]] = {}
    for key in fixed_by_key:
        fixed_by_name.setdefault(key[0], []).append(key)
    metadata_by_id = {
        package.get("id"): package
        for package in packages
        if isinstance(package, dict) and isinstance(package.get("id"), str)
    }
    metadata_keys: set[tuple[str, str, str | None, str | None]] = set()
    root_names = {"signed-native-consumer"}
    for package in packages:
        if not isinstance(package, dict):
            raise fail("cargo metadata contains a malformed package")
        key = _package_key(package)
        if key[0] in root_names and key[2] is None:
            continue
        if key not in fixed_by_key:
            raise fail(f"cargo metadata contains an unreviewed package identity: {key}")
        metadata_keys.add(key)
    if not metadata_keys:
        raise fail("cargo metadata has no reviewed source packages")
    nodes = resolve.get("nodes")
    if not isinstance(nodes, list):
        raise fail("cargo metadata resolve graph is missing")
    metadata_node_by_id = {
        node.get("id"): node
        for node in nodes
        if isinstance(node, dict) and isinstance(node.get("id"), str)
    }
    for package_id, package in metadata_by_id.items():
        key = _package_key(package)
        if key[0] in root_names and key[2] is None:
            continue
        fixed = fixed_by_key.get(key)
        if fixed is None:
            raise fail(f"cargo metadata package is not in the reviewed projection: {key}")
        node = metadata_node_by_id.get(package_id)
        if not isinstance(node, dict):
            raise fail(f"cargo metadata resolve node is missing for {package_id}")
        resolved = node.get("dependencies")
        if not isinstance(resolved, list):
            raise fail(f"cargo metadata dependency list is missing for {key[0]}")
        actual_dependencies: set[tuple[str, str, str | None, str | None]] = set()
        for dependency in resolved:
            dependency_id = dependency.get("pkg") if isinstance(dependency, dict) else None
            target = metadata_by_id.get(dependency_id)
            if target is None:
                raise fail(f"cargo metadata dependency target is missing for {key[0]}")
            target_key = _package_key(target)
            if target_key[0] in root_names and target_key[2] is None:
                continue
            actual_dependencies.add(target_key)
        expected_dependencies: set[tuple[str, str, str | None, str | None]] = set()
        for dependency in fixed.get("dependencies", []):
            if not isinstance(dependency, str):
                raise fail(f"reviewed projection dependency is malformed for {key[0]}")
            expected_dependencies.update(_fixed_dependency_keys(dependency, fixed_by_name))
        if actual_dependencies != expected_dependencies:
            raise fail(f"cargo metadata dependency edges differ for {key[0]} {key[1]}")
    return {
        "schema": "litchi_signed_native_harness_projection.v1",
        "source_side": "before" if kind == "f2" else "after",
        "reviewed_policy": reviewed["policy"],
        "reviewed_projection": reviewed["projection"],
        "reviewed_package_count": len(fixed_packages),
        "harness_package_count": len(packages),
        "reviewed_registry_and_dependency_policy": True,
    }


def _archive_name(name: str) -> str:
    if name.startswith("/"):
        raise fail(f"source archive contains an absolute member: {name}")
    pieces = Path(name.replace("/", "/")).parts
    if ".." in pieces:
        raise fail(f"source archive contains a parent traversal member: {name}")
    return "/".join(piece for piece in pieces if piece not in ("", "."))


def _archive_relative(name: str, prefix: str) -> str | None:
    name = _archive_name(name)
    if not prefix:
        return name
    if name == prefix:
        return None
    marker = prefix + "/"
    if not name.startswith(marker):
        raise fail(f"source archive contains a member outside its prefix: {name}")
    return name[len(marker) :]


def _verify_archive_one_pass(archive: Path, tree: dict[str, tuple[str, str]]) -> dict[str, Any]:
    reviewed = _w8_baseline()
    expected_directories = {""}
    for path in tree:
        parent = Path(path).parent
        while str(parent) not in ("", "."):
            expected_directories.add(str(parent).replace("\\", "/"))
            parent = parent.parent
    prefix: str | None = None
    seen_members: set[str] = set()
    seen_files: set[str] = set()
    pending_directories: list[str] = []
    try:
        with tarfile.open(archive, "r|*") as stream:
            for member in stream:
                name = _archive_name(member.name)
                if name in seen_members:
                    raise fail(f"source archive contains a duplicate member: {member.name}")
                seen_members.add(name)
                if prefix is None:
                    if member.isdir():
                        pending_directories.append(name)
                        continue
                    if not member.isfile():
                        raise fail(f"source archive begins with an unsupported member: {name}")
                    if name in tree:
                        prefix = ""
                    else:
                        candidate, separator, relative_candidate = name.partition("/")
                        if not separator or relative_candidate not in tree:
                            raise fail(f"source archive has no reviewed root prefix: {name}")
                        prefix = candidate
                    for directory in pending_directories:
                        relative_directory = _archive_relative(directory, prefix)
                        if relative_directory is not None and relative_directory not in expected_directories:
                            raise fail(f"source archive contains an unexpected directory: {relative_directory}")
                    pending_directories.clear()
                if prefix is None:
                    raise fail("source archive does not expose a root prefix before its first file")
                relative = _archive_relative(name, prefix)
                if relative is None:
                    if not member.isdir():
                        raise fail("source archive prefix member is not a directory")
                    continue
                if member.isdir():
                    if relative not in expected_directories:
                        raise fail(f"source archive contains an unexpected directory: {relative}")
                    continue
                if not member.isfile() or relative not in tree:
                    raise fail(f"source archive contains an unexpected non-regular member: {relative}")
                expected_mode, expected_blob = tree[relative]
                if (member.mode & 0o111) != (int(expected_mode, 8) & 0o111):
                    raise fail(f"source archive executable mode differs at {relative}")
                handle = stream.extractfile(member)
                if handle is None:
                    raise fail(f"source archive member is unreadable: {relative}")
                digest_value = hashlib.sha1(f"blob {member.size}\0".encode())
                for block in iter(lambda: handle.read(1024 * 1024), b""):
                    digest_value.update(block)
                if digest_value.hexdigest() != expected_blob:
                    raise fail(f"source archive blob differs at {relative}")
                seen_files.add(relative)
    except (OSError, tarfile.TarError) as error:
        raise fail(f"cannot inspect source archive {archive}: {error}") from error
    if prefix is None or seen_files != set(tree):
        raise fail("source archive file set differs from the reviewed source tree")
    return {
        "artifact": artifact(archive),
        "prefix": prefix,
        "file_count": len(tree),
        "tree_sha256": reviewed.tree_digest(tree),
        "production_delta": {"files_changed": 0, "api_overlay": False},
    }


def _verify_f2_extracted(source: Path, source_archive: Path) -> dict[str, Any]:
    reviewed = _w8_baseline()
    repository = _repository()
    tree = reviewed.git_tree(repository)
    if reviewed.tree_digest(tree) != "23e332254abb024cbf9b038b8219b861e363fe474423bccc6122c3183b4d1437":
        raise fail("reviewed f2 Git tree no longer matches the source-assembly closure")
    try:
        archive_info = _verify_archive_one_pass(source_archive, tree)
        probe = source / reviewed.PROBE_RELATIVE
        adapter = source / reviewed.ADAPTER_RELATIVE
        lock = source / reviewed.LOCK_RELATIVE
        workspace_info = reviewed.verify_workspace(
            source,
            tree,
            probe,
            adapter,
            lock,
            adapter.read_bytes(),
        )
    except Exception as error:
        raise fail(f"extracted f2 closure failed reviewed full-fence verification: {error}") from error
    if workspace_info.get("production_files") != 459676 or workspace_info.get("production_tree_sha256") != reviewed.tree_digest(tree):
        raise fail("extracted f2 production closure identity differs")
    if archive_info.get("file_count") != 459676 or archive_info.get("tree_sha256") != reviewed.tree_digest(tree):
        raise fail("f2 source archive closure identity differs")
    return {
        "production_tree": {
            "files": workspace_info["production_files"],
            "sha256": workspace_info["production_tree_sha256"],
            "source_role": "pristine_f2_before_leg",
        },
        "archive": archive_info,
        "workspace": workspace_info,
    }


def _verify_git_archive(source: Path, archive: Path, commit: str) -> dict[str, Any]:
    tree = _git_tree_records(source, commit)
    reviewed = _w8_baseline()
    try:
        archive_info = _verify_archive_one_pass(archive, tree)
    except Exception as error:
        raise fail(f"final source archive failed full-fence verification: {error}") from error
    return {
        "tree": {
            "files": len(tree),
            "sha256": reviewed.tree_digest(tree),
            "source_role": "pinned_git_checkout",
        },
        "archive": archive_info,
    }


def source_identity(
    source: Path,
    kind: str,
    *,
    source_archive: Path | None = None,
    source_assembly: Path | None = None,
    verify_closure: bool = True,
) -> dict[str, Any]:
    source_lexical = source.absolute()
    if source_lexical.is_symlink():
        raise fail(f"source checkout may not be a symlink: {source}")
    source = source.resolve()
    if source_archive is not None:
        source_archive = safe_external(source_archive, "source archive")
    if source_assembly is not None:
        source_assembly = safe_external(source_assembly, "source assembly")
    if kind not in SOURCE_KINDS:
        raise fail(f"unsupported source kind: {kind}")
    expected = SOURCE_KINDS[kind]
    if not source.is_dir() or not (source / "Cargo.toml").is_file():
        raise fail(f"source is not a usable checkout: {source}")
    git_head = subprocess.run(
        ["git", "rev-parse", "HEAD"], cwd=source, text=True,
        stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
    )
    git_checkout = git_head.returncode == 0
    tree_hash: str | None = None
    assembly_artifact: dict[str, Any] | None = None
    source_archive_artifact: dict[str, Any] | None = None
    source_tree: dict[str, Any] | None = None
    verified_closure: dict[str, Any] | None = None
    if git_head.returncode == 0:
        actual = git_head.stdout.strip()
        if actual != expected:
            raise fail(f"{kind} source HEAD {actual} differs from required {expected}")
        status = run(["git", "status", "--porcelain=v1"], source)
        if status:
            raise fail(f"{kind} source checkout is dirty:\n{status}")
        tree = subprocess.run(
            ["git", "ls-tree", "-r", "-z", expected],
            cwd=source,
            check=True,
            stdout=subprocess.PIPE,
        ).stdout
        tree_hash = hashlib.sha256(tree).hexdigest()
        source_tree = {
            "files": tree.count(b"\0"),
            "sha256": tree_hash,
            "source_role": "pinned_git_checkout",
        }
        if source_archive is not None:
            source_archive_artifact = artifact(source_archive)
    else:
        if kind != "f2" or source_archive is None or source_assembly is None:
            raise fail("non-git source requires the pinned f2 archive and source-assembly receipt")
        source_archive_artifact = artifact(source_archive)
        assembly_artifact = artifact(source_assembly)
        try:
            assembly = json.loads(source_assembly.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError) as error:
            raise fail(f"invalid source-assembly receipt: {error}") from error
        if assembly.get("schema") != "litchi_w8_source_assembly.v1":
            raise fail("f2 source-assembly receipt schema mismatch")
        if assembly.get("source_commit") != expected or assembly.get("resolved_commit") != expected:
            raise fail("f2 source-assembly receipt commit mismatch")
        recorded_archive = _verified_receipt_artifact(assembly.get("source_archive"), "f2 source archive")
        if recorded_archive != source_archive_artifact:
            raise fail("f2 source archive differs from source-assembly receipt")
        overlay = assembly.get("overlay")
        if not isinstance(overlay, dict) or overlay.get("kind") != "compatibility":
            raise fail("f2 source-assembly must declare compatibility-only overlay")
        if overlay.get("base_commit") != expected:
            raise fail("f2 compatibility overlay base commit differs")
        patch_value = overlay.get("patch")
        if isinstance(patch_value, dict) and isinstance(patch_value.get("path"), str):
            safe_external(Path(patch_value["path"]), "f2 adapter patch")
        patch = _verified_receipt_artifact(overlay.get("patch"), "f2 adapter patch")
        source_tree = {
            "files": 459676,
            "sha256": "23e332254abb024cbf9b038b8219b861e363fe474423bccc6122c3183b4d1437",
            "source_role": "pristine_f2_before_leg",
            "adapter_patch": patch,
        }
        tree_hash = source_tree["sha256"]
        if verify_closure:
            verified_closure = _verify_f2_extracted(source, source_archive)
    lock = source / "Cargo.lock"
    toolchain = source / "rust-toolchain.toml"
    cargo_config = source / ".cargo" / "config.toml"
    for path, label in ((lock, "Cargo.lock"), (toolchain, "rust-toolchain.toml"), (cargo_config, ".cargo/config.toml")):
        if not path.is_file():
            raise fail(f"{kind} source lacks {label}")
    lock_expected = REVIEWED_SOURCE_LOCKS[kind]
    toolchain_expected = REVIEWED_TOOLCHAIN
    config_expected = REVIEWED_CARGO_CONFIG
    lock_artifact = _checked_fixed_artifact(lock, lock_expected, f"{kind} source lock")
    toolchain_artifact = _checked_fixed_artifact(toolchain, toolchain_expected, f"{kind} rust toolchain")
    cargo_config_artifact = _checked_fixed_artifact(cargo_config, config_expected, f"{kind} cargo config")
    policy_identity = _reviewed_source_policy(kind)
    if git_checkout and verify_closure and source_archive is not None:
        verified_closure = _verify_git_archive(source, source_archive, expected)
        source_tree = verified_closure["tree"]
        tree_hash = source_tree["sha256"]
        source_archive_artifact = verified_closure["archive"]["artifact"]
    if source_archive_artifact is not None:
        expected_archive = REVIEWED_SOURCE_ARCHIVES[kind]
        _checked_fixed_artifact(
            Path(source_archive_artifact["path"]), expected_archive, f"{kind} source archive"
        )
    return {
        "kind": kind,
        "commit": expected,
        "source_ref": (
            f"git:{expected}"
            if git_checkout
            else f"archive:{source_archive_artifact['sha256']}"
        ),
        "source_root": str(source),
        "tree_sha256": tree_hash,
        "lock": lock_artifact,
        "toolchain": toolchain_artifact,
        "cargo_config": cargo_config_artifact,
        "source_archive": source_archive_artifact,
        "source_assembly": assembly_artifact,
        "source_tree": source_tree,
        "verified_closure": verified_closure,
        "reviewed_policy": policy_identity,
    }


def fixture_manifest() -> dict[str, Any]:
    if not FIXTURES.is_file():
        raise fail(f"missing signed fixture manifest: {FIXTURES}")
    try:
        value = json.loads(FIXTURES.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise fail(f"invalid signed fixture manifest: {error}") from error
    raw = FIXTURES.read_bytes()
    if not isinstance(value, dict) or value.get("schema") != "litchi_signed_fixture_manifest.v1":
        raise fail("signed fixture manifest schema mismatch")
    if canonical(value) != raw:
        raise fail("signed fixture manifest is not canonical JSON")
    rows = value.get("fixtures")
    if not isinstance(rows, list) or not rows:
        raise fail("signed fixture manifest has no fixtures")
    formats: set[str] = set()
    names: set[str] = set()
    for index, row in enumerate(rows):
        if not isinstance(row, dict):
            raise fail(f"fixture row {index} is not an object")
        name = row.get("name")
        relative = row.get("relative_path")
        if not isinstance(name, str) or not name or name in names:
            raise fail(f"fixture row {index} has a duplicate/empty name")
        if not isinstance(relative, str) or not relative:
            raise fail(f"fixture row {index} has no relative path")
        path = Path(relative)
        if path.is_absolute() or ".." in path.parts or "\\" in relative or "3rdparty" in path.parts:
            raise fail(f"fixture row {index} has an unsafe relative path")
        if row.get("origin") not in {"Opc", "Cfb"}:
            raise fail(f"fixture row {index} has an invalid origin")
        if row.get("fixture_mode") not in {"producer_signed", "synthetic_presence"}:
            raise fail(f"fixture row {index} has an invalid fixture mode")
        if row.get("signature_presence") != "detected" or row.get("cryptographic_verification") != "not_performed":
            raise fail(f"fixture row {index} has an invalid signature claim")
        if not isinstance(row.get("input_sha256"), str) or SHA256.fullmatch(row["input_sha256"]) is None:
            raise fail(f"fixture row {index} input hash is invalid")
        if not isinstance(row.get("target"), str) or not row["target"]:
            raise fail(f"fixture row {index} target is missing")
        if not isinstance(row.get("property"), str) or not row["property"]:
            raise fail(f"fixture row {index} property is missing")
        mode = row.get("fixture_mode")
        signed_hash = row.get("signed_sha256")
        if mode == "producer_signed":
            if not isinstance(signed_hash, str) or SHA256.fullmatch(signed_hash) is None:
                raise fail(f"fixture row {index} producer signature hash is invalid")
            if "signed_hash_mode" in row:
                raise fail(f"fixture row {index} producer fixture has a synthetic hash mode")
        elif mode == "synthetic_presence":
            if row.get("signed_hash_mode") != "probe_generated" or signed_hash is not None:
                raise fail(f"fixture row {index} synthetic signature hash policy is invalid")
        else:
            raise fail(f"fixture row {index} has an invalid fixture mode")
        formats.add(str(row.get("format")))
        names.add(name)
    if formats != {"DOCX", "PPTX", "XLSX", "XLSB", "DOC", "PPT", "XLS"}:
        raise fail(f"signed fixture formats are incomplete: {sorted(formats)}")
    return value


def adapter_source(kind: str) -> str:
    """Consumer-only API bridge, kept outside both production trees."""
    if kind == "f2":
        return """//! Consumer-only compatibility bridge for original f2 APIs.
#![allow(dead_code)]
pub fn exercise_api() {
    fn commit_fields<T>(commit: &litchi_ops::transaction::Commit<T>) {
        let _ = &commit.file;
        let _ = &commit.patch;
        let _ = &commit.diagnostics;
    }
    fn edit_entry(file: &litchi_ops::File) {
        let _ = file.edit_with_limits(
            &litchi_ops::transaction::SemanticLimits::default(),
            |_edit| Ok::<(), litchi_ops::Error>(()),
        );
    }
    fn xlsb_entry(package: &litchi_xlsb::Package) {
        let _ = package.edit_cell_values(0usize);
    }
    let _ = (commit_fields::<()>, edit_entry, xlsb_entry);
}
"""
    return """//! Consumer-only compatibility bridge for the final API.
#![allow(dead_code)]
pub fn exercise_api() {
    fn commit_methods<T>(commit: &litchi_ops::transaction::Commit<T>) {
        let _ = commit.file();
        let _ = commit.patch();
        let _ = commit.diagnostics();
    }
    fn edit_entry(file: &litchi_ops::File) {
        let options = litchi_ops::transaction::EditOptions::new()
            .semantic_limits(litchi_ops::transaction::SemanticLimits::default());
        let _ = file.edit_with_options(options, |_edit| Ok::<(), litchi_ops::Error>(()));
    }
    fn xlsb_entry(package: &litchi_xlsb::Package) {
        let _ = package.edit_cell_values(0usize).resolve();
    }
    let _ = (commit_methods::<()>, edit_entry, xlsb_entry);
}
"""


def probe_contract(source: str) -> dict[str, Any]:
    required = {
        "legacy_commit_fields": "#[cfg(feature = \"legacy_api\")]",
        "legacy_edit_limits": "edit_with_limits",
        "current_commit_accessors": "#[cfg(not(feature = \"legacy_api\"))]",
        "current_edit_options": "edit_with_options",
        "legacy_xlsb_option": "edit_cell_values(0)",
        "current_xlsb_resolve": ".resolve()",
        "xlsb_runtime_bridge": "fn xlsb_native_noop",
        "external_fixture_manifest": "LITCHI_SIGNED_FIXTURE_MANIFEST",
        "typed_publication_guard": "publication_attempts",
    }
    missing = [name for name, marker in required.items() if marker not in source]
    if missing:
        raise fail(f"signed probe is not the reviewed legacy/current source: missing {missing}")
    return {
        "schema": "litchi_signed_native_probe_contract.v1",
        "features": {"f2": ["legacy_api"], "472b": []},
        "markers": required,
        "source_sha256": hashlib.sha256(source.encode("utf-8")).hexdigest(),
    }


def project_toml(source: Path, kind: str) -> str:
    features = ", ".join(json.dumps(feature) for feature in _reviewed_harness_features(kind))
    deps = [
        ("litchi-ops", "crates/litchi-ops"),
        ("litchi-core", "crates/litchi-core"),
        ("litchi-opc", "crates/litchi-opc"),
        ("litchi-cfb", "crates/litchi-cfb"),
        ("litchi-docx", "crates/litchi-docx"),
        ("litchi-pptx", "crates/litchi-pptx"),
        ("litchi-xlsx", "crates/litchi-xlsx"),
        ("litchi-xlsb", "crates/litchi-xlsb"),
        ("litchi-doc", "crates/litchi-doc"),
        ("litchi-ppt", "crates/litchi-ppt"),
        ("litchi-xls", "crates/litchi-xls"),
        ("serde_json", None),
        ("sha2", None),
    ]
    lines = [
        "[package]",
        'name = "signed-native-consumer"',
        'version = "0.0.0"',
        'edition = "2024"',
        "",
        "[features]",
        'default = ["legacy_api"]' if kind == "f2" else 'default = []',
        'legacy_api = []',
        "",
        "[dependencies]",
    ]
    for name, relative in deps:
        if relative is None:
            version = "1" if name == "serde_json" else "0.10"
            lines.append(f'{name} = "{version}"')
        else:
            path = (source / relative).as_posix()
            features_clause = f", features = [{features}]" if name == "litchi-ops" else ""
            lines.append(f'{name} = {{ path = {json.dumps(path)}{features_clause} }}')
    return "\n".join(lines) + "\n"


def stage_fixtures(source: Path, fixture_root: Path, fixtures: dict[str, Any]) -> Path:
    fixture_root = safe_external(fixture_root, "fixture root")
    destination_root = fixture_root / "signed-native-fixtures"
    if destination_root.exists():
        raise fail(f"refusing to overwrite fixture root: {destination_root}")
    for row in fixtures["fixtures"]:
        relative = Path(row["relative_path"])
        source_path = source / relative
        try:
            source_path.resolve().relative_to(source.resolve())
        except ValueError:
            raise fail(f"fixture escapes the immutable source tree: {relative}")
        if not source_path.is_file() or source_path.is_symlink():
            raise fail(f"tracked fixture is missing: {source_path}")
        if digest(source_path) != row["input_sha256"]:
            raise fail(f"fixture input hash changed: {relative}")
        destination = destination_root / relative
        try:
            destination.resolve().relative_to(destination_root.resolve())
        except ValueError:
            raise fail(f"fixture destination escapes the staged root: {relative}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, destination)
    manifest_path = destination_root / "docs/unified-ops/results/0362ul24/fixtures.json"
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_bytes(canonical(fixtures))
    return manifest_path


def plan(
    source: Path,
    kind: str,
    output: Path,
    fixture_root: Path,
    *,
    source_archive: Path | None = None,
    source_assembly: Path | None = None,
    identity: dict[str, Any] | None = None,
    probe: Path | None = None,
) -> dict[str, Any]:
    if identity is None:
        identity = source_identity(
            source, kind, source_archive=source_archive, source_assembly=source_assembly
        )
    else:
        identity = json.loads(json.dumps(identity))
    fixtures = fixture_manifest()
    probe_path = _selected_probe(probe)
    output = safe_external(output, "output")
    fixture_root = safe_external(fixture_root, "fixture root")
    identity["source_capture"] = str((output / "source-identity.json").resolve())
    return {
        "schema": "litchi_signed_native_consumer_plan.v1",
        "source": identity,
        "probe": artifact(probe_path),
        "probe_source": artifact(probe_path),
        "fixtures": {
            "manifest": artifact(FIXTURES),
            "schema": fixtures["schema"],
            "formats": sorted({str(row["format"]) for row in fixtures["fixtures"]}),
            "count": len(fixtures["fixtures"]),
        },
        "adapter": {
            "kind": kind,
            "source": "consumer-only generated module",
            "sha256": hashlib.sha256(adapter_source(kind).encode()).hexdigest(),
            "f2_field_bridge": kind == "f2",
            "xlsb_option_bridge": kind == "f2",
        },
        "harness_features": _reviewed_harness_features(kind),
        "build_policy": {
            "cargo": "caller-supplied pinned memory-capped shim",
            "jobs": 1,
            "incremental": 0,
            "offline": True,
            "locked": True,
            "memory_cap": "20G",
            "target_cap_bytes": TARGET_CAP_BYTES,
            "tmp_cap_bytes": TARGET_CAP_BYTES,
            "target_dir": str(output / "target"),
            "tmp_dir": str(output / "tmp"),
            "fixture_root": str(fixture_root),
            "timing": "not run by this driver",
        },
        "commands": [
            "<cargo-shim> generate-lockfile --offline",
            "<cargo-shim> metadata --offline --locked --format-version=1",
            "<cargo-shim> build --offline --locked --release --bin signed-native-consumer",
        ],
        "source_closure": {
            "archive": identity["source_archive"],
            "assembly": identity["source_assembly"],
            "tree": identity["source_tree"],
            "lock": identity["lock"],
            "toolchain": identity["toolchain"],
            "cargo_config": identity["cargo_config"],
        },
    }


def write_external(path: Path, payload: bytes) -> dict[str, Any]:
    path = safe_external(path, "output")
    if path.exists():
        raise fail(f"refusing to overwrite existing output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)
    return artifact(path)


def prepare(
    source: Path,
    kind: str,
    output: Path,
    fixture_root: Path,
    *,
    source_archive: Path | None = None,
    source_assembly: Path | None = None,
    probe: Path | None = None,
) -> dict[str, Any]:
    probe_path = _selected_probe(probe)
    probe_text = probe_path.read_text(encoding="utf-8")
    contract = probe_contract(probe_text)
    output = safe_external(output, "output")
    fixture_root = safe_external(fixture_root, "fixture root")
    if output.exists():
        raise fail(f"refusing to overwrite existing preparation directory: {output}")
    output.mkdir(parents=True)
    project = output / "consumer"
    (project / "src").mkdir(parents=True)
    identity = source_identity(
        source, kind, source_archive=source_archive, source_assembly=source_assembly
    )
    identity["source_capture"] = str((output / "source-identity.json").resolve())
    fixtures = fixture_manifest()
    (project / "src" / "main.rs").write_text(
        "mod api_compat;\n" + probe_text, encoding="utf-8"
    )
    (project / "src" / "api_compat.rs").write_text(adapter_source(kind), encoding="utf-8")
    # Exercise the consumer-only bridge at startup; production format sources
    # remain untouched and the probe's measured routes are unchanged.
    main = project / "src" / "main.rs"
    text = main.read_text(encoding="utf-8")
    marker = "fn main() {"
    if marker not in text:
        raise fail("signed probe has no main function")
    main.write_text(text.replace(marker, marker + "\n    api_compat::exercise_api();", 1), encoding="utf-8")
    (project / "Cargo.toml").write_text(project_toml(source, kind), encoding="utf-8")
    fixture_manifest_path = stage_fixtures(source, fixture_root, fixtures)
    (output / "fixture-manifest-receipt.json").write_bytes(
        canonical(
            {
                "schema": "litchi_signed_native_fixture_binding.v1",
                "manifest": artifact(fixture_manifest_path),
                "root": str(fixture_manifest_path.parents[4].resolve()),
            }
        )
    )
    (output / "probe-input.json").write_bytes(
        canonical({"schema": "litchi_signed_native_probe_input.v1", "probe": artifact(probe_path)})
    )
    generated_archive = None
    if identity["source_archive"] is None:
        generated_archive = output / "source.tar"
        with generated_archive.open("wb") as stream:
            subprocess.run(
                ["git", "archive", "--format=tar", "--prefix=litchi-source/", identity["commit"]],
                cwd=source,
                check=True,
                stdout=stream,
            )
        verified_archive = _verify_git_archive(source, generated_archive, identity["commit"])
        identity["source_archive"] = verified_archive["archive"]["artifact"]
        identity["source_tree"] = verified_archive["tree"]
        identity["tree_sha256"] = verified_archive["tree"]["sha256"]
        identity["verified_closure"] = verified_archive
        _checked_fixed_artifact(
            generated_archive,
            REVIEWED_SOURCE_ARCHIVES[kind],
            f"{kind} generated source archive",
        )
    effective_archive = generated_archive or source_archive
    identity["adapter"] = artifact(project / "src/api_compat.rs")
    (output / "source-identity.json").write_bytes(canonical(identity))
    (output / "adapter-api.json").write_bytes(
        canonical({"schema": "litchi_signed_native_adapter.v1", "kind": kind, "source_commit": identity["commit"], "module": artifact(project / "src/api_compat.rs")})
    )
    (output / "probe-contract.json").write_bytes(canonical(contract))
    plan_value = plan(
        source, kind, output, fixture_manifest_path.parents[4],
        source_archive=effective_archive,
        source_assembly=source_assembly,
        identity=identity,
        probe=probe_path,
    )
    plan_value["consumer_manifest"] = artifact(project / "Cargo.toml")
    (output / "plan.json").write_bytes(
        canonical(plan_value)
    )
    return {
        "status": "prepared",
        "output": str(output),
        "source": identity,
        "fixture_manifest": str(fixture_manifest_path),
        "fixture_root": str(fixture_manifest_path.parents[4].resolve()),
    }


def _directory_bytes(path: Path, cwd: Path) -> int:
    if not path.is_dir() or path.is_symlink():
        raise fail(f"expected a fresh directory: {path}")
    try:
        return int(run(["du", "-sxB1", "--", str(path)], cwd).split()[0])
    except (IndexError, ValueError) as error:
        raise fail(f"could not measure directory: {path}") from error


def _load_object(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise fail(f"{label} is invalid JSON: {error}") from error
    if not isinstance(value, dict):
        raise fail(f"{label} is not a JSON object")
    return value


def _verify_staged_fixture_manifest(manifest_path: Path, root: Path) -> dict[str, Any]:
    manifest_path = safe_external(manifest_path, "staged fixture manifest")
    root = safe_external(root, "staged fixture root")
    expected = root / "docs/unified-ops/results/0362ul24/fixtures.json"
    if manifest_path != expected or manifest_path.is_symlink():
        raise fail("staged fixture manifest is not the explicit probe manifest")
    if not root.is_dir() or root.is_symlink() or not manifest_path.is_file():
        raise fail("staged fixture root or manifest is missing")
    value = _load_object(manifest_path, "staged fixture manifest")
    if canonical(value) != manifest_path.read_bytes():
        raise fail("staged fixture manifest is not canonical JSON")
    rows = value.get("fixtures")
    if not isinstance(rows, list):
        raise fail("staged fixture manifest has no fixture rows")
    for index, row in enumerate(rows):
        if not isinstance(row, dict):
            raise fail(f"staged fixture row {index} is not an object")
        relative = row.get("relative_path")
        if not isinstance(relative, str):
            raise fail(f"staged fixture row {index} has no relative path")
        path = Path(relative)
        if path.is_absolute() or ".." in path.parts or "\\" in relative or "3rdparty" in path.parts:
            raise fail(f"staged fixture row {index} has an unsafe relative path")
        fixture = root / path
        if fixture.is_symlink() or not fixture.is_file():
            raise fail(f"staged fixture row {index} is not a regular file")
        resolved = fixture.resolve()
        if not resolved.is_relative_to(root.resolve()):
            raise fail(f"staged fixture row {index} escapes the fixture root")
        if digest(resolved) != row.get("input_sha256"):
            raise fail(f"staged fixture row {index} hash differs from its manifest")
    return value


def _revalidate_preparation(prepared: Path) -> dict[str, Any]:
    project = prepared / "consumer"
    plan_path = prepared / "plan.json"
    source_identity_path = prepared / "source-identity.json"
    adapter_path = project / "src/api_compat.rs"
    main_path = project / "src/main.rs"
    contract_path = prepared / "probe-contract.json"
    probe_input_path = prepared / "probe-input.json"
    fixture_binding_path = prepared / "fixture-manifest-receipt.json"
    required = (
        (project, "consumer project"),
        (plan_path, "plan"),
        (source_identity_path, "source identity"),
        (adapter_path, "consumer adapter"),
        (main_path, "generated probe"),
        (contract_path, "probe contract"),
        (probe_input_path, "probe input"),
        (fixture_binding_path, "fixture binding"),
    )
    for path, label in required:
        if not path.exists() or path.is_symlink():
            raise fail(f"prepared output lacks immutable {label}: {path}")
    recorded_identity = _load_object(source_identity_path, "source identity")
    plan_value = _load_object(plan_path, "plan")
    kind = recorded_identity.get("kind")
    source_root = recorded_identity.get("source_root")
    source_archive = recorded_identity.get("source_archive")
    source_assembly = recorded_identity.get("source_assembly")
    if kind not in SOURCE_KINDS or not isinstance(source_root, str):
        raise fail("source identity kind/root is incomplete")
    if not isinstance(source_archive, dict) or not isinstance(source_archive.get("path"), str):
        raise fail("source identity archive is incomplete")
    archive_path = Path(source_archive["path"])
    assembly_path = None
    if source_assembly is not None:
        if not isinstance(source_assembly, dict) or not isinstance(source_assembly.get("path"), str):
            raise fail("source identity assembly is incomplete")
        assembly_path = Path(source_assembly["path"])
    actual = source_identity(
        Path(source_root),
        kind,
        source_archive=archive_path,
        source_assembly=assembly_path,
        verify_closure=True,
    )
    actual["source_capture"] = str(source_identity_path.resolve())
    actual["adapter"] = artifact(adapter_path)
    if actual != recorded_identity:
        raise fail("source identity changed after preparation")
    if plan_value.get("source") != recorded_identity:
        raise fail("plan source identity differs from preparation identity")
    if plan_value.get("schema") != "litchi_signed_native_consumer_plan.v1":
        raise fail("plan schema is invalid")
    consumer_manifest = plan_value.get("consumer_manifest")
    if not isinstance(consumer_manifest, dict) or _verified_receipt_artifact(
        consumer_manifest, "consumer manifest"
    ) != artifact(project / "Cargo.toml"):
        raise fail("consumer manifest changed after preparation")
    probe_input = _load_object(probe_input_path, "probe input")
    probe_artifact = probe_input.get("probe")
    if probe_input.get("schema") != "litchi_signed_native_probe_input.v1" or not isinstance(probe_artifact, dict):
        raise fail("probe input schema is invalid")
    probe_path = Path(str(probe_artifact.get("path", "")))
    if _verified_receipt_artifact(probe_artifact, "probe input") != probe_artifact:
        raise fail("probe input artifact is stale")
    probe_text = probe_path.read_text(encoding="utf-8")
    if plan_value.get("probe") != probe_artifact or plan_value.get("probe_source") != probe_artifact:
        raise fail("plan probe differs from the prepared reviewed probe")
    contract = _load_object(contract_path, "probe contract")
    current_contract = probe_contract(probe_text)
    if contract != current_contract:
        raise fail("probe contract changed after preparation")
    generated_text = main_path.read_text(encoding="utf-8")
    if generated_text != "mod api_compat;\n" + probe_text.replace(
        "fn main() {", "fn main() {\n    api_compat::exercise_api();", 1
    ):
        raise fail("generated probe is not the reviewed probe plus consumer adapter")
    fixture_binding = _load_object(fixture_binding_path, "fixture binding")
    fixture_manifest = fixture_binding.get("manifest")
    fixture_root = fixture_binding.get("root")
    if not isinstance(fixture_manifest, dict) or not isinstance(fixture_root, str):
        raise fail("fixture binding is incomplete")
    staged = _verify_staged_fixture_manifest(Path(str(fixture_manifest["path"])), Path(fixture_root))
    if artifact(Path(str(fixture_manifest["path"]))) != fixture_manifest:
        raise fail("fixture manifest changed after preparation")
    return {
        "source": recorded_identity,
        "plan": plan_value,
        "fixture_manifest": fixture_manifest,
        "fixture_root": fixture_root,
        "fixture_value": staged,
        "adapter": artifact(adapter_path),
        "probe": artifact(main_path),
        "source_identity_path": source_identity_path,
    }


def build(prepared: Path, cargo: Path | None) -> dict[str, Any]:
    prepared = safe_external(prepared, "prepared output")
    project = prepared / "consumer"
    plan_path = prepared / "plan.json"
    source_identity_path = prepared / "source-identity.json"
    fixture_binding_path = prepared / "fixture-manifest-receipt.json"
    for path, label in (
        (project, "consumer project"),
        (plan_path, "plan"),
        (source_identity_path, "source identity"),
        (fixture_binding_path, "fixture binding"),
    ):
        if not path.exists():
            raise fail(f"prepared output lacks {label}: {path}")
    if not project.is_dir() or project.is_symlink():
        raise fail(f"consumer project is not a regular directory: {project}")
    revalidated = _revalidate_preparation(prepared)
    fixture_manifest = revalidated["fixture_manifest"]
    fixture_root = revalidated["fixture_root"]
    fixture_manifest_path = Path(str(fixture_manifest["path"]))
    if cargo is None:
        raise fail("build requires the pinned memory-capped cargo shim (--cargo)")
    cargo = safe_external(cargo, "cargo shim")
    if not cargo.is_file() or cargo.is_symlink() or not os.access(cargo, os.X_OK):
        raise fail(f"cargo shim is not an executable regular file: {cargo}")

    target = prepared / "target"
    temporary = prepared / "tmp"
    for path, label in ((target, "target"), (temporary, "TMPDIR")):
        if path.exists() or path.is_symlink():
            raise fail(f"refusing to reuse {label} directory: {path}")
        path.mkdir()

    cargo_cmd = str(cargo)
    env = dict(os.environ)
    env.update(
        {
            "CARGO_TARGET_DIR": str(target),
            "CARGO_BUILD_JOBS": "1",
            "CARGO_INCREMENTAL": "0",
            "CARGO_NET_OFFLINE": "true",
            "CARGO_TERM_COLOR": "never",
            "LITCHI_MEM_MAX": "20G",
            "TMPDIR": str(temporary),
            "LITCHI_SIGNED_FIXTURE_ROOT": fixture_root,
            "LITCHI_SIGNED_FIXTURE_MANIFEST": str(fixture_manifest_path),
            "LITCHI_SIGNED_FIXTURE_MANIFEST_SHA256": fixture_manifest["sha256"],
        }
    )
    run([cargo_cmd, "generate-lockfile", "--offline"], project, environment=env)
    metadata_raw = run(
        [cargo_cmd, "metadata", "--offline", "--locked", "--format-version=1"],
        project,
        environment=env,
    )
    try:
        metadata_value = json.loads(metadata_raw)
    except json.JSONDecodeError as error:
        raise fail(f"cargo metadata returned invalid JSON: {error}") from error
    if not isinstance(metadata_value, dict):
        raise fail("cargo metadata did not return an object")
    harness_projection = _validate_harness_metadata(
        metadata_value, revalidated["source"]["kind"]
    )
    metadata_path = prepared / "metadata.json"
    metadata_path.write_bytes(canonical(metadata_value))
    harness_projection_path = prepared / "harness-projection.json"
    harness_projection_path.write_bytes(canonical(harness_projection))
    command = [
        cargo_cmd,
        "build",
        "--offline",
        "--locked",
        "--release",
        "--bin",
        "signed-native-consumer",
    ]
    try:
        subprocess.run(command, cwd=project, env=env, check=True)
    except (OSError, subprocess.CalledProcessError) as error:
        raise fail(f"release build failed: {error}") from error
    binary = target / "release/signed-native-consumer"
    if not binary.is_file() or binary.is_symlink() or not os.access(binary, os.X_OK):
        raise fail(f"release binary is missing or not executable: {binary}")
    target_bytes = _directory_bytes(target, project)
    temporary_bytes = _directory_bytes(temporary, project)
    for label, size in (("target", target_bytes), ("TMPDIR", temporary_bytes)):
        if size >= TARGET_CAP_BYTES:
            raise fail(f"{label} exceeds 20 GB cap: {size}")

    source_identity = revalidated["source"]
    source_archive = source_identity["source_archive"]
    adapter_artifact = artifact(project / "src/api_compat.rs")
    consumer_manifest = artifact(project / "Cargo.toml")
    if source_identity.get("adapter") != adapter_artifact:
        raise fail("source identity adapter differs from prepared consumer adapter")
    probe_artifact = artifact(project / "src/main.rs")
    lock_artifact = artifact(project / "Cargo.lock")
    metadata_artifact = artifact(metadata_path)
    binary_artifact = artifact(binary)
    receipt = {
        "schema": "litchi_signed_native_build.v1",
        "status": "built",
        "source_commit": source_identity.get("commit"),
        "source": source_identity,
        "source_lock": source_identity.get("lock"),
        "source_archive": source_archive,
        "source_archive_sha256": source_archive.get("sha256"),
        "source_identity": source_identity,
        "reviewed_source_policy": source_identity.get("reviewed_policy"),
        "reviewed_projection": source_identity.get("reviewed_policy", {}).get("projection"),
        "plan": artifact(plan_path),
        "consumer_manifest": consumer_manifest,
        "adapter": adapter_artifact,
        "adapter_sha256": adapter_artifact["sha256"],
        "probe": probe_artifact,
        "lock": lock_artifact,
        "lock_sha256": lock_artifact["sha256"],
        "fixtures": fixture_manifest,
        "fixture_manifest_sha256": fixture_manifest["sha256"],
        "probe_contract": artifact(prepared / "probe-contract.json"),
        "fixture_root": fixture_root,
        "metadata": metadata_artifact,
        "dependency_projection": metadata_artifact,
        "harness_projection": artifact(harness_projection_path),
        "binary": binary_artifact,
        "binary_sha256": binary_artifact["sha256"],
        "cargo": artifact(cargo),
        "command": command,
        "environment": {
            key: env[key]
            for key in (
                "CARGO_TARGET_DIR",
                "CARGO_BUILD_JOBS",
                "CARGO_INCREMENTAL",
                "CARGO_NET_OFFLINE",
                "CARGO_TERM_COLOR",
                "LITCHI_MEM_MAX",
                "TMPDIR",
                "LITCHI_SIGNED_FIXTURE_ROOT",
                "LITCHI_SIGNED_FIXTURE_MANIFEST",
                "LITCHI_SIGNED_FIXTURE_MANIFEST_SHA256",
            )
        },
        "target_bytes": target_bytes,
        "tmp_bytes": temporary_bytes,
        "timing": None,
    }
    receipt_path = prepared / "build-receipt.json"
    if receipt_path.exists() or receipt_path.is_symlink():
        raise fail(f"refusing to overwrite build receipt: {receipt_path}")
    receipt_path.write_bytes(canonical(receipt))
    return {
        "status": "built",
        "receipt": str(receipt_path),
        "binary": binary_artifact,
        "target_bytes": target_bytes,
        "tmp_bytes": temporary_bytes,
    }


def _capability_evidence(
    capabilities_path: Path,
    *,
    source_commit: str,
    binary: dict[str, Any],
    fixture_manifest: dict[str, Any],
    fixture_root: str,
    endpoint_role: str,
) -> tuple[dict[str, str], dict[str, Any]]:
    if endpoint_role not in {"before", "after"}:
        raise fail("capability endpoint role must be before or after")
    capabilities_path = safe_external(capabilities_path, "capability receipt")
    value = _load_object(capabilities_path, "capability receipt")
    if value.get("schema") != "litchi_signed_native_capabilities.v1":
        raise fail("capability receipt schema mismatch")
    if value.get("source_commit") != source_commit:
        raise fail("capability receipt source commit differs from build")
    if "endpoint" in value and value.get("endpoint") != endpoint_role:
        raise fail("capability receipt endpoint role differs from build")
    cpu = value.get("cpu")
    if not isinstance(cpu, int) or cpu < 0:
        raise fail("capability receipt CPU binding is missing")
    recorded_binary = value.get("binary")
    if not isinstance(recorded_binary, dict) or _verified_receipt_artifact(recorded_binary, "capability binary") != binary:
        raise fail("capability receipt binary differs from build")
    routes = value.get("routes")
    if not isinstance(routes, dict) or set(routes) != set(ROUTES):
        raise fail("capability receipt does not enumerate the signed route matrix")
    allowed = {"available", "capability_missing", "expected_refusal", "not_applicable"}
    if any(status not in allowed for status in routes.values()):
        raise fail("capability receipt contains an invalid status")
    evidence = value.get("evidence")
    if not isinstance(evidence, dict) or set(evidence) != set(ROUTES):
        raise fail("capability receipt lacks one validate receipt per signed route")
    validator = _reviewed_module(
        "docs/unified-ops/results/0362ul24/validate_signed_native.py",
        "validate_signed_native_capability",
    )
    validator_endpoint = "before" if endpoint_role == "before" else "final"
    for route in ROUTES:
        row = evidence[route]
        receipt_artifact = _verified_receipt_artifact(row, f"capability evidence for {route}")
        receipt_path = Path(receipt_artifact["path"])
        receipt_path = safe_external(receipt_path, f"capability evidence for {route}")
        try:
            receipt_path.relative_to(_repository())
        except ValueError:
            pass
        else:
            raise fail(f"capability evidence for {route} must be external to this checkout")
        receipt = _load_object(receipt_path, f"capability evidence for {route}")
        if canonical(receipt) != receipt_path.read_bytes():
            raise fail(f"capability child receipt is not canonical for {route}")
        if receipt.get("schema") != "litchi_signed_native_child.v1" or receipt.get("mode") != "validate":
            raise fail(f"capability evidence for {route} is not a validate receipt")
        if receipt.get("route") != route:
            raise fail(f"capability evidence route mismatch for {route}")
        if receipt.get("calls") is not None:
            raise fail(f"capability evidence unexpectedly carries measure calls for {route}")
        if receipt.get("timeout_seconds") != CHILD_TIMEOUT_SECONDS:
            raise fail(f"capability evidence timeout is not the pinned guard for {route}")
        if receipt.get("exit_code") != 0 or receipt.get("timed_out") is not False:
            raise fail(f"capability evidence did not exit successfully for {route}")
        if receipt.get("guard_error") is not None or receipt.get("after_guard_error") is not None:
            raise fail(f"capability evidence guard failed for {route}")
        if receipt.get("fixture_root") != fixture_root or receipt.get("fixture_manifest") != fixture_manifest:
            raise fail(f"capability evidence fixture binding differs for {route}")
        expected_command = [
            "taskset", "-c", str(cpu), str(binary["path"]), "--mode", "validate",
            "--route", route, "--blocks", "1",
        ]
        command = receipt.get("command")
        if command != expected_command:
            raise fail(f"capability evidence command is not the pinned typed validation for {route}")
        process_group = receipt.get("process_group")
        if not isinstance(process_group, int) or process_group <= 0:
            raise fail(f"capability evidence process group is missing for {route}")
        rss_max = receipt.get("rss_max_bytes")
        aggregate_rss = receipt.get("aggregate_rss_max_bytes")
        child_rss_limit = getattr(validator, "CHILD_RSS_LIMIT", TARGET_CAP_BYTES)
        child_output_limit = getattr(validator, "CHILD_OUTPUT_LIMIT", 4 * 1024 * 1024)
        if (
            not isinstance(rss_max, int) or rss_max < 0 or rss_max > child_rss_limit
            or not isinstance(aggregate_rss, int) or aggregate_rss < 0 or aggregate_rss > child_rss_limit
            or rss_max != aggregate_rss
            or receipt.get("rss_limit_bytes") != child_rss_limit
            or receipt.get("output_limit_bytes") != child_output_limit
        ):
            raise fail(f"capability evidence aggregate RSS is missing for {route}")
        failed_after = receipt.get("failed_after")
        if (
            not isinstance(failed_after, dict)
            or failed_after.get("process_group_empty") is not True
            or failed_after.get("guard_error") is not None
            or failed_after.get("after_guard_error") is not None
            or failed_after.get("timed_out") is not False
            or failed_after.get("rss_max_bytes") != rss_max
            or failed_after.get("aggregate_rss_max_bytes") != aggregate_rss
            or failed_after.get("rss_limit_bytes") != child_rss_limit
            or failed_after.get("output_limit_bytes") != child_output_limit
        ):
            raise fail(f"capability evidence process group was not reaped for {route}")
        snapshots = receipt.get("host_snapshots")
        if (
            not isinstance(snapshots, list)
            or len(snapshots) < 2
            or not isinstance(receipt.get("during_snapshots"), int)
            or receipt["during_snapshots"] < 1
            or len(snapshots) != receipt["during_snapshots"] + 2
        ):
            raise fail(f"capability evidence host snapshots are incomplete for {route}")
        receipt_root = receipt_path.parent.resolve()
        for snapshot in snapshots:
            snapshot_artifact = _verified_receipt_artifact(snapshot, f"{route} host snapshot")
            snapshot_path = safe_external(Path(snapshot_artifact["path"]), f"{route} host snapshot")
            if not snapshot_path.is_relative_to(receipt_root):
                raise fail(f"capability host snapshot escapes the child receipt directory for {route}")
            snapshot_value = _load_object(snapshot_path, f"{route} host snapshot")
            if snapshot_value.get("blockers") != []:
                raise fail(f"capability evidence host snapshot has blockers for {route}")
        child_binary = receipt.get("binary")
        if not isinstance(child_binary, dict) or _verified_receipt_artifact(child_binary, f"{route} binary") != binary:
            raise fail(f"capability evidence binary mismatch for {route}")
        stdout_artifact = _verified_receipt_artifact(receipt.get("stdout"), f"{route} stdout")
        stderr_artifact = _verified_receipt_artifact(receipt.get("stderr"), f"{route} stderr")
        stdout_path = safe_external(Path(stdout_artifact["path"]), f"{route} stdout")
        stderr_path = safe_external(Path(stderr_artifact["path"]), f"{route} stderr")
        if not stdout_path.is_relative_to(receipt_root) or not stderr_path.is_relative_to(receipt_root):
            raise fail(f"capability child output escapes the child receipt directory for {route}")
        if (
            not isinstance(failed_after.get("stdout_bytes"), int)
            or failed_after["stdout_bytes"] != stdout_artifact["bytes"]
            or failed_after["stdout_bytes"] > child_output_limit
            or not isinstance(failed_after.get("stderr_bytes"), int)
            or failed_after["stderr_bytes"] != stderr_artifact["bytes"]
            or failed_after["stderr_bytes"] > child_output_limit
            or failed_after.get("exit_code") != receipt.get("exit_code")
        ):
            raise fail(f"capability child output guard is incomplete for {route}")
        try:
            result = json.loads(stdout_path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, json.JSONDecodeError) as error:
            raise fail(f"capability evidence stdout is invalid for {route}: {error}") from error
        if not isinstance(result, dict) or result.get("route") != route or canonical(result) != stdout_path.read_bytes():
            raise fail(f"capability evidence output route mismatch for {route}")
        try:
            child_value = dict(result)
            child_value["_receipt"] = str(receipt_path)
            child_value["_receipt_artifact"] = receipt_artifact
            if hasattr(validator, "validate_child_receipt"):
                validator.validate_child_receipt(
                    child_value,
                    expected_route=route,
                    expected_mode="validate",
                    expected_calls=None,
                    expected_binary=binary,
                    expected_fixture_root=fixture_root,
                    expected_fixture_manifest=fixture_manifest,
                    expected_cpu=cpu,
                    output_path=receipt_root,
                    expected_endpoint=validator_endpoint,
                )
            else:
                validator.validate_route(result, route, endpoint=validator_endpoint)
        except Exception as error:
            raise fail(f"typed route validation failed for {route}: {error}") from error
        observed = result.get("capability")
        if route.endswith("/refusal"):
            rows = result.get("fixtures")
            if not isinstance(rows, list) or not rows or any(
                not isinstance(row, dict) or row.get("status") != "expected_refusal"
                for row in rows
            ):
                raise fail(f"refusal evidence is not an expected refusal for {route}")
            observed = "expected_refusal"
        if observed != routes[route]:
            raise fail(f"capability status for {route} is not bound to probe output")
    return {str(route): str(status) for route, status in routes.items()}, value


def endpoint(
    prepared: Path,
    capabilities: Path,
    endpoint_role: str,
    output: Path | None = None,
) -> dict[str, Any]:
    """Bind one post-build, untimed capability receipt into the endpoint schema.

    The capability receipt is produced by a guarded probe preflight.  Requiring
    every child receipt here prevents a caller from hand-writing endpoint
    capability claims while keeping build and timing stages separate.
    """
    if endpoint_role not in {"before", "after"}:
        raise fail("endpoint role must be before or after")
    prepared = safe_external(prepared, "prepared output")
    capabilities = safe_external(capabilities, "capability receipt")
    receipt_path = prepared / "build-receipt.json"
    source_identity_path = prepared / "source-identity.json"
    fixture_binding_path = prepared / "fixture-manifest-receipt.json"
    for path, label in (
        (receipt_path, "build receipt"),
        (source_identity_path, "source identity"),
        (fixture_binding_path, "fixture binding"),
    ):
        if not path.is_file() or path.is_symlink():
            raise fail(f"prepared output lacks {label}: {path}")
    receipt = _load_object(receipt_path, "build receipt")
    if receipt.get("schema") != "litchi_signed_native_build.v1" or receipt.get("status") != "built":
        raise fail("build receipt is not a completed signed-native build")
    source_identity = _load_object(source_identity_path, "source identity")
    commit = source_identity.get("commit")
    source_ref = source_identity.get("source_ref")
    source_capture = source_identity.get("source_capture")
    source_root = source_identity.get("source_root")
    tree_sha256 = source_identity.get("tree_sha256")
    if (
        not isinstance(commit, str)
        or not FULL_SHA.fullmatch(commit)
        or not isinstance(source_ref, str)
        or not source_ref
        or not isinstance(source_capture, str)
        or not isinstance(source_root, str)
        or not Path(source_root).is_dir()
        or not isinstance(tree_sha256, str)
        or SHA256.fullmatch(tree_sha256) is None
    ):
        raise fail("source identity lacks the required pinned fields")
    if Path(source_capture).resolve() != source_identity_path.resolve():
        raise fail("source capture is not the immutable source identity receipt")
    source_archive = receipt.get("source_archive")
    adapter = receipt.get("adapter")
    lock = receipt.get("lock")
    binary = receipt.get("binary")
    probe = receipt.get("probe")
    fixtures = receipt.get("fixtures")
    dependency_projection = receipt.get("dependency_projection")
    for value, label in (
        (source_archive, "source archive"),
        (adapter, "adapter"),
        (lock, "lock"),
        (binary, "binary"),
        (probe, "probe"),
        (fixtures, "fixtures"),
        (dependency_projection, "dependency projection"),
    ):
        if not isinstance(value, dict):
            raise fail(f"build receipt lacks {label} artifact")
        _verified_receipt_artifact(value, label)
    if source_archive.get("sha256") != receipt.get("source_archive_sha256"):
        raise fail("build receipt source archive hash is not bound")
    if source_identity.get("source_archive") != source_archive:
        raise fail("source identity archive differs from build receipt")
    if adapter.get("sha256") != receipt.get("adapter_sha256"):
        raise fail("build receipt adapter hash is not bound")
    if lock.get("sha256") != receipt.get("lock_sha256"):
        raise fail("build receipt lock hash is not bound")
    if binary.get("sha256") != receipt.get("binary_sha256"):
        raise fail("build receipt binary hash is not bound")
    if receipt.get("source_commit") != commit:
        raise fail("build receipt source commit differs")
    identity_adapter = source_identity.get("adapter")
    if identity_adapter != adapter:
        raise fail("source identity adapter differs from build receipt")
    fixture_binding = _load_object(fixture_binding_path, "fixture binding")
    fixture_manifest = fixture_binding.get("manifest")
    fixture_root = fixture_binding.get("root")
    if not isinstance(fixture_manifest, dict) or fixture_manifest != fixtures:
        raise fail("build receipt fixtures differ from staged fixture binding")
    if not isinstance(fixture_root, str) or not Path(fixture_root).is_dir():
        raise fail("staged fixture root is missing")
    statuses, capability_value = _capability_evidence(
        capabilities,
        source_commit=commit,
        binary=binary,
        fixture_manifest=fixtures,
        fixture_root=fixture_root,
        endpoint_role=endpoint_role,
    )
    endpoint_value = {
        "schema": "litchi_signed_native_endpoint.v1",
        "source_commit": commit,
        "source_identity": {
            "commit": commit,
            "source_ref": source_ref,
            "source_capture": source_capture,
            "source_root": source_root,
            "tree_sha256": tree_sha256,
            "archive": source_archive,
            "adapter": adapter,
        },
        "source": artifact(source_identity_path),
        "source_archive": source_archive,
        "adapter": adapter,
        "lock": lock,
        "binary": binary,
        "probe": probe,
        "fixtures": fixtures,
        "dependency_projection": dependency_projection,
        "harness_projection": receipt.get("harness_projection"),
        "reviewed_source_policy": receipt.get("reviewed_source_policy"),
        "reviewed_projection": receipt.get("reviewed_projection"),
        "build_receipt": artifact(receipt_path),
        "capabilities": statuses,
        "capabilities_receipt": artifact(capabilities),
        "fixture_root": fixture_root,
        "build": {
            "source_commit": receipt.get("source_commit"),
            "binary_sha256": receipt.get("binary_sha256"),
            "source_archive_sha256": receipt.get("source_archive_sha256"),
            "adapter_sha256": receipt.get("adapter_sha256"),
            "lock_sha256": receipt.get("lock_sha256"),
            "capabilities_schema": capability_value.get("schema"),
        },
    }
    endpoint_path = safe_external(output or (prepared / "endpoint.json"), "endpoint")
    if endpoint_path.exists() or endpoint_path.is_symlink():
        raise fail(f"refusing to overwrite endpoint manifest: {endpoint_path}")
    endpoint_path.parent.mkdir(parents=True, exist_ok=True)
    endpoint_path.write_bytes(canonical(endpoint_value))
    return {"status": "endpoint", "endpoint": str(endpoint_path), "binary": binary}


def main() -> int:
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("plan", "prepare"):
        p = sub.add_parser(name)
        p.add_argument("--source", required=True, type=Path)
        p.add_argument("--kind", required=True, choices=sorted(SOURCE_KINDS))
        p.add_argument("--output", required=True, type=Path)
        p.add_argument("--fixture-root", required=True, type=Path)
        p.add_argument("--probe", type=Path)
        p.add_argument("--source-archive", type=Path)
        p.add_argument("--source-assembly", type=Path)
    b = sub.add_parser("build")
    b.add_argument("--prepared", required=True, type=Path)
    b.add_argument("--cargo", required=True, type=Path)
    e = sub.add_parser("endpoint")
    e.add_argument("--prepared", required=True, type=Path)
    e.add_argument("--capabilities", required=True, type=Path)
    e.add_argument("--endpoint", required=True, choices=("before", "after"))
    e.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.command == "plan":
        value = plan(
            args.source, args.kind, args.output, args.fixture_root,
            source_archive=args.source_archive,
            source_assembly=args.source_assembly,
            probe=args.probe,
        )
    elif args.command == "prepare":
        value = prepare(
            args.source, args.kind, args.output, args.fixture_root,
            source_archive=args.source_archive,
            source_assembly=args.source_assembly,
            probe=args.probe,
        )
    elif args.command == "build":
        value = build(args.prepared, args.cargo)
    else:
        value = endpoint(args.prepared, args.capabilities, args.endpoint, args.output)
    print(json.dumps(value, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"signed-native preparation failed: {error}", file=sys.stderr)
        raise
