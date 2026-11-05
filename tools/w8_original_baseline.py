#!/usr/bin/env python3
"""Assemble the original W8 source without changing its production tree.

The W8 before leg is the pristine ``f2bbbc8d07`` tree.  That commit has the
old public result fields (``Commit.file`` and ``BatchCommit.file``), while the
reviewed 109-route probe was written against the later accessor spelling
``into_file()``.  This tool creates a standalone example adapter that changes
only those four consumer call sites.  It never creates a production patch.

The source archive is checked against every blob in the f2 tree.  The external
workspace is checked the same way, with only the corrected probe, the generated
adapter and an explicitly pinned external lockfile permitted as additions.
Preparation emits route, fixture and dependency descriptors only; compilation,
validation and performance qualification remain pending.
"""

from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import os
from pathlib import Path, PurePosixPath
import re
import subprocess
import sys
import tarfile
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools import w8_file_provenance as provenance


BASELINE_COMMIT = "f2bbbc8d07ab5da26c180f1ff74813d39f632d54"
PROBE_SOURCE_COMMIT = "56dbd37551ba204e99199feeaf61a8dab848046d"
PROBE_RELATIVE = Path("crates/litchi-ops/examples/file_w8_qualified_probe.rs")
ADAPTER_RELATIVE = Path("crates/litchi-ops/examples/file_w8_original_baseline_probe.rs")
LOCK_RELATIVE = Path("Cargo.lock")
REVIEWED_PROBE_SHA256 = "21c8e3833c96411a612a2631f98093bd70cf0a5c28b2952c3d15873f464e54b0"
REVIEWED_PROBE_BYTES = 75_035
REVIEWED_LOCK_SHA256 = "1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c"
REVIEWED_LOCK_BYTES = 139_441
TARGET = "x86_64-unknown-linux-gnu"
PROFILE = "release"
W8_FEATURES = (
    "docx", "xlsx", "xlsb", "pptx", "doc", "xls", "ppt", "rtf", "odt", "ods", "odp",
)


class OverlayError(RuntimeError):
    """An input is stale, ambiguous, or not bound to the before workload."""


def fail(message: str) -> OverlayError:
    return OverlayError(message)


def canonical(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def digest(path: Path) -> str:
    try:
        value = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                value.update(block)
        return value.hexdigest()
    except OSError as error:
        raise fail(f"cannot read {path}: {error}") from error


def artifact(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if not path.is_file():
        raise fail(f"artifact is not a regular file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": digest(path)}


def external_output(path: Path, label: str) -> Path:
    path = provenance.safe_location(path, label, output=True)
    try:
        path.relative_to(ROOT)
    except ValueError:
        return path
    raise fail(f"{label} must be outside the repository: {path}")


def write_new(path: Path, value: object) -> dict[str, Any]:
    path = external_output(path, "generated output")
    if path.exists():
        raise fail(f"refusing to overwrite generated output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = canonical(value)
    path.write_bytes(payload)
    return artifact(path)


def write_bytes_new(path: Path, payload: bytes) -> dict[str, Any]:
    path = external_output(path, "generated output")
    if path.exists():
        raise fail(f"refusing to overwrite generated output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)
    return artifact(path)


def run_git(root: Path, *arguments: str, binary: bool = False) -> bytes | str:
    try:
        result = subprocess.run(["git", "-C", str(root), *arguments], check=True, capture_output=True)
    except (OSError, subprocess.CalledProcessError) as error:
        detail = error.stderr.decode(errors="replace").strip() if isinstance(error, subprocess.CalledProcessError) else str(error)
        raise fail(f"git {' '.join(arguments)} failed: {detail}") from error
    return result.stdout if binary else result.stdout.decode()


def checked_commit(root: Path, value: str, label: str) -> None:
    actual = str(run_git(root, "rev-parse", "--verify", f"{value}^{{commit}}")).strip()
    if actual != value:
        raise fail(f"{label} is not the reviewed full commit: {actual}")


def object_bytes(root: Path, revision: str, relative: Path) -> bytes:
    try:
        return bytes(run_git(root, "show", f"{revision}:{relative.as_posix()}", binary=True))
    except OverlayError as error:
        raise fail(f"missing {relative} at {revision}: {error}") from error


def git_blob_sha1(payload: bytes) -> str:
    return hashlib.sha1(f"blob {len(payload)}\0".encode() + payload).hexdigest()


def git_tree(root: Path) -> dict[str, tuple[str, str]]:
    """Return the complete path -> (mode, blob id) map for the f2 tree."""

    raw = bytes(run_git(root, "ls-tree", "-r", "--full-tree", "-z", BASELINE_COMMIT, binary=True))
    records: dict[str, tuple[str, str]] = {}
    for entry in raw.split(b"\0"):
        if not entry:
            continue
        try:
            header, path_bytes = entry.split(b"\t", 1)
            mode, kind, object_id = header.decode("ascii").split(" ", 2)
            path = path_bytes.decode("utf-8")
        except (ValueError, UnicodeDecodeError) as error:
            raise fail("f2 tree contains an unreadable path record") from error
        if kind != "blob":
            raise fail(f"f2 tree contains unsupported {kind} entry: {path}")
        if path in records:
            raise fail(f"f2 tree contains duplicate path: {path}")
        records[path] = (mode, object_id)
    if not records:
        raise fail("f2 tree is empty")
    return records


def tree_digest(tree: dict[str, tuple[str, str]]) -> str:
    return digest_bytes(canonical([[path, mode, blob] for path, (mode, blob) in sorted(tree.items())]))


def verify_history(root: Path, tree: dict[str, tuple[str, str]]) -> dict[str, Any]:
    checked_commit(root, BASELINE_COMMIT, "baseline commit")
    checked_commit(root, PROBE_SOURCE_COMMIT, "corrected probe source commit")
    if PROBE_RELATIVE.as_posix() in tree or ADAPTER_RELATIVE.as_posix() in tree:
        raise fail("the pristine f2 tree unexpectedly contains a W8 probe or adapter")
    corrected = object_bytes(root, PROBE_SOURCE_COMMIT, PROBE_RELATIVE)
    if len(corrected) != REVIEWED_PROBE_BYTES or digest_bytes(corrected) != REVIEWED_PROBE_SHA256:
        raise fail("the corrected probe blob at its reviewed source commit changed")
    return {
        "baseline_commit": BASELINE_COMMIT,
        "probe_source_commit": PROBE_SOURCE_COMMIT,
        "tree_files": len(tree),
        "tree_sha256": tree_digest(tree),
        "corrected_probe": {"path": PROBE_RELATIVE.as_posix(), "bytes": REVIEWED_PROBE_BYTES, "sha256": REVIEWED_PROBE_SHA256},
    }


def _archive_name(name: str) -> str:
    if name.startswith("/"):
        raise fail(f"source archive has an absolute member: {name}")
    pieces = PurePosixPath(name).parts
    if ".." in pieces:
        raise fail(f"source archive has a parent traversal member: {name}")
    return "/".join(piece for piece in pieces if piece not in ("", "."))


def _archive_prefixes(members: list[tarfile.TarInfo]) -> set[str]:
    prefixes: set[str] = set()
    for member in members:
        if member.isfile():
            name = _archive_name(member.name)
            if name == "Cargo.toml":
                prefixes.add("")
            elif name.endswith("/Cargo.toml"):
                prefixes.add(name[: -len("/Cargo.toml")])
    return prefixes


def _archive_directories(tree: dict[str, tuple[str, str]]) -> set[str]:
    directories = {""}
    for path in tree:
        current = PurePosixPath(path).parent
        while str(current) not in ("", "."):
            directories.add(str(current))
            current = current.parent
    return directories


def _strip_prefix(name: str, prefix: str) -> str | None:
    name = _archive_name(name)
    if not prefix:
        return name
    if name == prefix:
        return None
    marker = prefix + "/"
    if not name.startswith(marker):
        return None
    return name[len(marker) :]


def verify_source_archive(archive: Path, tree: dict[str, tuple[str, str]]) -> dict[str, Any]:
    archive = provenance.safe_location(archive, "pristine source archive")
    if not archive.is_file():
        raise fail(f"pristine source archive is missing: {archive}")
    try:
        with tarfile.open(archive, "r:*") as stream:
            members = stream.getmembers()
            names: set[str] = set()
            for member in members:
                name = _archive_name(member.name)
                if name in names:
                    raise fail(f"source archive has a duplicate member: {member.name}")
                names.add(name)
                if not member.isfile() and not member.isdir():
                    raise fail(f"source archive has an unsupported non-regular member: {member.name}")
            prefixes = _archive_prefixes(members)
            expected_directories = _archive_directories(tree)
            matches: list[str] = []
            for prefix in sorted(prefixes):
                seen: set[str] = set()
                valid = True
                for member in members:
                    relative = _strip_prefix(member.name, prefix)
                    if relative is None:
                        if prefix and _archive_name(member.name) == prefix:
                            continue
                        valid = False
                        break
                    if member.isdir():
                        if relative not in expected_directories:
                            valid = False
                            break
                        continue
                    if relative in seen or relative not in tree:
                        valid = False
                        break
                    seen.add(relative)
                if valid and seen == set(tree):
                    matches.append(prefix)
            if len(matches) != 1:
                raise fail("pristine source archive must contain exactly the f2 file set under one prefix")
            prefix = matches[0]
            seen: set[str] = set()
            for member in members:
                if not member.isfile():
                    if member.isdir():
                        continue
                    raise fail(f"source archive has an unsupported non-regular member: {member.name}")
                relative = _strip_prefix(member.name, prefix)
                if relative is None:
                    continue
                seen.add(relative)
                expected_mode, expected_blob = tree[relative]
                if (member.mode & 0o111) != (int(expected_mode, 8) & 0o111):
                    raise fail(f"source archive executable mode differs at {relative}")
                handle = stream.extractfile(member)
                if handle is None:
                    raise fail(f"source archive member is unreadable: {relative}")
                if git_blob_sha1(handle.read()) != expected_blob:
                    raise fail(f"source archive blob differs at {relative}")
            if seen != set(tree):
                raise fail("source archive file set changed while checking blobs")
    except (OSError, tarfile.TarError) as error:
        raise fail(f"cannot inspect pristine source archive {archive}: {error}") from error
    return {"artifact": artifact(archive), "prefix": prefix, "file_count": len(tree), "tree_sha256": tree_digest(tree), "production_delta": {"files_changed": 0, "api_overlay": False}}


def _workspace_files(workspace: Path) -> set[str]:
    actual: set[str] = set()
    for base, directories, files in os.walk(workspace, followlinks=False):
        base_path = Path(base)
        for name in directories + files:
            path = base_path / name
            if path.is_symlink():
                raise fail(f"workspace contains an unsupported symlink: {path}")
        for name in files:
            actual.add((base_path / name).relative_to(workspace).as_posix())
    return actual


def _check_workspace_blob(path: Path, expected_mode: str, expected_blob: str, label: str) -> None:
    if not path.is_file() or path.is_symlink():
        raise fail(f"workspace {label} is not a regular file: {path}")
    # Git archive extraction and this worktree use a shared umask, so the
    # writable bits are not a source identity.  The Git blob identity below
    # is the production closure; the tree mode is retained in tree_digest.
    if (path.stat().st_mode & 0o111) != (int(expected_mode, 8) & 0o111):
        raise fail(f"workspace executable mode differs at {path}")
    try:
        payload = path.read_bytes()
    except OSError as error:
        raise fail(f"cannot read workspace {path}: {error}") from error
    if git_blob_sha1(payload) != expected_blob:
        raise fail(f"workspace production blob differs at {path}")


def verify_workspace(workspace: Path, tree: dict[str, tuple[str, str]], probe: Path, adapter: Path, lock: Path, adapter_bytes: bytes) -> dict[str, Any]:
    workspace = provenance.safe_location(workspace, "pristine f2 workspace")
    if not workspace.is_dir():
        raise fail(f"pristine f2 workspace is not a directory: {workspace}")
    if probe != workspace / PROBE_RELATIVE:
        raise fail(f"probe must be {PROBE_RELATIVE} inside the workspace")
    if adapter != workspace / ADAPTER_RELATIVE:
        raise fail(f"adapter must be {ADAPTER_RELATIVE} inside the workspace")
    if lock != workspace / LOCK_RELATIVE:
        raise fail(f"pinned lock must be {LOCK_RELATIVE} inside the workspace")
    for path, label in ((probe, "corrected probe"), (adapter, "consumer adapter"), (lock, "pinned lock")):
        provenance.safe_location(path, label)
        if not path.is_file():
            raise fail(f"workspace {label} is missing: {path}")
    if lock.stat().st_size != REVIEWED_LOCK_BYTES or digest(lock) != REVIEWED_LOCK_SHA256:
        raise fail("workspace pinned lock is not the reviewed historical f2 lock")
    actual = _workspace_files(workspace)
    allowed = {PROBE_RELATIVE.as_posix(), ADAPTER_RELATIVE.as_posix(), LOCK_RELATIVE.as_posix()}
    expected = set(tree) | allowed
    extra = sorted(actual - expected)
    missing = sorted(expected - actual)
    if extra:
        raise fail(f"workspace has files outside the f2 tree and adapter allowance: {extra[:5]!r}")
    if missing:
        raise fail(f"workspace is missing files from the f2 tree or adapter allowance: {missing[:5]!r}")
    for relative, (mode, blob) in tree.items():
        _check_workspace_blob(workspace / relative, mode, blob, relative)
    if probe.stat().st_size != REVIEWED_PROBE_BYTES or digest(probe) != REVIEWED_PROBE_SHA256:
        raise fail("workspace corrected probe is not the reviewed d39 blob")
    if adapter.read_bytes() != adapter_bytes:
        raise fail("workspace consumer adapter does not match the generated source")
    return {
        "workspace": str(workspace),
        "production_files": len(tree),
        "production_tree_sha256": tree_digest(tree),
        "probe": artifact(probe),
        "adapter": artifact(adapter),
        "lock": artifact(lock),
        "allowed_additions": [PROBE_RELATIVE.as_posix(), ADAPTER_RELATIVE.as_posix(), LOCK_RELATIVE.as_posix()],
    }


TRANSFORMS = (
    ("unified_batch", "BatchCommit.file"),
    ("unified_import", "Commit<Summary>.file"),
    ("unified_merge", "Commit<Summary>.file"),
    ("unified_refresh", "Commit<Summary>.file"),
)
XLSB_RESOLUTION_SEQUENCE = '\n                        .resolve()\n                        .expect("xlsb worksheet resolution")'
XLSB_RESOLUTION_LABEL = '.resolve().expect("xlsb worksheet resolution")'
XLSB_RESOLUTION_TRANSFORMS = (
    "native_set",
    "native_import",
    "native_merge",
)


def _method_segment(source: str, name: str) -> str:
    match = re.search(rf"(?m)^        fn\s+{re.escape(name)}\s*\(", source)
    if not match:
        raise fail(f"probe has no method {name}()")
    end = re.search(r"(?m)^        (?:fn|struct|enum|impl|#\[)", source[match.end() :])
    stop = match.end() + end.start() if end else len(source)
    return source[match.start() : stop]


def consumer_adapter(source: str) -> tuple[bytes, list[dict[str, Any]]]:
    resolution_count = source.count(".resolve()")
    if resolution_count not in (0, len(XLSB_RESOLUTION_TRANSFORMS)):
        raise fail("consumer adapter expected either the reviewed three XLSB resolution calls or none")
    if source.count(XLSB_RESOLUTION_SEQUENCE) != resolution_count:
        raise fail("consumer adapter found an unexpected XLSB resolution expression")
    if source.count(".into_file()") != len(TRANSFORMS):
        raise fail("corrected probe no longer has exactly four consumer into_file call sites")
    positions: list[tuple[int, str, str, str, str]] = []
    if resolution_count:
        for function in XLSB_RESOLUTION_TRANSFORMS:
            segment = _method_segment(source, function)
            if segment.count(".resolve()") != 1 or segment.count(XLSB_RESOLUTION_SEQUENCE) != 1:
                raise fail(f"consumer adapter expected one reviewed XLSB resolution call in {function}()")
            positions.append(
                (
                    source.index(segment) + segment.index(XLSB_RESOLUTION_SEQUENCE),
                    function,
                    XLSB_RESOLUTION_SEQUENCE,
                    "",
                    "xlsb_resolution",
                )
            )
    for function, old_api in TRANSFORMS:
        segment = provenance.function_segment(source, function)
        if segment.count(".into_file()") != 1:
            raise fail(f"consumer adapter expected one into_file call in {function}()")
        positions.append(
            (
                source.index(segment) + segment.index(".into_file()"),
                function,
                ".into_file()",
                ".file",
                old_api,
            )
        )
    transformed = source
    records: list[dict[str, Any]] = []
    for position, function, old_text, new_text, kind in sorted(positions, reverse=True):
        if transformed[position : position + len(old_text)] != old_text:
            raise fail(f"consumer adapter replacement offset changed in {function}()")
        transformed = transformed[:position] + new_text + transformed[position + len(old_text) :]
        if kind == "xlsb_resolution":
            records.append(
                {
                    "function": function,
                    "old_accessor": XLSB_RESOLUTION_LABEL,
                    "new_accessor": "",
                    "old_api": "f2 cell_values::edit_cell_values() -> Result<Option<Edit>>",
                }
            )
        else:
            records.append({"function": function, "old_accessor": old_text, "new_accessor": new_text, "old_api": kind})
    records.reverse()
    if ".into_file()" in transformed:
        raise fail("consumer adapter left an into_file call in the standalone probe")
    if ".resolve()" in transformed:
        raise fail("consumer adapter left an XLSB resolution call in the standalone probe")
    return transformed.encode(), records


def ensure_adapter(path: Path, payload: bytes) -> tuple[dict[str, Any], bool]:
    path = provenance.safe_location(path, "consumer adapter")
    if path.exists():
        if path.read_bytes() != payload:
            raise fail(f"refusing to use a stale consumer adapter: {path}")
        return artifact(path), False
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)
    return artifact(path), True


def adapter_patch(corrected: bytes, adapter: bytes) -> bytes:
    """Describe only the two added workload/consumer files, never production."""

    chunks: list[str] = []
    for relative, payload in ((PROBE_RELATIVE, corrected), (ADAPTER_RELATIVE, adapter)):
        text = payload.decode("utf-8")
        chunks.extend(
            difflib.unified_diff(
                [],
                text.splitlines(keepends=True),
                fromfile="/dev/null",
                tofile=f"b/{relative.as_posix()}",
            )
        )
    value = "".join(chunks).encode()
    headers = {
        "--- /dev/null",
        f"+++ b/{PROBE_RELATIVE.as_posix()}",
        f"+++ b/{ADAPTER_RELATIVE.as_posix()}",
    }
    if not value or any(
        line.startswith(("--- ", "+++ ")) and line.rstrip("\n") not in headers
        for line in value.decode("utf-8").splitlines(keepends=True)
    ):
        raise fail("consumer adapter patch contains a path outside the two added files")
    return value


def _route_binding(suffix: str) -> dict[str, Any]:
    names = {
        "open_get": "OpenGet", "save": "Save", "repeated_save": "RepeatedSave",
        "set_commit": "SetCommit", "set_publication": "SetPublication",
        "batch_commit": "BatchCommit", "batch_publication": "BatchPublication",
        "import_commit": "ImportCommit", "import_publication": "ImportPublication",
        "merge_commit": "MergeCommit", "merge_publication": "MergePublication",
        "refresh_commit": "RefreshCommit", "refresh_publication": "RefreshPublication",
    }
    enum = names[suffix]
    if suffix == "open_get":
        return {
            "phase": "read", "operation_count": 1, "changed": False, "exact_noop": True,
            "native_timing": "sample(calls, || { LeafPackage::open(family.name, &family.bytes); package.read_marker(family.name) })",
            "facade_timing": "sample(calls, || { source_facade(family); file.get(family.marker).view().text() })",
            "native_publication": "source_leaf.publish() in validation()", "facade_publication": "source_facade.to_bytes() in validation()",
            "native_dispatch": "LeafPackage::open(family.name, &family.bytes)", "facade_dispatch": "source_facade(family) then file.get(family.marker)",
            "semantic_marker": "family.marker", "adapter_api": None,
            "assertions": (("routes_for", "Route::OpenGet"), ("phase_once", "(Side::Leaf, Route::OpenGet)"), ("phase_once", "LeafPackage::open(family.name, &family.bytes)"), ("phase_once", "package.read_marker(family.name)"), ("phase_once", "(Side::Facade, Route::OpenGet)"), ("phase_once", "source_facade(family)"), ("phase_once", "file.get(family.marker)"), ("validation", "route.exact_noop()")),
        }
    if suffix == "save":
        return {
            "phase": "publication", "operation_count": 1, "changed": False, "exact_noop": True,
            "native_timing": "sample(calls, || leaf_source.publish().len())", "facade_timing": "sample(calls, || facade_source.to_bytes().expect(\"facade save\").len())",
            "native_publication": "leaf_source.publish().len()", "facade_publication": "facade_source.to_bytes().expect(\"facade save\").len()",
            "native_dispatch": "leaf_source.publish()", "facade_dispatch": "facade_source.to_bytes()", "semantic_marker": "family.marker", "adapter_api": None,
            "assertions": (("routes_for", "Route::Save"), ("phase_once", "(Side::Leaf, Route::Save)"), ("phase_once", "leaf_source.publish().len()"), ("phase_once", "(Side::Facade, Route::Save)"), ("phase_once", "facade_source.to_bytes().expect(\"facade save\")"), ("validation", "route.exact_noop()")),
        }
    if suffix == "repeated_save":
        return {
            "phase": "publication", "operation_count": 2, "changed": False, "exact_noop": True,
            "native_timing": "sample_with_operations(calls, 2, || { leaf_source.publish(); leaf_source.publish().len() })", "facade_timing": "sample_with_operations(calls, 2, || { facade_source.to_bytes(); facade_source.to_bytes().len() })",
            "native_publication": "leaf_source.publish() twice", "facade_publication": "facade_source.to_bytes() twice",
            "native_dispatch": "leaf_source.publish() then leaf_source.publish().len()", "facade_dispatch": "facade_source.to_bytes() then facade_source.to_bytes().len()", "semantic_marker": "family.marker", "adapter_api": None,
            "assertions": (("routes_for", "Route::RepeatedSave"), ("phase_once", "(Side::Leaf, Route::RepeatedSave)"), ("phase_once", "sample_with_operations(calls, 2"), ("phase_once", "(Side::Facade, Route::RepeatedSave)"), ("phase_once", "facade_source.to_bytes()"), ("validation", "route.exact_noop()")),
        }
    groups = {
        "set": ("SetCommit", "SetPublication", "source.native_set", "unified_set", "Commit.file"),
        "batch": ("BatchCommit", "BatchPublication", "source.native_batch", "unified_batch", "BatchCommit.file"),
        "import": ("ImportCommit", "ImportPublication", "source.native_import", "unified_import", "Commit<Summary>.file"),
        "merge": ("MergeCommit", "MergePublication", "source.native_merge", "unified_merge", "Commit<Summary>.file"),
        "refresh": ("RefreshCommit", "RefreshPublication", "source.native_refresh", "unified_refresh", "Commit<Summary>.file"),
    }
    group = next(key for key, values in groups.items() if suffix.startswith(key + "_"))
    commit, publication, native, facade, api = groups[group]
    is_publication = suffix.endswith("publication")
    changed = group != "refresh"
    marker_group = {"set": "Route::SetCommit | Route::SetPublication", "batch": "Route::BatchCommit | Route::BatchPublication", "import": "Route::ImportCommit | Route::ImportPublication", "merge": "Route::MergeCommit | Route::MergePublication", "refresh": "Route::RefreshCommit | Route::RefreshPublication"}[group]
    if is_publication:
        native_timing = "sample(calls, || prepared.publish().len())"
        facade_timing = "sample(calls, || prepared.to_bytes().expect(\"facade publication\").len())"
        native_dispatch = "leaf_commit(family, leaf_source, route, import_input) then prepared.publish()"
        facade_dispatch = "facade_commit(family, facade_source, route, import_input, record) then prepared.to_bytes()"
        phase_assertions = (("phase_once", "route.is_publication()"), ("phase_once", "prepared.publish().len()"), ("phase_once", "prepared.to_bytes().expect(\"facade publication\")"))
    else:
        native_timing = "sample(calls, || leaf_commit(family, leaf_source, route, import_input))"
        facade_timing = "sample(calls, || facade_commit(family, facade_source, route, import_input, record))"
        native_dispatch = f"leaf_commit {commit} arm: {native}"
        facade_dispatch = f"facade_commit {commit} arm: {facade}"
        phase_assertions = (("phase_once", "route.is_commit()"), ("phase_once", "leaf_commit(family, leaf_source, route, import_input)"), ("phase_once", "facade_commit(family, facade_source, route, import_input, record)"))
    semantic = {"set": "family.paths[family.paths.len() / 2].as_str()", "batch": "family.paths[0].as_str()", "import": '"/sheet[1]/A1"', "merge": "family.marker", "refresh": '"/body/p[1001]"'}[group]
    semantic_assertion = marker_group if group != "merge" else "family.marker"
    assertions = (("routes_for", f"Route::{enum}"), ("leaf_commit", f"Route::{commit} | Route::{publication}"), ("leaf_commit", native), ("facade_commit", f"Route::{commit} | Route::{publication}"), ("facade_commit", facade), ("semantic_marker", semantic_assertion), ("expected_semantic", f"Route::{commit}"), ("validation", "route.changed()")) + phase_assertions
    return {
        "phase": "publication" if is_publication else "commit", "operation_count": 1, "changed": changed, "exact_noop": False,
        "native_timing": native_timing, "facade_timing": facade_timing,
        "native_publication": "leaf_after.publish() in validation()", "facade_publication": "facade_after.to_bytes().expect(\"facade validation bytes\") in validation()",
        "native_dispatch": native_dispatch, "facade_dispatch": facade_dispatch, "semantic_marker": semantic,
        "adapter_api": {"result_type": api, "old_accessor": ".file", "consumer_transform": ".into_file() -> .file"}, "assertions": assertions,
    }


def _assert_route_source(source: str, assertions: tuple[tuple[str, str], ...], suffix: str) -> list[dict[str, str]]:
    checked: list[dict[str, str]] = []
    for function, needle in assertions:
        segment = provenance.function_segment(source, function)
        if needle not in segment:
            raise fail(f"route {suffix} source binding missing {function}() marker: {needle}")
        checked.append({"function": function, "contains": needle, "segment_sha256": digest_bytes(segment.encode())})
    return checked


def route_map(probe: Path) -> dict[str, Any]:
    source = provenance.read_text(probe, "corrected probe")
    try:
        descriptor = provenance.fixture_descriptor(probe)
        routes = provenance.fixture_routes(descriptor)
    except provenance.ProvenanceError as error:
        raise fail(f"corrected probe is not the reviewed 109-route source: {error}") from error
    function_names = ("family_names", "routes_for", "family", "leaf_commit", "facade_commit", "phase_once", "semantic_marker", "expected_semantic", "validation")
    fingerprints = {name: digest_bytes(provenance.function_segment(source, name).encode()) for name in function_names}
    rows: list[dict[str, Any]] = []
    for family in descriptor["families"]:
        for suffix in descriptor["routes"][family]:
            binding = _route_binding(suffix)
            checked = _assert_route_source(source, binding["assertions"], f"{family}/{suffix}")
            rows.append({
                "index": len(rows), "family": family, "route": f"{family}/{suffix}", "fixture": descriptor["recipe"][family],
                "route_enum": binding["assertions"][0][1], "phase": binding["phase"], "operation_count": binding["operation_count"],
                "changed": binding["changed"], "exact_noop": binding["exact_noop"], "native_timing": binding["native_timing"],
                "facade_timing": binding["facade_timing"], "native_publication": binding["native_publication"], "facade_publication": binding["facade_publication"],
                "native_dispatch": binding["native_dispatch"], "facade_dispatch": binding["facade_dispatch"], "semantic_marker": binding["semantic_marker"],
                "adapter_api": binding["adapter_api"], "source_assertions": checked,
            })
    if [row["route"] for row in rows] != routes:
        raise fail("generated route map does not preserve corrected probe route order")
    if len(rows) != 109 or len({row["route"] for row in rows}) != 109:
        raise fail("generated route map is not a complete unique 109-route matrix")
    return {
        "schema": "litchi_w8_original_baseline_route_map.v2", "status": "prepared_route_binding_only", "baseline_commit": BASELINE_COMMIT, "probe_source_commit": PROBE_SOURCE_COMMIT,
        "probe": {"path": PROBE_RELATIVE.as_posix(), "bytes": REVIEWED_PROBE_BYTES, "sha256": REVIEWED_PROBE_SHA256}, "families": descriptor["families"], "route_count": len(rows), "routes": rows,
        "function_fingerprints": fingerprints,
        "provenance": {"production_source_unchanged": True, "no_workload_substitution": True, "native_compile": "pending", "native_validation": "pending", "performance_qualification": "pending"},
    }


def prepare(args: argparse.Namespace) -> dict[str, Any]:
    root = args.git_root.resolve()
    tree = git_tree(root)
    history = verify_history(root, tree)
    workspace = args.workspace.resolve()
    probe = workspace / PROBE_RELATIVE
    adapter = workspace / ADAPTER_RELATIVE
    lock = workspace / LOCK_RELATIVE
    if args.probe is not None and args.probe.resolve() != probe:
        raise fail(f"probe must be {PROBE_RELATIVE} inside the workspace")
    if args.lock.resolve() != lock:
        raise fail(f"lock must be {LOCK_RELATIVE} inside the workspace")
    if args.lock_source_commit != BASELINE_COMMIT:
        raise fail("the pinned lock must declare f2 as its source commit")
    lock_origin = args.lock_origin.strip()
    if not lock_origin:
        raise fail("lock origin must explicitly describe the external f2 resolution")
    if "f2" not in lock_origin.lower() and BASELINE_COMMIT not in lock_origin:
        raise fail("lock origin must identify the pristine f2 dependency resolution")
    corrected = object_bytes(root, PROBE_SOURCE_COMMIT, PROBE_RELATIVE)
    if not probe.is_file() or probe.read_bytes() != corrected:
        raise fail("workspace corrected probe is not the reviewed d39 source")
    adapter_bytes, transforms = consumer_adapter(corrected.decode("utf-8"))
    archive_info = verify_source_archive(args.source_archive.resolve(), tree)
    adapter_artifact, adapter_created = ensure_adapter(adapter, adapter_bytes)
    workspace_info = verify_workspace(workspace, tree, probe, adapter, lock, adapter_bytes)
    output_dir = external_output(args.output_dir, "output directory")
    output_dir.mkdir(parents=True, exist_ok=True)
    descriptor = provenance.fixture_descriptor(probe)
    fixture_artifact = write_new(output_dir / "litchi_w8_fixture_descriptor.json", descriptor)
    mapped = route_map(probe)
    mapped["consumer_adapter"] = {
        "path": ADAPTER_RELATIVE.as_posix(),
        "artifact": adapter_artifact,
        "transforms": transforms,
        "source_sha256": digest_bytes(adapter_bytes),
        "production_source_unchanged": True,
    }
    route_artifact = write_new(output_dir / "original-baseline-route-map.json", mapped)
    patch_artifact = write_bytes_new(
        output_dir / "original-baseline-consumer-adapter.patch",
        adapter_patch(corrected, adapter_bytes),
    )
    source_assembly = provenance.source_assembly(
        argparse.Namespace(
            source=args.source_archive.resolve(),
            source_commit=BASELINE_COMMIT,
            overlay_kind="compatibility",
            overlay_base_commit=BASELINE_COMMIT,
            overlay_patch=Path(patch_artifact["path"]),
            description="standalone corrected-probe and old-API consumer additions; f2 production tree unchanged",
        )
    )
    source_assembly_artifact = write_new(output_dir / "original-baseline-source-assembly.json", source_assembly)
    projection_args = argparse.Namespace(workspace=workspace, package_manifest=workspace / provenance.PACKAGE_MANIFEST_RELATIVE, lock=lock, toolchain=workspace / "rust-toolchain.toml", probe=probe, fixture=Path(fixture_artifact["path"]), features=[], all_features=True, target=args.target, profile=args.profile)
    projection = provenance.dependency_projection(projection_args)
    selected_features = set(projection["features"]["selected"])
    if not set(W8_FEATURES).issubset(selected_features):
        raise fail("all-features projection omitted a format required by the corrected probe")
    projection_artifact = write_new(output_dir / "litchi_w8_dependency_projection.json", projection)
    assembly = {
        "schema": "litchi_w8_original_baseline_overlay.v2", "status": "prepared_consumer_adapter_only", "qualification": "not_run",
        "baseline": {"commit": BASELINE_COMMIT, "source_role": "pristine_f2_before_leg", "production_tree": {"files": history["tree_files"], "sha256": history["tree_sha256"], "source_archive": archive_info["artifact"], "source_archive_prefix": archive_info["prefix"], "production_delta": {"files_changed": 0, "api_overlay": False}}},
        "consumer_adapter": {"path": ADAPTER_RELATIVE.as_posix(), "artifact": adapter_artifact, "created": adapter_created, "source_probe": {"path": PROBE_RELATIVE.as_posix(), "source_commit": PROBE_SOURCE_COMMIT, "artifact": workspace_info["probe"]}, "transforms": transforms, "production_source_unchanged": True, "old_api_accessors": ["Commit.file", "BatchCommit.file"]},
        "workspace": workspace_info,
        "source_assembly": source_assembly_artifact,
        "adapter_patch": patch_artifact,
        "lock": {"path": LOCK_RELATIVE.as_posix(), "artifact": workspace_info["lock"], "source_commit": args.lock_source_commit, "origin": lock_origin, "intentional_external_delta": True},
        "fixture": fixture_artifact, "route_map": route_artifact, "dependency_projection": projection_artifact,
        "build_selection": {"features": projection["features"]["selected"], "all_features": True, "target": args.target, "profile": args.profile, "jobs": 1},
        "provenance_rules": {"no_production_diff": True, "no_workload_substitution": True, "native_compile": "pending", "native_validation": "pending", "performance_qualification": "pending", "acceptance_claim": False},
    }
    assembly_artifact = write_new(output_dir / "original-baseline-assembly.json", assembly)
    return {"assembly": assembly_artifact, "source_assembly": source_assembly_artifact, "adapter_patch": patch_artifact, "adapter": adapter_artifact, "route_map": route_artifact, "fixture": fixture_artifact, "dependency_projection": projection_artifact, "route_count": mapped["route_count"], "production_files": history["tree_files"], "adapter_created": adapter_created, "baseline_commit": BASELINE_COMMIT, "probe_sha256": REVIEWED_PROBE_SHA256}


def parser() -> argparse.ArgumentParser:
    command = argparse.ArgumentParser(description=__doc__.split("\n\n", 1)[0])
    command.add_argument("command", nargs="?", default="prepare", choices=("prepare",))
    command.add_argument("--git-root", required=True, type=Path)
    command.add_argument("--workspace", required=True, type=Path)
    command.add_argument("--probe", type=Path)
    command.add_argument("--lock", required=True, type=Path)
    command.add_argument("--lock-source-commit", required=True)
    command.add_argument("--lock-origin", required=True)
    command.add_argument("--source-archive", required=True, type=Path)
    command.add_argument("--output-dir", required=True, type=Path)
    command.add_argument("--target", default=TARGET)
    command.add_argument("--profile", default=PROFILE)
    return command


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    prepare(args)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OverlayError, provenance.ProvenanceError) as error:
        print(f"w8 original baseline: {error}", file=sys.stderr)
        raise SystemExit(2)
