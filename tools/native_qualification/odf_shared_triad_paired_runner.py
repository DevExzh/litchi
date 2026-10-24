#!/usr/bin/env python3
"""Run the paired ODF shared-triad qualification protocol.

The runner has two deliberately separate phases.  ``build`` prepares one
normal and one allocator-instrumented binary for the exact baseline overlay
and the selected candidate.  ``abba`` reuses those hashed binaries and runs
the route-specific probe in baseline/candidate/candidate/baseline order, with
fixed warmups and samples.  The baseline is not made into a candidate by
copying production files: its only allowed files outside the 173dd tree are
the reviewed probe adapter and the separately pinned Cargo.lock recorded by
the preparation receipt.

This is orchestration and validation only.  It makes no acceptance claim.
All output, logs, targets and temporary directories are external to the
source worktrees and every destination is fresh or rejected as stale.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import os
from pathlib import Path
import resource
import signal
import subprocess
import sys
import time
from typing import Any, Mapping, Sequence


TOOLS_ROOT = Path(__file__).resolve().parents[1]
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

import native_qualification_runner as generic  # noqa: E402
from native_qualification import odf_shared_triad_runner as shared  # noqa: E402


SCHEMA = "litchi.odf.shared-triad-qualification.paired-runner.v1"
PREPARATION_SCHEMA = "litchi.odf.shared-triad-qualification.preparation.v1"
ROLES = ("baseline", "candidate")
ABBA_ORDER = ("baseline", "candidate", "candidate", "baseline")
PHASES = ("normal", "allocator")
NORMAL_MODE = "timed"
ALLOCATOR_MODE = "allocator"
WARMUPS = 3
SAMPLES = 30
PROBE_ITERATIONS = 128
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_LIMIT_BYTES = 20_000_000_000
MEMORY_HARD_LIMIT_BYTES = 20_000_000_000
PROBE_MAX_BYTES = 1_048_576
PROBE_PATH = Path("crates/litchi-ops/examples/odf_shared_triad_qualification.rs")
LOCK_PATH = Path("Cargo.lock")
PACKAGE_MANIFEST = Path("crates/litchi-ops/Cargo.toml")
FEATURES = ("odc", "odi", "oth", "odm")
ALLOCATOR_COUNTERS = (
    "allocations",
    "deallocations",
    "allocated_bytes",
    "deallocated_bytes",
    "reallocations",
)


class RunnerError(RuntimeError):
    """A fail-closed manifest, source, process, or receipt error."""


def _build_keys() -> tuple[str, ...]:
    """Return the role/variant order used by both build and receipt phases."""

    return tuple(
        f"{role}/{variant}"
        for variant in ("normal", "allocator")
        for role in ROLES
    )


def _validate_build_event_journal(journal: Any) -> dict[str, Mapping[str, Any]]:
    expected_keys = _build_keys()
    if not isinstance(journal, Mapping) or journal.get("count") != len(expected_keys):
        raise RunnerError("build receipt event journal is incomplete")
    events = journal.get("events")
    if not isinstance(events, list) or len(events) != len(expected_keys):
        raise RunnerError("build receipt event journal events are incomplete")
    journal_hash = _sha256_bytes(json.dumps(events, sort_keys=True, separators=(",", ":")).encode())
    if journal.get("sha256") != journal_hash:
        raise RunnerError("build receipt event journal hash differs")
    expected_event_ids = [f"build-{key.replace('/', '-')}" for key in expected_keys]
    actual_event_ids = [event.get("event_id") for event in events if isinstance(event, Mapping)]
    if actual_event_ids != expected_event_ids:
        raise RunnerError("build receipt event journal order differs")
    events_by_id = {event.get("event_id"): event for event in events if isinstance(event, Mapping)}
    if len(events_by_id) != len(expected_keys):
        raise RunnerError("build receipt event journal has duplicate or malformed event IDs")
    return events_by_id


def _now() -> str:
    return _dt.datetime.now(_dt.timezone.utc).isoformat(timespec="milliseconds")


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _resolve_executable(name: str, cwd: Path, env: Mapping[str, str]) -> Path:
    """Resolve a child executable using the child's cwd and PATH."""

    path_value = env.get("PATH", os.defpath)
    for directory_name in path_value.split(os.pathsep):
        directory = Path(directory_name or ".")
        if not directory.is_absolute():
            directory = cwd / directory
        candidate = directory / name
        if candidate.is_file() and os.access(candidate, os.X_OK):
            return candidate.resolve(strict=True)
    raise RunnerError(f"{name} executable is not available in the build PATH")


def _sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _resolve(value: str | os.PathLike[str], *, base: Path | None = None) -> Path:
    path = Path(value)
    if not path.is_absolute():
        if base is None:
            raise RunnerError(f"path must be absolute: {value!r}")
        path = base / path
    probe = Path(path.anchor)
    for part in path.parts[1:]:
        probe /= part
        if probe.is_symlink():
            raise RunnerError(f"symlink path component is forbidden: {path}")
    path = path.expanduser().resolve(strict=False)
    if "3rd" + "party" in path.parts:
        raise RunnerError(f"untracked upstream checkout path is forbidden: {path}")
    return path


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _assert_disjoint(paths: Mapping[str, Path]) -> None:
    entries = list(paths.items())
    for index, (left_name, left) in enumerate(entries):
        for right_name, right in entries[index + 1 :]:
            if left == right or _under(left, right) or _under(right, left):
                raise RunnerError(
                    f"paths must be pairwise disjoint: {left_name}={left}, {right_name}={right}"
                )


def _reject_symlink_ancestors(path: Path, label: str) -> None:
    current = path.parent
    ancestors = []
    while True:
        ancestors.append(current)
        if current == current.parent:
            break
        current = current.parent
    for ancestor in reversed(ancestors):
        if ancestor.is_symlink():
            raise RunnerError(f"{label} ancestor may not be a symlink: {ancestor}")


def _directory_size(path: Path) -> int:
    return generic._directory_size(path)


def _git(root: Path, *args: str) -> str:
    try:
        return shared._git(root, *args)
    except shared.RunnerError as error:
        raise RunnerError(str(error)) from error


def _load_json(path: Path, label: str) -> tuple[dict[str, Any], str, int]:
    try:
        raw = path.read_bytes()
        data = json.loads(raw)
    except (OSError, json.JSONDecodeError) as error:
        raise RunnerError(f"cannot read {label} {path}: {error}") from error
    if not isinstance(data, dict):
        raise RunnerError(f"{label} root must be an object")
    return data, _sha256_bytes(raw), len(raw)


def _require_sha(value: Any, label: str) -> str:
    if not isinstance(value, str) or not shared.FULL_SHA.fullmatch(value):
        raise RunnerError(f"{label} must be a full 40-character lowercase commit SHA")
    return value


def _require_hash(value: Any, label: str) -> str:
    if not isinstance(value, str) or not shared.HASH.fullmatch(value):
        raise RunnerError(f"{label} must be a full 64-character lowercase SHA-256")
    return value


def validate_manifest(data: Mapping[str, Any]) -> dict[str, Any]:
    """Validate the shared manifest and its exact paired contract."""

    try:
        manifest = shared.validate_manifest(data)
    except shared.RunnerError as error:
        raise RunnerError(str(error)) from error
    paired = manifest.get("paired")
    if not isinstance(paired, Mapping):
        raise RunnerError("manifest.paired is required for the paired runner")
    if paired.get("abba_order") != list(ABBA_ORDER):
        raise RunnerError(f"paired ABBA order must be {list(ABBA_ORDER)!r}")
    if paired.get("warmup_iterations_per_leg") != WARMUPS:
        raise RunnerError(f"paired warmups must remain {WARMUPS}")
    if paired.get("samples_per_leg") != SAMPLES:
        raise RunnerError(f"paired samples must remain {SAMPLES}")
    if paired.get("probe_iterations") != PROBE_ITERATIONS:
        raise RunnerError(f"paired probe iterations must remain {PROBE_ITERATIONS}")
    if paired.get("families") != list(shared.FAMILIES):
        raise RunnerError("paired families differ from the reviewed route set")
    if paired.get("phases") != list(PHASES):
        raise RunnerError("paired phases must be normal and allocator")
    calibration = paired.get("calibration")
    if not isinstance(calibration, Mapping):
        raise RunnerError("paired.calibration is required")
    if calibration.get("mode") != "validate" or calibration.get("iterations") != 1:
        raise RunnerError("paired calibration must be one untimed validate iteration")
    if calibration.get("roles") != list(ROLES) or calibration.get("families") != list(shared.FAMILIES):
        raise RunnerError("paired calibration endpoint/family set differs")
    quiet = paired.get("quiet")
    if not isinstance(quiet, Mapping):
        raise RunnerError("paired.quiet is required")
    if quiet.get("process_guard") != "tools.native_qualification_runner.ProcessGuard":
        raise RunnerError("paired quiet process guard differs from the reviewed runner")
    if quiet.get("cpu_affinity_required") is not True:
        raise RunnerError("paired runner requires pinned CPU affinity")
    if quiet.get("resource_monitor") != "loadavg+proc-descendants+target-bytes":
        raise RunnerError("paired resource monitor contract differs")
    allocator = paired.get("allocator")
    if not isinstance(allocator, Mapping):
        raise RunnerError("paired.allocator is required")
    if allocator.get("scope") != "operation_global_system_allocator":
        raise RunnerError("allocator scope is not operation-global system allocator")
    if allocator.get("source_cfg") != "odf_shared_triad_allocator":
        raise RunnerError("allocator source cfg differs")
    if allocator.get("elapsed_claimable") is not False:
        raise RunnerError("allocator elapsed values must be non-claimable")
    if allocator.get("operation_metrics") != list(ALLOCATOR_COUNTERS):
        raise RunnerError("allocator operation metrics differ")
    overlay = paired.get("baseline_overlay")
    if not isinstance(overlay, Mapping):
        raise RunnerError("paired.baseline_overlay is required")
    if overlay.get("preparation_schema") != PREPARATION_SCHEMA:
        raise RunnerError("baseline preparation schema differs")
    if overlay.get("allowed_paths") != [str(LOCK_PATH), str(PROBE_PATH)]:
        raise RunnerError("baseline overlay paths differ")
    if overlay.get("production_tree") != "baseline HEAD tree must remain unchanged":
        raise RunnerError("baseline production-tree overlay policy differs")
    return manifest


def _status_paths(status: str) -> list[str]:
    result = []
    for line in status.splitlines():
        if len(line) < 4:
            raise RunnerError(f"malformed git status line: {line!r}")
        path = line[3:]
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        result.append(path)
    return result


def _production_tree(root: Path) -> str:
    return _sha256_bytes(_git(root, "ls-tree", "-r", "--full-tree", "HEAD").encode())


def _cargo_config_paths(root: Path) -> list[Path]:
    try:
        config_paths = list(generic._cargo_config_files(root))
    except generic.RunnerError as error:
        raise RunnerError(str(error)) from error
    home_cargo = Path.home() / ".cargo"
    for name in ("config.toml", "config"):
        candidate = home_cargo / name
        if candidate.exists() or candidate.is_symlink():
            config_paths.append(candidate)
    unique_paths = []
    for path in config_paths:
        path = _resolve(path)
        if path not in unique_paths:
            unique_paths.append(path)
    return sorted(unique_paths)


def _stat_identity(path: Path) -> dict[str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"expected a regular file: {path}")
    stat = path.stat()
    return {
        "device": stat.st_dev,
        "inode": stat.st_ino,
        "mode": stat.st_mode,
        "bytes": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
    }


def _cargo_config_identity(root: Path) -> dict[str, Any]:
    config_paths = _cargo_config_paths(root)
    files = []
    digest = hashlib.sha256()
    for path in config_paths:
        file_hash, file_bytes = _sha256_file(path)
        digest.update(str(path).encode())
        digest.update(b"\0")
        digest.update(file_hash.encode())
        digest.update(b"\0")
        digest.update(str(file_bytes).encode())
        files.append(
            {
                "path": str(path),
                "sha256": file_hash,
                "bytes": file_bytes,
                "stat": _stat_identity(path),
            }
        )
    return {"sha256": digest.hexdigest(), "files": files}


def _tracked_closure(root: Path) -> dict[str, Any]:
    sparse_config = subprocess.run(
        ["git", "-C", str(root), "config", "--bool", "core.sparseCheckout"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if sparse_config.returncode not in (0, 1) or sparse_config.stderr.strip():
        raise RunnerError(f"cannot inspect sparse-checkout configuration: {sparse_config.stderr.strip()}")
    sparse = sparse_config.stdout.strip().lower()
    if sparse == "true":
        raise RunnerError(f"sparse-checkout is enabled for source endpoint: {root}")
    head_records = _git(root, "ls-tree", "-r", "-z", "--full-tree", "HEAD").split("\0")
    head_paths: set[str] = set()
    for record in head_records:
        if not record:
            continue
        header, separator, relative = record.partition("\t")
        if not separator:
            raise RunnerError(f"malformed HEAD tree record: {record!r}")
        mode = header.split(" ", 1)[0]
        if mode not in {"100644", "100755"}:
            raise RunnerError(f"source closure contains unsupported tracked mode {mode}: {relative}")
        head_paths.add(relative)
    status_raw = _git(root, "ls-files", "-v", "-z")
    tracked_raw = _git(root, "ls-files", "--stage", "-z")
    status_records = status_raw.split("\0")
    tracked_records = tracked_raw.split("\0")
    files = []
    digest = hashlib.sha256()
    seen = set()
    for record in status_records:
        if not record:
            continue
        marker, separator, relative = record.partition(" ")
        if not separator or marker != "H":
            raise RunnerError(f"source endpoint has sparse/skip-worktree/assume-unchanged entry: {record!r}")
        seen.add(relative)
    for record in tracked_records:
        if not record:
            continue
        head, separator, relative = record.partition("\t")
        if not separator:
            raise RunnerError(f"malformed tracked-file record: {record!r}")
        mode = head.split(" ", 1)[0]
        if mode not in {"100644", "100755"}:
            raise RunnerError(f"source closure contains unsupported tracked mode {mode}: {relative}")
        if relative not in seen:
            raise RunnerError(f"tracked file is missing from full worktree status: {relative}")
        path = _resolve(relative, base=root)
        if path.is_symlink() or not path.is_file():
            raise RunnerError(f"tracked source file is absent or symlinked: {path}")
        file_hash, file_bytes = _sha256_file(path)
        digest.update(relative.encode())
        digest.update(b"\0")
        digest.update(mode.encode())
        digest.update(b"\0")
        digest.update(file_hash.encode())
        digest.update(b"\0")
        digest.update(str(file_bytes).encode())
        files.append({"path": relative, "mode": mode, "sha256": file_hash, "bytes": file_bytes})
    if seen != head_paths:
        missing = sorted(head_paths - seen)
        extra = sorted(seen - head_paths)
        raise RunnerError(f"source worktree is not a materialized full HEAD tree: missing={missing}, extra={extra}")
    if len(files) != len(seen):
        raise RunnerError("source worktree tracked-file closure is incomplete")
    return {
        "sha256": digest.hexdigest(),
        "file_count": len(files),
        "flags_sha256": _sha256_bytes(status_raw.encode()),
        "index_sha256": _sha256_bytes(tracked_raw.encode()),
    }


def _source_state(root: Path) -> dict[str, Any]:
    closure = _tracked_closure(root)
    configs = _cargo_config_identity(root)
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    diff = _git(root, "diff", "HEAD", "--no-ext-diff", "--binary")
    head = _git(root, "rev-parse", "HEAD").strip()
    tree = _git(root, "rev-parse", "HEAD^{tree}").strip()
    return {
        "root": str(root),
        "commit": head,
        "tree": tree,
        "production_tree_sha256": _production_tree(root),
        "status": status,
        "status_sha256": _sha256_bytes(status.encode()),
        "diff_sha256": _sha256_bytes(diff.encode()),
        "source_tree_sha256": _sha256_bytes(
            (tree + "\0" + status + "\0" + diff + "\0" + closure["sha256"] + "\0" + configs["sha256"]).encode()
        ),
        "status_paths": _status_paths(status),
        "tracked_closure": closure,
        "cargo_configs": configs,
    }


def _sparse_checkout_value(root: Path) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), "config", "--bool", "core.sparseCheckout"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode not in (0, 1) or result.stderr.strip():
        raise RunnerError(f"cannot inspect sparse-checkout configuration: {result.stderr.strip()}")
    return result.stdout.strip().lower()


def _git_diff_clean(root: Path, *args: str) -> bool:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode not in (0, 1):
        raise RunnerError(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.returncode == 0


def _cheap_source_snapshot(root: Path) -> dict[str, Any]:
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    tracked_flags = _git(root, "ls-files", "-v", "-z")
    tracked_index = _git(root, "ls-files", "--stage", "-z")
    config_stats = []
    for path in _cargo_config_paths(root):
        config_hash, config_bytes = _sha256_file(path)
        config_stats.append(
            {
                "path": str(path),
                "sha256": config_hash,
                "bytes": config_bytes,
                "stat": _stat_identity(path),
            }
        )
    file_stats = []
    for relative in (PACKAGE_MANIFEST, PROBE_PATH, LOCK_PATH):
        path = _resolve(relative, base=root)
        file_stats.append({"path": str(path), "stat": _stat_identity(path)})
    return {
        "head": _git(root, "rev-parse", "HEAD").strip(),
        "tree": _git(root, "rev-parse", "HEAD^{tree}").strip(),
        "status_sha256": _sha256_bytes(status.encode()),
        "status_paths": _status_paths(status),
        "worktree_clean": _git_diff_clean(root, "diff", "--quiet", "HEAD", "--"),
        "index_clean": _git_diff_clean(root, "diff", "--cached", "--quiet", "--"),
        "sparse_checkout": _sparse_checkout_value(root),
        "tracked_flags_sha256": _sha256_bytes(tracked_flags.encode()),
        "tracked_index_sha256": _sha256_bytes(tracked_index.encode()),
        "config_stats": config_stats,
        "file_stats": file_stats,
    }


def _package_identity(root: Path, manifest: Mapping[str, Any]) -> dict[str, Any]:
    path = _resolve(str(manifest["source"]["package_manifest"]), base=root)
    if path.relative_to(root).as_posix() != str(PACKAGE_MANIFEST):
        raise RunnerError("package manifest path is not the canonical litchi-ops path")
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"package manifest is not a regular file: {path}")
    try:
        import tomllib

        data = tomllib.loads(path.read_text(encoding="utf-8"))
        package = data["package"]
    except (OSError, KeyError, tomllib.TOMLDecodeError) as error:
        raise RunnerError(f"cannot read litchi-ops package manifest: {error}") from error
    if package.get("name") != "litchi-ops":
        raise RunnerError("canonical package manifest is not litchi-ops")
    digest, size = _sha256_file(path)
    return {"path": str(path), "sha256": digest, "bytes": size, "name": package["name"]}


def _file_identity(root: Path, relative: Path, label: str) -> dict[str, Any]:
    path = _resolve(relative, base=root)
    digest, size = _sha256_file(path)
    return {"path": str(path), "sha256": digest, "bytes": size, "label": label}


def _validate_preparation(
    preparation: Mapping[str, Any],
    *,
    preparation_hash: str,
    manifest: Mapping[str, Any],
    baseline_root: Path,
    candidate_root: Path,
    candidate_commit: str,
    candidate_lock_sha256: str,
) -> dict[str, Any]:
    if preparation.get("schema") != PREPARATION_SCHEMA or preparation.get("lane") != "odf-shared-triad":
        raise RunnerError("baseline preparation schema/lane differs")
    baseline = preparation.get("baseline")
    candidate = preparation.get("candidate")
    if not isinstance(baseline, Mapping) or not isinstance(candidate, Mapping):
        raise RunnerError("baseline preparation endpoints are missing")
    baseline_before = baseline.get("before")
    baseline_after = baseline.get("after_overlay")
    if not isinstance(baseline_before, Mapping) or not isinstance(baseline_after, Mapping):
        raise RunnerError("baseline preparation source states are missing")
    if _resolve(str(baseline_before.get("root"))) != baseline_root:
        raise RunnerError("preparation baseline root differs")
    if baseline_before.get("head") != manifest["baseline_commit"]:
        raise RunnerError("preparation baseline commit differs from manifest")
    if baseline_before.get("clean") is not True:
        raise RunnerError("preparation baseline was not clean before overlay")
    if baseline_after.get("tree") != baseline_before.get("tree"):
        raise RunnerError("preparation baseline tracked tree changed")
    if baseline.get("commit") != manifest["baseline_commit"]:
        raise RunnerError("preparation baseline commit record differs")
    overlay_paths = baseline.get("overlay_paths")
    if overlay_paths not in (
        [str(PROBE_PATH)],
        [str(LOCK_PATH), str(PROBE_PATH)],
        [str(PROBE_PATH), str(LOCK_PATH)],
    ):
        raise RunnerError("preparation overlay paths are not exactly probe plus Cargo.lock")
    if baseline.get("production_tree_sha256") != _production_tree(baseline_root):
        raise RunnerError("preparation production tree does not bind to the clean baseline")
    if _resolve(str(candidate.get("root"))) != candidate_root:
        raise RunnerError("preparation candidate root differs")
    if candidate.get("commit") != candidate_commit or candidate_commit == shared.HISTORICAL_CANDIDATE:
        raise RunnerError("preparation candidate commit is stale or differs from selected candidate")
    manifest_candidate = manifest.get("candidate_commit")
    if manifest_candidate is not None and manifest_candidate != candidate_commit:
        raise RunnerError("selected candidate differs from the manifest candidate")
    candidate_probe = candidate.get("probe")
    candidate_lock = candidate.get("lock")
    if not isinstance(candidate_probe, Mapping) or not isinstance(candidate_lock, Mapping):
        raise RunnerError("preparation candidate probe/lock identity is missing")
    _require_hash(candidate_probe.get("sha256"), "preparation candidate probe SHA")
    _require_hash(candidate_lock.get("sha256"), "preparation candidate lock SHA")
    if candidate_probe.get("sha256") != manifest["probe"]["sha256"]:
        raise RunnerError("preparation candidate probe differs from manifest")
    if candidate_lock.get("sha256") != candidate_lock_sha256:
        raise RunnerError("preparation candidate lock differs from selected identity")
    probe = baseline.get("probe")
    lock = baseline.get("lock")
    if not isinstance(probe, Mapping) or not isinstance(lock, Mapping):
        raise RunnerError("preparation baseline probe/lock overlay identity is missing")
    destination_probe = _resolve(str(probe.get("destination_path")))
    destination_lock = _resolve(str(lock.get("destination_path")))
    if destination_probe != _resolve(PROBE_PATH, base=baseline_root):
        raise RunnerError("preparation baseline probe destination is not canonical")
    if destination_lock != _resolve(LOCK_PATH, base=baseline_root):
        raise RunnerError("preparation baseline lock destination is not canonical")
    for mapping, label in ((probe, "preparation baseline probe"), (lock, "preparation baseline lock")):
        _require_hash(mapping.get("sha256"), f"{label} SHA")
        if mapping.get("bytes", 0) <= 0:
            raise RunnerError(f"{label} byte count is invalid")
    if probe.get("sha256") != manifest["probe"]["sha256"]:
        raise RunnerError("preparation baseline probe differs from manifest")
    if _resolve(str(probe.get("source_path"))) != _resolve(str(candidate_probe.get("path"))):
        raise RunnerError("preparation adapter source is not the selected candidate probe")
    if not _resolve(str(lock.get("source_path"))).is_file():
        raise RunnerError("preparation baseline lock source is missing")
    if not _resolve(str(probe.get("destination_path"))).is_file() or not _resolve(str(lock.get("destination_path"))).is_file():
        raise RunnerError("preparation overlay files are missing")
    observed_probe, probe_bytes = _sha256_file(destination_probe)
    observed_lock, lock_bytes = _sha256_file(destination_lock)
    if observed_probe != probe["sha256"] or probe_bytes != probe.get("bytes"):
        raise RunnerError("baseline probe overlay hash does not match preparation")
    if observed_lock != lock["sha256"] or lock_bytes != lock.get("bytes"):
        raise RunnerError("baseline lock overlay hash does not match preparation")
    return {
        "schema": preparation["schema"],
        "sha256": preparation_hash,
        "baseline_commit": baseline["commit"],
        "baseline_production_tree_sha256": baseline["production_tree_sha256"],
        "baseline_probe_sha256": probe["sha256"],
        "baseline_lock_sha256": lock["sha256"],
        "candidate_commit": candidate["commit"],
        "candidate_probe_sha256": candidate_probe["sha256"],
        "candidate_lock_sha256": candidate_lock["sha256"],
        "overlay_paths": sorted(overlay_paths),
    }


def _endpoint_identity(
    manifest: Mapping[str, Any],
    *,
    role: str,
    root: Path,
    candidate_commit: str,
    candidate_lock_sha256: str,
    preparation: Mapping[str, Any],
) -> dict[str, Any]:
    if role not in ROLES:
        raise RunnerError(f"unknown endpoint role: {role}")
    state = _source_state(root)
    package = _package_identity(root, manifest)
    probe = _file_identity(root, PROBE_PATH, "probe")
    lock = _file_identity(root, LOCK_PATH, "Cargo.lock")
    if role == "baseline":
        expected_commit = manifest["baseline_commit"]
        if state["commit"] != expected_commit:
            raise RunnerError(f"baseline HEAD {state['commit']} differs from {expected_commit}")
        if state["production_tree_sha256"] != preparation["baseline_production_tree_sha256"]:
            raise RunnerError("baseline production tree differs from preparation")
        if sorted(state["status_paths"]) != preparation["overlay_paths"]:
            raise RunnerError("baseline has paths outside the reviewed overlay")
        if probe["sha256"] != preparation["baseline_probe_sha256"]:
            raise RunnerError("baseline probe differs from preparation")
        if lock["sha256"] != preparation["baseline_lock_sha256"]:
            raise RunnerError("baseline lock differs from preparation")
        clean = False
    else:
        expected_commit = candidate_commit
        if state["commit"] != expected_commit:
            raise RunnerError(f"candidate HEAD {state['commit']} differs from selected {expected_commit}")
        if state["status_paths"]:
            raise RunnerError(f"candidate source worktree is dirty: {state['status']}")
        if probe["sha256"] != manifest["probe"]["sha256"]:
            raise RunnerError("candidate probe differs from manifest")
        if lock["sha256"] != candidate_lock_sha256:
            raise RunnerError("candidate lock differs from selected identity")
        clean = True
    return {
        "role": role,
        "root": str(root),
        "commit": state["commit"],
        "tree": state["tree"],
        "production_tree_sha256": state["production_tree_sha256"],
        "source_tree_sha256": state["source_tree_sha256"],
        "status_sha256": state["status_sha256"],
        "diff_sha256": state["diff_sha256"],
        "status_paths": state["status_paths"],
        "tracked_closure": state["tracked_closure"],
        "cargo_configs": state["cargo_configs"],
        "clean": clean,
        "package_manifest": package,
        "probe": probe,
        "lock": lock,
        "features": list(FEATURES),
        "routes": list(manifest["routes"]),
        "phases": {
            "setup_outside_clock": list(manifest["phases"]["setup_outside_clock"]),
            "inside_clock": list(manifest["phases"]["inside_clock"]),
        },
    }


def _proc_resources(pid: int) -> dict[str, int] | None:
    status = Path("/proc") / str(pid) / "status"
    try:
        values: dict[str, int] = {}
        for line in status.read_text(encoding="utf-8", errors="replace").splitlines():
            key, _, value = line.partition(":")
            if key in {"VmRSS", "VmHWM", "VmSize"}:
                number = value.strip().split()[0]
                values[key] = int(number) * 1024
        return values
    except (FileNotFoundError, PermissionError, OSError, ValueError, IndexError):
        return None


def _terminate(process: subprocess.Popen[Any]) -> None:
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except (ProcessLookupError, PermissionError):
        process.terminate()
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except (ProcessLookupError, PermissionError):
            process.kill()
        process.wait()


def _descendants_gone(pid: int) -> list[int]:
    deadline = time.monotonic() + 5
    remaining: set[int] = set()
    while time.monotonic() < deadline:
        snapshot = generic._proc_snapshot()
        descendants = generic._descendants(snapshot, pid)
        remaining = {
            child
            for child, entry in snapshot.items()
            if child != pid and (child in descendants or entry.get("pgrp") == pid)
        }
        if not remaining:
            return []
        time.sleep(0.1)
    return sorted(remaining)


class PairedRunner:
    def __init__(
        self,
        *,
        manifest_path: Path,
        preparation_path: Path,
        baseline_root: Path,
        candidate_root: Path,
        candidate_commit: str,
        candidate_lock_sha256: str,
        output: Path,
        target_root: Path,
        tmp_root: Path,
        cpu: int,
        phase: str,
    ) -> None:
        if phase not in {"build", "abba"}:
            raise RunnerError("phase must be build or abba")
        if not isinstance(cpu, int) or cpu < 0 or cpu >= (os.cpu_count() or 1):
            raise RunnerError("cpu must be an available non-negative CPU index")
        self.phase = phase
        self.manifest_path = _resolve(manifest_path)
        self.preparation_path = _resolve(preparation_path)
        self.baseline_root = _resolve(baseline_root)
        self.candidate_root = _resolve(candidate_root)
        self.candidate_commit = _require_sha(candidate_commit, "candidate_commit")
        if self.candidate_commit == shared.HISTORICAL_CANDIDATE:
            raise RunnerError("historical 15ab candidate is not admissible")
        self.candidate_lock_sha256 = _require_hash(candidate_lock_sha256, "candidate_lock_sha256")
        self.output = _resolve(output)
        self.target_root = _resolve(target_root)
        self.tmp_root = _resolve(tmp_root)
        self.cpu = cpu
        _assert_disjoint(
            {
                "baseline": self.baseline_root,
                "candidate": self.candidate_root,
                "output": self.output,
                "target_root": self.target_root,
                "tmp_root": self.tmp_root,
                "preparation": self.preparation_path,
            }
        )
        self.manifest, self.manifest_sha256, _ = _load_json(self.manifest_path, "manifest")
        self.manifest = validate_manifest(self.manifest)
        preparation, prep_hash, _ = _load_json(self.preparation_path, "preparation")
        self.preparation = _validate_preparation(
            preparation,
            preparation_hash=prep_hash,
            manifest=self.manifest,
            baseline_root=self.baseline_root,
            candidate_root=self.candidate_root,
            candidate_commit=self.candidate_commit,
            candidate_lock_sha256=self.candidate_lock_sha256,
        )
        self._endpoint_cache: dict[str, dict[str, Any]] = {}
        self.guard = generic.ProcessGuard()
        self.runner_hash, _ = _sha256_file(Path(__file__).resolve())
        if self.phase == "build":
            if self.output.exists():
                if self.output.is_symlink() or not self.output.is_dir() or any(self.output.iterdir()):
                    raise RunnerError("build output must be a fresh empty directory")
            else:
                self.output.mkdir(parents=True)
        else:
            if not self.output.is_dir() or self.output.is_symlink():
                raise RunnerError("ABBA output must be the existing build output directory")
            if (self.output / "build-receipt.json").is_symlink() or not (self.output / "build-receipt.json").is_file():
                raise RunnerError("ABBA requires the successful build receipt")
        self.target_root.mkdir(parents=True, exist_ok=True)
        self.tmp_root.mkdir(parents=True, exist_ok=True)
        if self.target_root.is_symlink() or self.tmp_root.is_symlink():
            raise RunnerError("target/tmp roots may not be symlinks")

    def _fresh_dir(self, path: Path, label: str) -> None:
        if _under(path, self.baseline_root) or _under(path, self.candidate_root):
            raise RunnerError(f"{label} must be outside source worktrees: {path}")
        _reject_symlink_ancestors(path, label)
        if path.exists():
            if path.is_symlink() or not path.is_dir() or any(path.iterdir()):
                raise RunnerError(f"{label} must be absent or empty: {path}")
        path.mkdir(parents=True, exist_ok=True)

    def _fresh_file(self, path: Path, label: str) -> None:
        if _under(path, self.baseline_root) or _under(path, self.candidate_root):
            raise RunnerError(f"{label} must be outside source worktrees: {path}")
        if path.is_symlink() or path.exists():
            raise RunnerError(f"{label} must not overwrite an existing path: {path}")
        _reject_symlink_ancestors(path, label)
        path.parent.mkdir(parents=True, exist_ok=True)

    def _target_path(self, role: str, variant: str) -> Path:
        return self.target_root / f"{role}-{variant}"

    def _tmp_path(self, role: str, variant: str) -> Path:
        return self.tmp_root / f"{role}-{variant}"

    def _environment(self, role: str, variant: str, target: Path, tmpdir: Path) -> dict[str, str]:
        allowed = {
            "AR",
            "CC",
            "CFLAGS",
            "CXX",
            "CXXFLAGS",
            "HOME",
            "LANG",
            "LD_LIBRARY_PATH",
            "LOGNAME",
            "MAKEFLAGS",
            "PATH",
            "PKG_CONFIG_PATH",
            "PKG_CONFIG_SYSROOT_DIR",
            "RANLIB",
            "RUSTUP_HOME",
            "RUSTUP_TOOLCHAIN",
            "SOURCE_DATE_EPOCH",
            "TEMP",
            "TMP",
            "USER",
        }
        env = {
            key: value
            for key, value in os.environ.items()
            if key in allowed or key.startswith("LC_")
        }
        env.update(
            {
                "CARGO_TARGET_DIR": str(target),
                "CARGO_BUILD_JOBS": "1",
                "CARGO_INCREMENTAL": "0",
                "TMPDIR": str(tmpdir),
                "LITCHI_MEM_MAX": "20G",
            }
        )
        if variant == "allocator":
            env["RUSTFLAGS"] = "--cfg odf_shared_triad_allocator"
        env["LITCHI_QUALIFICATION_ROLE"] = role
        env["LITCHI_QUALIFICATION_VARIANT"] = variant
        return env

    @staticmethod
    def _recorded_environment(env: Mapping[str, str]) -> dict[str, str]:
        # ``_environment`` deliberately constructs a minimal child
        # environment.  Record every value passed to the child so a receipt
        # cannot be replayed under a different linker, toolchain or locale.
        return {key: env[key] for key in sorted(env)}

    def _preexec(self) -> Any:
        cpu = self.cpu

        def configure() -> None:
            os.sched_setaffinity(0, {cpu})
            resource.setrlimit(resource.RLIMIT_AS, (MEMORY_HARD_LIMIT_BYTES, MEMORY_HARD_LIMIT_BYTES))

        return configure

    def _run_process(
        self,
        *,
        event_id: str,
        role: str,
        variant: str,
        argv: Sequence[str],
        cwd: Path,
        env: Mapping[str, str],
        target: Path,
        stdout_path: Path,
        stderr_path: Path,
    ) -> dict[str, Any]:
        self.guard.require_quiet()
        self._fresh_file(stdout_path, f"{event_id} stdout")
        self._fresh_file(stderr_path, f"{event_id} stderr")
        target_before = _directory_size(target)
        if target_before > TARGET_PREFLIGHT_BYTES:
            raise RunnerError(f"target exceeds preflight limit before {event_id}")
        start_ns = time.monotonic_ns()
        monitor: list[dict[str, Any]] = []
        violation: list[dict[str, Any]] = []
        returncode: int | None = None
        error: str | None = None
        process: subprocess.Popen[Any] | None = None
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            try:
                process = subprocess.Popen(
                    list(argv),
                    cwd=str(cwd),
                    env=dict(env),
                    stdin=subprocess.DEVNULL,
                    stdout=stdout,
                    stderr=stderr,
                    start_new_session=True,
                    preexec_fn=self._preexec(),
                )
                last_monitor = 0.0
                while True:
                    returncode = process.poll()
                    now = time.monotonic()
                    if now - last_monitor >= 0.25 or returncode is not None:
                        target_bytes = _directory_size(target)
                        monitor.append(
                            {
                                "utc": _now(),
                                "loadavg": self.guard.loadavg(),
                                "target_bytes": target_bytes,
                                "child": _proc_resources(process.pid),
                            }
                        )
                        last_monitor = now
                        if target_bytes > TARGET_HARD_LIMIT_BYTES:
                            error = f"target exceeded hard limit during {event_id}"
                            if process.poll() is None:
                                _terminate(process)
                            returncode = process.returncode
                            break
                    if returncode is not None:
                        break
                    offenders = self.guard.offenders(process.pid)
                    if offenders:
                        violation.extend(offenders)
                        error = "unrelated Litchi process appeared during child execution"
                        _terminate(process)
                        returncode = process.returncode
                        break
                    time.sleep(0.05)
            except BaseException as exc:
                error = str(exc)
                if process is not None and process.poll() is None:
                    _terminate(process)
                returncode = process.returncode if process is not None else None
        descendants = _descendants_gone(process.pid) if process is not None else []
        offenders = self.guard.offenders()
        if descendants:
            error = error or f"child descendants survived: {descendants}"
        if offenders:
            error = error or "unrelated Litchi process remained after child exit"
        target_after = _directory_size(target)
        if target_after > TARGET_HARD_LIMIT_BYTES:
            error = error or f"target exceeded hard limit after {event_id}"
        stdout_hash, stdout_bytes = _sha256_file(stdout_path)
        stderr_hash, stderr_bytes = _sha256_file(stderr_path)
        event = {
            "schema": SCHEMA,
            "event_id": event_id,
            "role": role,
            "variant": variant,
            "argv": list(argv),
            "cwd": str(cwd),
            "environment": self._recorded_environment(env),
            "cpu_affinity": self.cpu,
            "start_utc": _now(),
            "duration_ns": time.monotonic_ns() - start_ns,
            "returncode": returncode,
            "stdout": {"path": str(stdout_path), "sha256": stdout_hash, "bytes": stdout_bytes},
            "stderr": {"path": str(stderr_path), "sha256": stderr_hash, "bytes": stderr_bytes},
            "target": {"path": str(target), "bytes_before": target_before, "bytes_after": target_after},
            "monitor": monitor,
            "process_guard_violations": violation,
            "post_exit_descendants": descendants,
            "post_exit_offenders": offenders,
            "error": error,
        }
        if error:
            raise RunnerError(f"{error} during {event_id}; stderr={stderr_path}")
        if returncode != 0:
            raise RunnerError(f"child failed during {event_id} with exit {returncode}; stderr={stderr_path}")
        return event

    def _toolchain(self, cwd: Path, env: Mapping[str, str]) -> dict[str, Any]:
        values: dict[str, Any] = {}
        recorded_environment = self._recorded_environment(env)
        for command, label in ((["rustc", "-Vv"], "rustc"), (["cargo", "-V"], "cargo")):
            executable_path = _resolve_executable(command[0], cwd, env)
            if "3rd" + "party" in executable_path.parts or not executable_path.is_file():
                raise RunnerError(f"{label} executable is not a permitted regular file: {executable_path}")
            result = subprocess.run(
                command,
                cwd=str(cwd),
                env=dict(env),
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            if result.returncode:
                raise RunnerError(f"{label} identity command failed: {result.stderr.strip()}")
            executable_hash, executable_bytes = _sha256_file(executable_path)
            values[label] = {
                "argv": command,
                "cwd": str(cwd),
                "environment": recorded_environment,
                "returncode": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
                "executable": {
                    "path": str(executable_path),
                    "sha256": executable_hash,
                    "bytes": executable_bytes,
                },
            }
        return values

    def _endpoint(self, role: str, *, full: bool = False) -> dict[str, Any]:
        if not full and role in self._endpoint_cache:
            cached = self._endpoint_cache[role]
            current_snapshot = _cheap_source_snapshot(cached["root"])
            if current_snapshot != cached["cheap_snapshot"]:
                raise RunnerError(f"{role} cheap source identity changed between qualification probes")
            return cached["endpoint"]
        endpoint = _endpoint_identity(
            self.manifest,
            role=role,
            root=self.baseline_root if role == "baseline" else self.candidate_root,
            candidate_commit=self.candidate_commit,
            candidate_lock_sha256=self.candidate_lock_sha256,
            preparation=self.preparation,
        )
        self._endpoint_cache[role] = {
            "root": Path(endpoint["root"]),
            "endpoint": endpoint,
            "cheap_snapshot": _cheap_source_snapshot(Path(endpoint["root"])),
        }
        return endpoint

    def _assert_endpoint_unchanged(
        self,
        role: str,
        expected: Mapping[str, Any],
        *,
        full: bool = False,
    ) -> dict[str, Any]:
        actual = self._endpoint(role, full=full)
        for key in (
            "commit",
            "tree",
            "production_tree_sha256",
            "source_tree_sha256",
            "status_sha256",
            "diff_sha256",
            "tracked_closure",
            "cargo_configs",
            "probe",
            "lock",
            "features",
            "routes",
            "phases",
        ):
            if actual.get(key) != expected.get(key):
                raise RunnerError(f"{role} endpoint identity changed: {key}")
        return actual

    @staticmethod
    def _build_command() -> list[str]:
        return [
            "cargo",
            "build",
            "--offline",
            "--locked",
            "-p",
            "litchi-ops",
            "--example",
            PROBE_PATH.stem,
            "--features",
            ",".join(FEATURES),
        ]

    def _binary_path(self, role: str, variant: str) -> Path:
        return self._target_path(role, variant) / "debug" / "examples" / PROBE_PATH.stem

    def _build(self, role: str, variant: str) -> dict[str, Any]:
        root = self.baseline_root if role == "baseline" else self.candidate_root
        target = self._target_path(role, variant)
        tmpdir = self._tmp_path(role, variant)
        self._fresh_dir(target, f"{role}/{variant} target")
        self._fresh_dir(tmpdir, f"{role}/{variant} tmpdir")
        endpoint = self._endpoint(role)
        env = self._environment(role, variant, target, tmpdir)
        toolchain = self._toolchain(root, env)
        binary = self._binary_path(role, variant)
        command = self._build_command()
        stdout = self.output / "build" / f"{role}-{variant}.stdout"
        stderr = self.output / "build" / f"{role}-{variant}.stderr"
        event = self._run_process(
            event_id=f"build-{role}-{variant}",
            role=role,
            variant=variant,
            argv=command,
            cwd=root,
            env=env,
            target=target,
            stdout_path=stdout,
            stderr_path=stderr,
        )
        if binary.is_symlink() or not binary.is_file() or not os.access(binary, os.X_OK):
            raise RunnerError(f"built probe binary is missing or not executable: {binary}")
        binary_hash, binary_bytes = _sha256_file(binary)
        event["binary"] = {
            "path": str(binary),
            "sha256": binary_hash,
            "bytes": binary_bytes,
        }
        event["toolchain"] = toolchain
        endpoint_after = self._assert_endpoint_unchanged(role, endpoint, full=True)
        if _directory_size(target) > TARGET_PREFLIGHT_BYTES:
            raise RunnerError(f"{role}/{variant} target exceeded the 19.2 GB preflight limit")
        return {
            "role": role,
            "variant": variant,
            "command": command,
            "environment": self._recorded_environment(env),
            "toolchain": dict(toolchain),
            "source": endpoint_after,
            "binary": {"path": str(binary), "sha256": binary_hash, "bytes": binary_bytes},
            "target": {"path": str(target), "bytes": _directory_size(target)},
            "tmpdir": str(tmpdir),
            "event": event,
        }

    def build(self) -> None:
        endpoints: dict[str, Any] = {}
        builds: dict[str, Any] = {}
        receipt_path = self.output / "build-receipt.json"
        try:
            for role in ROLES:
                endpoints[role] = self._endpoint(role, full=True)
            for variant in ("normal", "allocator"):
                for role in ROLES:
                    builds[f"{role}/{variant}"] = self._build(role, variant)
            journal_events = [builds[key]["event"] for key in _build_keys()]
            receipt = {
                "schema": SCHEMA,
                "phase": "build",
                "status": "passed",
                "manifest": {"path": str(self.manifest_path), "sha256": self.manifest_sha256},
                "preparation": {"path": str(self.preparation_path), **self.preparation},
                "runner": {"path": str(Path(__file__).resolve()), "sha256": self.runner_hash},
                "toolchain": {key: builds[key]["toolchain"] for key in _build_keys()},
                "cpu_affinity": self.cpu,
                "features": list(FEATURES),
                "routes": list(self.manifest["routes"]),
                "phases": dict(self.manifest["paired"]),
                "endpoints": endpoints,
                "builds": builds,
                "event_journal": {
                    "count": len(journal_events),
                    "sha256": _sha256_bytes(
                        json.dumps(journal_events, sort_keys=True, separators=(",", ":")).encode()
                    ),
                    "events": journal_events,
                },
                "claim_scope": "build identity only; no native acceptance claim",
            }
            self._write_receipt(receipt_path, receipt)
        except BaseException as error:
            failure = {
                "schema": SCHEMA,
                "phase": "build",
                "status": "failed",
                "manifest_sha256": self.manifest_sha256,
                "preparation": self.preparation,
                "cpu_affinity": self.cpu,
                "endpoints": endpoints,
                "builds": builds,
                "error": str(error),
                "claim_scope": "failed build identity only; no native acceptance claim",
            }
            if not receipt_path.exists():
                self._write_receipt(receipt_path, failure)
            raise

    def _load_build_receipt(self) -> dict[str, Any]:
        receipt_path = self.output / "build-receipt.json"
        receipt, receipt_hash, _ = _load_json(receipt_path, "build receipt")
        if receipt.get("schema") != SCHEMA or receipt.get("phase") != "build" or receipt.get("status") != "passed":
            raise RunnerError("build receipt is not a passed paired build")
        manifest_record = receipt.get("manifest")
        preparation_record = receipt.get("preparation")
        runner_record = receipt.get("runner")
        if not isinstance(manifest_record, Mapping) or manifest_record.get("sha256") != self.manifest_sha256:
            raise RunnerError("build receipt manifest identity differs")
        if not isinstance(preparation_record, Mapping) or preparation_record.get("sha256") != self.preparation["sha256"]:
            raise RunnerError("build receipt preparation identity differs")
        if not isinstance(runner_record, Mapping) or runner_record.get("sha256") != self.runner_hash:
            raise RunnerError("paired runner changed after build")
        if receipt.get("cpu_affinity") != self.cpu or receipt.get("features") != list(FEATURES):
            raise RunnerError("build receipt CPU/features differ")
        if receipt.get("routes") != list(self.manifest["routes"]):
            raise RunnerError("build receipt route descriptors differ")
        if receipt.get("phases") != dict(self.manifest["paired"]):
            raise RunnerError("build receipt paired protocol differs")
        if receipt.get("claim_scope") != "build identity only; no native acceptance claim":
            raise RunnerError("build receipt claim scope is not non-acceptance")
        builds = receipt.get("builds")
        expected_keys = set(_build_keys())
        if not isinstance(builds, Mapping) or set(builds) != expected_keys:
            raise RunnerError("build receipt does not contain exactly four role/variant builds")
        endpoints = receipt.get("endpoints")
        if not isinstance(endpoints, Mapping) or set(endpoints) != set(ROLES):
            raise RunnerError("build receipt endpoint identities are incomplete")
        events_by_id = _validate_build_event_journal(receipt.get("event_journal"))
        toolchains = receipt.get("toolchain")
        if not isinstance(toolchains, Mapping) or set(toolchains) != expected_keys:
            raise RunnerError("build receipt per-endpoint toolchain identities are incomplete")
        for role in ROLES:
            current = self._endpoint(role, full=True)
            if endpoints[role] != current:
                raise RunnerError(f"build receipt endpoint identity changed for {role}")
            for variant in ("normal", "allocator"):
                key = f"{role}/{variant}"
                build = builds[key]
                if not isinstance(build, Mapping) or build.get("role") != role or build.get("variant") != variant:
                    raise RunnerError(f"build receipt role/variant identity is malformed for {key}")
                if build.get("source") != current:
                    raise RunnerError(f"build receipt source endpoint differs for {key}")
                target = self._target_path(role, variant)
                tmpdir = self._tmp_path(role, variant)
                expected_binary = self._binary_path(role, variant)
                expected_command = self._build_command()
                build_env = self._environment(role, variant, target, tmpdir)
                expected_env = self._recorded_environment(build_env)
                current_toolchain = self._toolchain(
                    self.baseline_root if role == "baseline" else self.candidate_root,
                    build_env,
                )
                if build.get("command") != expected_command:
                    raise RunnerError(f"build command differs for {key}")
                if build.get("environment") != expected_env:
                    raise RunnerError(f"build environment differs for {key}")
                if build.get("toolchain") != current_toolchain or toolchains[key] != current_toolchain:
                    raise RunnerError(f"build toolchain differs for {key}")
                if _resolve(str(build.get("tmpdir"))) != tmpdir:
                    raise RunnerError(f"build temporary directory differs for {key}")
                target_record = build.get("target")
                if not isinstance(target_record, Mapping) or _resolve(str(target_record.get("path"))) != target:
                    raise RunnerError(f"build target path differs for {key}")
                if target_record.get("bytes") != _directory_size(target) or target_record.get("bytes", 0) > TARGET_PREFLIGHT_BYTES:
                    raise RunnerError(f"build target identity differs for {key}")
                binary = build.get("binary")
                if not isinstance(binary, Mapping) or _resolve(str(binary.get("path"))) != expected_binary:
                    raise RunnerError(f"build binary path differs for {key}")
                if expected_binary.is_symlink() or not expected_binary.is_file() or not os.access(expected_binary, os.X_OK):
                    raise RunnerError(f"build binary is missing or not executable for {key}")
                binary_hash, binary_bytes = _sha256_file(expected_binary)
                if binary.get("sha256") != binary_hash or binary.get("bytes") != binary_bytes:
                    raise RunnerError(f"build binary hash differs for {key}")
                event = build.get("event")
                event_id = f"build-{role}-{variant}"
                if not isinstance(event, Mapping) or events_by_id.get(event_id) != event:
                    raise RunnerError(f"build event is missing or not journal-bound for {key}")
                if (
                    event.get("schema") != SCHEMA
                    or event.get("event_id") != event_id
                    or event.get("role") != role
                    or event.get("variant") != variant
                ):
                    raise RunnerError(f"build event role/variant differs for {key}")
                if event.get("argv") != expected_command or _resolve(str(event.get("cwd"))) != (self.baseline_root if role == "baseline" else self.candidate_root):
                    raise RunnerError(f"build event command/cwd differs for {key}")
                if event.get("environment") != expected_env or event.get("cpu_affinity") != self.cpu:
                    raise RunnerError(f"build event environment differs for {key}")
                if event.get("returncode") != 0 or event.get("error") is not None:
                    raise RunnerError(f"build event did not complete successfully for {key}")
                if event.get("process_guard_violations") or event.get("post_exit_descendants") or event.get("post_exit_offenders"):
                    raise RunnerError(f"build event has quiet-host violations for {key}")
                event_target = event.get("target")
                if not isinstance(event_target, Mapping) or _resolve(str(event_target.get("path"))) != target:
                    raise RunnerError(f"build event target differs for {key}")
                if event_target.get("bytes_before") != 0 or event_target.get("bytes_after") != target_record["bytes"]:
                    raise RunnerError(f"build event target freshness differs for {key}")
                if not isinstance(event.get("monitor"), list) or not event["monitor"]:
                    raise RunnerError(f"build event lacks live resource monitor samples for {key}")
                event_binary = event.get("binary")
                if not isinstance(event_binary, Mapping) or _resolve(str(event_binary.get("path"))) != expected_binary:
                    raise RunnerError(f"build event binary path differs for {key}")
                if (
                    event_binary.get("sha256") != binary_hash
                    or event_binary.get("bytes") != binary_bytes
                ):
                    raise RunnerError(f"build event binary hash differs for {key}")
                if event.get("toolchain") != current_toolchain:
                    raise RunnerError(f"build event toolchain differs for {key}")
                for stream_name in ("stdout", "stderr"):
                    stream = event.get(stream_name)
                    if not isinstance(stream, Mapping):
                        raise RunnerError(f"build event {stream_name} identity is missing for {key}")
                    stream_path = _resolve(str(stream.get("path")))
                    expected_stream = self.output / "build" / f"{role}-{variant}.{stream_name}"
                    if stream_path != expected_stream:
                        raise RunnerError(f"build event {stream_name} path differs for {key}")
                    stream_hash, stream_bytes = _sha256_file(expected_stream)
                    if stream.get("sha256") != stream_hash or stream.get("bytes") != stream_bytes:
                        raise RunnerError(f"build event {stream_name} hash differs for {key}")
        receipt["receipt_sha256"] = receipt_hash
        return receipt

    def _validate_probe_payload(self, path: Path, mode: str, expected_iterations: int) -> dict[str, Any]:
        raw = path.read_bytes()
        if not raw or len(raw) > PROBE_MAX_BYTES:
            raise RunnerError(f"probe output is empty or exceeds {PROBE_MAX_BYTES} bytes: {path}")
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as error:
            raise RunnerError(f"probe output is not JSON: {path}: {error}") from error
        if not isinstance(payload, Mapping):
            raise RunnerError("probe output root must be an object")
        if payload.get("iterations") != expected_iterations:
            raise RunnerError(
                f"probe output iterations {payload.get('iterations')!r} differ from requested {expected_iterations}"
            )
        try:
            shared.validate_probe_output(self.manifest, payload, mode)
        except shared.RunnerError as error:
            raise RunnerError(str(error)) from error
        if [row.get("family") for row in payload["families"]] != list(shared.FAMILIES):
            raise RunnerError("probe output did not contain all four route families in order")
        if mode == ALLOCATOR_MODE:
            for row in payload["families"]:
                timing = row["timing"]
                if timing.get("allocator_enabled") is not True or timing.get("elapsed_claimable") is not False:
                    raise RunnerError("allocator output did not carry the non-claimable observer contract")
        return dict(payload)

    def _run_probe(
        self,
        *,
        build: Mapping[str, Any],
        phase: str,
        role: str,
        kind: str,
        index: int,
        iterations: int,
    ) -> dict[str, Any]:
        mode = {
            "calibration": "validate",
            "normal": NORMAL_MODE,
            "allocator": ALLOCATOR_MODE,
        }[phase]
        root = self.baseline_root if role == "baseline" else self.candidate_root
        variant = "allocator" if phase == "allocator" else "normal"
        if build.get("role") != role or build.get("variant") != variant:
            raise RunnerError(f"build handoff role/variant differs for {role}/{phase}")
        target = self._target_path(role, variant)
        tmpdir = self._tmp_path(role, variant)
        binary = self._binary_path(role, variant)
        handoff_binary = build.get("binary")
        if not isinstance(handoff_binary, Mapping) or _resolve(str(handoff_binary.get("path"))) != binary:
            raise RunnerError(f"{role}/{phase} build handoff binary path is not the fixed endpoint binary")
        binary_hash, binary_bytes = _sha256_file(binary)
        if binary_hash != handoff_binary.get("sha256") or binary_bytes != handoff_binary.get("bytes"):
            raise RunnerError(f"{role}/{phase} binary changed before {kind} {index}")
        env = self._environment(role, variant, target, tmpdir)
        base = self.output / "probe" / phase / role
        output = base / f"{kind}-{index:03d}.json"
        stderr = base / f"{kind}-{index:03d}.stderr"
        command = [
            str(binary),
            "--mode",
            mode,
            "--family",
            "all",
            "--iterations",
            str(iterations),
        ]
        event = self._run_process(
            event_id=f"probe-{phase}-{role}-{kind}-{index:03d}",
            role=role,
            variant=variant,
            argv=command,
            cwd=root,
            env=env,
            target=target,
            stdout_path=output,
            stderr_path=stderr,
        )
        payload = self._validate_probe_payload(output, mode, iterations)
        endpoint = self._assert_endpoint_unchanged(role, build["source"])
        output_hash, output_bytes = _sha256_file(output)
        return {
            "role": role,
            "phase": phase,
            "kind": kind,
            "index": index,
            "mode": mode,
            "iterations": iterations,
            "command": command,
            "source": endpoint,
            "probe": {
                "path": str(output),
                "sha256": output_hash,
                "bytes": output_bytes,
                "stderr": event["stderr"],
                "families": list(shared.FAMILIES),
            },
            "monitor": event["monitor"],
            "claim_scope": "paired correctness/protocol evidence only; no native acceptance claim",
            "payload_identity": {
                "schema": payload["schema"],
                "mode": payload["mode"],
                "claim_scope": payload["claim_scope"],
                "routes": [row["descriptor"] for row in payload["families"]],
            },
        }

    def _calibrate(self, builds: Mapping[str, Mapping[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for role in ROLES:
            build = builds[f"{role}/normal"]
            result[role] = self._run_probe(
                build=build,
                phase="calibration",
                role=role,
                kind="validate",
                index=0,
                iterations=self.manifest["paired"]["calibration"]["iterations"],
            )
        for role in ROLES:
            self._assert_endpoint_unchanged(role, builds[f"{role}/normal"]["source"], full=True)
        return result

    def _abba_phase(self, builds: Mapping[str, Mapping[str, Any]], phase: str) -> dict[str, Any]:
        records: list[dict[str, Any]] = []
        role_indices: dict[str, int] = {role: 0 for role in ROLES}
        for leg, role in enumerate(ABBA_ORDER):
            build = builds[f"{role}/{phase}"]
            for warmup in range(WARMUPS):
                role_indices[role] += 1
                records.append(
                    self._run_probe(
                        build=build,
                        phase=phase,
                        role=role,
                        kind=f"warmup-leg{leg + 1}",
                        index=role_indices[role],
                        iterations=PROBE_ITERATIONS,
                    )
                )
            for sample in range(SAMPLES):
                role_indices[role] += 1
                records.append(
                    self._run_probe(
                        build=build,
                        phase=phase,
                        role=role,
                        kind=f"sample-leg{leg + 1}",
                        index=role_indices[role],
                        iterations=PROBE_ITERATIONS,
                    )
                )
        expected = (WARMUPS + SAMPLES) * len(ABBA_ORDER)
        if len(records) != expected:
            raise RunnerError(f"{phase} ABBA record count {len(records)} differs from {expected}")
        for role in ROLES:
            self._assert_endpoint_unchanged(role, builds[f"{role}/{phase}"]["source"], full=True)
        return {
            "phase": phase,
            "mode": NORMAL_MODE if phase == "normal" else ALLOCATOR_MODE,
            "abba_order": list(ABBA_ORDER),
            "warmup_iterations_per_leg": WARMUPS,
            "samples_per_leg": SAMPLES,
            "records": records,
            "allocator_elapsed_claimable": phase != "allocator",
        }

    def _write_receipt(self, path: Path, receipt: Mapping[str, Any]) -> None:
        self._fresh_file(path, "receipt")
        with path.open("x", encoding="utf-8") as stream:
            stream.write(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
            stream.flush()
            os.fsync(stream.fileno())

    def abba(self) -> None:
        receipt_path = self.output / "abba-receipt.json"
        build_receipt = self._load_build_receipt()
        builds = build_receipt.get("builds")
        if not isinstance(builds, Mapping) or set(builds) != {f"{role}/{variant}" for role in ROLES for variant in ("normal", "allocator")}:
            raise RunnerError("build receipt does not contain all normal/allocator endpoints")
        calibration: dict[str, Any] = {}
        phases: dict[str, Any] = {}
        try:
            calibration = self._calibrate(builds)
            for phase in PHASES:
                phases[phase] = self._abba_phase(builds, phase)
            endpoints = {
                key: {
                    "role": value["role"],
                    "variant": value["variant"],
                    "source": value["source"],
                    "binary": value["binary"],
                    "target": value["target"],
                    "features": list(FEATURES),
                    "routes": list(self.manifest["routes"]),
                    "phases": dict(self.manifest["paired"]),
                    "toolchain": value["toolchain"],
                }
                for key, value in builds.items()
            }
            receipt = {
                "schema": SCHEMA,
                "phase": "abba",
                "status": "passed",
                "manifest": {"path": str(self.manifest_path), "sha256": self.manifest_sha256},
                "preparation": {"path": str(self.preparation_path), **self.preparation},
                "build_receipt": {
                    "path": str(self.output / "build-receipt.json"),
                    "sha256": build_receipt["receipt_sha256"],
                },
                "runner": {"path": str(Path(__file__).resolve()), "sha256": self.runner_hash},
                "cpu_affinity": self.cpu,
                "features": list(FEATURES),
                "routes": list(self.manifest["routes"]),
                "phases": dict(self.manifest["paired"]),
                "endpoints": endpoints,
                "calibration": calibration,
                "abba": phases,
                "claim_scope": "paired correctness/protocol evidence only; no native acceptance claim",
            }
            self._write_receipt(receipt_path, receipt)
        except BaseException as error:
            failure = {
                "schema": SCHEMA,
                "phase": "abba",
                "status": "failed",
                "manifest_sha256": self.manifest_sha256,
                "preparation": self.preparation,
                "build_receipt_sha256": build_receipt.get("receipt_sha256"),
                "cpu_affinity": self.cpu,
                "calibration": calibration,
                "abba": phases,
                "error": str(error),
                "claim_scope": "failed paired protocol only; no native acceptance claim",
            }
            if not receipt_path.exists():
                self._write_receipt(receipt_path, failure)
            raise


def _parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--preparation", type=Path, required=True)
    parser.add_argument("--baseline-root", type=Path, required=True)
    parser.add_argument("--candidate-root", type=Path, required=True)
    parser.add_argument("--candidate-commit", required=True)
    parser.add_argument("--candidate-lock-sha256", required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--target-root", type=Path, required=True)
    parser.add_argument("--tmp-root", type=Path, required=True)
    parser.add_argument("--cpu", type=int, required=True)
    parser.add_argument("--phase", choices=("build", "abba"), required=True)
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parse_args(argv or sys.argv[1:])
    try:
        runner = PairedRunner(
            manifest_path=args.manifest,
            preparation_path=args.preparation,
            baseline_root=args.baseline_root,
            candidate_root=args.candidate_root,
            candidate_commit=args.candidate_commit,
            candidate_lock_sha256=args.candidate_lock_sha256,
            output=args.output,
            target_root=args.target_root,
            tmp_root=args.tmp_root,
            cpu=args.cpu,
            phase=args.phase,
        )
        if args.phase == "build":
            runner.build()
        else:
            runner.abba()
        return 0
    except (RunnerError, OSError, json.JSONDecodeError) as error:
        print(f"ODF shared-triad paired qualification refused: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
