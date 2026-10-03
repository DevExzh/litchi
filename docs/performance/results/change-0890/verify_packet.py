#!/usr/bin/env python3
"""Replay a retained 0890 paired packet without rebuilding or benchmarking.

This adapter reuses the reviewed 0887 packet verifier for command, receipt,
reduction, and binding checks.  Its binary seam accepts the retained managed
copy descriptors after those copies have been removed, while source replay
also permits explicitly inventoried new MCE test files.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path
import re
import sys
import tarfile
import tempfile
from typing import Any


PACKET = Path(__file__).resolve().parent
sys.path.insert(0, str(PACKET))
import experiment  # noqa: E402
import reduce as reducer  # noqa: E402

_SOURCE = PACKET.parent / "change-0887" / "verify_packet.py"
_SPEC = importlib.util.spec_from_file_location("litchi_mce_verify_packet_0887", _SOURCE)
if _SPEC is None or _SPEC.loader is None:
    raise ImportError(f"cannot load reviewed packet verifier: {_SOURCE}")
_IMPL = importlib.util.module_from_spec(_SPEC)
sys.modules[_SPEC.name] = _IMPL
_SPEC.loader.exec_module(_IMPL)
_IMPL.PACKET = PACKET
_IMPL.ROOT = PACKET.parents[3]
_IMPL.experiment = experiment
_IMPL.reducer = reducer
_IMPL.SCHEMA = "litchi.xlsx-mce-expanded-name-cache-verification-0890.v1"
_ACTIVE_PLAN: Path | None = None


def _patch_paths(path: Path) -> set[str]:
    paths: set[str] = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        match = re.fullmatch(r"diff --git a/(.+) b/(.+)", line)
        if match is None:
            continue
        left, right = match.groups()
        name = right if left == "/dev/null" else left
        _IMPL.require(left == right or left == "/dev/null", f"candidate patch path mismatch: {line}")
        _IMPL.require(not Path(name).is_absolute() and ".." not in Path(name).parts, f"unsafe candidate patch path: {name}")
        paths.add(name)
    _IMPL.require(paths, "candidate patch contains no paths")
    return paths


def verify_source_reconstruction() -> dict[str, Any]:
    packet = PACKET
    freeze = _IMPL.read(packet / "candidate-freeze-v2.json")
    _IMPL.require(freeze.get("schema") == "litchi.mce-candidate-freeze-0890.v1", "candidate freeze schema mismatch")
    archive_record = _IMPL.descriptor(packet.parent / "change-0885" / "source.tar.gz")
    manifest_record = _IMPL.descriptor(packet.parent / "change-0885" / "source.json")
    patch_record = freeze.get("patch")
    _IMPL.require(isinstance(patch_record, dict), "candidate freeze patch record is missing")
    archive = _IMPL.resolve_path(archive_record["path"])
    manifest_path = _IMPL.resolve_path(manifest_record["path"])
    patch_raw = Path(patch_record.get("path", ""))
    patch = (patch_raw if patch_raw.is_absolute() else _IMPL.ROOT / patch_raw).resolve()
    _IMPL.check_record(archive_record)
    _IMPL.check_record(manifest_record)
    patch_check = dict(patch_record)
    patch_check["path"] = str(patch)
    _IMPL.check_record(patch_check)
    manifest = _IMPL.read(manifest_path)
    files = manifest.get("files")
    frozen_files = freeze.get("files")
    _IMPL.require(isinstance(frozen_files, list), "candidate freeze file inventory is missing")
    overlays = {
        entry["path"]: entry
        for entry in frozen_files
        if isinstance(entry, dict) and isinstance(entry.get("path"), str)
    }
    _IMPL.require(isinstance(files, dict), "baseline source manifest is incomplete")
    _IMPL.require(isinstance(overlays, dict) and overlays, "candidate overlays are missing")
    candidate_paths = freeze.get("candidate_paths")
    _IMPL.require(isinstance(candidate_paths, list) and set(candidate_paths) == set(overlays), "candidate freeze path inventory differs from file inventory")
    _IMPL.require(_patch_paths(patch) == set(overlays), "candidate patch paths differ from overlays")

    with tempfile.TemporaryDirectory(prefix="litchi-0890-replay-") as temporary:
        root = Path(temporary)
        with tarfile.open(archive, "r:*") as handle:
            handle.extractall(root, filter="data")
        check = _IMPL.subprocess.run(
            ["git", "apply", "--check", str(patch)],
            cwd=root,
            stdout=_IMPL.subprocess.PIPE,
            stderr=_IMPL.subprocess.PIPE,
            text=True,
            check=False,
        )
        _IMPL.require(check.returncode == 0, f"candidate patch check failed: {check.stderr.strip()}")
        applied = _IMPL.subprocess.run(
            ["git", "apply", str(patch)],
            cwd=root,
            stdout=_IMPL.subprocess.PIPE,
            stderr=_IMPL.subprocess.PIPE,
            text=True,
            check=False,
        )
        _IMPL.require(applied.returncode == 0, f"candidate patch apply failed: {applied.stderr.strip()}")
        expected_names = set(files) | set(overlays)
        actual_names: set[str] = set()
        for path in root.rglob("*"):
            _IMPL.require(not path.is_symlink(), f"source reconstruction contains a symlink: {path}")
            if path.is_file():
                actual_names.add(path.relative_to(root).as_posix())
        _IMPL.require(actual_names == expected_names, "reconstructed source file set differs from inventory")
        for name in sorted(expected_names):
            path = root / name
            expected = overlays.get(name, files.get(name))
            _IMPL.require(path.stat().st_size == expected["bytes"], f"source byte count mismatch: {name}")
            _IMPL.require(_IMPL.sha256(path) == expected["sha256"], f"source hash mismatch: {name}")
        overlay_identities = {
            name: {"bytes": (root / name).stat().st_size, "sha256": _IMPL.sha256(root / name)}
            for name in sorted(overlays)
        }
    return {
        "status": "pass",
        "baseline_archive": {"path": str(archive), "bytes": archive.stat().st_size, "sha256": _IMPL.sha256(archive)},
        "candidate_patch": {"path": str(patch), "bytes": patch.stat().st_size, "sha256": _IMPL.sha256(patch)},
        "manifest": str(manifest_path),
        "manifest_files": len(expected_names),
        "overlay_files": sorted(overlays),
        "overlay_identities": overlay_identities,
        "git_apply_check": "pass",
        "git_apply": "pass",
        "temporary_directory_cleaned": True,
    }


def run_freeze_verify() -> dict[str, Any]:
    _IMPL.require(_ACTIVE_PLAN is not None, "active plan is missing")
    plan = experiment.load_plan(_ACTIVE_PLAN)
    return {
        "status": "pass",
        "plan": str(_ACTIVE_PLAN.resolve()),
        "plan_sha256": _IMPL.sha256(_ACTIVE_PLAN),
        "retained_only": bool(plan.get("retained_only")),
        "workload_launched": False,
    }


_IMPL.verify_source_reconstruction = verify_source_reconstruction
_IMPL.run_freeze_verify = run_freeze_verify
_ORIGINAL_VERIFY = _IMPL.verify


def verify(plan_path: Path, experiment_root: Path, reduction_path: Path, output: Path) -> dict[str, Any]:
    global _ACTIVE_PLAN
    _ACTIVE_PLAN = plan_path
    return _ORIGINAL_VERIFY(plan_path, experiment_root, reduction_path, output)


_IMPL.verify = verify


if __name__ == "__main__":
    raise SystemExit(_IMPL.main())
