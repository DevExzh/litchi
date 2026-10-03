#!/usr/bin/env python3
"""Verify retained 0891 evidence without rebuilding or launching workloads.

The reviewed 0887 verifier checks every captured command, semantic report, and
fresh reduction. This adapter supplies 0891 identity and replaces its source
replay with a reconstruction from the retained 0885 archive and sealed patch.
Managed executable copies are never recreated; after cleanup the retained
binary descriptors are sufficient for replay of the packet metadata.
"""

from __future__ import annotations

import argparse
import importlib.util
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys
import tarfile
import tempfile
from typing import Any

sys.dont_write_bytecode = True
PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
sys.path.insert(0, str(PACKET))
import experiment  # noqa: E402
import reduce as reducer  # noqa: E402

SOURCE = PACKET.parent / "change-0887" / "verify_packet.py"
SPEC = importlib.util.spec_from_file_location("litchi_x14ac_verify_0887", SOURCE)
if SPEC is None or SPEC.loader is None:
    raise ImportError(f"cannot load reviewed packet verifier: {SOURCE}")
IMPL = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = IMPL
SPEC.loader.exec_module(IMPL)
IMPL.PACKET = PACKET
IMPL.ROOT = ROOT
IMPL.experiment = experiment
IMPL.reducer = reducer
IMPL.SCHEMA = "litchi.xlsx-x14ac-verification-0891.v1"


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def read(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def resolve(raw: Any, base: Path = PACKET) -> Path:
    require(isinstance(raw, str), "descriptor path is missing")
    value = Path(raw)
    if value.is_absolute():
        return value.resolve()
    if value.parts[:2] == ("docs", "performance"):
        return (ROOT / value).resolve()
    return (base / value).resolve()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    info = path.stat()
    require(path.is_file() and not path.is_symlink(), f"expected retained regular file: {path}")
    return {"path": str(path), "bytes": info.st_size, "sha256": sha256(path), "mode": info.st_mode & 0o777, "executable": bool(info.st_mode & 0o100)}


def check_descriptor(record: dict[str, Any], label: str, *, base: Path = PACKET) -> dict[str, Any]:
    path = resolve(record.get("path"), base)
    actual = descriptor(path)
    identity = record.get("identity") if isinstance(record.get("identity"), dict) else record
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in identity:
            require(identity[key] == actual[key], f"{label} {key} differs from retained file")
    return actual


def iter_records(value: Any, prefix: str = ""):
    if isinstance(value, dict) and isinstance(value.get("path"), str) and isinstance(value.get("sha256"), str):
        yield prefix, value
    elif isinstance(value, dict):
        for key, child in value.items():
            name = f"{prefix}.{key}" if prefix else key
            yield from iter_records(child, name)


def patch_paths(path: Path) -> set[str]:
    paths: set[str] = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        match = re.fullmatch(r"diff --git a/(.+) b/(.+)", line)
        if match is None:
            continue
        left, right = match.groups()
        name = right if left == "/dev/null" else left
        require(left == right or left == "/dev/null", f"candidate patch path is asymmetric: {line}")
        require(not Path(name).is_absolute() and ".." not in Path(name).parts, f"candidate patch path is unsafe: {name}")
        paths.add(name)
    require(paths, "candidate patch has no paths")
    return paths


def verify_bindings(plan_path: Path, plan: dict[str, Any]) -> dict[str, Any]:
    bindings = plan.get("bindings")
    require(isinstance(bindings, dict), "plan bindings are missing")
    checked = []
    skipped = []
    for name, record in iter_records(bindings):
        if "candidate_overlay" in name:
            skipped.append({"name": name, "path": record.get("path"), "reason": "sealed overlay checked by archive plus patch reconstruction"})
            continue
        path = resolve(record.get("path"), plan_path.parent)
        if not path.exists() and ("managed" in name or path.parent == PACKET / "executables"):
            skipped.append({"name": name, "path": str(path), "reason": "managed copy removed after capture"})
            continue
        actual = check_descriptor(record, name, base=plan_path.parent)
        checked.append({"name": name, **actual})
    return {"status": "pass", "records": checked, "skipped": skipped, "record_count": len(checked) + len(skipped)}


def run_freeze_verify() -> dict[str, Any]:
    command = [sys.executable, "-B", str(PACKET / "freeze.py"), "--verify", "--retained-only"]
    completed = subprocess.run(command, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=False)
    require(completed.returncode == 0, f"freeze retained-only verification failed: {completed.stderr.strip()}")
    return {"argv": command, "cwd": str(ROOT), "exit_code": completed.returncode, "stdout_sha256": hashlib.sha256(completed.stdout.encode()).hexdigest(), "stderr_sha256": hashlib.sha256(completed.stderr.encode()).hexdigest()}


def verify_source_reconstruction() -> dict[str, Any]:
    freeze = read(PACKET / "candidate-freeze-v3.json")
    require(freeze.get("schema") == "litchi.xlsx-x14ac-candidate-freeze-0891.v1", "candidate freeze schema mismatch")
    predecessor = freeze.get("supersedes")
    require(isinstance(predecessor, dict), "candidate freeze v3 predecessor descriptor is missing")
    predecessor_path = resolve(predecessor.get("path"))
    predecessor_file = PACKET / "candidate-freeze-v2.json"
    require(predecessor_path == predecessor_file.resolve(), "candidate freeze v3 predecessor path mismatch")
    predecessor_actual = descriptor(predecessor_file)
    require(predecessor.get("bytes") == predecessor_actual["bytes"] and predecessor.get("sha256") == predecessor_actual["sha256"], "candidate freeze v2 predecessor identity mismatch")
    archive = PACKET.parent / "change-0885" / "source.tar.gz"
    manifest_path = PACKET.parent / "change-0885" / "source.json"
    patch_record = freeze.get("patch")
    require(isinstance(patch_record, dict), "candidate patch record is missing")
    patch = resolve(patch_record.get("path"))
    check_descriptor({"path": str(archive), "bytes": archive.stat().st_size, "sha256": sha256(archive)}, "baseline archive")
    check_descriptor({"path": str(manifest_path), "bytes": manifest_path.stat().st_size, "sha256": sha256(manifest_path)}, "baseline manifest")
    check_descriptor({**patch_record, "path": str(patch)}, "candidate patch")
    manifest = read(manifest_path)
    files = manifest.get("files")
    require(isinstance(files, dict) and len(files) == 10_765, "baseline source manifest file count mismatch")
    overlays = {entry["path"]: entry for entry in freeze.get("files", []) if isinstance(entry, dict) and isinstance(entry.get("path"), str)}
    require(set(overlays) == set(freeze.get("candidate_paths", [])), "candidate overlay inventory mismatch")
    require(patch_paths(patch) == set(overlays), "candidate patch paths differ from frozen overlays")

    with tempfile.TemporaryDirectory(prefix="litchi-0891-replay-") as temporary:
        root = Path(temporary)
        with tarfile.open(archive, "r:*") as handle:
            handle.extractall(root, filter="data")
        check = subprocess.run(["git", "apply", "--check", str(patch)], cwd=root, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=False)
        require(check.returncode == 0, f"candidate patch check failed: {check.stderr.strip()}")
        applied = subprocess.run(["git", "apply", str(patch)], cwd=root, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=False)
        require(applied.returncode == 0, f"candidate patch apply failed: {applied.stderr.strip()}")
        expected_names = set(files) | set(overlays)
        actual_names = {path.relative_to(root).as_posix() for path in root.rglob("*") if path.is_file()}
        require(actual_names == expected_names, f"reconstructed source census differs: {len(actual_names)} vs {len(expected_names)}")
        for name in sorted(expected_names):
            path = root / name
            expected = overlays.get(name, files.get(name))
            require(path.stat().st_size == expected["bytes"], f"source byte count mismatch: {name}")
            require(sha256(path) == expected["sha256"], f"source hash mismatch: {name}")
    return {"status": "pass", "baseline_archive": {"path": str(archive.resolve()), "bytes": archive.stat().st_size, "sha256": sha256(archive)}, "baseline_manifest": {"path": str(manifest_path.resolve()), "bytes": manifest_path.stat().st_size, "sha256": sha256(manifest_path)}, "candidate_patch": {"path": str(patch), "bytes": patch.stat().st_size, "sha256": sha256(patch)}, "manifest_files": len(expected_names), "overlay_files": sorted(overlays), "git_apply_check": "pass", "git_apply": "pass", "temporary_directory_cleaned": True}


IMPL.verify_bindings = verify_bindings
IMPL.run_freeze_verify = run_freeze_verify
IMPL.verify_source_reconstruction = verify_source_reconstruction
ORIGINAL_VERIFY_COMMANDS = IMPL.verify_commands


def verify_commands(plan_path: Path, experiment_root: Path, plan: dict[str, Any]) -> dict[str, Any]:
    """Replay command custody with retained managed copies after cleanup.

    The reviewed 0887 verifier loads the plan for validation but then passes
    its original JSON object to ``verify_commands``.  Supply the adapter's
    retained-only marker on that replay object so removed managed copies use
    their sealed descriptors, while partial custody remains an error.
    """
    managed = plan.get("managed_binaries")
    require(isinstance(managed, dict), "plan has no managed executable descriptors")
    paths: list[Path] = []
    for role in ("before", "after"):
        side = managed.get(role)
        require(isinstance(side, dict), f"plan has no managed {role} binaries")
        for mode in ("native", "allocator"):
            item = experiment.managed_item(side.get(mode))
            path = Path(str(item.get("path", "")))
            if not path.is_absolute():
                path = PACKET / path
            paths.append(path.resolve())
    present = [path.is_file() for path in paths]
    require(not any(present) or all(present), "managed executable custody is partial")
    replay_plan = dict(plan)
    replay_plan["retained_only"] = not any(present)
    return ORIGINAL_VERIFY_COMMANDS(plan_path, experiment_root, replay_plan)


IMPL.verify_commands = verify_commands
ORIGINAL_VERIFY = IMPL.verify


def verify(plan_path: Path, experiment_root: Path, reduction_path: Path, output: Path) -> dict[str, Any]:
    require(not output.exists(), f"verification output already exists: {output}")
    require(reduction_path.is_file(), "reduction.json is required")
    plan = read(plan_path)
    require(isinstance(plan, dict), "plan must be an object")
    # The reviewed verifier performs the complete child/semantic/reduction
    # checks; our replacements above change only packet custody and source
    # reconstruction.
    result = ORIGINAL_VERIFY(plan_path, experiment_root, reduction_path, output)
    return result


IMPL.verify = verify


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--experiment-root", type=Path, required=True)
    parser.add_argument("--reduction", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(verify(args.plan.resolve(), args.experiment_root.resolve(), args.reduction.resolve(), args.output.resolve()), indent=2, sort_keys=True))
