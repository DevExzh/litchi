#!/usr/bin/env python3
"""Verify the retained 0887 capture packet without rerunning its workload.

The verifier checks frozen bindings, the complete command schedule, the
retained reducer output, and a fresh reconstruction of the candidate source.
It deliberately does not build, benchmark, or remove the owned build roots.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import sys
import tarfile
import tempfile
from typing import Any, Iterator


sys.dont_write_bytecode = True
PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
sys.path.insert(0, str(PACKET))
import experiment  # noqa: E402
import reduce as reducer  # noqa: E402


SCHEMA = "litchi.xlsx-mce-inline4-verification-0887.v1"
MANIFEST_FILE_COUNT = 10_765
OWNER_MARKER = ".owner-0887"
OWNER_MARKER_TEXT = "MCE attribute candidate 0887\n"


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


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    info = path.stat()
    require(stat.S_ISREG(info.st_mode), f"expected regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": sha256(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def resolve_path(raw: str, base: Path = PACKET) -> Path:
    path = Path(raw)
    return (path if path.is_absolute() else base / path).resolve()


def check_record(record: dict[str, Any], *, base: Path = PACKET) -> dict[str, Any]:
    require(isinstance(record, dict) and isinstance(record.get("path"), str), "invalid retained descriptor")
    path = resolve_path(record["path"], base)
    actual = descriptor(path)
    identity = record.get("identity") if isinstance(record.get("identity"), dict) else record
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in identity:
            require(identity[key] == actual[key], f"{path} {key} does not match its retained descriptor")
    return actual


def iter_binding_records(value: Any, name: str = "") -> Iterator[tuple[str, dict[str, Any]]]:
    if isinstance(value, dict):
        if isinstance(value.get("path"), str) and isinstance(value.get("sha256"), str):
            yield name, value
            return
        for key, item in value.items():
            child = f"{name}.{key}" if name else key
            yield from iter_binding_records(item, child)


def verify_bindings(plan_path: Path, plan: dict[str, Any]) -> dict[str, Any]:
    bindings = plan.get("bindings")
    require(isinstance(bindings, dict), "plan bindings are missing")
    records = list(iter_binding_records(bindings))
    require(records, "plan has no binding records")
    checked = []
    skipped = []
    for name, record in records:
        if name.startswith("candidate_overlay/"):
            skipped.append({
                "name": name,
                "path": record.get("path"),
                "reason": "candidate overlay is verified by independent baseline archive plus candidate patch replay",
            })
            continue
        actual = check_record(record, base=plan_path.parent)
        checked.append({"name": name, "path": actual["path"], "bytes": actual["bytes"], "sha256": actual["sha256"]})

    status_checks = []
    for _, record in records:
        path = resolve_path(record["path"], plan_path.parent)
        if path.name == "build.json":
            build = read(path)
            require(build.get("status") == "pass", "candidate build receipt is not passing")
            status_checks.append({"path": str(path), "status": build["status"]})
        elif path.name in {"native-build.json", "allocator-build.json"}:
            build = read(path)
            require(build.get("exit_code") == 0, f"{path.name} did not exit successfully")
            status_checks.append({"path": str(path), "exit_code": build["exit_code"]})
        elif path.name in {"methodology-review.json", "semantic-review.json"}:
            review = read(path)
            require(review.get("status") == "pass", f"{path.name} is not passing")
            status_checks.append({"path": str(path), "status": review["status"]})
    return {"records": checked, "skipped": skipped, "record_count": len(records), "status_checks": status_checks}


def run_freeze_verify() -> dict[str, Any]:
    freeze = PACKET / "freeze.py"
    require(freeze.is_file(), "freeze.py is missing")
    command = [sys.executable, "-B", str(freeze), "--verify", "--retained-only"]
    completed = subprocess.run(
        command,
        cwd=ROOT,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    require(completed.returncode == 0, f"freeze --verify failed: {completed.stderr.strip()}")
    return {
        "argv": command,
        "cwd": str(ROOT),
        "exit_code": completed.returncode,
        "stdout_sha256": hashlib.sha256(completed.stdout.encode()).hexdigest(),
        "stderr_sha256": hashlib.sha256(completed.stderr.encode()).hexdigest(),
    }


def expected_rows() -> list[dict[str, Any]]:
    rows = experiment._schedule()
    return [
        {
            "path": str(Path("runs") / f"{row['ordinal']:03d}-{row['lane']}-{row['pair']}-{row['role']}" / "receipt.json"),
            "row": row,
        }
        for row in rows
    ]


def verify_commands(plan_path: Path, experiment_root: Path, plan: dict[str, Any]) -> dict[str, Any]:
    root = experiment_root.resolve()
    receipt_path = root / "experiment.json"
    require(receipt_path.is_file(), "experiment.json is missing")
    receipt = read(receipt_path)
    rows = experiment._schedule()
    require(receipt.get("schema") == experiment.EXPERIMENT_SCHEMA, "experiment receipt schema mismatch")
    require(receipt.get("status") == "pass", "experiment receipt is not passing")
    require(receipt.get("plan") == str(plan_path.resolve()), "experiment plan path mismatch")
    require(receipt.get("plan_sha256") == sha256(plan_path), "experiment plan hash mismatch")
    require(receipt.get("rows") == expected_rows(), "experiment receipt rows differ from the frozen schedule")
    require(receipt.get("children") == len(rows) == experiment.TOTAL_CHILDREN, "experiment child count mismatch")
    require(receipt.get("native_children") == experiment.TOTAL_NATIVE_CHILDREN, "native child count mismatch")
    require(receipt.get("allocator_children") == experiment.TOTAL_ALLOCATOR_CHILDREN, "allocator child count mismatch")
    require(receipt.get("no_retries_or_replacements") is True, "capture retry/replacement policy is not frozen")
    require(receipt.get("performance_claim") is None, "capture receipt makes a performance claim")

    binaries = {
        (role, mode): experiment._binary(plan, role, mode)
        for role in ("before", "after")
        for mode in ("native", "allocator")
    }
    command_count = 0
    semantic_count = 0
    for row in rows:
        mode = "native" if row["lane"] == "native" else "allocator"
        binary = binaries[(row["role"], mode)]
        folder = root / "runs" / f"{row['ordinal']:03d}-{row['lane']}-{row['pair']}-{row['role']}"
        command_path = folder / "command.json"
        child_receipt_path = folder / "receipt.json"
        report_path = folder / "report.json"
        producer_path = folder / "producer.json"
        peak_path = folder / "peak-rss-kib.txt"
        command = read(command_path)
        child_receipt = read(child_receipt_path)
        expected_argv = experiment._command(
            binary["path"], row["case"], row["warmup"], row["samples"], report_path, producer_path, peak_path
        )
        require(command.get("schema") == experiment.COMMAND_SCHEMA, f"command schema mismatch: {row['ordinal']}")
        require(command.get("row") == row, f"command row mismatch: {row['ordinal']}")
        require(command.get("argv") == expected_argv, f"command argv mismatch: {row['ordinal']}")
        require(command.get("cwd") == str(experiment.PACKET), f"command cwd mismatch: {row['ordinal']}")
        require(command.get("environment") == experiment.CHILD_ENV, f"command environment mismatch: {row['ordinal']}")
        require(command.get("binary") == binary, f"command binary descriptor mismatch: {row['ordinal']}")
        require(child_receipt.get("command", {}).get("sha256") == sha256(command_path), f"command custody mismatch: {row['ordinal']}")
        for key in ("row", "argv", "cwd", "environment", "binary", "plan_sha256", "started_utc", "finished_utc", "exit_code"):
            require(child_receipt.get(key) == command.get(key), f"receipt/command {key} mismatch: {row['ordinal']}")
        semantic_path = folder / "semantic.json"
        semantic = read(semantic_path)
        require(semantic.get("status") == "pass", f"semantic receipt is not passing: {row['ordinal']}")
        require(semantic.get("report_sha256") == sha256(report_path), f"semantic report hash mismatch: {row['ordinal']}")
        expected_producer_sha = sha256(producer_path) if producer_path.is_file() else None
        require(semantic.get("producer_sha256") == expected_producer_sha, f"semantic producer hash mismatch: {row['ordinal']}")
        command_count += 1
        semantic_count += 1

    return {
        "status": "pass",
        "experiment": str(receipt_path),
        "plan_sha256": receipt["plan_sha256"],
        "rows": len(rows),
        "commands_checked": command_count,
        "semantic_receipts_checked": semantic_count,
        "cwd": str(experiment.PACKET),
        "environment": experiment.CHILD_ENV,
    }


def patch_paths(patch: Path) -> set[str]:
    paths = set()
    pattern = re.compile(r"^diff --git a/(.+) b/(.+)$")
    for line in patch.read_text(encoding="utf-8").splitlines():
        match = pattern.fullmatch(line)
        if match is None:
            continue
        left, right = match.groups()
        require(left == right and not Path(left).is_absolute() and ".." not in Path(left).parts, "candidate patch has an unsafe or asymmetric path")
        paths.add(left)
    require(paths, "candidate patch contains no file paths")
    return paths


def verify_source_reconstruction() -> dict[str, Any]:
    source_record = read(PACKET / "candidate-source.json")
    archive_record = source_record.get("baseline_archive")
    manifest_record = source_record.get("baseline_manifest")
    patch_record = source_record.get("patch")
    require(isinstance(archive_record, dict) and isinstance(manifest_record, dict) and isinstance(patch_record, dict), "candidate source records are incomplete")
    archive = resolve_path(archive_record["path"])
    manifest_path = resolve_path(manifest_record["path"])
    patch = resolve_path(patch_record["path"])
    check_record(archive_record)
    check_record(manifest_record)
    check_record(patch_record)
    manifest = read(manifest_path)
    files = manifest.get("files")
    require(isinstance(files, dict) and len(files) == MANIFEST_FILE_COUNT, "baseline manifest file count mismatch")
    overlays = source_record.get("overlays")
    require(isinstance(overlays, dict) and overlays, "candidate overlays are missing")
    changed = patch_paths(patch)
    require(changed == set(overlays), "candidate patch paths differ from retained overlay records")

    with tempfile.TemporaryDirectory(prefix="litchi-0887-replay-") as temporary:
        root = Path(temporary)
        with tarfile.open(archive, "r:*") as handle:
            handle.extractall(root, filter="data")
        check = subprocess.run(
            ["git", "apply", "--check", str(patch)],
            cwd=root,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )
        require(check.returncode == 0, f"candidate patch check failed: {check.stderr.strip()}")
        apply = subprocess.run(
            ["git", "apply", str(patch)],
            cwd=root,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )
        require(apply.returncode == 0, f"candidate patch apply failed: {apply.stderr.strip()}")

        actual_files: set[str] = set()
        for path in root.rglob("*"):
            if path.is_symlink():
                raise ValueError(f"source reconstruction contains a symlink: {path}")
            if path.is_file():
                actual_files.add(path.relative_to(root).as_posix())
        require(actual_files == set(files), "source reconstruction file set differs from manifest")
        for name, record in files.items():
            path = root / name
            actual = descriptor(path)
            expected = overlays.get(name, record)
            require(actual["bytes"] == expected["bytes"], f"source byte count mismatch: {name}")
            require(actual["sha256"] == expected["sha256"], f"source hash mismatch: {name}")
            # ``git apply`` rewrites touched regular files with the process
            # umask (0664 on this host), even when the patch has no mode
            # header.  The retained source manifest's byte/hash identity is
            # the reconstruction contract; mode is checked for the retained
            # descriptors and is intentionally not treated as a patch-content
            # mismatch here.
        reconstructed_files = len(actual_files)
        overlay_identities = {
            name: {
                "bytes": (root / name).stat().st_size,
                "sha256": sha256(root / name),
            }
            for name in sorted(overlays)
        }
    require(not root.exists(), "temporary source replay directory was not cleaned")
    return {
        "status": "pass",
        "baseline_archive": {"path": str(archive), "sha256": sha256(archive), "bytes": archive.stat().st_size},
        "candidate_patch": {"path": str(patch), "sha256": sha256(patch), "bytes": patch.stat().st_size},
        "manifest": str(manifest_path),
        "manifest_files": reconstructed_files,
        "overlay_files": sorted(overlays),
        "overlay_identities": overlay_identities,
        "git_apply_check": "pass",
        "git_apply": "pass",
        "temporary_directory_cleaned": True,
    }


def verify(plan_path: Path, experiment_root: Path, reduction_path: Path, output: Path) -> dict[str, Any]:
    require(reduction_path.is_file(), "reduction.json is required before packet verification")
    require(not output.exists(), f"verification output already exists: {output}")
    plan = read(plan_path)
    require(isinstance(plan, dict), "plan must be a JSON object")
    experiment.load_plan(plan_path)
    freeze = run_freeze_verify()
    bindings = verify_bindings(plan_path, plan)
    commands = verify_commands(plan_path, experiment_root, plan)
    source = verify_source_reconstruction()
    retained = read(reduction_path)
    computed = reducer.reduce(plan_path.resolve(), experiment_root.resolve())
    require(retained == computed, "retained reduction.json differs from a fresh reducer recomputation")
    result = {
        "schema": SCHEMA,
        "status": "pass",
        "plan": str(plan_path.resolve()),
        "plan_sha256": sha256(plan_path),
        "experiment_root": str(experiment_root.resolve()),
        "reduction": str(reduction_path.resolve()),
        "reduction_sha256": sha256(reduction_path),
        "freeze_verify": freeze,
        "bindings": bindings,
        "commands": commands,
        "source_reconstruction": source,
        "reduction_recomputed": True,
        "performance_claim": None,
        "adoption": "not evaluated; root review remains separate",
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as handle:
        json.dump(result, handle, indent=2, sort_keys=True)
        handle.write("\n")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--experiment-root", type=Path, required=True)
    parser.add_argument("--reduction", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = verify(args.plan.resolve(), args.experiment_root.resolve(), args.reduction.resolve(), args.output.resolve())
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
