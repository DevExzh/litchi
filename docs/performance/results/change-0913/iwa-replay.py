#!/usr/bin/env python3
"""Replay the retained iWork resolution against its immutable commit.

This verifier reads packet receipts and Git objects only.  It does not inspect
the current working-tree source, the current index, a worktree, a build tree,
or any live process.  The requested output is the only new file it creates.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]

COMMIT = "1b12cef9a97bbe3597039f7e3caadc2c45c905ce"
PARENT = "59f75c7f79026a4d124b61e97c178c7b45de19ab"
CHERRY_PICK = "f2869418f0038098f4e4934b406c892ca0fe78fd"
INVENTORY_PREREQUISITE = "c09fce06c7"

SOURCE_FILES = (
    "crates/litchi-iwa-archive/src/ingress.rs",
    "crates/litchi-iwa-archive/src/lib.rs",
    "crates/litchi-iwa-archive/src/package.rs",
    "crates/litchi-iwa-core/src/archive.rs",
    "crates/litchi-iwa-core/src/archive/ingress.rs",
    "crates/litchi-iwa-core/src/archive/inventory.rs",
    "crates/litchi-iwa-core/src/lib.rs",
    "crates/litchi-iwa-core/src/snappy.rs",
)

FINAL_QUALITY = "iwa-quality-v4"
FINAL_QUALITY_KINDS = ("tests", "doctests", "clippy", "docs", "boundaries")
FAILED_ATTEMPT_DIRS = ("iwa-quality", "iwa-quality-v2", "iwa-quality-v3")


class ReplayError(ValueError):
    """The retained iWork evidence is missing or contradictory."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReplayError(message)


def load(name: str) -> Any:
    path = HERE / name
    require(path.is_file() and not path.is_symlink(), f"missing packet JSON: {name}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ReplayError(f"invalid packet JSON {name}: {error}") from error


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_path(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing packet file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def packet_path(raw: str, label: str) -> Path:
    path = Path(raw)
    require(not path.is_absolute() and ".." not in path.parts, f"{label}: unsafe packet path")
    resolved = (ROOT / path).resolve()
    try:
        resolved.relative_to(ROOT.resolve())
    except ValueError as error:
        raise ReplayError(f"{label}: path escapes workspace") from error
    return resolved


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    try:
        relative = str(path.relative_to(ROOT.resolve()))
    except ValueError as error:
        raise ReplayError(f"descriptor path escapes workspace: {path}") from error
    require(path.is_file() and not path.is_symlink(), f"descriptor target is not a regular file: {path}")
    return {"path": relative, "bytes": path.stat().st_size, "sha256": sha256_path(path)}


def descriptor_shape(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: descriptor is malformed")
    path = value.get("path")
    require(isinstance(path, str) and path, f"{label}: descriptor path is missing")
    size = value.get("bytes")
    require(type(size) is int and size >= 0, f"{label}: descriptor byte count is invalid")
    digest = value.get("sha256")
    require(
        isinstance(digest, str)
        and len(digest) == 64
        and all(char in "0123456789abcdef" for char in digest),
        f"{label}: descriptor SHA-256 is invalid",
    )
    return {"path": path, "bytes": size, "sha256": digest}


def check_packet_descriptor(value: Any, label: str) -> dict[str, Any]:
    expected = descriptor_shape(value, label)
    path = packet_path(expected["path"], label)
    actual = descriptor(path)
    require(actual == expected, f"{label}: packet bytes changed")
    return actual


def check_packet_descriptor_tree(value: Any, label: str) -> int:
    """Check nested descriptors whose paths refer to retained packet files."""
    count = 0
    if isinstance(value, dict):
        if {"path", "bytes", "sha256"}.issubset(value):
            check_packet_descriptor(value, label)
            return 1
        for key, child in value.items():
            count += check_packet_descriptor_tree(child, f"{label}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            count += check_packet_descriptor_tree(child, f"{label}[{index}]")
    return count


def git(*args: str) -> bytes:
    try:
        return subprocess.check_output(["git", *args], cwd=ROOT)
    except subprocess.CalledProcessError as error:
        raise ReplayError(f"git command failed: git {' '.join(args)}") from error


def git_text(*args: str) -> str:
    return git(*args).decode("utf-8").strip()


def commit_blob(path: str) -> bytes:
    require("\x00" not in path and not path.startswith("/"), f"unsafe commit path: {path}")
    kind = git_text("cat-file", "-t", f"{COMMIT}:{path}")
    require(kind == "blob", f"immutable commit path is not a blob: {path}")
    return git("show", f"{COMMIT}:{path}")


def immutable_commit_check(resolution: dict[str, Any], source: dict[str, Any], formatted: dict[str, Any]) -> dict[str, Any]:
    require(git_text("rev-parse", f"{COMMIT}^{{commit}}") == COMMIT, "immutable iWork commit is unavailable")
    parents = git_text("show", "-s", "--format=%P", COMMIT).split()
    require(parents == [PARENT], "immutable iWork commit parent differs")
    prerequisite = git_text("rev-parse", f"{INVENTORY_PREREQUISITE}^{{commit}}")
    require(prerequisite.startswith(INVENTORY_PREREQUISITE), "inventory prerequisite cannot be resolved")

    changed = git_text("diff-tree", "--no-commit-id", "--name-only", "-r", COMMIT).splitlines()
    require(sorted(changed) == list(SOURCE_FILES), "immutable iWork commit changed-file set differs")

    rows = resolution.get("files")
    require(isinstance(rows, list), "resolution file inventory is missing")
    by_path = {row.get("path"): row for row in rows if isinstance(row, dict)}
    require(len(by_path) == len(rows) and tuple(sorted(by_path)) == tuple(sorted(SOURCE_FILES)), "resolution file inventory differs")

    source_rows = source.get("files")
    format_rows = formatted.get("files")
    require(isinstance(source_rows, list) and isinstance(format_rows, list), "quality or format source inventory is missing")
    source_by_path = {row.get("path"): row for row in source_rows if isinstance(row, dict)}
    format_by_path = {row.get("path"): row for row in format_rows if isinstance(row, dict)}
    require(tuple(sorted(source_by_path)) == tuple(sorted(SOURCE_FILES)), "quality source inventory differs")
    require(tuple(sorted(format_by_path)) == tuple(sorted(SOURCE_FILES)), "format source inventory differs")

    checked = []
    for path in SOURCE_FILES:
        blob = commit_blob(path)
        expected = {"path": path, "bytes": len(blob), "sha256": sha256_bytes(blob)}
        for label, row in (("resolution", by_path[path]), ("quality source", source_by_path[path]), ("format", format_by_path[path])):
            require(
                isinstance(row, dict)
                and row.get("path") == expected["path"]
                and row.get("bytes") == expected["bytes"]
                and row.get("sha256") == expected["sha256"],
                f"{label} blob differs: {path}",
            )
        checked.append(expected)

    return {
        "commit": COMMIT,
        "parent": PARENT,
        "inventory_prerequisite": prerequisite,
        "changed_files": list(SOURCE_FILES),
        "source_blobs_checked": len(checked),
    }


def check_resolution_before(before: dict[str, Any]) -> None:
    require(before.get("status") == "authorized", "iWork resolution authorization receipt is missing")
    require(before.get("head_before") == PARENT, "iWork resolution starting HEAD differs")
    require(before.get("cherry_pick") == CHERRY_PICK, "iWork cherry-pick identity differs")
    staged = before.get("staged_paths")
    require(isinstance(staged, list) and set(staged) <= set(SOURCE_FILES), "iWork pre-resolution staged paths differ")
    unmerged = before.get("unmerged_paths")
    require(isinstance(unmerged, list) and set(unmerged) <= set(SOURCE_FILES), "iWork pre-resolution conflict paths differ")
    index = before.get("index_before_sha256")
    require(isinstance(index, str) and len(index) == 64, "iWork pre-resolution index receipt is malformed")


def check_resolution_after(after: dict[str, Any]) -> None:
    require(after.get("status") == "pass", "iWork resolution did not pass")
    require(after.get("commit") == COMMIT and after.get("parent") == PARENT, "iWork resolution commit differs")
    require(after.get("cherry_pick") == CHERRY_PICK, "iWork resolution cherry-pick differs")
    require(after.get("inventory_prerequisite") == INVENTORY_PREREQUISITE, "iWork inventory prerequisite differs")
    require(after.get("tests_passed") == 302 and after.get("tests_failed") == 0, "iWork test totals differ")
    require(after.get("cherry_pick_state_cleared") is True, "iWork cherry-pick state was not cleared")
    require(after.get("unmerged_paths") == [], "iWork resolution retains unmerged paths")
    index = after.get("index_after_sha256")
    require(isinstance(index, str) and len(index) == 64 and all(char in "0123456789abcdef" for char in index), "iWork post-resolution index receipt is malformed")
    require(after.get("scope") == "Only the eight validated iWork files committed; performance work remains unstaged.", "iWork resolution scope differs")


def check_quality_source(source: dict[str, Any]) -> None:
    require(source.get("status") == "pass" and source.get("base") == PARENT, "final iWork quality source differs")
    check_packet_descriptor(source.get("helper"), "final iWork quality helper")


def check_quality_receipt(path: Path, expected_argv: list[str], source_descriptor: dict[str, Any], helper: dict[str, Any], lock_update: dict[str, Any]) -> dict[str, Any]:
    receipt = load(str(path.relative_to(HERE)))
    require(receipt.get("exit_code") == 0, f"{path.name}: final quality command did not pass")
    require(receipt.get("argv") == expected_argv, f"{path}: command differs")
    require(receipt.get("source") == source_descriptor, f"{path}: source binding differs")
    require(receipt.get("helper") == helper, f"{path}: helper binding differs")
    require(receipt.get("quality_lock") == lock_update, f"{path}: lock binding differs")
    require(check_packet_descriptor_tree(receipt, str(path.relative_to(HERE))) >= 3, f"{path}: descriptor bindings are incomplete")
    logs = receipt.get("logs")
    require(isinstance(logs, list) and len(logs) == 2, f"{path}: logs are incomplete")
    return receipt


def check_quality(after: dict[str, Any]) -> dict[str, Any]:
    quality = load("iwa-quality.json")
    require(quality.get("status") == "pass", "final iWork quality receipt did not pass")
    final_source = descriptor(HERE / "iwa-quality-source-v4.json")
    final_helper = descriptor(HERE / "iwa-quality-v4.py")
    lock_update = descriptor(HERE / "iwa-lock-update/command.json")
    expected_commands = quality.get("commands")
    require(isinstance(expected_commands, list) and len(expected_commands) == len(FINAL_QUALITY_KINDS), "final iWork command list is incomplete")
    passing_receipts = []
    for kind, expected_record in zip(FINAL_QUALITY_KINDS, expected_commands, strict=True):
        require(isinstance(expected_record, dict) and isinstance(expected_record.get("argv"), list), f"final iWork {kind} record is malformed")
        receipt = check_quality_receipt(HERE / FINAL_QUALITY / kind / "command.json", expected_record["argv"], final_source, final_helper, lock_update)
        require(receipt == expected_record, f"final iWork {kind} record differs from its command receipt")
        passing_receipts.append(receipt)

    failed_receipts = 0
    failed_attempt_commands = 0
    for attempt in FAILED_ATTEMPT_DIRS:
        root = HERE / attempt
        require(root.is_dir(), f"retained failed quality attempt is missing: {attempt}")
        commands = sorted(root.glob("*/command.json"))
        require(commands, f"retained failed quality attempt has no commands: {attempt}")
        for command in commands:
            receipt = load(str(command.relative_to(HERE)))
            require(isinstance(receipt.get("exit_code"), int), f"{command}: exit code is missing")
            require(check_packet_descriptor_tree(receipt, str(command.relative_to(HERE))) >= 3, f"{command}: descriptor bindings are incomplete")
            failed_attempt_commands += 1
            if receipt["exit_code"] != 0:
                failed_receipts += 1
    require(failed_receipts >= 3, "retained failed iWork quality attempts are incomplete")

    return {
        "final_attempt": FINAL_QUALITY,
        "final_commands": len(passing_receipts),
        "final_commands_passed": all(receipt.get("exit_code") == 0 for receipt in passing_receipts),
        "failed_attempt_directories": list(FAILED_ATTEMPT_DIRS),
        "failed_attempt_commands_checked": failed_attempt_commands,
        "failed_receipts_retained": failed_receipts,
        "tests_passed": after["tests_passed"],
    }


def check_lock_and_format() -> dict[str, Any]:
    lock = load("iwa-quality-lock.json")
    require(lock.get("status") == "pass", "iWork quality lock receipt did not pass")
    retained = lock.get("retained")
    require(isinstance(retained, dict) and retained.get("path") == "docs/performance/results/change-0913/iwa-Cargo.lock", "iWork retained lock path differs")
    check_packet_descriptor(retained, "iWork retained lock")

    update = load("iwa-lock-update/command.json")
    require(update.get("status") == "pass" and update.get("exit_code") == 0, "iWork lock update did not pass")
    require(update.get("argv") == ["cargo", "update", "--offline", "--workspace"], "iWork lock update command differs")
    require(update.get("baseline_registry_packages_not_retained") == [], "iWork lock update dropped registry packages")
    require(update.get("before", {}).get("path") == "docs/performance/results/change-0913/iwa-Cargo.lock", "iWork lock update baseline differs")
    require(update.get("after", {}).get("path") == "docs/performance/results/change-0913/iwa-Cargo-v3.lock", "iWork lock update result differs")
    require(check_packet_descriptor_tree(update, "iwa-lock-update") >= 4, "iWork lock update descriptors are incomplete")

    final_format = load("iwa-format-final.json")
    require(final_format.get("status") == "pass", "final iWork format receipt did not pass")
    scoped_format = load("iwa-format-scoped.json")
    require(scoped_format.get("status") == "fail", "scoped iWork format receipt changed unexpectedly")
    require("No new rustfmt deltas" in str(scoped_format.get("scope")), "scoped iWork format qualification is missing")
    return {
        "retained_lock": descriptor(HERE / "iwa-Cargo.lock"),
        "updated_lock": descriptor(HERE / "iwa-Cargo-v3.lock"),
        "lock_update": descriptor(HERE / "iwa-lock-update/command.json"),
        "final_format": descriptor(HERE / "iwa-format-final.json"),
        "scoped_format_status": scoped_format["status"],
    }


def check_cleanup() -> dict[str, Any]:
    cleanup = load("iwa-cleanup.json")
    require(cleanup.get("status") == "pass", "iWork cleanup did not pass")
    require(cleanup.get("worktrees_created") == 1 and cleanup.get("worktrees_removed") == 1, "iWork cleanup worktree counts differ")
    removed = cleanup.get("removed")
    require(isinstance(removed, list) and len(removed) == 2, "iWork cleanup receipt is incomplete")
    expected_paths = {
        "/home/zhuhe/code/litchi-worktrees/0913-iwa-cherry-resolution",
        "/home/zhuhe/code/litchi/target/litchi-0913-iwa-build",
    }
    require({row.get("path") for row in removed if isinstance(row, dict)} == expected_paths, "iWork cleanup roots differ")
    for row in removed:
        require(isinstance(row, dict) and type(row.get("files")) is int and row["files"] >= 0, "iWork cleanup file count is malformed")
        require(type(row.get("logical_bytes")) is int and row["logical_bytes"] >= 0, "iWork cleanup byte count is malformed")
    return {"status": cleanup["status"], "roots_removed": len(removed), "logical_bytes_removed": sum(row["logical_bytes"] for row in removed)}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True, help="new replay result JSON")
    args = parser.parse_args()
    output = args.output.resolve()
    try:
        require(not output.exists() and not output.is_symlink(), f"refusing to overwrite replay output: {output}")
        before = load("iwa-resolution-before.json")
        after = load("iwa-resolution-after.json")
        source = load("iwa-quality-source-v4.json")
        formatted = load("iwa-format-final.json")
        check_resolution_before(before)
        check_resolution_after(after)
        check_quality_source(source)
        immutable = immutable_commit_check(after, source, formatted)
        quality = check_quality(after)
        lock_and_format = check_lock_and_format()
        cleanup = check_cleanup()
        result = {
            "schema": "litchi.performance.0913.iwa-replay.v1",
            "status": "pass",
            "packet": descriptor(HERE / "iwa-resolution-after.json"),
            "immutable_commit": immutable,
            "quality": quality,
            "lock_and_format": lock_and_format,
            "cleanup": cleanup,
            "current_worktree_source_checked": False,
            "current_index_checked": False,
            "build_or_test_or_workload_run": False,
        }
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2, sort_keys=True)
            stream.write("\n")
    except (ReplayError, OSError, subprocess.CalledProcessError, json.JSONDecodeError) as error:
        print(f"iwa-replay.py: error: {error}", file=sys.stderr)
        return 2
    print("0913 iWork immutable-commit replay passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
