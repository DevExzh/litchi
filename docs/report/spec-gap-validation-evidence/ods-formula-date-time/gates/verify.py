#!/usr/bin/env python3
"""Fail-closed verifier for the date/time isolated gate receipts.

The verifier accepts no result while the root-owned freeze is absent.  Once a
freeze exists it checks the preparation commit, retained lock, staged source
closure, source stability, exact command set, and dynamic test summaries.  It
does not contain historical test totals or lookup-batch identifiers.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
from typing import Any


HERE = Path(__file__).resolve().parent
EVIDENCE = HERE.parent
ROOT = HERE.parents[4]
BASELINE_SCHEMA = "ods-formula-date-time-baseline-v1"
STAGE_SCHEMA = "ods-formula-date-time-stage-v1"
SUMMARY = re.compile(r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored")
INCLUDE = re.compile(r'include_(?:bytes|str)!\(\s*"([^"\n]+)')
SHA256 = re.compile(r"[0-9a-f]{64}\Z")


class VerificationError(RuntimeError):
    """A retained input or receipt is malformed or inconsistent."""


class PendingReceipt(RuntimeError):
    """Preparation is complete only when root has not yet produced a receipt."""


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def read_json(path: Path) -> Any:
    if not path.is_file():
        raise PendingReceipt(f"missing gate receipt: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise VerificationError(f"invalid JSON receipt {path}: {error}") from error


def load(name: str) -> Any:
    return read_json(HERE / name)


def equal(label: str, observed: Any, expected: Any) -> None:
    if observed != expected:
        raise VerificationError(f"{label}: expected {expected!r}, observed {observed!r}")


def safe_relative(value: object, label: str) -> str:
    if not isinstance(value, str) or not value or Path(value).is_absolute() or ".." in Path(value).parts:
        raise VerificationError(f"{label} is not safely repository-relative: {value!r}")
    return value


def safe_repo_path(repo: Path, relative: str) -> Path:
    safe_relative(relative, "selected path")
    path = (repo / relative).resolve()
    try:
        path.relative_to(repo.resolve())
    except ValueError as error:
        raise VerificationError(f"selected path escapes isolated checkout: {relative}") from error
    return path


def baseline() -> dict[str, Any]:
    value = read_json(EVIDENCE / "baseline.json")
    if not isinstance(value, dict) or value.get("schema") != BASELINE_SCHEMA:
        raise VerificationError("date/time baseline has an unexpected schema")
    if not isinstance(value.get("preparation_commit"), str) or not value["preparation_commit"]:
        raise VerificationError("date/time preparation_commit is missing")
    lock = value.get("isolated_lock")
    if not isinstance(lock, dict) or not isinstance(lock.get("path"), str) or not isinstance(lock.get("sha256"), str):
        raise VerificationError("date/time isolated_lock is malformed")
    safe_relative(lock["path"], "isolated lock path")
    if SHA256.fullmatch(lock["sha256"]) is None:
        raise VerificationError("date/time isolated lock hash is malformed")
    return value


def git_bytes(commit: str, relative: str) -> bytes | None:
    try:
        return subprocess.check_output(["git", "show", f"{commit}:{relative}"], cwd=ROOT, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError:
        return None


def include_dependencies(commit: str, selected: dict[str, str]) -> set[str]:
    listing = subprocess.check_output(["git", "ls-tree", "-r", "--name-only", commit, "crates"], cwd=ROOT, text=True)
    source_paths = {path for path in listing.splitlines() if path.startswith("crates/litchi-ods/") and path.endswith(".rs")}
    source_paths.update(path for path in selected if path.startswith("crates/litchi-ods/") and path.endswith(".rs"))
    dependencies: set[str] = set()
    for relative in sorted(source_paths):
        data = git_bytes(commit, relative)
        current = ROOT / relative
        if relative in selected and current.is_file():
            data = current.read_bytes()
        if data is None:
            continue
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            continue
        for included in INCLUDE.findall(text):
            path = Path(os.path.normpath(str(Path(relative).parent / included)))
            if path.is_absolute() or ".." in path.parts:
                continue
            resolved = str(path)
            if (ROOT / resolved).is_file() or git_bytes(commit, resolved) is not None:
                dependencies.add(resolved)
    return dependencies


def expected_workspace_paths(commit: str, selected: dict[str, str]) -> set[str]:
    paths = {"Cargo.toml", "Cargo.lock"}
    listing = subprocess.check_output(["git", "ls-tree", "-r", "--name-only", commit, "crates"], cwd=ROOT, text=True)
    paths.update(path for path in listing.splitlines() if path.endswith((".rs", "Cargo.toml", "build.rs")))
    # The runner adds every frozen selected input to its custody manifest,
    # including the global/package feature matrices and authored evidence.
    paths.update(selected)
    paths.update(include_dependencies(commit, selected))
    return paths


def validate_map(value: Any, label: str) -> dict[str, str]:
    if not isinstance(value, dict) or not value:
        raise VerificationError(f"{label} is empty or malformed")
    output: dict[str, str] = {}
    for relative, checksum in value.items():
        safe_relative(relative, label)
        if SHA256.fullmatch(checksum) is None:
            raise VerificationError(f"{label} hash is malformed: {relative}")
        output[relative] = checksum
    return output


def verify_sources(config: dict[str, Any], freeze: dict[str, Any], stage: dict[str, Any], before: dict[str, Any], after: dict[str, Any]) -> None:
    preparation = config["preparation_commit"]
    lock_hash = config["isolated_lock"]["sha256"]
    equal("freeze schema", freeze.get("schema"), "ods-formula-date-time-freeze-v1")
    equal("freeze base commit", freeze.get("base_commit"), preparation)
    equal("stage schema", stage.get("schema"), STAGE_SCHEMA)
    equal("stage base commit", stage.get("base_commit"), preparation)
    equal("freeze isolated lock", freeze.get("isolated_lock_sha256"), lock_hash)
    selected = validate_map(freeze.get("selected_files"), "freeze selected_files")
    staged = validate_map(stage.get("selected_files"), "stage selected_files")
    equal("freeze/stage selected path set", set(selected), set(staged))
    for relative, checksum in selected.items():
        equal(f"staged source {relative}", staged[relative], checksum)
        source = ROOT / relative
        lock = config["isolated_lock"]
        if relative == "Cargo.lock":
            source = ROOT / lock["path"]
        if not source.is_file():
            raise VerificationError(f"selected source is absent from working tree: {relative}")
        equal(f"working source {relative}", digest(source), checksum)

    if not isinstance(before, dict) or not isinstance(after, dict):
        raise VerificationError("source receipts are malformed")
    equal("isolated source stability", before, after)
    equal("source-before head", before.get("git_head"), preparation)
    equal("source-after head", after.get("git_head"), preparation)
    before_selected = validate_map(before.get("source_sha256"), "source-before selected_files")
    equal("source-before selected path set", set(before_selected), set(selected))
    for relative, checksum in selected.items():
        equal(f"source-before {relative}", before_selected[relative], checksum)

    workspace = validate_map(before.get("workspace_source_sha256"), "workspace source map")
    expected = expected_workspace_paths(preparation, selected)
    equal("workspace source path set", set(workspace), expected)
    for relative, observed in workspace.items():
        source = ROOT / relative
        if relative == "Cargo.lock":
            source = ROOT / config["isolated_lock"]["path"]
        if relative in selected and relative != "Cargo.lock":
            expected_hash = selected[relative]
        elif relative == "Cargo.lock":
            expected_hash = lock_hash
        else:
            data = git_bytes(preparation, relative)
            expected_hash = digest(source) if data is None and source.is_file() else hashlib.sha256(data).hexdigest() if data is not None else None
        if expected_hash is None:
            raise VerificationError(f"workspace source lacks a baseline hash: {relative}")
        equal(f"workspace source {relative}", observed, expected_hash)
    boundary = ROOT / "tools/check_crate_boundaries.py"
    if not boundary.is_file():
        raise VerificationError("boundary checker is absent")
    equal("boundary checker hash stability", before.get("boundary_tool_sha256"), after.get("boundary_tool_sha256"))
    equal("boundary checker hash", before.get("boundary_tool_sha256"), digest(boundary))
    equal("workspace lock hash", before.get("workspace_lock_sha256"), lock_hash)


def expected_commands(batch: list[str]) -> dict[str, list[str]]:
    return {
        "ods-tests": ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods"],
        "clippy": ["cargo", "clippy", "--locked", "--offline", "-p", "litchi-ods", "--all-targets", "--", "-D", "warnings"],
        "rustdoc": ["cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--no-deps"],
        "format": ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"],
        "batch-format": ["rustfmt", "--edition", "2024", "--check", "--config", "skip_children=true", *batch],
        "boundaries": ["python3", "tools/check_crate_boundaries.py"],
        "diff-check": ["git", "diff", "--check"],
    }


def verify_receipts(config: dict[str, Any], freeze: dict[str, Any], stage: dict[str, Any], batch: list[str]) -> dict[str, Any]:
    environment = load("environment.json")
    equal("gate environment head", environment.get("head"), config["preparation_commit"])
    equal("gate RUSTFLAGS", environment.get("RUSTFLAGS"), None)
    equal("gate RUSTDOCFLAGS", environment.get("RUSTDOCFLAGS"), "-D warnings")
    equal("gate lock environment hash", environment.get("gate_lock_sha256"), config["isolated_lock"]["sha256"])
    lock = HERE / "Cargo.lock"
    if lock.is_file():
        equal("gate lock file hash", digest(lock), config["isolated_lock"]["sha256"])
    equal("gate batch declaration", batch, stage.get("batch_files"))
    selected = validate_map(freeze.get("selected_files"), "freeze selected_files")
    if any(path not in selected for path in batch):
        raise VerificationError("batch-format file is outside the frozen source map")
    commands = expected_commands(batch)
    results = load("results.json")
    if not isinstance(results, list) or len(results) != len(commands) or not all(isinstance(row, dict) for row in results):
        raise VerificationError("gate results do not contain the complete command set")
    names = [row.get("name") for row in results]
    equal("gate result name set", set(names), set(commands))
    for row in results:
        if row.get("exit_code") != 0:
            raise VerificationError(f"retained gate did not pass: {row!r}")
        name = row.get("name")
        if name not in commands:
            raise VerificationError(f"unknown gate result name: {name!r}")
        equal(f"{name} command", row.get("command"), commands[name])
        log = HERE / f"{name}.log"
        if not log.is_file():
            raise VerificationError(f"missing gate log: {name}")
        equal(f"{name} log hash", digest(log), row.get("log_sha256"))
    test_log = HERE / "ods-tests.log"
    if not test_log.is_file():
        raise VerificationError("ods-tests.log is absent")
    summaries = SUMMARY.findall(test_log.read_text(encoding="utf-8"))
    if not summaries or any(status != "ok" or int(failed) != 0 for status, _, failed, _ in summaries):
        raise VerificationError("ods-tests.log has no wholly passing test summaries")
    verification = load("verification.json")
    equal("stable source receipt", verification.get("stable_sources"), True)
    equal("required gate receipt", verification.get("all_required_checks_passed"), True)
    return {
        "commands": len(results),
        "test_summaries": len(summaries),
        "tests": {
            "passed": sum(int(passed) for _, passed, _, _ in summaries),
            "failed": sum(int(failed) for _, _, failed, _ in summaries),
            "ignored": sum(int(ignored) for _, _, _, ignored in summaries),
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--allow-pending", action="store_true", help="report missing receipts without claiming verification")
    args = parser.parse_args()
    try:
        config = baseline()
        required = (
            "freeze.json",
            "stage-manifest.json",
            "staged-profile-sources.json",
            "batch-files.json",
            "environment.json",
            "source-before.json",
            "source-after.json",
            "results.json",
            "verification.json",
            "ods-tests.log",
        )
        missing = [name for name in required if not (HERE / name).is_file()]
        if missing:
            raise PendingReceipt("gate receipts are absent: " + ", ".join(missing))
        freeze = load("freeze.json")
        stage = load("stage-manifest.json")
        before = load("source-before.json")
        after = load("source-after.json")
        batch = load("batch-files.json")
        if not isinstance(batch, list) or not all(isinstance(path, str) for path in batch):
            raise VerificationError("batch-files.json is malformed")
        verify_sources(config, freeze, stage, before, after)
        receipt = verify_receipts(config, freeze, stage, batch)
        result = {"status": "ok", "verified": True, "receipt": receipt}
    except PendingReceipt as error:
        if not args.allow_pending:
            raise SystemExit(f"date/time gate verification failed: {error}") from error
        result = {"status": "pending", "verified": False, "pending": [str(error)]}
    except (OSError, VerificationError, subprocess.CalledProcessError) as error:
        raise SystemExit(f"date/time gate verification failed: {error}") from error
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
