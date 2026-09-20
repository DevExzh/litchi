#!/usr/bin/env python3
"""Fail-closed verifier for the retained seven lookup gate receipts."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess


HERE = Path(__file__).resolve().parent
EVIDENCE = HERE.parent
ROOT = HERE.parents[4]
SUMMARY = re.compile(r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored")
INCLUDE = re.compile(r'include_(?:bytes|str)!\(\s*"([^"\n]+)')


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def load(name: str):
    path = HERE / name
    if not path.is_file():
        raise RuntimeError(f"missing gate receipt: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise RuntimeError(f"invalid gate receipt: {path}: {error}") from error


def baseline() -> dict:
    value = json.loads((EVIDENCE / "baseline.json").read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise RuntimeError("baseline.json is not an object")
    return value


CONFIG = baseline()
BASELINE_COMMIT = CONFIG.get("commit")
GATE_LOCK_SHA256 = CONFIG.get("gate_lock_sha256")
if not isinstance(BASELINE_COMMIT, str) or not isinstance(GATE_LOCK_SHA256, str):
    raise RuntimeError("baseline commit or gate lock hash is missing")


def equal(label: str, observed, expected) -> None:
    if observed != expected:
        raise RuntimeError(f"{label}: expected {expected!r}, observed {observed!r}")


def validate_relative(relative: object, label: str) -> str:
    if not isinstance(relative, str) or not relative or Path(relative).is_absolute() or ".." in Path(relative).parts:
        raise RuntimeError(f"{label} is not safely repository-relative: {relative!r}")
    return relative


def git_bytes(commit: str, relative: str) -> bytes | None:
    try:
        return subprocess.check_output(["git", "show", f"{commit}:{relative}"], cwd=ROOT, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError:
        return None


def include_dependencies(commit: str, selected: dict[str, str]) -> set[str]:
    listing = subprocess.check_output(["git", "ls-tree", "-r", "--name-only", commit, "crates"], cwd=ROOT, text=True)
    source_paths = {path for path in listing.splitlines() if path.startswith("crates/litchi-ods/") and path.endswith(".rs")}
    source_paths.update(relative for relative in selected if relative.startswith("crates/litchi-ods/") and relative.endswith(".rs"))
    dependencies: set[str] = set()
    for relative in sorted(source_paths):
        data = git_bytes(commit, relative)
        candidate = ROOT / relative
        selected_source = relative in selected and candidate.is_file()
        if selected_source:
            data = candidate.read_bytes()
        if data is None:
            continue
        try:
            source = data.decode("utf-8")
        except UnicodeDecodeError:
            continue
        for included in INCLUDE.findall(source):
            resolved = Path(os.path.normpath(str(Path(relative).parent / included)))
            if resolved.is_absolute() or ".." in resolved.parts:
                continue
            dependency = str(resolved)
            exists = (ROOT / dependency).is_file() if selected_source else git_bytes(commit, dependency) is not None
            if exists:
                dependencies.add(dependency)
    return dependencies


def expected_workspace_paths(commit: str, selected: dict[str, str]) -> set[str]:
    for relative in selected:
        validate_relative(relative, "selected workspace path")
    paths = {"Cargo.toml", "Cargo.lock"}
    listing = subprocess.check_output(["git", "ls-tree", "-r", "--name-only", commit, "crates"], cwd=ROOT, text=True)
    paths.update(path for path in listing.splitlines() if path.endswith((".rs", "Cargo.toml", "build.rs")))
    # The candidate profile may add new Rust modules, manifests, or build
    # scripts that do not exist at the retained baseline commit.  They are
    # still part of the isolated workspace closure and must be present in the
    # observed run manifest.  Include selected workspace files explicitly
    # before adding include! dependencies.
    paths.update(
        relative
        for relative in selected
        if relative in {"Cargo.toml", "Cargo.lock"}
        or (relative.startswith("crates/") and relative.endswith((".rs", "Cargo.toml", "build.rs")))
    )
    paths.update(include_dependencies(commit, selected))
    return paths


def verify_workspace_path_set(observed: set[str], expected: set[str]) -> None:
    equal("workspace source path set", observed, expected)


def verify_sources(freeze: dict, staged: dict, before: dict, after: dict) -> list[str]:
    if not all(isinstance(value, dict) for value in (freeze, staged, before, after)):
        raise RuntimeError("gate source receipts must be objects")
    equal("freeze base commit", freeze.get("base_commit"), BASELINE_COMMIT)
    equal("source snapshot stability", before, after)
    equal("boundary tool stability", before.get("boundary_tool_sha256"), after.get("boundary_tool_sha256"))
    boundary = ROOT / "tools/check_crate_boundaries.py"
    if not boundary.is_file():
        raise RuntimeError("boundary checker is absent")
    equal("boundary tool hash", before.get("boundary_tool_sha256"), digest(boundary))
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict) or not selected or not all(isinstance(relative, str) for relative in selected):
        raise RuntimeError("freeze selected_files is empty")
    for relative in selected:
        validate_relative(relative, "frozen selected path")
    if not isinstance(staged, dict):
        raise RuntimeError("staged source map is malformed")
    equal("full staged source path set", set(staged), set(selected))
    manifest = before.get("source_sha256")
    if not isinstance(manifest, dict) or not all(isinstance(relative, str) for relative in manifest):
        raise RuntimeError("source-before selected map is malformed")
    equal("source manifest selected path set", set(manifest), set(selected))
    for relative, expected in selected.items():
        if not isinstance(expected, str):
            raise RuntimeError(f"selected hash is malformed: {relative}")
        equal(f"staged source {relative}", staged.get(relative), expected)
        source = HERE / "Cargo.lock" if relative == "Cargo.lock" else ROOT / relative
        if not source.is_file():
            raise RuntimeError(f"selected source is absent: {relative}")
        equal(f"working source {relative}", digest(source), expected)
        equal(f"manifest selected source {relative}", manifest.get(relative), expected)
    workspace = before.get("workspace_source_sha256")
    if not isinstance(workspace, dict) or not all(isinstance(relative, str) for relative in workspace):
        raise RuntimeError("workspace source map is malformed")
    verify_workspace_path_set(set(workspace), expected_workspace_paths(BASELINE_COMMIT, selected))
    for relative, observed in workspace.items():
        validate_relative(relative, "workspace source path")
        expected = selected.get(relative)
        if expected is None:
            data = git_bytes(BASELINE_COMMIT, relative)
            expected = hashlib.sha256(data).hexdigest() if data is not None else None
        if relative == "Cargo.lock":
            expected = GATE_LOCK_SHA256
        if expected is None:
            raise RuntimeError(f"workspace source lacks baseline hash: {relative}")
        equal(f"workspace source {relative}", observed, expected)
    return sorted(relative for relative, expected in selected.items() if git_bytes(BASELINE_COMMIT, relative) is None or hashlib.sha256(git_bytes(BASELINE_COMMIT, relative)).hexdigest() != expected)


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


def verify_receipts(freeze: dict, staged: dict, before: dict, after: dict) -> dict[str, object]:
    environment = load("environment.json")
    equal("gate environment head", environment.get("head"), BASELINE_COMMIT)
    equal("gate RUSTFLAGS", environment.get("RUSTFLAGS"), None)
    equal("gate RUSTDOCFLAGS", environment.get("RUSTDOCFLAGS"), "-D warnings")
    equal("gate lock environment hash", environment.get("gate_lock_sha256"), GATE_LOCK_SHA256)
    equal("gate lock file hash", digest(HERE / "Cargo.lock"), GATE_LOCK_SHA256)
    batch = load("batch-files.json")
    if not isinstance(batch, list) or not all(isinstance(path, str) for path in batch):
        raise RuntimeError("batch-files.json is malformed")
    if len(batch) != len(set(batch)) or any(".." in Path(path).parts or Path(path).is_absolute() for path in batch):
        raise RuntimeError("batch-files.json contains duplicate or escaping paths")
    if not all(path in freeze.get("selected_files", {}) for path in batch):
        raise RuntimeError("batch-format file is outside the frozen source map")
    commands = expected_commands(batch)
    results = load("results.json")
    if not isinstance(results, list) or len(results) != len(commands) or not all(isinstance(row, dict) for row in results):
        raise RuntimeError("gate results do not contain the complete command set")
    names = [row.get("name") for row in results]
    equal("gate result name set", set(names), set(commands))
    for row in results:
        if not isinstance(row, dict) or row.get("exit_code") != 0:
            raise RuntimeError(f"retained gate did not pass: {row!r}")
        name = row.get("name")
        if name not in commands:
            raise RuntimeError(f"unknown gate result name: {name!r}")
        equal(f"{name} command", row.get("command"), commands[name])
        log = HERE / f"{name}.log"
        if not log.is_file():
            raise RuntimeError(f"missing gate log: {name}")
        equal(f"{name} log hash", digest(log), row.get("log_sha256"))
    text = (HERE / "ods-tests.log").read_text(encoding="utf-8")
    summaries = SUMMARY.findall(text)
    if not summaries or any(status != "ok" or int(failed) != 0 for status, _, failed, _ in summaries):
        raise RuntimeError("ods-tests.log has no wholly passing Cargo test summaries")
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
    parser = argparse.ArgumentParser()
    parser.add_argument("--allow-pending", action="store_true")
    args = parser.parse_args()
    required = ["freeze.json", "staged-profile-sources.json", "environment.json", "source-before.json", "source-after.json", "batch-files.json", "results.json", "verification.json", "ods-tests.log"]
    missing = [name for name in required if not (HERE / name).is_file()]
    if missing:
        result = {"status": "pending", "verified": False, "pending": ["gate receipts are absent: " + ", ".join(missing)]}
        if args.allow_pending:
            print(json.dumps(result, sort_keys=True))
            return 0
        raise SystemExit("lookup gate verification failed: " + result["pending"][0])
    try:
        freeze, staged, before, after = (load(name) for name in ("freeze.json", "staged-profile-sources.json", "source-before.json", "source-after.json"))
        excluded = verify_sources(freeze, staged, before, after)
        receipt = verify_receipts(freeze, staged, before, after)
        result = {"status": "ok", "verified": True, "gate_only_source_paths": excluded, "receipt": receipt}
    except (OSError, RuntimeError, subprocess.CalledProcessError) as error:
        raise SystemExit(f"lookup gate verification failed: {error}") from error
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
