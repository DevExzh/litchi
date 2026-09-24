#!/usr/bin/env python3
"""Verify retained reference-metadata gate receipts.

No expected test total is embedded here.  The verifier checks every retained
Cargo command, log digest, summary line, focused target, source map, and
workspace closure.  ``--allow-pending`` is intended only for the pre-freeze
state, when stage/run receipts do not exist yet; it never converts a present
but failing receipt into success.
"""

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
GATE_LOCK_SHA256 = "58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3"
SUMMARY = re.compile(r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored")


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
    return json.loads(path.read_text(encoding="utf-8"))


def equal(label: str, observed, expected) -> None:
    if observed != expected:
        raise RuntimeError(f"{label}: expected {expected!r}, observed {observed!r}")


def git_digest(commit: str, relative: str) -> str | None:
    try:
        content = subprocess.check_output(
            ["git", "show", f"{commit}:{relative}"],
            cwd=ROOT,
            stderr=subprocess.DEVNULL,
        )
    except subprocess.CalledProcessError:
        return None
    return hashlib.sha256(content).hexdigest()


INCLUDE = re.compile(r'include_(?:bytes|str)!\(\s*"([^"\n]+)"')


def git_bytes(commit: str, relative: str) -> bytes | None:
    try:
        return subprocess.check_output(
            ["git", "show", f"{commit}:{relative}"],
            cwd=ROOT,
            stderr=subprocess.DEVNULL,
        )
    except subprocess.CalledProcessError:
        return None


def include_dependencies(commit: str, selected: dict[str, str]) -> set[str]:
    """Derive literal include files from baseline plus selected source text."""

    source_paths: set[str] = set()
    listing = subprocess.check_output(
        ["git", "ls-tree", "-r", "--name-only", commit, "crates"],
        cwd=ROOT,
        text=True,
    )
    # run.manifest scans the litchi-ods candidate package for literal
    # dependencies.  Its workspace map still contains every crate source, but
    # include! closure is intentionally derived from baseline litchi-ods text
    # plus selected litchi-ods candidate text.
    source_paths.update(
        path
        for path in listing.splitlines()
        if path.startswith("crates/litchi-ods/") and path.endswith(".rs")
    )
    source_paths.update(
        relative
        for relative in selected
        if relative.startswith("crates/litchi-ods/") and relative.endswith(".rs")
    )
    dependencies: set[str] = set()
    for relative in sorted(source_paths):
        data = git_bytes(commit, relative)
        candidate = ROOT / relative
        selected_text = relative in selected and candidate.is_file()
        if selected_text:
            data = candidate.read_bytes()
        if data is None:
            continue
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            continue
        for included in INCLUDE.findall(text):
            resolved = Path(os.path.normpath(str(Path(relative).parent / included)))
            if resolved.is_absolute() or ".." in resolved.parts:
                continue
            dependency = str(resolved)
            exists = (
                (ROOT / dependency).is_file()
                if selected_text
                else git_bytes(commit, dependency) is not None
            )
            if exists:
                dependencies.add(dependency)
    return dependencies


def expected_workspace_paths(commit: str, selected: dict[str, str]) -> set[str]:
    paths = {"Cargo.toml", "Cargo.lock"}
    listing = subprocess.check_output(
        ["git", "ls-tree", "-r", "--name-only", commit, "crates"],
        cwd=ROOT,
        text=True,
    )
    for relative in listing.splitlines():
        if relative.endswith((".rs", "Cargo.toml", "build.rs")):
            paths.add(relative)
    paths.update(include_dependencies(commit, selected))
    return paths


def verify_sources(freeze: dict, staged: dict, before: dict, after: dict) -> list[str]:
    baseline = json.loads((EVIDENCE / "baseline.json").read_text(encoding="utf-8"))["commit"]
    equal("freeze base commit", freeze.get("base_commit"), baseline)
    equal("source snapshot stability", before, after)
    equal("boundary tool stability", before.get("boundary_tool_sha256"), after.get("boundary_tool_sha256"))
    equal(
        "boundary tool hash",
        before.get("boundary_tool_sha256"),
        digest(ROOT / "tools/check_crate_boundaries.py"),
    )
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict) or not selected:
        raise RuntimeError("freeze selected_files is empty")
    equal("full staged source path set", set(staged), set(selected))
    for relative, expected in selected.items():
        if not isinstance(expected, str):
            raise RuntimeError(f"missing hash for selected source {relative}")
        equal(f"staged source {relative}", staged.get(relative), expected)
        source = HERE / "Cargo.lock" if relative == "Cargo.lock" else ROOT / relative
        if not source.is_file():
            raise RuntimeError(f"selected source is absent: {relative}")
        equal(f"working source {relative}", digest(source), expected)

    manifest = before.get("source_sha256")
    if not isinstance(manifest, dict):
        raise RuntimeError("source-before source_sha256 is malformed")
    for relative, expected in selected.items():
        equal(f"manifest selected source {relative}", manifest.get(relative), expected)

    workspace = before.get("workspace_source_sha256")
    if not isinstance(workspace, dict):
        raise RuntimeError("source-before workspace_source_sha256 is malformed")
    expected_paths = expected_workspace_paths(baseline, selected)
    expected_paths.update(
        relative
        for relative in selected
        if relative.startswith("crates/")
        and relative.endswith((".rs", "Cargo.toml", "build.rs"))
    )
    equal("workspace source path set", set(workspace), expected_paths)
    changed: list[str] = []
    for relative, observed in workspace.items():
        expected = selected.get(relative) if relative in selected else git_digest(baseline, relative)
        if relative == "Cargo.lock":
            expected = GATE_LOCK_SHA256
        if expected is None:
            raise RuntimeError(f"workspace source has no baseline or selected hash: {relative}")
        equal(f"workspace source {relative}", observed, expected)
    for relative, expected in selected.items():
        if git_digest(baseline, relative) != expected:
            changed.append(relative)
    return sorted(changed)


def expected_commands(batch: list[str]) -> dict[str, list[str]]:
    return {
        "ods-tests": ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods"],
        "clippy": [
            "cargo", "clippy", "--locked", "--offline", "-p", "litchi-ods",
            "--all-targets", "--", "-D", "warnings",
        ],
        "rustdoc": ["cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--no-deps"],
        "format": ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"],
        "batch-format": [
            "rustfmt", "--edition", "2024", "--check", "--config", "skip_children=true", *batch,
        ],
        "boundaries": ["python3", "tools/check_crate_boundaries.py"],
        "diff-check": ["git", "diff", "--check"],
    }


def verify_receipts(freeze: dict, staged: dict, before: dict, after: dict) -> dict[str, object]:
    baseline = json.loads((EVIDENCE / "baseline.json").read_text(encoding="utf-8"))["commit"]
    environment = load("environment.json")
    equal("gate environment head", environment.get("head"), baseline)
    equal("gate RUSTFLAGS", environment.get("RUSTFLAGS"), None)
    equal("gate RUSTDOCFLAGS", environment.get("RUSTDOCFLAGS"), "-D warnings")
    equal("gate lock environment hash", environment.get("gate_lock_sha256"), GATE_LOCK_SHA256)
    equal("gate lock file hash", digest(HERE / "Cargo.lock"), GATE_LOCK_SHA256)

    batch = load("batch-files.json")
    if not isinstance(batch, list) or len(batch) != len(set(batch)):
        raise RuntimeError("batch-files.json is not a unique path list")
    if not all(relative in freeze["selected_files"] for relative in batch):
        raise RuntimeError("batch-format file is outside the frozen source map")
    commands = expected_commands(batch)
    results = load("results.json")
    if not isinstance(results, list):
        raise RuntimeError("results.json is not a list")
    equal("gate result names", {row.get("name") for row in results}, set(commands))
    if len(results) != len(commands):
        raise RuntimeError("gate result count does not match the seven-command protocol")
    for row in results:
        name = row["name"]
        equal(f"{name} command", row.get("command"), commands[name])
        equal(f"{name} exit code", row.get("exit_code"), 0)
        log = HERE / f"{name}.log"
        if not log.is_file():
            raise RuntimeError(f"missing gate log: {log}")
        equal(f"{name} log hash", digest(log), row.get("log_sha256"))

    log_text = (HERE / "ods-tests.log").read_text(encoding="utf-8")
    summaries = SUMMARY.findall(log_text)
    if not summaries:
        raise RuntimeError("ods-tests.log has no Cargo test summary")
    if any(status != "ok" for status, *_ in summaries):
        raise RuntimeError("ods-tests.log contains a failed test summary")
    totals = {
        "passed": sum(int(passed) for _, passed, _, _ in summaries),
        "failed": sum(int(failed) for _, _, failed, _ in summaries),
        "ignored": sum(int(ignored) for _, _, _, ignored in summaries),
    }
    focused: dict[str, int] = {}
    for relative in batch:
        path = Path(relative)
        if not (relative.startswith("crates/litchi-ods/tests/ods_formula_reference_metadata_") and path.suffix == ".rs"):
            continue
        marker = f"Running tests/{path.name}"
        if log_text.count(marker) != 1:
            raise RuntimeError(f"focused target marker count for {path.name} is not one")
        section = log_text.split(marker, 1)[1].split("Running ", 1)[0]
        rows = SUMMARY.findall(section)
        if len(rows) != 1 or rows[0][0] != "ok":
            raise RuntimeError(f"focused target summary missing for {path.name}")
        focused[path.stem] = int(rows[0][1])
    if not focused:
        raise RuntimeError("no focused reference metadata target was retained")

    verification = load("verification.json")
    equal("stable source receipt", verification.get("stable_sources"), True)
    equal("required gate receipt", verification.get("all_required_checks_passed"), True)
    return {"totals": totals, "focused": focused, "gates": len(commands)}


def pending(missing: list[str]) -> int:
    print(json.dumps({"status": "pending", "verified": False, "missing": sorted(missing)}, sort_keys=True))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--allow-pending",
        action="store_true",
        help="return a pending receipt when stage/run outputs do not exist yet",
    )
    args = parser.parse_args()
    required = [
        "freeze.json", "staged-profile-sources.json", "environment.json",
        "source-before.json", "source-after.json", "batch-files.json",
        "results.json", "verification.json", "ods-tests.log",
    ]
    missing = [name for name in required if not (HERE / name).is_file()]
    if missing:
        if args.allow_pending:
            return pending(missing)
        raise RuntimeError("gate receipts are not complete: " + ", ".join(missing))

    freeze = load("freeze.json")
    staged = load("staged-profile-sources.json")
    before = load("source-before.json")
    after = load("source-after.json")
    changed = verify_sources(freeze, staged, before, after)
    receipt = verify_receipts(freeze, staged, before, after)
    receipt.update({"status": "ok", "verified": True, "changed_paths": changed})
    print(json.dumps(receipt, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, subprocess.CalledProcessError) as error:
        raise SystemExit(f"gate verification failed: {error}")
