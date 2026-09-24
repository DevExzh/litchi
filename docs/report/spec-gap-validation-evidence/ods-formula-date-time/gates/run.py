#!/usr/bin/env python3
"""Run the locked/offline date/time gates after root promotes a freeze.

The runner is intentionally inert during preparation: without ``freeze.json``
it reports pending with ``--allow-pending`` and otherwise exits before any
command is started.  It records dynamic test totals and log hashes; no
historical outcome is embedded in this file.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import time
from typing import Any


HERE = Path(__file__).resolve().parent
EVIDENCE = HERE.parent
ROOT = HERE.parents[4]
STAGE_SCHEMA = "ods-formula-date-time-stage-v1"
BASELINE_SCHEMA = "ods-formula-date-time-baseline-v1"


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"invalid gate input {path}: {error}") from error


def load(name: str) -> Any:
    path = HERE / name
    if not path.is_file():
        raise RuntimeError(f"missing gate input: {path}")
    return read_json(path)


def baseline() -> dict[str, Any]:
    value = read_json(EVIDENCE / "baseline.json")
    if not isinstance(value, dict) or value.get("schema") != BASELINE_SCHEMA:
        raise RuntimeError("date/time baseline has an unexpected schema")
    return value


def safe_relative(value: object, label: str) -> str:
    if not isinstance(value, str) or not value or Path(value).is_absolute() or ".." in Path(value).parts:
        raise RuntimeError(f"{label} is not safely repository-relative: {value!r}")
    return value


def safe_repo_path(repo: Path, relative: str) -> Path:
    safe_relative(relative, "selected path")
    path = (repo / relative).resolve()
    try:
        path.relative_to(repo.resolve())
    except ValueError as error:
        raise RuntimeError(f"path escapes isolated checkout: {relative}") from error
    return path


def selected_source_paths(repo: Path, selected_raw: object) -> dict[str, Path]:
    if not isinstance(selected_raw, dict) or not selected_raw or not all(isinstance(path, str) for path in selected_raw):
        raise RuntimeError("frozen selected_files is malformed")
    selected = {path: safe_repo_path(repo, path) for path in selected_raw}
    missing = [path for path, value in selected.items() if not value.is_file()]
    if missing:
        raise RuntimeError("selected source files are absent from isolated checkout:\n" + "\n".join(sorted(missing)))
    return selected


def verify_selected_hashes(selected: dict[str, Path], expected: object) -> None:
    if not isinstance(expected, dict) or set(expected) != set(selected):
        raise RuntimeError("frozen selected_files do not match the isolated source paths")
    for relative, path in selected.items():
        checksum = expected.get(relative)
        if not isinstance(checksum, str) or len(checksum) != 64:
            raise RuntimeError(f"frozen selected hash is malformed: {relative}")
        if digest(path) != checksum:
            raise RuntimeError(f"isolated selected source hash does not match freeze: {relative}")


def include_dependencies(repo: Path) -> set[Path]:
    dependencies: set[Path] = set()
    crates = repo / "crates/litchi-ods"
    if not crates.is_dir():
        return dependencies
    include = re.compile(r'include_(?:bytes|str)!\(\s*"([^"\n]+)')
    for rust in crates.rglob("*.rs"):
        try:
            text = rust.read_text(encoding="utf-8")
        except (OSError, UnicodeDecodeError):
            continue
        for relative in include.findall(text):
            candidate = (rust.parent / relative).resolve()
            try:
                candidate.relative_to(repo.resolve())
            except ValueError:
                continue
            if candidate.is_file():
                dependencies.add(candidate)
    return dependencies


def manifest(repo: Path, selected_raw: object) -> dict[str, Any]:
    selected_paths = selected_source_paths(repo, selected_raw)
    workspace_paths: set[Path] = {repo / "Cargo.toml", repo / "Cargo.lock"}
    crates = repo / "crates"
    if crates.is_dir():
        workspace_paths.update(
            path
            for path in crates.rglob("*")
            if path.is_file() and (path.suffix == ".rs" or path.name in {"Cargo.toml", "build.rs"})
        )
    workspace_paths.update(include_dependencies(repo))
    workspace_paths.update(selected_paths.values())
    absent = [path for path in workspace_paths if not path.is_file()]
    if absent:
        raise RuntimeError("workspace source is absent:\n" + "\n".join(str(path) for path in sorted(absent)))
    return {
        "git_head": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=repo, text=True).strip(),
        "source_sha256": {relative: digest(path) for relative, path in sorted(selected_paths.items())},
        "workspace_source_sha256": {
            str(path.relative_to(repo)): digest(path) for path in sorted(workspace_paths)
        },
        "workspace_lock_sha256": digest(repo / "Cargo.lock"),
        "boundary_tool_sha256": digest(repo / "tools/check_crate_boundaries.py"),
    }


def trim_log(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    if text:
        path.write_text("\n".join(line.rstrip(" \t") for line in text.splitlines()) + "\n", encoding="utf-8")


def frozen_inputs(config: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    freeze_path = HERE / "freeze.json"
    if not freeze_path.is_file():
        raise FileNotFoundError("date/time freeze.json is absent; root must promote the staged manifest first")
    freeze = read_json(freeze_path)
    stage = load("stage-manifest.json")
    if not isinstance(freeze, dict) or not isinstance(stage, dict):
        raise RuntimeError("freeze or stage manifest is malformed")
    selected = validate_freeze_identity(config, freeze, stage)
    return freeze, stage, selected


def validate_freeze_identity(config: dict[str, Any], freeze: dict[str, Any], stage: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(freeze, dict) or not isinstance(stage, dict):
        raise RuntimeError("freeze or stage manifest is malformed")
    if freeze.get("schema") != "ods-formula-date-time-freeze-v1":
        raise RuntimeError("freeze schema is not the date/time freeze schema")
    preparation = config.get("preparation_commit")
    if freeze.get("base_commit") != preparation or stage.get("base_commit") != preparation:
        raise RuntimeError("freeze/stage base commit does not match preparation_commit")
    production = config.get("production_commit")
    if freeze.get("production_commit") != production or stage.get("production_commit") != production:
        raise RuntimeError("freeze/stage production commit does not match production_commit")
    selected = freeze.get("selected_files")
    if selected != stage.get("selected_files"):
        raise RuntimeError("freeze selected_files differ from the reviewed stage manifest")
    lock = config.get("isolated_lock")
    if not isinstance(lock, dict) or freeze.get("isolated_lock_sha256") != lock.get("sha256"):
        raise RuntimeError("freeze isolated lock hash does not match the date/time baseline")
    return selected


def expected_commands(batch: list[str]) -> dict[str, list[str]]:
    return {
        "ods-tests": ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods"],
        "clippy": [
            "cargo",
            "clippy",
            "--locked",
            "--offline",
            "-p",
            "litchi-ods",
            "--all-targets",
            "--",
            "-D",
            "warnings",
        ],
        "rustdoc": ["cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--no-deps"],
        "format": ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"],
        "batch-format": ["rustfmt", "--edition", "2024", "--check", "--config", "skip_children=true", *batch],
        "boundaries": ["python3", "tools/check_crate_boundaries.py"],
        "diff-check": ["git", "diff", "--check"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("repo", type=Path, help="staged isolated checkout")
    parser.add_argument("target", type=Path, help="isolated Cargo target directory")
    parser.add_argument("--allow-pending", action="store_true", help="report absent freeze without running gates")
    args = parser.parse_args()

    try:
        config = baseline()
        freeze, stage, selected = frozen_inputs(config)
    except FileNotFoundError as error:
        result = {"status": "pending", "verified": False, "pending": [str(error)]}
        if args.allow_pending:
            print(json.dumps(result, sort_keys=True))
            return 0
        raise SystemExit(str(error)) from error

    repo = args.repo.resolve()
    target = args.target.resolve()
    if repo == ROOT:
        raise RuntimeError("refusing to run gates in the working tree")
    if not repo.is_dir():
        raise RuntimeError(f"isolated checkout is absent: {repo}")
    preparation = config["preparation_commit"]
    lock_hash = config["isolated_lock"]["sha256"]
    head = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=repo, text=True).strip()
    if head != preparation or digest(repo / "Cargo.lock") != lock_hash:
        raise RuntimeError("isolated checkout head or retained lock does not match the staged preparation")
    if freeze.get("isolated_lock_sha256") != lock_hash or stage.get("isolated_lock", {}).get("sha256") != lock_hash:
        raise RuntimeError("freeze/stage lock identity does not match the baseline")
    if not isinstance(selected, dict):
        raise RuntimeError("frozen selected_files is malformed")
    batch = stage.get("batch_files")
    if not isinstance(batch, list) or not all(isinstance(path, str) for path in batch) or len(batch) != len(set(batch)):
        raise RuntimeError("stage batch-files declaration is malformed")
    if any(path not in selected for path in batch):
        raise RuntimeError("batch-format file is outside the frozen source map")
    selected_paths = selected_source_paths(repo, selected)
    verify_selected_hashes(selected_paths, selected)

    metadata = {
        "repo": str(repo),
        "target": str(target),
        "head": head,
        "rustc": subprocess.check_output(["rustc", "-Vv"], cwd=repo, text=True).strip(),
        "cargo": subprocess.check_output(["cargo", "-V"], cwd=repo, text=True).strip(),
        "kernel": subprocess.check_output(["uname", "-a"], text=True).strip(),
        "RUSTFLAGS": os.environ.get("RUSTFLAGS"),
        "RUSTDOCFLAGS": "-D warnings",
        "gate_lock_sha256": lock_hash,
        "freeze_base_commit": freeze["base_commit"],
    }
    (HERE / "environment.json").write_text(json.dumps(metadata, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    before = manifest(repo, selected)
    (HERE / "source-before.json").write_text(json.dumps(before, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    commands = expected_commands(batch)
    results: list[dict[str, Any]] = []
    env = dict(os.environ, CARGO_TARGET_DIR=str(target))
    for name, command in commands.items():
        current_env = dict(env)
        if name == "rustdoc":
            current_env["RUSTDOCFLAGS"] = "-D warnings"
        log = HERE / f"{name}.log"
        started = time.monotonic()
        with log.open("w", encoding="utf-8") as stream:
            completed = subprocess.run(command, cwd=repo, env=current_env, stdout=stream, stderr=subprocess.STDOUT)
        trim_log(log)
        results.append(
            {
                "name": name,
                "command": command,
                "exit_code": completed.returncode,
                "seconds": time.monotonic() - started,
                "log_sha256": digest(log),
            }
        )
        (HERE / "results.json").write_text(json.dumps(results, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(f"{name}: {completed.returncode}", flush=True)

    after = manifest(repo, selected)
    (HERE / "source-after.json").write_text(json.dumps(after, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    stable = before == after
    passed = stable and all(int(row["exit_code"]) == 0 for row in results)
    (HERE / "verification.json").write_text(
        json.dumps({"stable_sources": stable, "all_required_checks_passed": passed}, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
