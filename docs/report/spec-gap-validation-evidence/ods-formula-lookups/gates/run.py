#!/usr/bin/env python3
"""Run the locked/offline ODS gates in a staged isolated checkout.

The checkout is populated by ``stage.py``.  This runner records source maps,
command lines, logs, and dynamic test summaries; it embeds no historical
outcome totals.
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


HERE = Path(__file__).resolve().parent
EVIDENCE = HERE.parent
ROOT = HERE.parents[4]


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def load(name: str):
    return json.loads((HERE / name).read_text(encoding="utf-8"))


def baseline() -> dict:
    value = json.loads((EVIDENCE / "baseline.json").read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise RuntimeError("baseline.json is not an object")
    return value


def safe_repo_path(repo: Path, relative: str) -> Path:
    if not isinstance(relative, str) or not relative or Path(relative).is_absolute() or ".." in Path(relative).parts:
        raise RuntimeError(f"selected path is not safely repository-relative: {relative!r}")
    path = (repo / relative).resolve()
    try:
        path.relative_to(repo.resolve())
    except ValueError as error:
        raise RuntimeError(f"path escapes isolated checkout: {relative}") from error
    return path


def selected_source_paths(repo: Path, selected_raw: object) -> dict[str, Path]:
    if not isinstance(selected_raw, dict) or not all(isinstance(relative, str) for relative in selected_raw):
        raise RuntimeError("freeze selected_files is malformed")
    selected = {relative: safe_repo_path(repo, relative) for relative in selected_raw}
    missing = [relative for relative, path in selected.items() if not path.is_file()]
    if missing:
        raise RuntimeError("selected source files are absent from isolated checkout:\n" + "\n".join(sorted(missing)))
    return selected


def manifest(repo: Path) -> dict[str, object]:
    freeze = load("freeze.json")
    selected_raw = freeze.get("selected_files")
    selected_paths = selected_source_paths(repo, selected_raw)
    workspace_paths: set[Path] = {repo / "Cargo.toml", repo / "Cargo.lock"}
    crates = repo / "crates"
    if crates.is_dir():
        workspace_paths.update(
            path
            for path in crates.rglob("*")
            if path.is_file() and (path.suffix == ".rs" or path.name in {"Cargo.toml", "build.rs"})
        )
    ods = repo / "crates/litchi-ods"
    for rust in ods.rglob("*.rs") if ods.is_dir() else ():
        try:
            text = rust.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        for included in re.findall(r'include_(?:bytes|str)!\(\s*"([^"\n]+)"', text):
            candidate = (rust.parent / included).resolve()
            try:
                candidate.relative_to(repo)
            except ValueError:
                continue
            if candidate.is_file():
                workspace_paths.add(candidate)
    selected = {
        relative: digest(path)
        for relative, path in sorted(selected_paths.items())
    }
    workspace = {
        str(path.relative_to(repo)): digest(path)
        for path in sorted(workspace_paths)
        if path.is_file()
    }
    return {
        "git_head": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=repo, text=True).strip(),
        "source_sha256": selected,
        "workspace_source_sha256": workspace,
        "workspace_lock_sha256": digest(repo / "Cargo.lock"),
        "boundary_tool_sha256": digest(repo / "tools/check_crate_boundaries.py"),
    }


def trim_log(path: Path) -> None:
    text = path.read_text(encoding="utf-8")
    if text:
        path.write_text("\n".join(line.rstrip(" \t") for line in text.splitlines()) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("repo", type=Path, help="staged isolated checkout")
    parser.add_argument("target", type=Path, help="isolated Cargo target directory")
    args = parser.parse_args()
    repo = args.repo.resolve()
    target = args.target.resolve()
    if repo == ROOT:
        raise RuntimeError("refusing to run gates in the working tree")
    config = baseline()
    commit = config.get("commit")
    lock_hash = config.get("gate_lock_sha256")
    if not isinstance(commit, str) or not isinstance(lock_hash, str):
        raise RuntimeError("baseline commit or gate lock hash is missing")
    freeze = load("freeze.json")
    head = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=repo, text=True).strip()
    if head != commit or freeze.get("base_commit") != commit:
        raise RuntimeError(f"isolated checkout head/base mismatch: {head} / {freeze.get('base_commit')}")
    if digest(repo / "Cargo.lock") != lock_hash:
        raise RuntimeError("isolated checkout does not contain the retained gate lock")
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
    before = manifest(repo)
    (HERE / "source-before.json").write_text(json.dumps(before, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    batch = load("batch-files.json")
    if not isinstance(batch, list) or not all(isinstance(path, str) for path in batch) or len(batch) != len(set(batch)):
        raise RuntimeError("batch-files.json is malformed")
    commands = [
        ("ods-tests", ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods"]),
        ("clippy", ["cargo", "clippy", "--locked", "--offline", "-p", "litchi-ods", "--all-targets", "--", "-D", "warnings"]),
        ("rustdoc", ["cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--no-deps"]),
        ("format", ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"]),
        ("batch-format", ["rustfmt", "--edition", "2024", "--check", "--config", "skip_children=true", *batch]),
        ("boundaries", ["python3", "tools/check_crate_boundaries.py"]),
        ("diff-check", ["git", "diff", "--check"]),
    ]
    results: list[dict[str, object]] = []
    env = dict(os.environ, CARGO_TARGET_DIR=str(target))
    for name, command in commands:
        current_env = dict(env)
        if name == "rustdoc":
            current_env["RUSTDOCFLAGS"] = "-D warnings"
        log = HERE / f"{name}.log"
        started = time.monotonic()
        with log.open("w", encoding="utf-8") as stream:
            completed = subprocess.run(command, cwd=repo, env=current_env, stdout=stream, stderr=subprocess.STDOUT)
        trim_log(log)
        results.append({"name": name, "command": command, "exit_code": completed.returncode, "seconds": time.monotonic() - started, "log_sha256": digest(log)})
        (HERE / "results.json").write_text(json.dumps(results, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(f"{name}: {completed.returncode}", flush=True)
    after = manifest(repo)
    (HERE / "source-after.json").write_text(json.dumps(after, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    stable = before == after
    passed = stable and all(int(row["exit_code"]) == 0 for row in results)
    (HERE / "verification.json").write_text(json.dumps({"stable_sources": stable, "all_required_checks_passed": passed}, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
