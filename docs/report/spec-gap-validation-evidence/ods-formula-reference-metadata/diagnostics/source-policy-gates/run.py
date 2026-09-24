#!/usr/bin/env python3
"""Run the seven locked/offline integration gates in an isolated checkout.

The checkout must already have been populated by ``stage.py``.  All output is
written below this gate directory so the verifier can hash every command log
and source snapshot.  The command intentionally performs no source edits.
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
GATE_LOCK_SHA256 = "58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3"


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def load(name: str):
    return json.loads((HERE / name).read_text(encoding="utf-8"))


def manifest(repo: Path) -> dict[str, object]:
    freeze = load("freeze.json")
    selected_paths: set[Path] = {
        repo / relative for relative in freeze["selected_files"]
    }
    workspace_paths: set[Path] = {repo / "Cargo.toml", repo / "Cargo.lock"}
    crates = repo / "crates"
    if crates.is_dir():
        # Include every Rust/build manifest in the candidate workspace.  The
        # closure is intentionally wider than src/tests so examples, benches,
        # and newly landed modules cannot evade the source-stability check.
        workspace_paths.update(
            path
            for path in crates.rglob("*")
            if path.is_file() and (path.suffix == ".rs" or path.name in {"Cargo.toml", "build.rs"})
        )
    for rust in (repo / "crates/litchi-ods").rglob("*.rs"):
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
        str(path.relative_to(repo)): digest(path)
        for path in sorted(selected_paths)
        if path.is_file()
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
    if not text:
        return
    path.write_text(
        "\n".join(line.rstrip(" \t") for line in text.splitlines()) + "\n",
        encoding="utf-8",
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("repo", type=Path, help="staged isolated checkout")
    parser.add_argument("target", type=Path, help="isolated Cargo target directory")
    args = parser.parse_args()
    repo = args.repo.resolve()
    target = args.target.resolve()
    if repo == ROOT:
        raise RuntimeError("refusing to run gates in the working tree")
    freeze = load("freeze.json")
    baseline = json.loads((EVIDENCE / "baseline.json").read_text(encoding="utf-8"))["commit"]
    head = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=repo, text=True).strip()
    if head != baseline:
        raise RuntimeError(f"isolated checkout head {head} differs from baseline {baseline}")
    if digest(repo / "Cargo.lock") != GATE_LOCK_SHA256:
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
        "gate_lock_sha256": GATE_LOCK_SHA256,
        "freeze_base_commit": freeze["base_commit"],
    }
    (HERE / "environment.json").write_text(
        json.dumps(metadata, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    before = manifest(repo)
    (HERE / "source-before.json").write_text(
        json.dumps(before, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )

    batch = load("batch-files.json")
    commands = [
        ("ods-tests", ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods"]),
        (
            "clippy",
            [
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
        ),
        (
            "rustdoc",
            ["cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--no-deps"],
        ),
        ("format", ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"]),
        (
            "batch-format",
            ["rustfmt", "--edition", "2024", "--check", "--config", "skip_children=true", *batch],
        ),
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
            completed = subprocess.run(
                command,
                cwd=repo,
                env=current_env,
                stdout=stream,
                stderr=subprocess.STDOUT,
            )
        trim_log(log)
        result = {
            "name": name,
            "command": command,
            "exit_code": completed.returncode,
            "seconds": time.monotonic() - started,
            "log_sha256": digest(log),
        }
        results.append(result)
        (HERE / "results.json").write_text(
            json.dumps(results, indent=2, sort_keys=True) + "\n", encoding="utf-8"
        )
        print(f"{name}: {completed.returncode}", flush=True)

    after = manifest(repo)
    (HERE / "source-after.json").write_text(
        json.dumps(after, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    stable = before == after
    passed = stable and all(int(row["exit_code"]) == 0 for row in results)
    (HERE / "verification.json").write_text(
        json.dumps(
            {
                "stable_sources": stable,
                "all_required_checks_passed": passed,
            },
            indent=2,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
    )
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
