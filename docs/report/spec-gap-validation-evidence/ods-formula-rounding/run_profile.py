#!/usr/bin/env python3
"""Build and capture the bounded ODS rounding profile.

The release binary is built in a fresh external target directory. Every lane
is a fresh process wrapped by ``/usr/bin/time -v`` so maximum RSS is scoped to
one case and phase. The target directory is removed after a successful or
failed run; source and result evidence stays in this directory.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
from typing import Any


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
HARNESS = HERE / "harness" / "Cargo.toml"
RESULTS = HERE / "results"
SOURCE_FILES = [
    "crates/litchi-ods/src/codec/formula/evaluation.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/rounding.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value.rs",
    "crates/litchi-ods/Cargo.toml",
    "crates/litchi-core/Cargo.toml",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/harness/Cargo.toml",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/harness/Cargo.lock",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/harness/src/main.rs",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/run_profile.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/verify.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-rounding/summarize.py",
]
CASES = [
    "scalar-control",
    "scalar-int",
    "scalar-floor",
    "scalar-round",
    "scalar-rounddown",
    "scalar-ceiling",
    "scalar-mround",
    "scalar-roundup",
    "scalar-trunc",
    "array-control-4x4",
    "array-int-4x4",
    "array-floor-4x4",
    "array-round-4x4",
    "array-rounddown-4x4",
    "array-ceiling-4x4",
    "array-mround-4x4",
    "array-roundup-4x4",
    "array-trunc-4x4",
    "array-roundup-16x16",
]
PHASES = ["evaluate", "parse-evaluate"]


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def source_snapshot() -> dict[str, Any]:
    raw_status = subprocess.check_output(
        ["git", "status", "--short", "--untracked-files=all"], cwd=ROOT
    )
    evidence_prefix = str(HERE.relative_to(ROOT)).encode() + b"/"
    status = b"\n".join(
        line
        for line in raw_status.splitlines()
        if not line[3:].startswith(evidence_prefix)
    )
    if status:
        status += b"\n"
    return {
        "git_head": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip(),
        "dirty_status_sha256": hashlib.sha256(status).hexdigest(),
        "dirty_status_bytes": len(status),
        "source_sha256": {name: digest(ROOT / name) for name in SOURCE_FILES},
    }


def command_text(command: list[str]) -> str:
    return " ".join(subprocess.list2cmdline([part]) for part in command)


def run_checked(command: list[str], *, cwd: Path, env: dict[str, str], stdout: Path | None = None) -> None:
    output = None if stdout is None else stdout.open("w", encoding="utf-8")
    try:
        result = subprocess.run(command, cwd=cwd, env=env, stdout=output, text=True)
    finally:
        if output is not None:
            output.close()
    if result.returncode != 0:
        raise RuntimeError(f"command failed with {result.returncode}: {command_text(command)}")


def rss_kib(path: Path) -> int:
    text = path.read_text(encoding="utf-8")
    match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", text)
    if match is None:
        raise RuntimeError(f"missing maximum RSS in {path}")
    return int(match.group(1))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--iterations", type=int, default=20)
    args = parser.parse_args()
    if args.warmups < 0 or args.iterations <= 0:
        parser.error("--warmups must be nonnegative and --iterations must be positive")

    RESULTS.mkdir(parents=True, exist_ok=True)
    for path in RESULTS.iterdir():
        if path.is_file():
            path.unlink()
        elif path.is_dir():
            shutil.rmtree(path)

    before = source_snapshot()
    target = Path(tempfile.mkdtemp(prefix="litchi-ods-rounding-target-", dir="/var/tmp"))
    environment = {
        "rustc": subprocess.check_output(["rustc", "--version"], text=True).strip(),
        "cargo": subprocess.check_output(["cargo", "--version"], text=True).strip(),
        "kernel": os.uname().release,
        "cpu": subprocess.check_output(["lscpu"], text=True),
        "target_dir": str(target),
        "warmups": args.warmups,
        "iterations": args.iterations,
        "cases": CASES,
        "phases": PHASES,
    }
    (RESULTS / "environment.json").write_text(json.dumps(environment, indent=2) + "\n")
    build_env = os.environ.copy()
    build_env["CARGO_TARGET_DIR"] = str(target)
    build_command = [
        "cargo",
        "build",
        "--manifest-path",
        str(HARNESS),
        "--release",
    ]
    (RESULTS / "commands.txt").write_text(command_text(build_command) + "\n")
    try:
        run_checked(build_command, cwd=ROOT, env=build_env, stdout=RESULTS / "build.log")
        binary = target / "release" / "ods-formula-rounding-performance"
        if not binary.is_file():
            raise RuntimeError(f"missing release binary: {binary}")
        binary_hash = digest(binary)
        records: list[dict[str, Any]] = []
        for case in CASES:
            for phase in PHASES:
                stem = f"{case}.{phase}"
                stdout_path = RESULTS / f"{stem}.stdout.json"
                stderr_path = RESULTS / f"{stem}.stderr.log"
                time_path = RESULTS / f"{stem}.time.txt"
                command = [
                    "/usr/bin/time",
                    "-v",
                    "-o",
                    str(time_path),
                    str(binary),
                    "--case",
                    case,
                    "--phase",
                    phase,
                    "--warmups",
                    str(args.warmups),
                    "--iterations",
                    str(args.iterations),
                ]
                with (RESULTS / "commands.txt").open("a", encoding="utf-8") as commands:
                    commands.write(command_text(command) + "\n")
                with stdout_path.open("w", encoding="utf-8") as stdout, stderr_path.open(
                    "w", encoding="utf-8"
                ) as stderr:
                    completed = subprocess.run(
                        command, cwd=ROOT, env=build_env, stdout=stdout, stderr=stderr, text=True
                    )
                if completed.returncode != 0:
                    raise RuntimeError(f"lane failed with {completed.returncode}: {stem}")
                lines = [line for line in stdout_path.read_text().splitlines() if line.strip()]
                if len(lines) != 1:
                    raise RuntimeError(f"lane {stem} produced {len(lines)} JSON lines")
                row = json.loads(lines[0])
                row["rss_kib"] = rss_kib(time_path)
                row["binary_sha256"] = binary_hash
                row["source_git_head"] = before["git_head"]
                records.append(row)
        after = source_snapshot()
        source_hashes_unchanged = before["source_sha256"] == after["source_sha256"]
        if not source_hashes_unchanged:
            raise RuntimeError("rounding sources changed during capture")
        source_receipt = {
            "before": before,
            "after": after,
            "source_hashes_unchanged": source_hashes_unchanged,
            "unchanged": source_hashes_unchanged and before["git_head"] == after["git_head"],
            "binary_sha256": binary_hash,
        }
        (RESULTS / "source-manifest.json").write_text(
            json.dumps(source_receipt, indent=2, sort_keys=True) + "\n"
        )
        (RESULTS / "measurements.jsonl").write_text(
            "".join(json.dumps(record, sort_keys=True) + "\n" for record in records)
        )
    finally:
        shutil.rmtree(target, ignore_errors=True)
    cleanup = {"target_dir": str(target), "removed": not target.exists()}
    (RESULTS / "target-cleanup.json").write_text(json.dumps(cleanup, indent=2) + "\n")
    if not cleanup["removed"]:
        raise RuntimeError(f"temporary target was not removed: {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
