#!/usr/bin/env python3
"""Capture the bounded ODS paired-statistics profile in fresh child processes.

The committed baseline is used for all matched controls; paired-statistics
rows are candidate-only additions while shared descriptive-regression rows
are measured on both sides. Every lane hashes the
compiled workspace closure and authored profile inputs before and after capture.
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
ROOT = HERE.parents[4]
REL_PROFILE = HERE.relative_to(ROOT)
REL_HARNESS = REL_PROFILE / "harness"
GATE_LOCK = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-order-statistics/gates/Cargo.lock"
GATE_LOCK_SHA256 = "58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3"
BASELINE_COMMIT = "aa48eee68cb6ee0904523394d4ff8015dce6e595"
BINARY_NAME = "ods-formula-paired-statistics-performance"
# Reviewed paired contract marker. Capture still fails closed if the file or
# any companion oracle/golden input is missing or changes.
CONTRACT_SHA256: str | None = "9ce79b7b5014c6e29c1d30c61a77187d6dc4bce9be7810e1f5006e8ef147aadd"
PHASES = ("evaluate", "parse-evaluate")
CONTROL_CASES = (
    "scalar-control-arithmetic",
    "scalar-control-sin",
    "scalar-control-imsum",
    "database-control-dsum",
    "scalar-control-average",
    "scalar-control-counta",
    "scalar-control-var",
    "scalar-control-stdev",
    "database-control-dvar",
    "database-control-dstdev",
    "array-control-4x4-arithmetic",
    "array-control-4x4-sin",
    "array-control-16x16-arithmetic",
    "array-control-16x16-sin",
    "reference-array-16x4-arithmetic",
    "scalar-aggregate-sum",
    "literal-aggregate-4x1-sum",
    "reference-aggregate-64x4-sum",
    "reference-conditional-256x4-sumifs",
    "reference-control-average",
    "reference-control-counta",
)
REPRESENTATIVE_CONTROLS = (
    "representative-median",
    "representative-rank",
    "representative-percentrank",
)
MATCHED_DESCRIPTIVE_CONTROLS = tuple(
    name
    for function in ("avedev", "devsq", "kurt", "skew", "skewp")
    for name in (
        f"matched-descriptive-reference-{function}",
        f"matched-descriptive-sensitive-inline-{function}",
        f"matched-descriptive-extreme-inline-{function}",
    )
)
MATCHED_CONTROL_CASES = CONTROL_CASES + REPRESENTATIVE_CONTROLS + MATCHED_DESCRIPTIVE_CONTROLS

# The selected manifest includes the value dispatch, scalar bridge, shared
# statistical/paired kernels, existing aggregate/conditional/database paths,
# and every authored paired-statistics test/oracle input that can affect the
# measured evaluator. Optional module paths are retained so the manifest
# remains valid if the implementation splits the paired owner.
SOURCE_FILES = (
    "Cargo.toml",
    "Cargo.lock",
    "crates/litchi-core/Cargo.toml",
    "crates/litchi-ods/Cargo.toml",
    "crates/litchi-ods/src/codec/formula/evaluation.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/aggregate.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/paired.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/order.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/statistical.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/descriptive.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/descriptive/harmonic.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/descriptive/moments.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/descriptive/reciprocal_sum.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/dyadic.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/numerics.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/rounding.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/aggregate.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/paired.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/matrix.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/order.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/owned.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/references.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/conditional.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/database.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/criteria.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/statistical.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/descriptive.rs",
    "crates/litchi-ods/tests/ods_formula_conditional_evaluation.rs",
    "crates/litchi-ods/tests/ods_formula_conditional_oracle.rs",
    "crates/litchi-ods/tests/ods_formula_conditional_limits.rs",
    "crates/litchi-ods/tests/ods_formula_conditional_native.rs",
    "crates/litchi-ods/tests/ods_formula_criterion_text_profile.rs",
    "crates/litchi-ods/tests/ods_formula_statistical_evaluation.rs",
    "crates/litchi-ods/tests/ods_formula_statistical_limits.rs",
    "crates/litchi-ods/tests/ods_formula_statistical_native.rs",
    "crates/litchi-ods/tests/ods_formula_statistical_oracle.rs",
    "crates/litchi-ods/tests/ods_formula_descriptive_evaluation.rs",
    "crates/litchi-ods/tests/ods_formula_descriptive_limits.rs",
    "crates/litchi-ods/tests/ods_formula_descriptive_native.rs",
    "crates/litchi-ods/tests/ods_formula_descriptive_oracle.rs",
    "crates/litchi-ods/tests/ods_formula_evaluation.rs",
    "crates/litchi-ods/tests/ods_formula_functions.rs",
    "crates/litchi-ods/tests/ods_formula_database_numerics.rs",
    "crates/litchi-ods/tests/ods_formula_order_evaluation.rs",
    "crates/litchi-ods/tests/ods_formula_order_limits.rs",
    "crates/litchi-ods/tests/ods_formula_order_native.rs",
    "crates/litchi-ods/tests/ods_formula_order_oracle.rs",
    "crates/litchi-ods/tests/ods_formula_paired_evaluation.rs",
    "crates/litchi-ods/tests/ods_formula_paired_limits.rs",
    "crates/litchi-ods/tests/ods_formula_paired_native.rs",
    "crates/litchi-ods/tests/ods_formula_paired_oracle.rs",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/contract.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/numeric-goldens.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/numeric_oracle.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/README.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/baseline.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/cached-results.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/provenance.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/README.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/extract.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/reproduce.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-database-functions/criteria-text-profile.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-database-functions/implementation-profile.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/contract.md",
    "crates/litchi-ods/docs/FEATURE_MATRIX.md",
)

# New owner code and authored inputs are discovered by narrow globs at the
# freeze handoff. This keeps every landed paired-statistics file in the closure
# without guessing a module split before the implementation owner chooses it.
SOURCE_FILE_GLOBS = (
    "crates/litchi-ods/src/codec/formula/evaluation/paired*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/paired/**/*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/paired*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/paired/**/*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/descriptive*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/descriptive/**/*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/descriptive*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/descriptive/**/*.rs",
    "crates/litchi-ods/tests/ods_formula_paired_*.rs",
    "crates/litchi-ods/tests/ods_formula_descriptive_*.rs",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/contract.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/numeric-goldens.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/numeric_oracle.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/native/**/*",
)



def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def command_text(command: list[str]) -> str:
    return subprocess.list2cmdline(command)


def tool_text(command: list[str]) -> str:
    return subprocess.check_output(command, text=True, stderr=subprocess.STDOUT).strip()


def git_head(source_root: Path) -> str:
    return subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=source_root, text=True).strip()


def filtered_status(source_root: Path) -> bytes:
    raw = subprocess.check_output(
        ["git", "status", "--short", "--untracked-files=all"], cwd=source_root
    )
    profile_prefix = str(REL_PROFILE).encode() + b"/"
    lines = [line for line in raw.splitlines() if not line[3:].startswith(profile_prefix)]
    return b"\n".join(lines) + (b"\n" if lines else b"")


def profile_input_snapshot() -> dict[str, str]:
    excluded = {"results", "target", "__pycache__"}
    hashes: dict[str, str] = {}
    for path in sorted(HERE.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(HERE)
        if any(part in excluded for part in relative.parts):
            continue
        hashes[str(relative)] = digest(path)
    if not hashes:
        raise RuntimeError("profile input directory is empty")
    return hashes


def workspace_source_files(source_root: Path) -> list[Path]:
    paths: set[Path] = set()
    for relative in (Path("Cargo.toml"), Path("Cargo.lock")):
        path = source_root / relative
        if path.is_file():
            paths.add(path)
    crates = source_root / "crates"
    if crates.is_dir():
        for path in crates.rglob("*"):
            if not path.is_file() or "target" in path.parts:
                continue
            if path.suffix == ".rs" or path.name in {"Cargo.toml", "build.rs"}:
                paths.add(path)
    return sorted(paths)


def harness_hashes(source_root: Path) -> dict[str, str | None]:
    hashes: dict[str, str | None] = {}
    for relative in ("Cargo.toml", "Cargo.lock", "src/main.rs"):
        path = source_root / REL_HARNESS / relative
        hashes[relative] = digest(path) if path.is_file() else None
    return hashes


def source_files(source_root: Path) -> list[Path]:
    """Return the selected source closure, including recursively landed owners."""
    paths: set[Path] = {source_root / relative for relative in SOURCE_FILES}
    for pattern in SOURCE_FILE_GLOBS:
        paths.update(source_root.glob(pattern))
    return sorted(path for path in paths if path.is_file())


def source_snapshot(source_root: Path, profile_hashes: dict[str, str]) -> dict[str, Any]:
    status = filtered_status(source_root)
    selected: dict[str, str | None] = {}
    for relative in SOURCE_FILES:
        path = source_root / relative
        selected[relative] = digest(path) if path.is_file() else None
    for path in source_files(source_root):
        relative = str(path.relative_to(source_root))
        selected[relative] = digest(path) if path.is_file() else None
    closure = {
        str(path.relative_to(source_root)): digest(path)
        for path in workspace_source_files(source_root)
    }
    root_lock = source_root / "Cargo.lock"
    return {
        "git_head": git_head(source_root),
        "dirty_status_sha256_excluding_profile": hashlib.sha256(status).hexdigest(),
        "dirty_status_bytes_excluding_profile": len(status),
        "source_sha256": selected,
        "workspace_source_sha256": closure,
        "workspace_lock_sha256": digest(root_lock) if root_lock.is_file() else None,
        "harness_sha256": harness_hashes(source_root),
        "profile_input_sha256": profile_hashes,
    }


def copy_harness(source_root: Path) -> Path:
    destination = source_root / REL_HARNESS
    destination.parent.mkdir(parents=True, exist_ok=True)
    if source_root == ROOT:
        return destination
    if destination.exists():
        shutil.rmtree(destination)
    shutil.copytree(HERE / "harness", destination, ignore=shutil.ignore_patterns("target"))
    return destination


def run_checked(command: list[str], *, cwd: Path, env: dict[str, str], stdout: Path) -> None:
    with stdout.open("w", encoding="utf-8") as stream:
        completed = subprocess.run(command, cwd=cwd, env=env, stdout=stream, stderr=subprocess.STDOUT)
    if completed.returncode != 0:
        raise RuntimeError(f"command failed ({completed.returncode}): {command_text(command)}")


def run_preflight(
    *,
    binary: Path,
    source_root: Path,
    env: dict[str, str],
    output_dir: Path,
    cases: list[str],
) -> dict[str, Any]:
    """Run one untimed correctness preflight for every selected case."""
    stdout_path = output_dir / "preflight.stdout.log"
    receipt: dict[str, Any] = {
        "binary": str(binary),
        "cases": cases,
        "phase": "evaluate",
        "warmups": 0,
        "iterations": 1,
        "commands": [],
        "reference_reads": {},
    }
    with stdout_path.open("w", encoding="utf-8") as stream:
        for case in cases:
            command = [
                str(binary),
                "--case",
                case,
                "--phase",
                "evaluate",
                "--warmups",
                "0",
                "--iterations",
                "1",
                "--preflight-only",
            ]
            receipt["commands"].append(command_text(command))
            completed = subprocess.run(
                command,
                cwd=source_root,
                env=env,
                capture_output=True,
                text=True,
            )
            stream.write(completed.stdout)
            stream.write(completed.stderr)
            if completed.returncode != 0:
                raise RuntimeError(f"preflight failed ({completed.returncode}): {case}")
            lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
            measurement = next(
                (line for line in lines if line.startswith(f"preflight case={case} ")),
                None,
            )
            if measurement is None:
                raise RuntimeError(f"preflight omitted measurement for {case}")
            match = re.search(r"reference_reads=(\d+)", measurement)
            if match is None:
                raise RuntimeError(f"preflight omitted reference reads for {case}")
            reads = int(match.group(1))
            expected_reads = preflight_read_bound(case)
            if reads != expected_reads:
                raise RuntimeError(
                    f"preflight read bound failed for {case}: expected {expected_reads}, observed {reads}"
                )
            receipt["reference_reads"][case] = reads
    receipt["stdout"] = str(stdout_path.relative_to(output_dir))
    receipt["status"] = "ok"
    (output_dir / "preflight.json").write_text(
        json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    return receipt


def rss_kib(path: Path) -> int:
    text = path.read_text(encoding="utf-8")
    match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", text)
    if match is None:
        raise RuntimeError(f"missing maximum RSS in {path}")
    return int(match.group(1))


def staged_cases(binary: Path, *, source_root: Path, env: dict[str, str]) -> list[str]:
    output = subprocess.check_output([str(binary), "--list"], cwd=source_root, env=env, text=True)
    cases = [line.strip() for line in output.splitlines() if line.strip()]
    if not cases:
        raise RuntimeError("harness returned no cases")
    return cases


def preflight_read_bound(case: str) -> int:
    """Return the exact fixture read count expected before timing."""
    if case in {"database-control-dsum", "database-control-dvar", "database-control-dstdev"}:
        return 7
    if case == "reference-array-16x4-arithmetic":
        return 64
    if case == "reference-aggregate-64x4-sum" or case.startswith("reference-control-"):
        return 256
    if case == "reference-conditional-256x4-sumifs":
        return 1792
    if case.startswith("matched-descriptive-reference-"):
        return 512 if case.endswith("-avedev") else 256
    if case.startswith((
        "matched-descriptive-sensitive-inline-",
        "matched-descriptive-extreme-inline-",
        "extreme-paired-",
        "shape-reject-paired-",
        "resource-paired-",
    )):
        return 0
    if case.startswith((
        "small-paired-",
        "offset-paired-",
        "pairwise-skip-paired-",
        "forecast-query-array",
    )):
        return 128
    if case.startswith("large-paired-"):
        return 2048
    if case.startswith("cancellation-paired-"):
        return 1
    return 0


def preflight_candidate_tree(
    *, candidate_root: Path, output_dir: Path, profile_hashes: dict[str, str]
) -> dict[str, Any]:
    """Validate every candidate case before either side enters timing."""
    harness_manifest = copy_harness(candidate_root) / "Cargo.toml"
    target = Path(
        tempfile.mkdtemp(prefix="litchi-ods-paired-statistics-candidate-preflight-", dir="/var/tmp")
    )
    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    build_env = os.environ.copy()
    build_env["CARGO_TARGET_DIR"] = str(target)
    build_env["CARGO_TERM_COLOR"] = "never"
    build_command = [
        "cargo",
        "build",
        "--locked",
        "--offline",
        "--manifest-path",
        str(harness_manifest),
        "--release",
    ]
    try:
        run_checked(build_command, cwd=candidate_root, env=build_env, stdout=output_dir / "build.log")
        binary = target / "release" / BINARY_NAME
        if not binary.is_file():
            raise RuntimeError(f"missing release binary: {binary}")
        cases = staged_cases(binary, source_root=candidate_root, env=build_env)
        preflight = run_preflight(
            binary=binary,
            source_root=candidate_root,
            env=build_env,
            output_dir=output_dir,
            cases=cases,
        )
        return {
            "status": "ok",
            "cases": len(cases),
            "binary_sha256": digest(binary),
            "contract_sha256": CONTRACT_SHA256,
            "profile_input_sha256": profile_hashes,
            "output_dir": str(output_dir.relative_to(HERE)),
            "preflight": preflight,
        }
    finally:
        shutil.rmtree(target, ignore_errors=True)
        (output_dir / "target-cleanup.json").write_text(
            json.dumps({"target_dir": str(target), "removed": not target.exists()}, indent=2) + "\n",
            encoding="utf-8",
        )


def verify_candidate_freeze(candidate_root: Path, freeze_path: Path) -> dict[str, Any]:
    freeze = json.loads(freeze_path.read_text(encoding="utf-8"))
    frozen_source_root = freeze.get("source_root")
    if frozen_source_root is not None:
        frozen_root = Path(frozen_source_root).resolve()
        if frozen_root != candidate_root:
            raise RuntimeError(f"candidate root differs from freeze: {candidate_root} != {frozen_root}")
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict) or not selected:
        raise RuntimeError("candidate freeze has no selected file hashes")
    mismatches: list[str] = []
    for relative, expected in selected.items():
        path = candidate_root / relative
        actual = digest(path) if path.is_file() else None
        if actual != expected:
            mismatches.append(f"{relative}: expected {expected}, observed {actual}")
    if mismatches:
        raise RuntimeError("candidate freeze hash mismatch:\n" + "\n".join(mismatches))
    base_commit = str(freeze.get("base_commit", ""))
    baseline = subprocess.check_output(["git", "rev-parse", BASELINE_COMMIT], cwd=ROOT, text=True).strip()
    if not (base_commit == baseline or base_commit.startswith(baseline)):
        raise RuntimeError(f"candidate freeze base differs from {baseline}")
    return freeze


def capture_tree(
    *,
    label: str,
    source_root: Path,
    output_dir: Path,
    warmups: int,
    samples: int,
    requested_cases: tuple[str, ...] | None,
    profile_hashes: dict[str, str],
) -> dict[str, Any]:
    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    raw_dir = output_dir / "raw"
    raw_dir.mkdir()
    harness_manifest = copy_harness(source_root) / "Cargo.toml"
    target = Path(tempfile.mkdtemp(prefix=f"litchi-ods-paired-statistics-{label}-", dir="/var/tmp"))
    before = source_snapshot(source_root, profile_hashes)
    build_env = os.environ.copy()
    build_env["CARGO_TARGET_DIR"] = str(target)
    build_env["CARGO_TERM_COLOR"] = "never"
    build_command = ["cargo", "build", "--locked", "--offline", "--manifest-path", str(harness_manifest), "--release"]
    (output_dir / "build-command.json").write_text(
        json.dumps({"command": build_command, "cwd": str(source_root), "target": str(target)}, indent=2) + "\n",
        encoding="utf-8",
    )
    try:
        run_checked(build_command, cwd=source_root, env=build_env, stdout=output_dir / "build.log")
        binary = target / "release" / BINARY_NAME
        if not binary.is_file():
            raise RuntimeError(f"missing release binary: {binary}")
        binary_hash = digest(binary)
        available = staged_cases(binary, source_root=source_root, env=build_env)
        cases = list(available if requested_cases is None else requested_cases)
        missing = [case for case in cases if case not in available]
        if missing:
            raise RuntimeError(f"requested cases missing from harness: {missing}")
        preflight = run_preflight(
            binary=binary,
            source_root=source_root,
            env=build_env,
            output_dir=output_dir,
            cases=cases,
        )
        environment = {
            "label": label,
            "source_root": str(source_root),
            "source_git_head": before["git_head"],
            "rustc": subprocess.check_output(["rustc", "--version"], text=True).strip(),
            "rustc_verbose": tool_text(["rustc", "-Vv"]),
            "cargo": subprocess.check_output(["cargo", "--version"], text=True).strip(),
            "libc": tool_text(["getconf", "GNU_LIBC_VERSION"]),
            "rustflags": {
                "RUSTFLAGS": build_env.get("RUSTFLAGS", ""),
                "CARGO_ENCODED_RUSTFLAGS": build_env.get("CARGO_ENCODED_RUSTFLAGS", ""),
                "CARGO_PROFILE_RELEASE_LTO": build_env.get("CARGO_PROFILE_RELEASE_LTO", ""),
                "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": build_env.get("CARGO_PROFILE_RELEASE_CODEGEN_UNITS", ""),
            },
            "kernel": os.uname().release,
            "cpu": subprocess.check_output(["lscpu"], text=True),
            "warmups_per_child": warmups,
            "samples_per_group": samples,
            "cases": cases,
            "available_cases": available,
            "case_scope": "matched-controls" if requested_cases is not None else "all-named-cases",
            "preflight": preflight,
            "phases": list(PHASES),
            "binary_sha256": binary_hash,
            "contract_sha256": CONTRACT_SHA256,
            "profile_input_sha256": profile_hashes,
        }
        (output_dir / "environment.json").write_text(json.dumps(environment, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        records: list[dict[str, Any]] = []
        commands_path = output_dir / "commands.txt"
        commands_path.write_text("", encoding="utf-8")
        for case in cases:
            for phase in PHASES:
                for sample_index in range(samples):
                    stem = f"{case}.{phase}.sample-{sample_index + 1:02d}"
                    stdout_path = raw_dir / f"{stem}.stdout.json"
                    stderr_path = raw_dir / f"{stem}.stderr.log"
                    time_path = raw_dir / f"{stem}.time.txt"
                    command = [
                        "/usr/bin/time", "-v", "-o", str(time_path), str(binary),
                        "--case", case, "--phase", phase, "--warmups", str(warmups),
                        "--iterations", "1",
                    ]
                    with commands_path.open("a", encoding="utf-8") as commands:
                        commands.write(command_text(command) + "\n")
                    with stdout_path.open("w", encoding="utf-8") as stdout, stderr_path.open("w", encoding="utf-8") as stderr:
                        completed = subprocess.run(command, cwd=source_root, env=build_env, stdout=stdout, stderr=stderr, text=True)
                    if completed.returncode != 0:
                        raise RuntimeError(f"lane failed ({completed.returncode}): {stem}")
                    lines = [line for line in stdout_path.read_text(encoding="utf-8").splitlines() if line.strip()]
                    if len(lines) != 1:
                        raise RuntimeError(f"lane {stem} emitted {len(lines)} JSON lines")
                    row = json.loads(lines[0])
                    stderr_text = stderr_path.read_text(encoding="utf-8")
                    raw_stderr: str | None = str(stderr_path.relative_to(output_dir))
                    if not stderr_text:
                        stderr_path.unlink()
                        raw_stderr = None
                    row.update({
                        "capture": label,
                        "sample_index": sample_index + 1,
                        "rss_kib": rss_kib(time_path),
                        "binary_sha256": binary_hash,
                        "source_git_head": before["git_head"],
                        "raw_stdout": str(stdout_path.relative_to(output_dir)),
                        "raw_stderr": raw_stderr,
                        "raw_time": str(time_path.relative_to(output_dir)),
                    })
                    records.append(row)
        after = source_snapshot(source_root, profile_hashes)
        current_profile = profile_input_snapshot()
        if before["source_sha256"] != after["source_sha256"]:
            raise RuntimeError(f"selected source changed during {label} capture")
        if before["workspace_source_sha256"] != after["workspace_source_sha256"]:
            raise RuntimeError(f"compiled workspace source changed during {label} capture")
        if profile_hashes != current_profile:
            raise RuntimeError(f"profile inputs changed during {label} capture")
        (output_dir / "source-manifest.json").write_text(json.dumps({
            "before": before,
            "after": after,
            "source_sha256_unchanged": before["source_sha256"] == after["source_sha256"],
            "workspace_source_sha256_unchanged": before["workspace_source_sha256"] == after["workspace_source_sha256"],
            "git_head_unchanged": before["git_head"] == after["git_head"],
            "profile_input_sha256_unchanged": before["profile_input_sha256"] == after["profile_input_sha256"],
            "binary_sha256": binary_hash,
        }, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        (output_dir / "measurements.jsonl").write_text("".join(json.dumps(row, sort_keys=True) + "\n" for row in records), encoding="utf-8")
        return {"label": label, "cases": cases, "phases": list(PHASES), "records": len(records), "source_git_head": before["git_head"], "binary_sha256": binary_hash, "case_scope": environment["case_scope"], "preflight": preflight}
    finally:
        shutil.rmtree(target, ignore_errors=True)
        (output_dir / "target-cleanup.json").write_text(json.dumps({"target_dir": str(target), "removed": not target.exists()}, indent=2) + "\n", encoding="utf-8")


def add_baseline_worktree(path: Path, lock_source: Path) -> None:
    subprocess.run(["git", "worktree", "add", "--detach", str(path), BASELINE_COMMIT], check=True, cwd=ROOT)
    shutil.copy2(lock_source, path / "Cargo.lock")


def remove_worktree(path: Path) -> bool:
    if not path.exists():
        return True
    completed = subprocess.run(["git", "worktree", "remove", "--force", str(path)], cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    if completed.returncode == 0:
        return not path.exists()
    shutil.rmtree(path, ignore_errors=True)
    return not path.exists()


def require_capture_inputs() -> None:
    """Require the reviewed contract and retained oracle inputs exactly."""
    if CONTRACT_SHA256 is None:
        raise RuntimeError(
            "paired-statistics capture is contract-gated; the reviewed "
            "contract hash is not configured"
        )
    contract = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/contract.md"
    oracle = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/numeric_oracle.py"
    goldens = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-paired-statistics/numeric-goldens.json"
    missing = [str(path.relative_to(ROOT)) for path in (contract, oracle, goldens) if not path.is_file()]
    if missing:
        raise RuntimeError(
            "paired-statistics profile is contract-gated; missing reviewed inputs: "
            + ", ".join(missing)
        )
    observed_contract = digest(contract)
    if observed_contract != CONTRACT_SHA256:
        raise RuntimeError(
            "paired-statistics contract hash differs from the reviewed input: "
            f"expected {CONTRACT_SHA256}, observed {observed_contract}"
        )
    contract_text = contract.read_text(encoding="utf-8")
    missing_functions = [
        function
        for function in ("CORREL", "COVAR", "PEARSON", "RSQ", "SLOPE", "INTERCEPT", "STEYX", "FORECAST")
        if function not in contract_text
    ]
    if missing_functions:
        raise RuntimeError(f"contract omits paired functions: {missing_functions}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline-only", action="store_true")
    parser.add_argument("--candidate-root", type=Path)
    parser.add_argument("--candidate-freeze", type=Path)
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--samples", type=int, default=15)
    args = parser.parse_args()
    if args.baseline_only and (args.candidate_root is not None or args.candidate_freeze is not None):
        parser.error("--baseline-only cannot be combined with candidate options")
    if not args.baseline_only and (args.candidate_root is None or args.candidate_freeze is None):
        parser.error("candidate capture requires --candidate-root and --candidate-freeze")
    if args.warmups < 0 or args.samples <= 0:
        parser.error("warmups must be nonnegative and samples must be positive")
    require_capture_inputs()
    if not GATE_LOCK.is_file():
        raise RuntimeError(f"missing retained gate Cargo.lock: {GATE_LOCK}")
    observed_gate_lock = digest(GATE_LOCK)
    if observed_gate_lock != GATE_LOCK_SHA256:
        raise RuntimeError(
            "retained paired-statistics gate Cargo.lock hash differs: "
            f"expected {GATE_LOCK_SHA256}, observed {observed_gate_lock}"
        )
    profile_before = profile_input_snapshot()
    results_root = HERE / "results"
    results_root.mkdir(parents=True, exist_ok=True)
    (results_root / "profile-inputs-before.json").write_text(json.dumps(profile_before, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    summaries: list[dict[str, Any]] = []
    candidate_preflight_gate: dict[str, Any] | None = None
    candidate_root: Path | None = None
    freeze: dict[str, Any] | None = None
    if not args.baseline_only:
        candidate_root = args.candidate_root.resolve()
        freeze = verify_candidate_freeze(candidate_root, args.candidate_freeze.resolve())
        candidate_preflight_gate = preflight_candidate_tree(
            candidate_root=candidate_root,
            output_dir=results_root / "preflight-before-timing",
            profile_hashes=profile_before,
        )
    baseline_worktree = Path(tempfile.mkdtemp(prefix="litchi-ods-paired-statistics-baseline-", dir="/tmp"))
    baseline_worktree.rmdir()
    cleanup: dict[str, Any] = {"baseline_worktree": str(baseline_worktree)}
    try:
        add_baseline_worktree(baseline_worktree, GATE_LOCK)
        summaries.append(capture_tree(label=f"baseline-{BASELINE_COMMIT}", source_root=baseline_worktree, output_dir=results_root / f"baseline-{BASELINE_COMMIT}", warmups=args.warmups, samples=args.samples, requested_cases=MATCHED_CONTROL_CASES, profile_hashes=profile_before))
    finally:
        cleanup["baseline_removed"] = remove_worktree(baseline_worktree)
    if not args.baseline_only:
        assert candidate_root is not None and freeze is not None
        summaries.append(capture_tree(label="candidate-final", source_root=candidate_root, output_dir=results_root / "candidate-final", warmups=args.warmups, samples=args.samples, requested_cases=None, profile_hashes=profile_before))
        summaries[-1]["freeze_path"] = str(args.candidate_freeze.resolve())
        summaries[-1]["freeze_base_commit"] = freeze["base_commit"]
    profile_after = profile_input_snapshot()
    if profile_before != profile_after:
        raise RuntimeError("profile inputs changed between captures")
    (results_root / "profile-inputs-after.json").write_text(json.dumps(profile_after, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    (results_root / "capture-summary.json").write_text(json.dumps({
        "baseline_commit": BASELINE_COMMIT,
        "contract_sha256": CONTRACT_SHA256,
        "controls": list(MATCHED_CONTROL_CASES),
        "phases": list(PHASES),
        "warmups_per_child": args.warmups,
        "samples_per_group": args.samples,
        "profile_input_sha256_before": profile_before,
        "profile_input_sha256_after": profile_after,
        "candidate_preflight_gate": candidate_preflight_gate,
        "captures": summaries,
        "cleanup": cleanup,
    }, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
