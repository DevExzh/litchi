#!/usr/bin/env python3
"""Capture the bounded ODS reference-metadata profile in fresh child processes.

The committed baseline is used for all matched controls; reference-metadata rows
are candidate-only additions. Every lane hashes the
compiled workspace closure and authored profile inputs before and after capture.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
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
GATE_LOCK = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/gates/Cargo.lock"
GATE_LOCK_SHA256 = "58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3"
BASELINE_COMMIT = "049c09cdde3978593149079c4257df047a3fa419"
BINARY_NAME = "ods-formula-reference-metadata-performance"
# The source freeze records the reviewed contract digest. Keeping it in the
# runner makes a capture fail closed if the contract is changed afterward.
CONTRACT_SHA256 = "87f20c419139bf721221d06c8714edd900ae7a57cbedecf070878e4f6cf64d8a"
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
MATCHED_CONTROL_CASES = CONTROL_CASES + REPRESENTATIVE_CONTROLS
MATCHED_CONTROL_CASES += (
    "concat-borrowed-literals",
    "concat-owned-left",
    "concat-owned-right",
    "concat-growth-chain",
)

REFERENCE_METADATA_FUNCTIONS = (
    "AREAS", "COLUMN", "COLUMNS", "ISREF", "ROW", "ROWS", "SHEET", "SHEETS",
)

REFERENCE_METADATA_CASES = (
    "reference-metadata-areas-range",
    "reference-metadata-areas-3d",
    "reference-metadata-areas-list",
    "reference-metadata-areas-duplicate-list",
    "reference-metadata-column-current",
    "reference-metadata-column-cell",
    "reference-metadata-column-3d-cell",
    "reference-metadata-column-matrix",
    "reference-metadata-column-scalar-vector",
    "reference-metadata-columns-range",
    "reference-metadata-columns-3d",
    "reference-metadata-columns-array",
    "reference-metadata-columns-one-entry-list",
    "reference-metadata-isref-reference",
    "reference-metadata-isref-list",
    "reference-metadata-isref-array",
    "reference-metadata-isref-error",
    "reference-metadata-isref-number",
    "reference-metadata-isref-source",
    "reference-metadata-isref-selected-source",
    "reference-metadata-isref-selected-local",
    "reference-metadata-isref-iferror-source",
    "reference-metadata-isref-ifna-source",
    "reference-metadata-isref-iferror-source-arithmetic",
    "reference-metadata-isref-ifna-source-arithmetic",
    "reference-metadata-isref-iferror-source-fallback",
    "reference-metadata-isref-ifna-source-fallback",
    "reference-metadata-row-current",
    "reference-metadata-row-cell",
    "reference-metadata-row-3d-cell",
    "reference-metadata-row-matrix",
    "reference-metadata-row-scalar-vector",
    "reference-metadata-rows-range",
    "reference-metadata-rows-3d",
    "reference-metadata-rows-array",
    "reference-metadata-rows-one-entry-list",
    "reference-metadata-sheet-current",
    "reference-metadata-sheet-text",
    "reference-metadata-sheet-hidden-text",
    "reference-metadata-sheet-number-coercion",
    "reference-metadata-sheet-logical-coercion",
    "reference-metadata-sheet-unknown-text",
    "reference-metadata-sheet-array-matrix",
    "reference-metadata-sheet-reference",
    "reference-metadata-sheet-source",
    "reference-metadata-sheet-selected-source",
    "reference-metadata-sheet-selected-local",
    "reference-metadata-sheet-iferror-source",
    "reference-metadata-sheet-ifna-source",
    "reference-metadata-sheet-iferror-source-fallback",
    "reference-metadata-sheet-ifna-source-fallback",
    "reference-metadata-sheet-computed-scalar",
    "reference-metadata-sheet-3d",
    "reference-metadata-sheets-current",
    "reference-metadata-sheets-reference",
    "reference-metadata-sheets-source",
    "reference-metadata-sheets-selected-source",
    "reference-metadata-sheets-selected-local",
    "reference-metadata-sheets-iferror-source",
    "reference-metadata-sheets-ifna-source",
    "reference-metadata-sheets-iferror-source-fallback",
    "reference-metadata-sheets-ifna-source-fallback",
    "reference-metadata-sheets-one-entry-list",
    "reference-metadata-sheets-3d",
    "reference-metadata-large-range",
    "reference-metadata-row-projected",
    "reference-metadata-column-projected",
    "reference-metadata-rows-projected",
    "reference-metadata-columns-projected",
    "reference-metadata-isref-source-array-true-matrix",
    "reference-metadata-isref-source-array-false-matrix",
    "reference-metadata-sheet-source-array-true-matrix",
    "reference-metadata-sheet-source-array-false-matrix",
    "reference-metadata-sheets-source-array-true-matrix",
    "reference-metadata-sheets-source-array-false-matrix",
    "reference-metadata-columns-formula-error",
    "reference-metadata-sheets-formula-error",
    "reference-metadata-areas-refusal",
    "reference-metadata-areas-array-refusal",
    "reference-metadata-column-list-refusal",
    "reference-metadata-column-one-entry-list",
    "reference-metadata-column-array-refusal",
    "reference-metadata-row-list-refusal",
    "reference-metadata-row-one-entry-list",
    "reference-metadata-row-array-refusal",
    "reference-metadata-sheet-list-refusal",
    "reference-metadata-sheet-array-refusal",
    "reference-metadata-rows-list-refusal",
    "reference-metadata-rows-scalar-refusal",
    "reference-metadata-columns-list-refusal",
    "reference-metadata-columns-scalar-refusal",
    "reference-metadata-sheets-list-refusal",
    "reference-metadata-sheets-array-refusal",
    "reference-metadata-areas-arity",
    "reference-metadata-columns-arity",
    "reference-metadata-isref-arity",
    "reference-metadata-areas-resource",
    "reference-metadata-rows-resource",
    "reference-metadata-sheet-cancel",
)

MATCHED_CONTROL_CASES = CONTROL_CASES + REPRESENTATIVE_CONTROLS + (
    "concat-borrowed-literals",
    "concat-owned-left",
    "concat-owned-right",
    "concat-growth-chain",
)

FEATURE_MATRIX = (
    ("descriptor-shape", "single-area geometry without cell reads", 256, "metadata"),
    ("three-dimensional", "ordered three-sheet geometry and sheet counts", 768, "metadata"),
    ("reference-list", "ordered union area records", 2, "metadata"),
    ("current-coordinate", "ROW/COLUMN current position and first-cell coordinates", 0, "metadata"),
    ("array-descriptor", "ROWS/COLUMNS over inline arrays", 0, "metadata"),
    ("projected-matrix", "projected metadata outputs preserve coordinates", 0, "matrix"),
    ("shape-refusal", "scalar arguments return formula VALUE before reads", 0, "formula-error"),
    ("large-geometry", "bounded large descriptor metadata", 32768, "metadata"),
    ("resource", "reference geometry cap returns typed resource failure", 0, "typed-failure"),
    ("cancellation", "sticky cancellation remains typed", 0, "typed-failure"),
)

# The selected manifest includes the value dispatch, scalar/text bridge, the
# existing controls used for matching, metadata owners and tests, and every
# authored contract/oracle/native input that can affect the measured evaluator.
SOURCE_FILES = (
    "Cargo.toml",
    "Cargo.lock",
    "crates/litchi-core/Cargo.toml",
    "crates/litchi-ods/Cargo.toml",
    "crates/litchi-ods/src/codec/formula/evaluation.rs",
    "crates/litchi-ods/src/codec/formula/functions.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/aggregate.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/paired.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/order.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/statistical.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/descriptive.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/dyadic.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/numerics.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/rounding.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/calendar.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/inspection.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/references.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/inspection.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/reference_metadata.rs",
    "crates/litchi-ods/tests/ods_formula_evaluation.rs",
    "crates/litchi-ods/tests/ods_formula_functions.rs",
    "crates/litchi-ods/docs/FEATURE_MATRIX.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/baseline.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/integration-plan.md",
)

SOURCE_FILE_GLOBS = (
    "crates/litchi-ods/src/codec/formula/evaluation/reference_metadata*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/reference_metadata*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/matrix*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/owned*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/*.rs",
    "crates/litchi-ods/tests/ods_formula_reference_metadata_*.rs",
    "docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/contract.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/*_oracle.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/*-goldens.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata/native/**/*",
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


def preserve_existing_output(path: Path) -> Path | None:
    """Move a prior capture aside so a retry cannot erase its evidence."""
    if not path.exists():
        return None
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    destination = path.with_name(f"diagnostic-{path.name}-{stamp}")
    suffix = 1
    while destination.exists():
        destination = path.with_name(f"diagnostic-{path.name}-{stamp}-{suffix}")
        suffix += 1
    path.rename(destination)
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
    observed = [line.strip() for line in output.splitlines() if line.strip()]
    if not observed:
        raise RuntimeError("harness returned no cases")
    expected = list(MATCHED_CONTROL_CASES) + list(REFERENCE_METADATA_CASES)
    if set(observed) != set(expected) or len(observed) != len(expected):
        missing = sorted(set(expected) - set(observed))
        unexpected = sorted(set(observed) - set(expected))
        raise RuntimeError(
            "harness case set differs from the frozen matrix: "
            f"missing={missing}, unexpected={unexpected}, "
            f"observed_count={len(observed)}, expected_count={len(expected)}"
        )
    # The harness may enumerate enum variants in implementation order. Keep
    # capture, preflight, environment, and verifier rows in matrix order.
    return expected


def preflight_read_bound(case: str) -> int:
    """Return the exact fixture read count expected before timing."""
    if case == "reference-metadata-sheet-computed-scalar":
        # SHEET itself is metadata-only, but ABS must scalar-intersect its
        # computed range argument and read the selected child cell first.
        return 1
    if case.startswith("reference-metadata-"):
        # Metadata functions consume descriptors/provider metadata only. They
        # never read cells; even the cap and cancellation lanes must report 0.
        return 0
    if case in {"database-control-dsum", "database-control-dvar", "database-control-dstdev"}:
        return 7
    if case == "reference-array-16x4-arithmetic":
        return 64
    if case == "reference-aggregate-64x4-sum" or case.startswith("reference-control-"):
        return 256
    if case == "reference-conditional-256x4-sumifs":
        return 1792
    return 0


def preflight_candidate_tree(
    *, candidate_root: Path, output_dir: Path, profile_hashes: dict[str, str]
) -> dict[str, Any]:
    """Validate every candidate case before either side enters timing."""
    harness_manifest = copy_harness(candidate_root) / "Cargo.toml"
    target = Path(
        tempfile.mkdtemp(prefix="litchi-ods-reference-metadata-candidate-preflight-", dir="/var/tmp")
    )
    preserve_existing_output(output_dir)
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
    preserve_existing_output(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    raw_dir = output_dir / "raw"
    raw_dir.mkdir()
    harness_manifest = copy_harness(source_root) / "Cargo.toml"
    target = Path(tempfile.mkdtemp(prefix=f"litchi-ods-reference-metadata-{label}-", dir="/var/tmp"))
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
    """Require the frozen reference-metadata contract and oracle inputs."""
    if CONTRACT_SHA256 is None:
        raise RuntimeError(
            "reference-metadata profile is contract-gated; configure the frozen "
            "contract digest before capture"
        )
    evidence = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-reference-metadata"
    contract = evidence / "contract.md"
    oracle_candidates = sorted(evidence.glob("*_oracle.py"))
    golden_candidates = sorted(evidence.glob("*-goldens.json"))
    missing = [str(path.relative_to(ROOT)) for path in [contract] if not path.is_file()]
    if not oracle_candidates:
        missing.append(str((evidence / "*_oracle.py").relative_to(ROOT)))
    if not golden_candidates:
        missing.append(str((evidence / "*-goldens.json").relative_to(ROOT)))
    if missing:
        raise RuntimeError("reference-metadata profile missing reviewed inputs: " + ", ".join(missing))
    observed_contract = digest(contract)
    if observed_contract != CONTRACT_SHA256:
        raise RuntimeError(
            "reference-metadata contract hash differs from reviewed input: "
            f"expected {CONTRACT_SHA256}, observed {observed_contract}"
        )
    contract_text = contract.read_text(encoding="utf-8")
    missing_functions = [function for function in REFERENCE_METADATA_FUNCTIONS if function not in contract_text]
    if missing_functions:
        raise RuntimeError(f"contract omits reference metadata functions: {missing_functions}")
    oracle = oracle_candidates[0]
    oracle_check = subprocess.run(["python3", str(oracle), "--check"], cwd=ROOT, capture_output=True, text=True)
    if oracle_check.returncode != 0:
        raise RuntimeError("reference-metadata oracle/goldens check failed: " + (oracle_check.stdout + oracle_check.stderr).strip())


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
            "retained reference-metadata gate Cargo.lock hash differs: "
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
    # A full detached baseline checkout is larger than the host's shared /tmp
    # tmpfs in this workspace. Keep it beside the build targets on /var/tmp so
    # a low tmpfs watermark cannot invalidate an otherwise frozen capture.
    baseline_worktree = Path(
        tempfile.mkdtemp(prefix="litchi-ods-reference-metadata-baseline-", dir="/var/tmp")
    )
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
