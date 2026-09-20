#!/usr/bin/env python3
"""Bounded matched baseline/candidate profile for the ODF date/time family.

The runner is capture-only.  It builds an isolated harness in a fresh target,
preflights every named row against the independent case matrix, then records
raw JSON and /usr/bin/time receipts.  It never changes production or the
frozen candidate checkout.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import math
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
GATE_LOCK = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-date-time/gates/Cargo.lock"
GATE_LOCK_SHA256 = "58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3"
BASELINE_COMMIT = "6fa3b8af6aa6b5cdce348609a40a34a30cf94fd6"
PREPARATION_COMMIT = "45e56928edfc334b45132330d405bc4c7b6f7b4c"
BINARY_NAME = "ods-formula-date-time-performance"
CONTRACT_SHA256 = "cc77d41f487993b3438f817dd62a359ba4ec4b3ca2359a893b31aecc79bc2c7f"
ORACLE_SHA256 = "b33011089974b18b1fcc0b984adab839600acc98f09dd850e02522348316259f"
ORACLE_VERIFY_SHA256 = "4884b8a0ff7fc7dc2a0259b4e7b4e24c4be169a166cd92982b728b2ed77eb64b"
MATRIX_PATH = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-date-time/performance/case-matrix.json"
BASELINE_PATH = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-date-time/baseline.json"
ORACLE_PATH = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-date-time/oracle-vectors.json"
ORACLE_VERIFY_PATH = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-date-time/oracle_verify.py"
CONTRACT_PATH = ROOT / "docs/report/spec-gap-validation-evidence/ods-formula-date-time/contract.md"
PHASES = ("evaluate", "parse-evaluate")

# The matched controls are unchanged from the predecessor VALUE/statistical
# process profile, with its VALUE/date-fraction and projected lazy controls
# retained so candidate date kernels cannot hide a regression in the bridge.
MATCHED_CONTROL_CASES = (
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
    "representative-median",
    "representative-rank",
    "representative-percentrank",
    "concat-borrowed-literals",
    "concat-owned-left",
    "concat-owned-right",
    "concat-growth-chain",
    "value-date-fraction-value-date",
    "value-date-fraction-value-time",
    "value-date-fraction-value-mixed-fraction",
    "lazy-if-cache-isnumber",
    "lazy-if-cache-isblank",
    "lazy-if-cache-istext",
)

# This declaration is intentionally literal so staging and review tools can
# inspect the complete named date workload without importing the harness.
DATE_CASES = (
    "date-core-date", "date-core-datedif", "date-core-datevalue", "date-core-day",
    "date-core-days", "date-core-days360", "date-core-eastersunday", "date-core-edate",
    "date-core-eomonth", "date-core-hour", "date-core-isoweeknum", "date-core-minute",
    "date-core-month", "date-core-networkdays", "date-core-now", "date-core-second",
    "date-core-time", "date-core-timevalue", "date-core-today", "date-core-weekday",
    "date-core-weeknum", "date-core-workday", "date-core-year", "date-core-yearfrac",
    "date-parser-datevalue-iso", "date-parser-datevalue-datetime", "date-parser-datevalue-enus",
    "date-parser-datevalue-month", "date-parser-timevalue-clock", "date-parser-timevalue-datetime",
    "date-parser-timevalue-fraction",
    "date-parser-datevalue-number", "date-parser-timevalue-number",
    "date-boundary-date-rollover", "date-boundary-date-1900", "date-boundary-days360-eu",
    "date-boundary-edate-clamp", "date-boundary-eomonth-clamp", "date-boundary-yearfrac-basis1",
    "date-boundary-weeknum-noninteger", "date-boundary-weekday-17",
    "date-networkdays-holidays-0", "date-networkdays-holidays-1", "date-networkdays-holidays-16",
    "date-networkdays-holidays-64", "date-networkdays-holidays-256",
    "date-workday-holidays-0", "date-workday-holidays-1", "date-workday-holidays-16",
    "date-workday-holidays-64", "date-workday-holidays-256",
    "date-networkdays-workweek-default", "date-networkdays-workweek-custom",
    "date-networkdays-workweek-alloff", "date-networkdays-workweek-malformed",
    "date-workday-workweek-default", "date-workday-workweek-custom",
    "date-workday-workweek-alloff", "date-workday-workweek-malformed",
    "date-networkdays-interval-7", "date-networkdays-interval-31",
    "date-networkdays-interval-365", "date-networkdays-interval-4096",
    "date-workday-offset-0", "date-workday-offset-1", "date-workday-offset-64",
    "date-workday-offset-256", "date-workday-offset-1024",
    "date-sequence-networkdays-projection", "date-sequence-workday-projection",
    "date-projected-day", "date-lazy-unselected",
    "date-timestamp-now", "date-timestamp-today", "date-timestamp-eastersunday",
    "date-missing-now", "date-missing-today", "date-missing-eastersunday",
    "date-list-refusal", "date-direct-sequence-refusal", "date-cancel-networkdays",
    "date-cancel-workday", "date-resource-networkdays", "date-work-limit-networkdays",
    "date-formula-error-sequence", "date-dateparam-error",
)

# Source closure is deliberately broad around the date kernel and the value
# bridge.  The recursive globs cover module splits landed after this plan.
SOURCE_FILES = (
    "Cargo.toml",
    "Cargo.lock",
    "crates/litchi-core/Cargo.toml",
    "crates/litchi-ods/Cargo.toml",
    "crates/litchi-ods/src/codec/formula/evaluation.rs",
    "crates/litchi-ods/src/codec/formula/functions.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/scalar.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/calendar.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/date_time.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/timestamp.rs",
    "crates/litchi-ods/tests/ods_formula_evaluation.rs",
    "crates/litchi-ods/tests/ods_formula_functions.rs",
    "crates/litchi-ods/tests/ods_formula_value_evaluation.rs",
    "crates/litchi-ods/tests/ods_formula_date_time_evaluation.rs",
    "crates/litchi-ods/tests/ods_formula_date_time_limits.rs",
    "crates/litchi-ods/tests/ods_formula_date_time_oracle.rs",
    "crates/litchi-ods/docs/FEATURE_MATRIX.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-date-time/contract.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-date-time/baseline.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-date-time/oracle-vectors.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-date-time/performance/case-matrix.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-date-time/oracle_verify.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-date-time/native/native-results.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-date-time/native/provenance.json",
)
SOURCE_FILE_GLOBS = (
    "crates/litchi-ods/src/codec/formula/evaluation/calendar*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/calendar/**/*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/date_time*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/date_time/**/*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/timestamp*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/timestamp/**/*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/text*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/text/**/*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/**/*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/aggregate*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/statistical*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/order*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/conditional*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/lookup*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/lookup/**/*.rs",
    "crates/litchi-ods/tests/ods_formula_*date*time*.rs",
    "crates/litchi-ods/tests/ods_formula_*value*.rs",
    "docs/report/spec-gap-validation-evidence/ods-formula-date-time/native/**/*",
)
FEATURE_MATRIX = (
    ("date-core", "24 scalar date/time functions with fixed serial profile", 0, "scalar"),
    ("parser", "DATEVALUE/TIMEVALUE ISO, en_US, clock and numeric fallback", 0, "scalar"),
    ("boundary", "rollover, clamps, week modes, basis and 1900 epoch", 0, "scalar"),
    ("holiday-scaling", "NETWORKDAYS/WORKDAY streamed holiday references", "N", "reference"),
    ("workweek", "default/custom/all-off/malformed seven-cell workweek", 7, "reference"),
    ("interval-scaling", "date-day spans and WORKDAY offsets", 0, "scalar"),
    ("projection", "sequence consumers and projected DAY values", "N*(H+W)", "matrix"),
    ("lazy", "unselected IF branch has zero reference reads", 0, "matrix"),
    ("timestamp", "injected NOW/TODAY/EASTERSUNDAY and typed absence", 0, "capability"),
    ("refusal", "direct list/sequence pseudotype refusal before reads", 0, "typed-error"),
    ("resource", "reference-cell budget refusal and sticky cancellation", "case", "typed-failure"),
    ("formula-error", "retained formula errors after complete sequence scan", "case", "formula-error"),
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


def measurement_context(env: dict[str, str]) -> dict[str, Any]:
    """Record only the approved compiler settings and host load indicators."""
    allowed_keys = {"RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS"}
    allowed_keys.update(key for key in env if key.startswith("CARGO_PROFILE_RELEASE_"))
    return {
        "allowlisted_environment": {
            key: env[key] for key in sorted(allowed_keys) if key in env
        },
        "load_average": list(os.getloadavg()),
        "cpu_count": os.cpu_count(),
    }


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
        raise RuntimeError("date/time profile input directory is empty")
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
        selected[str(path.relative_to(source_root))] = digest(path)
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


def matrix() -> dict[str, Any]:
    value = json.loads(MATRIX_PATH.read_text(encoding="utf-8"))
    if value.get("contract_sha256") != CONTRACT_SHA256:
        raise RuntimeError("case matrix contract hash is stale")
    if value.get("oracle_vectors_sha256") != ORACLE_SHA256:
        raise RuntimeError("case matrix oracle hash is stale")
    return value


def expected_cases() -> list[str]:
    value = matrix()
    observed = [row["name"] for row in value["matched_controls"] + value["candidate_cases"]]
    expected = list(MATCHED_CONTROL_CASES) + list(DATE_CASES)
    if observed != expected:
        raise RuntimeError("case matrix order/set differs from runner declarations")
    return expected


def preflight_read_bound(case: str) -> int:
    return int(preflight_matrix_row(case)["reference_reads"])


def preflight_matrix_row(case: str) -> dict[str, Any]:
    value = matrix()
    for row in value["matched_controls"] + value["candidate_cases"]:
        if row["name"] == case:
            return row
    raise RuntimeError(f"case has no matrix row: {case}")


def same_expected_outcome(observed: Any, expected: Any) -> bool:
    """Compare typed oracle outcomes with the harness numeric tolerance."""
    if isinstance(expected, bool):
        return isinstance(observed, bool) and observed == expected
    if isinstance(expected, (int, float)):
        return (
            isinstance(observed, (int, float))
            and not isinstance(observed, bool)
            and math.isfinite(observed)
            and math.isfinite(expected)
            and abs(observed - expected) <= max(abs(expected), 1.0) * 1.0e-10
        )
    if isinstance(expected, list):
        return isinstance(observed, list) and len(observed) == len(expected) and all(
            same_expected_outcome(actual, wanted) for actual, wanted in zip(observed, expected)
        )
    if isinstance(expected, dict):
        return isinstance(observed, dict) and observed.keys() == expected.keys() and all(
            same_expected_outcome(observed[key], value) for key, value in expected.items()
        )
    return type(observed) is type(expected) and observed == expected


def run_preflight(*, binary: Path, source_root: Path, env: dict[str, str], output_dir: Path, cases: list[str]) -> dict[str, Any]:
    stdout_path = output_dir / "preflight.stdout.log"
    receipt: dict[str, Any] = {
        "binary": str(binary), "cases": cases, "phase": "evaluate",
        "warmups": 0, "iterations": 1, "commands": [], "reference_reads": {}, "descriptions": {},
    }
    with stdout_path.open("w", encoding="utf-8") as stream:
        for case in cases:
            command = [str(binary), "--case", case, "--phase", "evaluate", "--warmups", "0", "--iterations", "1", "--preflight-only"]
            receipt["commands"].append(command_text(command))
            completed = subprocess.run(command, cwd=source_root, env=env, capture_output=True, text=True)
            stream.write(completed.stdout)
            stream.write(completed.stderr)
            if completed.returncode != 0:
                raise RuntimeError(f"preflight failed ({completed.returncode}): {case}")
            lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
            measurement = next((line for line in lines if line.startswith(f"preflight case={case} ")), None)
            if measurement is None:
                raise RuntimeError(f"preflight omitted measurement for {case}")
            description_line = next((line for line in lines if line.startswith("preflight-json ")), None)
            if description_line is None:
                raise RuntimeError(f"preflight omitted formula description for {case}")
            try:
                description = json.loads(description_line.removeprefix("preflight-json "))
            except json.JSONDecodeError as error:
                raise RuntimeError(f"preflight emitted invalid formula description for {case}: {error}") from error
            row = preflight_matrix_row(case)
            for field in ("source", "evaluation_path", "shape"):
                if description.get(field) != row.get(field):
                    raise RuntimeError(
                        f"preflight {field} mismatch for {case}: matrix={row.get(field)!r}, observed={description.get(field)!r}"
                    )
            if "expected" not in description or not same_expected_outcome(description["expected"], row["expected"]):
                raise RuntimeError(
                    f"preflight expected outcome mismatch for {case}: "
                    f"matrix={row['expected']!r}, harness={description.get('expected')!r}"
                )
            match = re.search(r"reference_reads=(\d+)", measurement)
            if match is None:
                raise RuntimeError(f"preflight omitted reference reads for {case}")
            reads = int(match.group(1))
            expected = preflight_read_bound(case)
            if description.get("reference_reads") != reads:
                raise RuntimeError(f"preflight JSON/text read count differs for {case}")
            if reads != expected:
                raise RuntimeError(f"preflight read bound failed for {case}: expected {expected}, observed {reads}")
            receipt["reference_reads"][case] = reads
            receipt["descriptions"][case] = {
                "source": description["source"],
                "evaluation_path": description["evaluation_path"],
                "shape": description["shape"],
                "expected": description["expected"],
            }
    receipt["stdout"] = str(stdout_path.relative_to(output_dir))
    receipt["status"] = "ok"
    (output_dir / "preflight.json").write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return receipt


def rss_kib(path: Path) -> int:
    match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", path.read_text(encoding="utf-8"))
    if match is None:
        raise RuntimeError(f"missing maximum RSS in {path}")
    return int(match.group(1))


def staged_cases(binary: Path, *, source_root: Path, env: dict[str, str]) -> list[str]:
    output = subprocess.check_output([str(binary), "--list"], cwd=source_root, env=env, text=True)
    observed = [line.strip() for line in output.splitlines() if line.strip()]
    expected = expected_cases()
    if set(observed) != set(expected) or len(observed) != len(expected):
        raise RuntimeError(f"harness case set differs from matrix: observed={len(observed)}, expected={len(expected)}")
    return expected


def preflight_candidate_tree(*, candidate_root: Path, output_dir: Path, profile_hashes: dict[str, str]) -> dict[str, Any]:
    preserve_existing_output(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    harness_manifest = candidate_root / REL_HARNESS / "Cargo.toml"
    before = source_snapshot(candidate_root, profile_hashes)
    target = Path(tempfile.mkdtemp(prefix="litchi-ods-date-time-candidate-preflight-", dir="/var/tmp"))
    build_env = os.environ.copy()
    build_env["CARGO_TARGET_DIR"] = str(target)
    build_env["CARGO_TERM_COLOR"] = "never"
    command = [
        "cargo", "build", "--locked", "--offline", "--manifest-path", str(harness_manifest),
        "--release", "--features", "date-time-candidate",
    ]
    try:
        run_checked(command, cwd=candidate_root, env=build_env, stdout=output_dir / "build.log")
        binary = target / "release" / BINARY_NAME
        if not binary.is_file():
            raise RuntimeError(f"missing release binary: {binary}")
        cases = staged_cases(binary, source_root=candidate_root, env=build_env)
        preflight = run_preflight(binary=binary, source_root=candidate_root, env=build_env, output_dir=output_dir, cases=cases)
        assert_unchanged(before, source_snapshot(candidate_root, profile_hashes), "candidate preflight")
        result = {"status": "ok", "cases": len(cases), "binary_sha256": digest(binary), "contract_sha256": CONTRACT_SHA256, "profile_input_sha256": profile_hashes, "preflight": preflight}
        (output_dir / "summary.json").write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        return result
    finally:
        shutil.rmtree(target, ignore_errors=True)
        (output_dir / "target-cleanup.json").write_text(json.dumps({"target_dir": str(target), "removed": not target.exists()}, indent=2) + "\n", encoding="utf-8")


def verify_candidate_freeze(candidate_root: Path, freeze_path: Path) -> dict[str, Any]:
    baseline = json.loads(BASELINE_PATH.read_text(encoding="utf-8"))
    if baseline.get("production_commit") != BASELINE_COMMIT:
        raise RuntimeError("date/time baseline production_commit differs from the matched baseline")
    if baseline.get("preparation_commit") != PREPARATION_COMMIT:
        raise RuntimeError("date/time baseline preparation_commit differs from the staged source")
    freeze = json.loads(freeze_path.read_text(encoding="utf-8"))
    frozen_source_root = freeze.get("source_root")
    if frozen_source_root is not None and Path(frozen_source_root).resolve() != candidate_root:
        raise RuntimeError(f"candidate root differs from freeze: {candidate_root} != {frozen_source_root}")
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict) or not selected:
        raise RuntimeError("candidate freeze has no selected file hashes")
    mismatches = []
    for relative, expected in selected.items():
        path = candidate_root / relative
        actual = digest(path) if path.is_file() else None
        if actual != expected:
            mismatches.append(f"{relative}: expected {expected}, observed {actual}")
    if mismatches:
        raise RuntimeError("candidate freeze hash mismatch:\n" + "\n".join(mismatches))
    preparation = subprocess.check_output(["git", "rev-parse", PREPARATION_COMMIT], cwd=ROOT, text=True).strip()
    if freeze.get("base_commit") != preparation:
        raise RuntimeError("candidate freeze base differs from recorded preparation commit")
    if freeze.get("production_commit") != BASELINE_COMMIT:
        raise RuntimeError("candidate freeze production commit differs from matched baseline")
    if git_head(candidate_root) != preparation:
        raise RuntimeError("candidate checkout HEAD differs from frozen preparation base")
    for relative in ("Cargo.toml", "Cargo.lock", "src/main.rs"):
        key = str(REL_HARNESS / relative)
        if key not in selected:
            raise RuntimeError(f"candidate freeze does not bind harness input: {key}")
    return freeze


def assert_unchanged(before: dict[str, Any], after: dict[str, Any], context: str) -> None:
    for field in ("git_head", "source_sha256", "workspace_source_sha256",
                  "workspace_lock_sha256", "harness_sha256", "profile_input_sha256"):
        if before[field] != after[field]:
            raise RuntimeError(f"{field} changed during {context}")


def capture_tree(*, label: str, source_root: Path, output_dir: Path, warmups: int, samples: int, requested_cases: tuple[str, ...] | None, profile_hashes: dict[str, str]) -> dict[str, Any]:
    preserve_existing_output(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    raw_dir = output_dir / "raw"
    raw_dir.mkdir()
    harness_root = copy_harness(source_root) if requested_cases is not None else source_root / REL_HARNESS
    harness_manifest = harness_root / "Cargo.toml"
    target = Path(tempfile.mkdtemp(prefix=f"litchi-ods-date-time-{label}-", dir="/var/tmp"))
    before = source_snapshot(source_root, profile_hashes)
    build_env = os.environ.copy()
    build_env["CARGO_TARGET_DIR"] = str(target)
    build_env["CARGO_TERM_COLOR"] = "never"
    build_command = ["cargo", "build", "--locked", "--offline", "--manifest-path", str(harness_manifest), "--release"]
    if requested_cases is None:
        build_command.extend(["--features", "date-time-candidate"])
    (output_dir / "build-command.json").write_text(json.dumps({"command": build_command, "cwd": str(source_root), "target": str(target)}, indent=2) + "\n", encoding="utf-8")
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
        preflight = run_preflight(binary=binary, source_root=source_root, env=build_env, output_dir=output_dir, cases=cases)
        environment = {
            "label": label, "source_root": str(source_root), "source_git_head": before["git_head"],
            "rustc": subprocess.check_output(["rustc", "--version"], text=True).strip(),
            "rustc_verbose": tool_text(["rustc", "-Vv"]), "cargo": subprocess.check_output(["cargo", "--version"], text=True).strip(),
            "libc": tool_text(["getconf", "GNU_LIBC_VERSION"]), "kernel": os.uname().release,
            "cpu": subprocess.check_output(["lscpu"], text=True), "warmups_per_child": warmups,
            "samples_per_group": samples, "cases": cases, "phases": list(PHASES),
            "case_scope": "matched-controls" if requested_cases is not None else "all-date-time-cases",
            "preflight": preflight, "binary_sha256": binary_hash, "contract_sha256": CONTRACT_SHA256,
            "profile_input_sha256": profile_hashes,
        }
        environment["measurement_context_before"] = measurement_context(build_env)
        environment["measurement_context_after"] = None
        (output_dir / "environment.json").write_text(json.dumps(environment, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        records: list[dict[str, Any]] = []
        (output_dir / "commands.txt").write_text("", encoding="utf-8")
        for case in cases:
            for phase in PHASES:
                for sample_index in range(samples):
                    stem = f"{case}.{phase}.sample-{sample_index + 1:02d}"
                    stdout_path = raw_dir / f"{stem}.stdout.json"
                    stderr_path = raw_dir / f"{stem}.stderr.log"
                    time_path = raw_dir / f"{stem}.time.txt"
                    command = ["/usr/bin/time", "-v", "-o", str(time_path), str(binary), "--case", case, "--phase", phase, "--warmups", str(warmups), "--iterations", "1"]
                    with (output_dir / "commands.txt").open("a", encoding="utf-8") as stream:
                        stream.write(command_text(command) + "\n")
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
                    row.update({"capture": label, "sample_index": sample_index + 1, "rss_kib": rss_kib(time_path), "binary_sha256": binary_hash, "source_git_head": before["git_head"], "raw_stdout": str(stdout_path.relative_to(output_dir)), "raw_stderr": raw_stderr, "raw_time": str(time_path.relative_to(output_dir))})
                    records.append(row)
        environment["measurement_context_after"] = measurement_context(build_env)
        (output_dir / "environment.json").write_text(json.dumps(environment, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        after = source_snapshot(source_root, profile_hashes)
        assert_unchanged(before, after, f"{label} capture")
        if profile_hashes != profile_input_snapshot():
            raise RuntimeError(f"profile inputs changed during {label} capture")
        (output_dir / "source-manifest.json").write_text(json.dumps({"before": before, "after": after, "source_sha256_unchanged": True, "workspace_source_sha256_unchanged": True, "profile_input_sha256_unchanged": True, "binary_sha256": binary_hash}, indent=2, sort_keys=True) + "\n", encoding="utf-8")
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
    baseline = json.loads(BASELINE_PATH.read_text(encoding="utf-8"))
    if baseline.get("production_commit") != BASELINE_COMMIT:
        raise RuntimeError("date/time baseline production_commit differs from the matched baseline")
    if baseline.get("preparation_commit") != PREPARATION_COMMIT:
        raise RuntimeError("date/time baseline preparation_commit differs from the staged source")
    for path in (CONTRACT_PATH, MATRIX_PATH, ORACLE_PATH, ORACLE_VERIFY_PATH):
        if not path.is_file():
            raise RuntimeError(f"missing reviewed date/time input: {path}")
    if digest(CONTRACT_PATH) != CONTRACT_SHA256:
        raise RuntimeError("date/time contract hash differs from the reviewed input")
    if digest(ORACLE_PATH) != ORACLE_SHA256 or digest(ORACLE_VERIFY_PATH) != ORACLE_VERIFY_SHA256:
        raise RuntimeError("date/time oracle input hash differs from the reviewed input")
    oracle_check = subprocess.run(["python3", str(ORACLE_VERIFY_PATH), "--check"], cwd=ROOT, capture_output=True, text=True)
    if oracle_check.returncode != 0:
        raise RuntimeError("date/time oracle check failed: " + (oracle_check.stdout + oracle_check.stderr).strip())
    expected_cases()


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
    if not GATE_LOCK.is_file() or digest(GATE_LOCK) != GATE_LOCK_SHA256:
        raise RuntimeError(f"retained date/time gate Cargo.lock missing or hash differs: {GATE_LOCK}")
    profile_before = profile_input_snapshot()
    results_root = HERE / "results"
    results_root.mkdir(parents=True, exist_ok=True)
    (results_root / "profile-inputs-before.json").write_text(json.dumps(profile_before, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    summaries: list[dict[str, Any]] = []
    candidate_preflight: dict[str, Any] | None = None
    freeze: dict[str, Any] | None = None
    if not args.baseline_only:
        candidate_root = args.candidate_root.resolve()
        freeze = verify_candidate_freeze(candidate_root, args.candidate_freeze.resolve())
        candidate_preflight = preflight_candidate_tree(candidate_root=candidate_root, output_dir=results_root / "preflight-before-timing", profile_hashes=profile_before)
    baseline_worktree = Path(tempfile.mkdtemp(prefix="litchi-ods-date-time-baseline-", dir="/var/tmp"))
    baseline_worktree.rmdir()
    cleanup: dict[str, Any] = {"baseline_worktree": str(baseline_worktree)}
    try:
        add_baseline_worktree(baseline_worktree, GATE_LOCK)
        summaries.append(capture_tree(label=f"baseline-{BASELINE_COMMIT}", source_root=baseline_worktree, output_dir=results_root / f"baseline-{BASELINE_COMMIT}", warmups=args.warmups, samples=args.samples, requested_cases=MATCHED_CONTROL_CASES, profile_hashes=profile_before))
    finally:
        cleanup["baseline_removed"] = remove_worktree(baseline_worktree)
    if not args.baseline_only:
        candidate_root = args.candidate_root.resolve()
        verify_candidate_freeze(candidate_root, args.candidate_freeze.resolve())
        summaries.append(capture_tree(label="candidate-final", source_root=candidate_root, output_dir=results_root / "candidate-final", warmups=args.warmups, samples=args.samples, requested_cases=None, profile_hashes=profile_before))
        summaries[-1]["freeze_path"] = str(args.candidate_freeze.resolve())
        summaries[-1]["freeze_base_commit"] = freeze["base_commit"] if freeze else None
    profile_after = profile_input_snapshot()
    if profile_before != profile_after:
        raise RuntimeError("profile inputs changed between captures")
    (results_root / "profile-inputs-after.json").write_text(json.dumps(profile_after, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    (results_root / "capture-summary.json").write_text(json.dumps({"baseline_commit": BASELINE_COMMIT, "contract_sha256": CONTRACT_SHA256, "oracle_sha256": ORACLE_SHA256, "controls": list(MATCHED_CONTROL_CASES), "date_cases": list(DATE_CASES), "phases": list(PHASES), "warmups_per_child": args.warmups, "samples_per_group": args.samples, "profile_input_sha256_before": profile_before, "profile_input_sha256_after": profile_after, "candidate_preflight": candidate_preflight, "captures": summaries, "cleanup": cleanup}, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
