#!/usr/bin/env python3
"""Bounded baseline/candidate runner for the ODS financial profile.

The runner is capture-only.  It builds the public financial profile example in
isolated checkouts, preflights the complete matrix, and retains raw child
receipts.  A baseline Unsupported result is a capability receipt and is never
treated as a timed zero-cost financial implementation.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
import json
import math
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
from typing import Any

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[4]
PROFILE_REL = HERE.relative_to(ROOT)
MATRIX_PATH = HERE / "case-matrix.json"
EXAMPLE_REL = Path("crates/litchi-ods/examples/ods_formula_financial_profile.rs")
BASELINE_COMMIT = "f945a5129ad8967a863cf4937de01ee2152b9fbf"
CONTRACT_SHA256 = "fbe2283886326582b0e34fd277b38e6fb3e1bdeb2ecfbec5522da0f6f1467e4b"
BINARY_NAME = "ods_formula_financial_profile"
PHASES = ("evaluate", "parse-evaluate")

# Keep these declarations literal: staging and review tools consume them with
# ast.literal_eval without importing this runner.
SOURCE_FILES = (
    "Cargo.toml",
    "Cargo.lock",
    "crates/litchi-ods/Cargo.toml",
    "crates/litchi-ods/src/codec/formula/functions.rs",
    "crates/litchi-ods/src/codec/formula/evaluation.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/numerics.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/dyadic.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/date_time.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/paired.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/reference_metadata.rs",
    "crates/litchi-ods/examples/ods_formula_financial_profile.rs",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/contract.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/numerical-design.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/oracle_scalar.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/oracle_reducers.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/oracle-scalar-vectors.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/oracle-reducer-vectors.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/native/provenance.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/performance/plan.md",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/performance/case-matrix.json",
)
SOURCE_FILE_GLOBS = (
    "crates/litchi-ods/src/codec/formula/evaluation/**/*.rs",
    "crates/litchi-ods/src/codec/formula/evaluation*.rs",
    "crates/litchi-ods/tests/ods_formula_*financial*.rs",
    "crates/litchi-ods/tests/ods_formula_*value*.rs",
    "crates/litchi-core/src/budget.rs",
    "crates/litchi-core/src/execution.rs",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/native/**/*",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/oracle*.json",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/oracle*.py",
    "docs/report/spec-gap-validation-evidence/ods-formula-financial/*.md",
)
FEATURE_MATRIX = (
    ("financial-scalar", "eleven fixed-size scalar kernels", 0, "scalar"),
    ("financial-stream", "FVSCHEDULE/NPV/XNPV ordered streaming", "N", "reference"),
    ("financial-array", "MIRR matrix input and fixed numeric state", "N", "array"),
    ("financial-solver", "IRR/RATE/XIRR bounded repeated roots", "N", "solver"),
    ("financial-refusal", "known wrong sequence shape before reads", 0, "typed-error"),
    ("financial-resource", "reference/work/storage limits", "case", "typed-failure"),
    ("financial-cancellation", "sticky and mid-scan cancellation", "case", "typed-failure"),
    ("financial-error", "formula-error retention and typed precedence", "case", "error"),
)


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def command_text(command: list[str]) -> str:
    return subprocess.list2cmdline(command)


def load_matrix() -> dict[str, Any]:
    value = json.loads(MATRIX_PATH.read_text(encoding="utf-8"))
    if value.get("contract_sha256") != CONTRACT_SHA256:
        raise RuntimeError("financial case matrix contract hash is stale")
    if value.get("phases") != list(PHASES):
        raise RuntimeError("financial case matrix phase set is stale")
    controls = value.get("matched_controls")
    candidate = value.get("candidate_cases")
    if not isinstance(controls, list) or not isinstance(candidate, list):
        raise RuntimeError("financial case matrix has malformed case lists")
    rows = controls + candidate
    names = [row.get("name") for row in rows]
    if any(not isinstance(name, str) for name in names) or len(set(names)) != len(names):
        raise RuntimeError("financial case matrix has missing or duplicate names")
    counts = value.get("counts", {})
    expected_counts = {
        "matched_controls": len(controls),
        "candidate_cases": len(candidate),
        "baseline_timed_groups": len(controls) * len(PHASES),
        "candidate_timed_groups": len(rows) * len(PHASES),
    }
    if any(counts.get(key) != expected for key, expected in expected_counts.items()):
        raise RuntimeError(f"financial case matrix counts differ: {counts!r}")
    for row in rows:
        for key in ("name", "source", "evaluation_path", "shape", "expected", "reference_reads", "repeat", "class"):
            if key not in row:
                raise RuntimeError(f"matrix row {row.get('name')!r} lacks {key}")
        if not isinstance(row["reference_reads"], int) or row["reference_reads"] < 0:
            raise RuntimeError(f"matrix row {row['name']} has invalid read bound")
        if not isinstance(row["repeat"], int) or row["repeat"] <= 0:
            raise RuntimeError(f"matrix row {row['name']} has invalid repeat")
    return value


def same_outcome(observed: Any, expected: Any) -> bool:
    if isinstance(expected, bool):
        return isinstance(observed, bool) and observed == expected
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return (
            isinstance(observed, (int, float))
            and not isinstance(observed, bool)
            and math.isfinite(float(observed))
            and math.isfinite(float(expected))
            and abs(float(observed) - float(expected)) <= max(abs(float(expected)), 1.0) * 1.0e-10
        )
    if isinstance(expected, list):
        return isinstance(observed, list) and len(observed) == len(expected) and all(
            same_outcome(actual, wanted) for actual, wanted in zip(observed, expected)
        )
    if isinstance(expected, dict):
        return isinstance(observed, dict) and observed.keys() == expected.keys() and all(
            same_outcome(observed[key], value) for key, value in expected.items()
        )
    return type(observed) is type(expected) and observed == expected


def profile_input_snapshot() -> dict[str, str]:
    hashes: dict[str, str] = {}
    for path in sorted(HERE.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(HERE)
        if any(part in {"results", "target", "__pycache__"} for part in relative.parts):
            continue
        hashes[str(relative)] = digest(path)
    if not hashes:
        raise RuntimeError("financial profile input directory is empty")
    return hashes


def source_files(source_root: Path) -> list[Path]:
    paths = {source_root / relative for relative in SOURCE_FILES}
    for pattern in SOURCE_FILE_GLOBS:
        paths.update(source_root.glob(pattern))
    return sorted(path for path in paths if path.is_file())


def source_snapshot(source_root: Path, profile_hashes: dict[str, str]) -> dict[str, Any]:
    example = source_root / EXAMPLE_REL
    if not example.is_file():
        raise RuntimeError(f"financial profile example is missing from source snapshot: {example}")
    selected = {
        str(path.relative_to(source_root)): digest(path)
        for path in source_files(source_root)
    }
    selected[str(EXAMPLE_REL)] = digest(example)
    return {
        "git_head": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=source_root, text=True).strip(),
        "source_sha256": selected,
        "profile_input_sha256": profile_hashes,
        "lock_sha256": digest(source_root / "Cargo.lock") if (source_root / "Cargo.lock").is_file() else None,
    }


def stage_example(source_root: Path) -> Path | None:
    source = ROOT / EXAMPLE_REL
    if not source.is_file():
        raise RuntimeError(f"financial profile example is missing from root: {source}")
    destination = source_root / EXAMPLE_REL
    if destination.exists() or destination.is_symlink():
        if not destination.is_file():
            raise RuntimeError(f"financial profile example destination is not a file: {destination}")
        expected = digest(source)
        observed = digest(destination)
        if observed != expected:
            raise RuntimeError(
                f"existing financial profile example differs from root: {destination}"
            )
        return None
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    if digest(destination) != digest(source):
        destination.unlink(missing_ok=True)
        raise RuntimeError(f"staged financial profile example failed hash verification: {destination}")
    return destination


def build_binary(source_root: Path, label: str, output_dir: Path) -> tuple[Path, Path, dict[str, Any]]:
    staged: Path | None = None
    target: Path | None = None
    succeeded = False
    try:
        staged = stage_example(source_root)
        target = Path(tempfile.mkdtemp(prefix=f"litchi-ods-financial-{label}-", dir="/var/tmp"))
        env = os.environ.copy()
        env.update({"CARGO_TARGET_DIR": str(target), "CARGO_TERM_COLOR": "never"})
        manifest = source_root / "crates/litchi-ods/Cargo.toml"
        command = [
            "cargo", "build", "--locked", "--offline", "--manifest-path", str(manifest),
            "--example", "ods_formula_financial_profile", "--release",
        ]
        log = output_dir / f"{label}-build.log"
        output_dir.mkdir(parents=True, exist_ok=True)
        with log.open("w", encoding="utf-8") as stream:
            completed = subprocess.run(command, cwd=source_root, env=env, stdout=stream, stderr=subprocess.STDOUT)
        if completed.returncode != 0:
            raise RuntimeError(f"{label} build failed ({completed.returncode}); see {log}")
        binary = target / "release" / "examples" / BINARY_NAME
        if not binary.is_file():
            raise RuntimeError(f"{label} build omitted {binary}")
        succeeded = True
        return binary, target, {
            "command": command,
            "cwd": str(source_root),
            "target": str(target),
            "binary_sha256": digest(binary),
            "build_log": str(log.relative_to(output_dir)),
        }
    finally:
        if not succeeded and target is not None:
            shutil.rmtree(target, ignore_errors=True)
        if staged is not None:
            staged.unlink(missing_ok=True)


def parse_json_line(stdout: str, prefix: str) -> dict[str, Any]:
    lines = [line.strip() for line in stdout.splitlines() if line.strip()]
    candidates = [line.removeprefix(prefix) for line in lines if line.startswith(prefix)]
    if not candidates:
        raise RuntimeError(f"profile child omitted {prefix!r}: {stdout!r}")
    value = json.loads(candidates[-1])
    if not isinstance(value, dict):
        raise RuntimeError("profile child emitted a non-object receipt")
    return value


def matrix_rows(matrix: dict[str, Any], controls_only: bool = False) -> list[dict[str, Any]]:
    return list(matrix["matched_controls"] if controls_only else matrix["matched_controls"] + matrix["candidate_cases"])


def preflight(binary: Path, source_root: Path, rows: list[dict[str, Any]], output_dir: Path, *, allow_unsupported: bool) -> dict[str, Any]:
    receipt: dict[str, Any] = {"status": "ok", "rows": {}, "unsupported": [], "commands": []}
    stdout_log = output_dir / "preflight.stdout.log"
    with stdout_log.open("w", encoding="utf-8") as log:
        for row in rows:
            command = [str(binary), "--case", row["name"], "--phase", "evaluate", "--warmups", "0", "--iterations", "1", "--preflight-only"]
            if allow_unsupported:
                command.append("--allow-unsupported")
            receipt["commands"].append(command_text(command))
            completed = subprocess.run(command, cwd=source_root, capture_output=True, text=True)
            log.write(completed.stdout)
            log.write(completed.stderr)
            if completed.returncode != 0:
                raise RuntimeError(f"preflight failed for {row['name']}: {completed.stderr or completed.stdout}")
            description = parse_json_line(completed.stdout, "preflight-json ")
            observed_supported = description.get("supported")
            if observed_supported is False:
                if not allow_unsupported:
                    raise RuntimeError(f"candidate row {row['name']} is unsupported")
                receipt["unsupported"].append(row["name"])
                receipt["rows"][row["name"]] = {"supported": False, "description": description}
                continue
            for field in ("case", "source", "evaluation_path", "shape", "class", "repeat"):
                expected = row["name"] if field == "case" else row[field]
                if description.get(field) != expected:
                    raise RuntimeError(f"preflight {field} mismatch for {row['name']}: expected {expected!r}, observed {description.get(field)!r}")
            if "max_steps" in row and description.get("max_steps") != row["max_steps"]:
                raise RuntimeError(
                    f"preflight max_steps mismatch for {row['name']}: "
                    f"expected {row['max_steps']!r}, observed {description.get('max_steps')!r}"
                )
            if not same_outcome(description.get("expected"), row["expected"]):
                raise RuntimeError(f"preflight expected mismatch for {row['name']}")
            observed_reads = description.get("reference_reads")
            if observed_reads != row["reference_reads"]:
                raise RuntimeError(f"preflight read mismatch for {row['name']}: expected {row['reference_reads']}, observed {observed_reads}")
            receipt["rows"][row["name"]] = {"supported": True, "description": description}
    receipt["stdout"] = str(stdout_log.relative_to(output_dir))
    (output_dir / "preflight.json").write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return receipt


def rss_kib(path: Path) -> int | None:
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line.startswith("Maximum resident set size (kbytes):"):
            return int(line.split(":", 1)[1].strip())
    return None


def preserve_existing_output(path: Path) -> Path | None:
    if not path.exists() or not any(path.iterdir()):
        return None
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    destination = path.with_name(f"diagnostic-{path.name}-{stamp}")
    suffix = 1
    while destination.exists():
        destination = path.with_name(f"diagnostic-{path.name}-{stamp}-{suffix}")
        suffix += 1
    path.rename(destination)
    return destination


def capture_side(
    source_root: Path,
    label: str,
    rows: list[dict[str, Any]],
    output_dir: Path,
    warmups: int,
    samples: int,
    profile_hashes: dict[str, str],
    *,
    allow_unsupported: bool,
    preflight_rows: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    staged = stage_example(source_root)
    before = source_snapshot(source_root, profile_hashes)
    binary, target, build = build_binary(source_root, label, output_dir)
    try:
        receipt_rows = rows if preflight_rows is None else preflight_rows
        preflight_receipt = preflight(binary, source_root, receipt_rows, output_dir, allow_unsupported=allow_unsupported)
        supported_rows = [row for row in rows if row["name"] not in preflight_receipt["unsupported"]]
        raw_dir = output_dir / "raw"
        raw_dir.mkdir(parents=True, exist_ok=True)
        records = []
        for row in supported_rows:
            for phase in PHASES:
                for sample_index in range(samples):
                    stem = f"{row['name']}.{phase}.sample-{sample_index + 1:02d}"
                    stdout_path = raw_dir / f"{stem}.stdout.log"
                    stderr_path = raw_dir / f"{stem}.stderr.log"
                    time_path = raw_dir / f"{stem}.time.txt"
                    command = ["/usr/bin/time", "-v", "-o", str(time_path), str(binary), "--case", row["name"], "--phase", phase, "--warmups", str(warmups), "--iterations", "1"]
                    with stdout_path.open("w", encoding="utf-8") as stdout, stderr_path.open("w", encoding="utf-8") as stderr:
                        completed = subprocess.run(command, cwd=source_root, stdout=stdout, stderr=stderr, text=True)
                    if completed.returncode != 0:
                        raise RuntimeError(f"timed row failed: {stem}")
                    output = parse_json_line(stdout_path.read_text(encoding="utf-8"), "")
                    observed_rss = rss_kib(time_path)
                    for sample in output.get("samples", []):
                        if isinstance(sample, dict):
                            sample["rss_kib"] = observed_rss
                    output.update({
                        "capture": label,
                        "sample_index": sample_index + 1,
                        "case": row["name"],
                        "phase": phase,
                        "rss_kib": observed_rss,
                        "raw_stdout": str(stdout_path.relative_to(output_dir)),
                        "raw_stderr": str(stderr_path.relative_to(output_dir)),
                        "raw_time": str(time_path.relative_to(output_dir)),
                        "binary_sha256": build["binary_sha256"],
                    })
                    records.append(output)
        (output_dir / "measurements.jsonl").write_text(
            "".join(json.dumps(record, sort_keys=True) + "\n" for record in records), encoding="utf-8"
        )
        after = source_snapshot(source_root, profile_hashes)
        if before != after:
            raise RuntimeError(f"{label} source identity changed during capture")
        result = {
            "label": label,
            "source": before,
            "build": build,
            "preflight": preflight_receipt,
            "records": len(records),
            "warmups": warmups,
            "samples": samples,
        }
        (output_dir / "capture-summary.json").write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        return result
    finally:
        shutil.rmtree(target, ignore_errors=True)
        if staged is not None:
            staged.unlink(missing_ok=True)
        (output_dir / "target-cleanup.json").write_text(
            json.dumps({"target": str(target), "removed": not target.exists()}, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )


def verify_freeze(candidate_root: Path, freeze_path: Path) -> dict[str, Any]:
    freeze = json.loads(freeze_path.read_text(encoding="utf-8"))
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict) or not selected:
        raise RuntimeError("candidate freeze has no selected_files map")
    mismatches = []
    for relative, expected in selected.items():
        path = candidate_root / relative
        observed = digest(path) if path.is_file() else None
        if observed != expected:
            mismatches.append(f"{relative}: expected {expected}, observed {observed}")
    if mismatches:
        raise RuntimeError("candidate freeze mismatch:\n" + "\n".join(mismatches))
    return freeze


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--candidate-root", type=Path, default=ROOT)
    parser.add_argument("--baseline-root", type=Path)
    parser.add_argument("--candidate-freeze", type=Path)
    parser.add_argument("--results", type=Path, default=HERE / "results")
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--samples", type=int, default=15)
    parser.add_argument("--preflight-only", action="store_true")
    parser.add_argument("--capture", action="store_true")
    args = parser.parse_args()
    matrix = load_matrix()
    profile_hashes = profile_input_snapshot()
    if args.candidate_freeze:
        verify_freeze(args.candidate_root, args.candidate_freeze)
    args.results.mkdir(parents=True, exist_ok=True)
    if args.preflight_only:
        binary, target, build = build_binary(args.candidate_root, "candidate-preflight", args.results)
        try:
            result = preflight(binary, args.candidate_root, matrix_rows(matrix), args.results, allow_unsupported=False)
            print(json.dumps({"status": "ok", "build": build, "preflight": result}, indent=2, sort_keys=True))
        finally:
            shutil.rmtree(target, ignore_errors=True)
        return 0
    if not args.capture:
        parser.error("choose --preflight-only or --capture")
    if args.baseline_root is None:
        parser.error("--baseline-root is required for --capture")
    preserve_existing_output(args.results / "baseline")
    preserve_existing_output(args.results / "candidate")
    capture_side(
        args.baseline_root,
        "baseline",
        matrix_rows(matrix, controls_only=True),
        args.results / "baseline",
        args.warmups,
        args.samples,
        profile_hashes,
        allow_unsupported=True,
        preflight_rows=matrix_rows(matrix),
    )
    capture_side(
        args.candidate_root,
        "candidate",
        matrix_rows(matrix),
        args.results / "candidate",
        args.warmups,
        args.samples,
        profile_hashes,
        allow_unsupported=False,
    )
    print("financial capture complete")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
