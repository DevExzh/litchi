#!/usr/bin/env python3
"""Build and capture the isolated ODS array/reference candidate serially.

This is an orchestration tool for a diagnostic candidate capture.  It never
cleans a Cargo target, overwrites an evidence directory, or turns a measured
review trigger into an acceptance decision.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import importlib.util
import json
import os
import platform
import subprocess
import sys
import time
from collections import Counter
from pathlib import Path
from typing import Any, Callable, Iterable


PERFORMANCE = Path(__file__).resolve().parent
EVIDENCE = PERFORMANCE.parent
# ``PERFORMANCE`` is already a directory; the repository is four parents up
# from it (the equivalent file-based expression is ``Path(__file__).parents[5]``).
CANONICAL_ROOT = PERFORMANCE.parents[4]
GATES_RUNNER = EVIDENCE / "gates/run.py"
COMPARE = PERFORMANCE / "compare.py"
BASELINE_DEFAULT = PERFORMANCE / "baseline/raw.csv"
TARGET_DEFAULT = Path("/home/zhuhe/code/litchi-array-target")
TMP_DEFAULT = Path("/home/zhuhe/code/litchi-array-tmp")
BASELINE_BINARY_DEFAULT = TARGET_DEFAULT / "retained/scalar-baseline"
PROFILE_CPU = 6

PERFORMANCE_REL = PERFORMANCE.relative_to(CANONICAL_ROOT)
HARNESS_FILES = tuple(
    PERFORMANCE_REL / name
    for name in (
        "scalar-harness/Cargo.toml",
        "scalar-harness/Cargo.lock",
        "scalar-harness/run.py",
        "scalar-harness/src/main.rs",
        "value-harness/Cargo.toml",
        "value-harness/Cargo.lock",
        "value-harness/run.py",
        "value-harness/src/main.rs",
    )
)
TOOL_FILES = (
    PERFORMANCE_REL / "run_candidate.py",
    PERFORMANCE_REL / "README.md",
    PERFORMANCE_REL / "compare.py",
    GATES_RUNNER.relative_to(CANONICAL_ROOT),
)

SCALAR_RUNNER = PERFORMANCE_REL / "scalar-harness/run.py"
VALUE_RUNNER = PERFORMANCE_REL / "value-harness/run.py"
SCALAR_MANIFEST = PERFORMANCE_REL / "scalar-harness/Cargo.toml"
VALUE_MANIFEST = PERFORMANCE_REL / "value-harness/Cargo.toml"
SCALAR_BINARY_NAME = "ods-formula-array-reference-evaluation-scalar-profile"
VALUE_BINARY_NAME = "ods-formula-array-reference-evaluation-value-profile"


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="milliseconds").replace(
        "+00:00", "Z"
    )


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def hash_paths(root: Path, paths: Iterable[Path]) -> dict[str, str]:
    result: dict[str, str] = {}
    for relative in paths:
        path = root / relative
        if not path.is_file():
            raise RuntimeError(f"required input is missing: {path}")
        result[relative.as_posix()] = sha256_file(path)
    return result


def mapping_digest(mapping: dict[str, str]) -> str:
    digest = hashlib.sha256()
    for name, value in sorted(mapping.items()):
        digest.update(name.encode("utf-8"))
        digest.update(b"\0")
        digest.update(value.encode("ascii"))
        digest.update(b"\n")
    return digest.hexdigest()


def write_json(path: Path, value: Any) -> None:
    path.write_text(
        json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n",
        encoding="utf-8",
    )


def load_gate_hashes() -> Callable[[Path], dict[str, str]]:
    spec = importlib.util.spec_from_file_location(
        "ods_candidate_gate_hashes", GATES_RUNNER
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load gate hash function: {GATES_RUNNER}")
    module = importlib.util.module_from_spec(spec)
    write_bytecode = sys.dont_write_bytecode
    sys.dont_write_bytecode = True
    try:
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = write_bytecode
    hashes = getattr(module, "hashes", None)
    if not callable(hashes):
        raise RuntimeError("gates/run.py does not expose callable hashes(root)")
    return hashes


def git_head(root: Path) -> str:
    completed = subprocess.run(
        ["git", "rev-parse", "HEAD"],
        cwd=root,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    if completed.returncode != 0:
        detail = completed.stderr.strip() or "git rev-parse failed"
        raise RuntimeError(f"{root}: {detail}")
    return completed.stdout.strip()


def version_command(command: list[str], cwd: Path) -> dict[str, Any]:
    completed = subprocess.run(
        command,
        cwd=cwd,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    return {
        "command": command,
        "status": completed.returncode,
        "stdout": completed.stdout,
        "stderr": completed.stderr,
    }


def host_metadata() -> dict[str, Any]:
    # Keep this intentionally small.  In particular, do not collect process
    # tables or /proc command lines: those can contain unrelated secrets.
    return {
        "platform": {
            "system": platform.system(),
            "release": platform.release(),
            "machine": platform.machine(),
        },
        "python": platform.python_version(),
        "cpu_count": os.cpu_count(),
        "profile_cpu": PROFILE_CPU,
        "profile_affinity": [PROFILE_CPU],
    }


def ensure_fresh_output(
    path: Path,
    baseline: Path,
    baseline_binary: Path,
    target: Path,
    temp: Path,
) -> None:
    if path.exists():
        raise RuntimeError(f"output root already exists; choose a fresh path: {path}")
    baseline_dir = baseline.resolve().parent
    resolved = path.resolve()
    protected = {
        baseline_dir,
        baseline_binary.resolve(),
        target.resolve(),
        temp.resolve(),
    }
    if resolved in protected or any(item in resolved.parents for item in protected):
        raise RuntimeError(f"output root overlaps retained/build state: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.mkdir()


def ensure_baseline_binary_is_distinct(baseline_binary: Path, target: Path) -> None:
    """Reject a baseline ELF that Cargo could overwrite as a candidate ELF."""
    baseline_resolved = baseline_binary.resolve()
    candidates = (
        target / "release" / SCALAR_BINARY_NAME,
        target / "release" / VALUE_BINARY_NAME,
    )
    for candidate in candidates:
        if baseline_resolved == candidate.resolve():
            raise RuntimeError(
                "retained baseline ELF aliases a candidate executable: "
                f"{baseline_binary} == {candidate}"
            )
        if baseline_binary.is_file() and candidate.is_file():
            try:
                same_file = baseline_binary.samefile(candidate)
            except OSError:
                same_file = False
            if same_file:
                raise RuntimeError(
                    "retained baseline ELF is the same file as a candidate executable: "
                    f"{baseline_binary} and {candidate}"
                )


def binary_state(paths: dict[str, Path]) -> dict[str, dict[str, Any]]:
    state: dict[str, dict[str, Any]] = {}
    for name, path in paths.items():
        item: dict[str, Any] = {"path": str(path), "exists": path.is_file()}
        if path.is_file():
            item["bytes"] = path.stat().st_size
            item["sha256"] = sha256_file(path)
        else:
            item["bytes"] = None
            item["sha256"] = None
        state[name] = item
    return state


def retained_baseline_state(
    baseline: Path, baseline_binary: Path
) -> dict[str, dict[str, Any]]:
    paths = {
        "raw_csv": baseline,
        "binary": baseline_binary,
        "binary_sha256": baseline.parent / "binary-sha256.txt",
    }
    return binary_state(paths)


def run_logged(
    name: str,
    command: list[str],
    cwd: Path,
    env: dict[str, str],
    log_path: Path,
) -> dict[str, Any]:
    started = time.monotonic()
    started_at = utc_now()
    try:
        with log_path.open("w", encoding="utf-8") as log:
            completed = subprocess.run(
                command,
                cwd=cwd,
                env=env,
                stdout=log,
                stderr=subprocess.STDOUT,
                check=False,
            )
        status = completed.returncode
    except OSError as error:
        log_path.write_text(f"{type(error).__name__}: {error}\n", encoding="utf-8")
        status = 127
    return {
        "name": name,
        "command": command,
        "cwd": str(cwd),
        "log": str(log_path),
        "started_at": started_at,
        "finished_at": utc_now(),
        "seconds": time.monotonic() - started,
        "status": status,
    }


def validate_capture(path: Path, expected_rows: int) -> dict[str, Any]:
    result: dict[str, Any] = {
        "raw_csv": str(path / "raw.csv"),
        "expected_rows": expected_rows,
        "exists": False,
        "row_count": 0,
        "row_count_ok": False,
        "status_counts": {},
        "all_status_zero": False,
        "duplicate_phase_case_rows": [],
        "malformed_rows": [],
        "error": None,
    }
    raw = path / "raw.csv"
    if not raw.is_file():
        result["error"] = "raw.csv is missing"
        return result
    result["exists"] = True
    try:
        with raw.open(newline="", encoding="utf-8") as stream:
            rows = list(csv.DictReader(stream))
    except (OSError, csv.Error, UnicodeError) as error:
        result["error"] = f"cannot read raw.csv: {error}"
        return result
    result["row_count"] = len(rows)
    result["row_count_ok"] = len(rows) == expected_rows
    statuses = Counter(str(row.get("status") or "") for row in rows)
    result["status_counts"] = dict(sorted(statuses.items()))
    result["all_status_zero"] = bool(rows) and all(
        row.get("status") == "0" for row in rows
    )
    seen: set[tuple[str, str]] = set()
    duplicates: list[list[str]] = []
    malformed: list[dict[str, Any]] = []
    required_fields = ("phase", "case", "status", "p50_ns", "failure")
    for index, row in enumerate(rows, start=2):
        key = (str(row.get("phase") or ""), str(row.get("case") or ""))
        if key in seen:
            duplicates.append(list(key))
        seen.add(key)
        missing = [field for field in required_fields if not row.get(field)]
        if missing or None in row:
            malformed.append({"csv_line": index, "missing_fields": missing})
    result["duplicate_phase_case_rows"] = duplicates
    result["malformed_rows"] = malformed
    return result


def extract_review_triggers(compare_path: Path, output: Path) -> dict[str, Any]:
    review: dict[str, Any] = {
        "automatic_acceptance": False,
        "threshold_pct": 5,
        "source": str(compare_path),
        "row_count": 0,
        "rows": [],
        "error": None,
    }
    try:
        report = json.loads(compare_path.read_text(encoding="utf-8"))
        rows = report.get("rows")
        if not isinstance(rows, list):
            raise ValueError("comparison report has no rows array")
        review_rows = [row for row in rows if row.get("review_triggers")]
        review["threshold_pct"] = report.get("threshold_pct", 5)
        review["row_count"] = len(review_rows)
        review["rows"] = review_rows
    except (OSError, TypeError, ValueError, json.JSONDecodeError) as error:
        review["error"] = str(error)
    write_json(output / "review-triggers.json", review)
    return review


def source_snapshot(
    gate_hashes: Callable[[Path], dict[str, str]], workspace: Path
) -> dict[str, Any]:
    workspace_source = gate_hashes(workspace)
    canonical_source = gate_hashes(CANONICAL_ROOT)
    if len(workspace_source) != len(canonical_source):
        raise RuntimeError(
            "isolated and canonical source closures have different sizes: "
            f"{len(workspace_source)} vs {len(canonical_source)}"
        )
    workspace_harness = hash_paths(workspace, HARNESS_FILES)
    canonical_harness = hash_paths(CANONICAL_ROOT, HARNESS_FILES)
    if workspace_harness != canonical_harness:
        raise RuntimeError(
            "isolated workspace harness inputs differ from canonical harness inputs"
        )
    canonical_tools = hash_paths(CANONICAL_ROOT, TOOL_FILES)
    gate_runner_sha256 = sha256_file(GATES_RUNNER)
    return {
        "workspace_source": workspace_source,
        "canonical_source": canonical_source,
        "workspace_harness": workspace_harness,
        "canonical_harness": canonical_harness,
        "canonical_tools": canonical_tools,
        "gate_runner_sha256": gate_runner_sha256,
    }


def input_snapshot(
    gate_hashes: Callable[[Path], dict[str, str]], workspace: Path, baseline: Path
) -> dict[str, Any]:
    snapshot = source_snapshot(gate_hashes, workspace)
    snapshot["baseline_raw_sha256"] = sha256_file(baseline)
    return snapshot


def source_identity(workspace: Path, snapshot: dict[str, Any]) -> dict[str, Any]:
    return {
        "kind": "dirty-copied-snapshot",
        "description": (
            "candidate bytes in an isolated copied workspace; git HEAD values are "
            "supplemental provenance and do not identify the source"
        ),
        "workspace": str(workspace),
        "workspace_head": git_head(workspace),
        "canonical_head": git_head(CANONICAL_ROOT),
        "source_closure_file_count": len(snapshot["workspace_source"]),
        "workspace_source_closure_sha256": mapping_digest(snapshot["workspace_source"]),
        "canonical_source_closure_sha256": mapping_digest(snapshot["canonical_source"]),
        "harness_inputs_sha256": mapping_digest(snapshot["workspace_harness"]),
        "gate_hash_function": str(GATES_RUNNER),
        "gate_runner_sha256": snapshot["gate_runner_sha256"],
    }


def build_environment(target: Path, temp: Path) -> tuple[dict[str, str], dict[str, Any]]:
    env = os.environ.copy()
    overrides = {
        "CARGO_TARGET_DIR": str(target),
        "TMPDIR": str(temp),
        "CARGO_INCREMENTAL": "0",
    }
    env.update(overrides)
    inherited_flags: dict[str, str] = {}
    for key in (
        "RUSTFLAGS",
        "CARGO_ENCODED_RUSTFLAGS",
        "CARGO_PROFILE_RELEASE_DEBUG",
        "CARGO_PROFILE_RELEASE_LTO",
        "CARGO_PROFILE_RELEASE_CODEGEN_UNITS",
        "CARGO_PROFILE_RELEASE_PANIC",
        "CARGO_BUILD_JOBS",
    ):
        if key in os.environ:
            inherited_flags[key] = os.environ[key]
    return env, {
        "overrides": overrides,
        "inherited_build_flags": inherited_flags,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--workspace",
        type=Path,
        required=True,
        help="isolated copied candidate workspace (must differ from this checkout)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="new output root; an existing path is rejected",
    )
    parser.add_argument(
        "--baseline",
        type=Path,
        default=BASELINE_DEFAULT,
        help="retained scalar baseline raw.csv",
    )
    parser.add_argument(
        "--baseline-binary",
        type=Path,
        default=BASELINE_BINARY_DEFAULT,
        help="retained scalar baseline ELF (outside the evidence directory)",
    )
    parser.add_argument(
        "--target",
        type=Path,
        default=TARGET_DEFAULT,
        help="existing on-disk Cargo target; never cleaned by this tool",
    )
    parser.add_argument(
        "--tmpdir",
        type=Path,
        default=TMP_DEFAULT,
        help="existing on-disk temporary directory used by Cargo and children",
    )
    parser.add_argument("--warmups", type=int, default=3)
    parser.add_argument("--iterations", type=int, default=15)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    workspace = args.workspace.resolve(strict=True)
    baseline = args.baseline.resolve(strict=True)
    baseline_binary = args.baseline_binary.resolve()
    target = args.target.resolve()
    temp = args.tmpdir.resolve()
    output = args.output.resolve()

    if workspace == CANONICAL_ROOT:
        raise SystemExit("--workspace must be an isolated path, not the canonical checkout")
    if not (workspace / "crates/litchi-ods/Cargo.toml").is_file():
        raise SystemExit("--workspace does not contain crates/litchi-ods/Cargo.toml")
    if not baseline_binary.is_file():
        raise SystemExit(
            f"retained scalar baseline ELF is missing: {baseline_binary}"
        )
    if args.warmups < 0 or args.iterations <= 0:
        raise SystemExit("warmups must be nonnegative and iterations must be positive")
    if not target.is_dir():
        raise SystemExit(f"existing Cargo target directory is missing: {target}")
    if not temp.is_dir():
        raise SystemExit(f"existing temporary directory is missing: {temp}")
    ensure_baseline_binary_is_distinct(baseline_binary, target)
    ensure_fresh_output(output, baseline, baseline_binary, target, temp)

    started_at = utc_now()
    receipt: dict[str, Any] = {
        "schema": 1,
        "status": "running",
        "automatic_acceptance": False,
        "started_at": started_at,
        "script_sha256": sha256_file(Path(__file__)),
        "canonical_root": str(CANONICAL_ROOT),
        "workspace": str(workspace),
        "output": str(output),
        "baseline": str(baseline),
        "baseline_binary": str(baseline_binary),
        "target": str(target),
        "tmpdir": str(temp),
        "profile_cpu": PROFILE_CPU,
        "warmups": args.warmups,
        "iterations": args.iterations,
        "diagnostic_scope": (
            "one serial candidate capture against a retained scalar baseline; "
            "performance acceptance requires repeated paired AB captures"
        ),
        "commands": [],
        "captures": [],
    }

    def save() -> None:
        receipt["finished_at"] = utc_now()
        write_json(output / "run.json", receipt)

    try:
        if not baseline.is_file():
            raise RuntimeError(f"retained baseline raw.csv is missing: {baseline}")
        if not GATES_RUNNER.is_file() or not COMPARE.is_file():
            raise RuntimeError("gates/run.py or compare.py is missing")

        gate_hashes = load_gate_hashes()
        before = input_snapshot(gate_hashes, workspace, baseline)
        if before["workspace_source"] != before["canonical_source"]:
            raise RuntimeError(
                "isolated workspace source closure differs from canonical source closure"
            )
        receipt["source_identity"] = source_identity(workspace, before)
        receipt["source_closure"] = {
            "hash_function": str(GATES_RUNNER),
            "file_count": len(before["canonical_source"]),
            "workspace_sha256": mapping_digest(before["workspace_source"]),
            "canonical_sha256": mapping_digest(before["canonical_source"]),
            "matches_canonical_before": True,
        }
        receipt["input_hashes"] = {
            "before_file": "input-hashes-before.json",
            "after_file": "input-hashes-after.json",
        }
        write_json(
            output / "source-closure-before.json",
            {
                "hash_function": str(GATES_RUNNER),
                "file_count": len(before["canonical_source"]),
                "workspace": before["workspace_source"],
                "canonical": before["canonical_source"],
            },
        )
        write_json(
            output / "input-hashes-before.json",
            {
                "workspace_harness": before["workspace_harness"],
                "canonical_harness": before["canonical_harness"],
                "canonical_tools": before["canonical_tools"],
                "gate_runner_sha256": before["gate_runner_sha256"],
                "baseline_raw_sha256": before["baseline_raw_sha256"],
            },
        )
        write_json(output / "source-identity.json", receipt["source_identity"])

        env, env_metadata = build_environment(target, temp)
        receipt["build_environment"] = env_metadata
        receipt["toolchain"] = {
            "rustc": version_command(["rustc", "-Vv"], workspace),
            "cargo": version_command(["cargo", "-V"], workspace),
        }
        receipt["toolchain_ok"] = all(
            item["status"] == 0 for item in receipt["toolchain"].values()
        )
        receipt["host"] = host_metadata()
        receipt["retained_baseline_before"] = retained_baseline_state(
            baseline, baseline_binary
        )

        binaries = {
            "scalar": target / "release" / SCALAR_BINARY_NAME,
            "value": target / "release" / VALUE_BINARY_NAME,
        }
        receipt["candidate_binaries_before_build"] = binary_state(binaries)
        build_specs = (
            ("build-scalar", SCALAR_MANIFEST, binaries["scalar"]),
            ("build-value", VALUE_MANIFEST, binaries["value"]),
        )
        built_ok: dict[str, bool] = {}
        for name, manifest, binary in build_specs:
            command = [
                "cargo",
                "build",
                "--locked",
                "--offline",
                "--release",
                "--manifest-path",
                str(workspace / manifest),
            ]
            binary_before = binary_state({name: binary})[name]
            build = run_logged(name, command, workspace, env, output / f"{name}.log")
            build["binary_before"] = binary_before
            build["binary_after"] = binary_state({name: binary})[name]
            receipt["commands"].append(build)
            if build["status"] != 0:
                built_ok["scalar" if name == "build-scalar" else "value"] = False
                continue
            if not build["binary_after"]["exists"]:
                build["status"] = 127
                build["error"] = f"release executable was not produced: {binary}"
            built_ok["scalar" if name == "build-scalar" else "value"] = (
                build["status"] == 0 and build["binary_after"]["exists"]
            )
        receipt["candidate_binaries_after_build"] = binary_state(binaries)
        receipt["build_flags"] = {
            "profile": "release",
            "locked": True,
            "offline": True,
            "environment": env_metadata,
            "target": str(target),
        }

        source_identity_arg = (
            "dirty-copied-snapshot:" + receipt["source_identity"]["workspace_source_closure_sha256"]
        )
        capture_specs = (
            (
                "scalar",
                workspace / SCALAR_RUNNER,
                binaries["scalar"],
                282,
                ["--revision", "candidate", "--group", "all", "--phase", "all"],
            ),
            (
                "value",
                workspace / VALUE_RUNNER,
                binaries["value"],
                396,
                [
                    "--revision",
                    "candidate",
                    "--adapter",
                    "value",
                    "--group",
                    "all",
                    "--phase",
                    "all",
                    "--source-identity",
                    source_identity_arg,
                    "--source-manifest",
                    str(output / "source-identity.json"),
                ],
            ),
            (
                "worksheet",
                workspace / VALUE_RUNNER,
                binaries["value"],
                88,
                [
                    "--revision",
                    "candidate",
                    "--adapter",
                    "worksheet",
                    "--group",
                    "all",
                    "--phase",
                    "all",
                    "--source-identity",
                    source_identity_arg,
                    "--source-manifest",
                    str(output / "source-identity.json"),
                ],
            ),
            (
                "worksheet-instrumented",
                workspace / VALUE_RUNNER,
                binaries["value"],
                88,
                [
                    "--revision",
                    "candidate",
                    "--adapter",
                    "worksheet",
                    "--instrumented",
                    "--group",
                    "all",
                    "--phase",
                    "all",
                    "--source-identity",
                    source_identity_arg,
                    "--source-manifest",
                    str(output / "source-identity.json"),
                ],
            ),
        )
        capture_dirs = {
            "scalar": output / "scalar",
            "value": output / "value",
            "worksheet": output / "worksheet",
            "worksheet-instrumented": output / "worksheet-instrumented",
        }
        capture_statuses: list[bool] = []
        for name, runner, binary, expected_rows, options in capture_specs:
            capture_dir = capture_dirs[name]
            binary_key = "scalar" if name == "scalar" else "value"
            capture_dir.mkdir()
            command = [
                sys.executable,
                "-B",
                str(runner),
                "--binary",
                str(binary),
                "--output",
                str(capture_dir),
                *options,
                "--warmups",
                str(args.warmups),
                "--iterations",
                str(args.iterations),
            ]
            if not built_ok.get(binary_key, False) or not binary.is_file():
                capture = {
                    "name": name,
                    "command": command,
                    "status": 127,
                    "skipped": True,
                    "error": f"binary is missing: {binary}",
                }
                (output / f"capture-{name}.log").write_text(
                    capture["error"] + "\n", encoding="utf-8"
                )
            else:
                capture = run_logged(
                    f"capture-{name}",
                    command,
                    workspace,
                    env,
                    output / f"capture-{name}.log",
                )
                capture["skipped"] = False
            capture["expected_rows"] = expected_rows
            capture["validation"] = validate_capture(capture_dir, expected_rows)
            capture["binary_after"] = binary_state({"binary": binary})["binary"]
            capture_statuses.append(
                capture["status"] == 0
                and capture["validation"]["row_count_ok"]
                and capture["validation"]["all_status_zero"]
                and not capture["validation"]["duplicate_phase_case_rows"]
                and not capture["validation"]["malformed_rows"]
                and capture["validation"]["error"] is None
                and capture["binary_after"]
                == receipt["candidate_binaries_after_build"][binary_key]
            )
            receipt["captures"].append(capture)

        scalar_raw = capture_dirs["scalar"] / "raw.csv"
        compare_path = output / "compare.json"
        compare_command = [
            sys.executable,
            "-B",
            str(COMPARE),
            str(baseline),
            str(scalar_raw),
            str(compare_path),
        ]
        comparison = run_logged(
            "compare",
            compare_command,
            CANONICAL_ROOT,
            env,
            output / "compare.log",
        )
        comparison["report"] = str(compare_path)
        receipt["commands"].append(comparison)
        receipt["review_triggers"] = extract_review_triggers(compare_path, output)

        after = input_snapshot(gate_hashes, workspace, baseline)
        write_json(
            output / "source-closure-after.json",
            {
                "hash_function": str(GATES_RUNNER),
                "file_count": len(after["canonical_source"]),
                "workspace": after["workspace_source"],
                "canonical": after["canonical_source"],
            },
        )
        write_json(
            output / "input-hashes-after.json",
            {
                "workspace_harness": after["workspace_harness"],
                "canonical_harness": after["canonical_harness"],
                "canonical_tools": after["canonical_tools"],
                "gate_runner_sha256": after["gate_runner_sha256"],
                "baseline_raw_sha256": after["baseline_raw_sha256"],
            },
        )
        source_unchanged = (
            before["workspace_source"] == after["workspace_source"]
            and before["canonical_source"] == after["canonical_source"]
            and after["workspace_source"] == after["canonical_source"]
        )
        harness_unchanged = (
            before["workspace_harness"] == after["workspace_harness"]
            and before["canonical_harness"] == after["canonical_harness"]
            and after["workspace_harness"] == after["canonical_harness"]
        )
        tools_unchanged = before["canonical_tools"] == after["canonical_tools"]
        baseline_unchanged = before["baseline_raw_sha256"] == after["baseline_raw_sha256"]
        retained_baseline_after = retained_baseline_state(baseline, baseline_binary)
        retained_baseline_unchanged = (
            receipt["retained_baseline_before"] == retained_baseline_after
        )
        binary_after_capture = binary_state(binaries)
        binary_unchanged = (
            binary_after_capture == receipt["candidate_binaries_after_build"]
        )
        receipt["source_closure"]["matches_canonical_after"] = (
            after["workspace_source"] == after["canonical_source"]
        )
        receipt["source_unchanged"] = source_unchanged
        receipt["harness_unchanged"] = harness_unchanged
        receipt["tools_unchanged"] = tools_unchanged
        receipt["baseline_raw_unchanged"] = baseline_unchanged
        receipt["retained_baseline_unchanged"] = retained_baseline_unchanged
        receipt["candidate_binaries_after_capture"] = binary_after_capture
        receipt["candidate_binaries_unchanged"] = binary_unchanged
        receipt["retained_baseline_after"] = retained_baseline_after
        receipt["capture_statuses"] = capture_statuses
        receipt["status"] = (
            "complete"
            if (
                source_unchanged
                and harness_unchanged
                and tools_unchanged
                and baseline_unchanged
                and retained_baseline_unchanged
                and binary_unchanged
                and receipt["toolchain_ok"]
                and receipt["review_triggers"]["error"] is None
                and all(command["status"] == 0 for command in receipt["commands"])
                and all(capture_statuses)
            )
            else "failed"
        )
        save()
        return 0 if receipt["status"] == "complete" else 1
    except Exception as error:  # retain a reviewable receipt for preflight failures
        receipt["status"] = "failed-preflight"
        receipt["error"] = f"{type(error).__name__}: {error}"
        save()
        print(receipt["error"], file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
