#!/usr/bin/env python3
"""Finalize an already-completed 0894 probe corpus.

The original corpus run completed both probes and both census passes, but its
host comparator raised while normalizing the scalar ``summary.legs`` field.
This continuation verifies those existing artifacts, runs only the corrected
host comparator, and writes a canonical run receipt that retains the failed
comparison receipt and logs as recovery evidence.  It never invokes either
probe or Cargo.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import stat
import subprocess
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
DEFAULT_CORPUS = HERE / "corpus-v1"
DEFAULT_BASELINE_BUILD = ROOT / "target" / "litchi-0894-probe-baseline" / "build.json"
DEFAULT_CANDIDATE_BUILD = ROOT / "target" / "litchi-0894-probe-candidate" / "build.json"
INPUT_CENSUS = HERE / "probe" / "input_census.py"
ORIGINAL_COMPARATOR = HERE / "probe" / "compare.py"
CORRECTED_COMPARATOR = HERE / "compare-v2.py"
ORIGINAL_RUNNER = HERE / "probe" / "run.py"
TOOLCHAIN = "1.95.0"
PROBE_SCHEMA = "litchi.xlsx-observable-differential-0894.v1"
CENSUS_SCHEMA = "litchi.xlsx-input-census-0894.v1"
COMMAND_SCHEMA = "litchi.xlsx-observable-probe-command-receipt-0894.v1"
RUN_SCHEMA = "litchi.xlsx-observable-probe-run-0894.v1"
EXPECTED = {
    "inputs": 180,
    "worksheets": 394,
    "legs": 1576,
    "discovery_errors": 0,
    "source_open_errors": 0,
    "operation_errors": 12,
}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"expected regular non-symlink file: {path}")
    path = path.resolve()
    before = path.stat()
    digest = sha(path)
    after = path.stat()
    require(
        before.st_size == after.st_size and before.st_mtime_ns == after.st_mtime_ns,
        f"artifact changed while hashing: {path}",
    )
    mode = stat.S_IMODE(after.st_mode)
    return {
        "path": str(path),
        "bytes": after.st_size,
        "sha256": digest,
        "mode": mode,
        "executable": bool(mode & stat.S_IXUSR),
    }


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite receipt: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def read_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as stream:
        return json.load(stream)


def expect_descriptor(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} descriptor is missing")
    path_value = value.get("path")
    require(isinstance(path_value, str) and path_value, f"{label} descriptor path is missing")
    current = descriptor(Path(path_value))
    require(current == value, f"{label} descriptor changed")
    return current


def load_build(path: Path, role: str) -> tuple[dict[str, Any], dict[str, Any]]:
    value = read_json(path)
    require(value.get("schema") == "litchi.xlsx-observable-probe-build-0894.v1", f"{role} build schema mismatch")
    require(value.get("status") == "pass" and value.get("role") == role, f"{role} build is not passing")
    require(value.get("toolchain") == TOOLCHAIN and value.get("profile") == "dev", f"{role} build configuration mismatch")
    require(value.get("debug") == 0 and value.get("incremental") is False, f"{role} build profile is not frozen diagnostic dev")
    binary = expect_descriptor(value.get("binary"), f"{role} binary")
    lock = expect_descriptor(value.get("lock"), f"{role} Cargo.lock")
    require(Path(binary["path"]).resolve().is_file(), f"{role} binary disappeared")
    require(Path(lock["path"]).resolve().is_file(), f"{role} lock disappeared")
    return value, {"receipt": descriptor(path), "binary": binary, "lock": lock}


def load_probe_jsonl(path: Path, role: str) -> tuple[dict[str, Any], dict[str, Any]]:
    inputs: dict[str, dict[str, Any]] = {}
    summary: dict[str, Any] | None = None
    with path.open(encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, 1):
            require(line.strip(), f"{path}:{line_number} is blank")
            try:
                value = json.loads(line)
            except json.JSONDecodeError as error:
                raise ValueError(f"{path}:{line_number} is not complete JSON: {error}") from error
            require(isinstance(value, dict), f"{path}:{line_number} is not an object")
            require(value.get("schema") == PROBE_SCHEMA, f"{path}:{line_number} schema mismatch")
            kind = value.get("kind")
            if kind == "input":
                input_value = value.get("input")
                require(isinstance(input_value, dict), f"{path}:{line_number} input metadata missing")
                name = input_value.get("path")
                require(isinstance(name, str) and name, f"{path}:{line_number} input path missing")
                require(name not in inputs, f"{path}:{line_number} duplicate input {name}")
                require(isinstance(value.get("sheets"), list), f"{path}:{line_number} sheets missing")
                require(isinstance(value.get("legs"), list), f"{path}:{line_number} legs missing")
                inputs[name] = value
            elif kind == "summary":
                require(summary is None, f"{path}:{line_number} duplicate summary")
                summary = value
            else:
                raise ValueError(f"{path}:{line_number} unexpected record kind {kind!r}")
    require(summary is not None, f"{path} has no summary")
    worksheets = sum(len(value["sheets"]) for value in inputs.values())
    legs = sum(len(value["legs"]) for value in inputs.values())
    source_open_errors = sum(value.get("status") == "source_open_error" for value in inputs.values())
    operation_errors = sum(
        leg.get("operation_error") is not None
        for value in inputs.values()
        for leg in value["legs"]
    )
    stats = {
        "inputs": len(inputs),
        "worksheets": worksheets,
        "legs": legs,
        "discovery_errors": summary.get("discovery_errors"),
        "source_open_errors": source_open_errors,
        "operation_errors": operation_errors,
        "summary": {
            key: summary.get(key)
            for key in ("inputs", "worksheets", "legs", "discovery_errors", "source_open_errors", "operation_errors")
        },
        "input_paths": sorted(inputs),
    }
    for key in ("inputs", "worksheets", "legs", "discovery_errors", "source_open_errors", "operation_errors"):
        require(summary.get(key) == stats[key], f"{role} summary {key} disagrees with records")
    require({key: stats[key] for key in EXPECTED} == EXPECTED, f"{role} counts do not match expected corpus: {stats}")
    return stats, {"descriptor": descriptor(path), "summary": summary}


def load_census(path: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    inputs: list[str] = []
    discovery_errors = 0
    summary: dict[str, Any] | None = None
    with path.open(encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, 1):
            require(line.strip(), f"{path}:{line_number} is blank")
            try:
                value = json.loads(line)
            except json.JSONDecodeError as error:
                raise ValueError(f"{path}:{line_number} is not complete JSON: {error}") from error
            require(isinstance(value, dict) and value.get("schema") == CENSUS_SCHEMA, f"{path}:{line_number} census schema mismatch")
            if value.get("kind") == "input":
                input_value = value.get("input")
                require(isinstance(input_value, dict) and isinstance(input_value.get("path"), str), f"{path}:{line_number} census input missing")
                inputs.append(input_value["path"])
            elif value.get("kind") == "discovery_error":
                discovery_errors += 1
            elif value.get("kind") == "summary":
                require(summary is None, f"{path}:{line_number} duplicate census summary")
                summary = value
            else:
                raise ValueError(f"{path}:{line_number} unexpected census kind")
    require(summary is not None, f"{path} has no census summary")
    require(len(inputs) == len(set(inputs)), f"{path} has duplicate census inputs")
    require(summary.get("inputs") == len(inputs), f"{path} census input count disagrees")
    require(summary.get("discovery_errors") == discovery_errors, f"{path} census discovery count disagrees")
    stats = {
        "inputs": len(inputs),
        "discovery_errors": discovery_errors,
        "input_paths": sorted(inputs),
    }
    require(stats["inputs"] == EXPECTED["inputs"] and discovery_errors == 0, f"{path} census counts are not frozen: {stats}")
    return stats, {"descriptor": descriptor(path), "summary": summary}


def load_command(
    path: Path,
    label: str,
    expected_output: Path,
    expected_binary: Path | None = None,
    output_flag: str = "--jsonl",
) -> tuple[dict[str, Any], dict[str, Any]]:
    value = read_json(path)
    require(value.get("schema") == COMMAND_SCHEMA and value.get("label") == label, f"{label} command receipt schema/label mismatch")
    require(value.get("exit_code") == 0 and value.get("status") == "pass", f"{label} command did not pass")
    expect_descriptor(value.get("stdout"), f"{label} stdout")
    expect_descriptor(value.get("stderr"), f"{label} stderr")
    argv = value.get("argv")
    require(isinstance(argv, list) and all(isinstance(arg, str) for arg in argv), f"{label} argv missing")
    require(output_flag in argv and Path(argv[argv.index(output_flag) + 1]).resolve() == expected_output.resolve(), f"{label} output binding mismatch")
    if expected_binary is not None:
        require(Path(argv[0]).resolve() == expected_binary.resolve(), f"{label} binary binding mismatch")
        require("manifest" in argv, f"{label} is not a manifest probe")
    return value, {"receipt": descriptor(path), "stdout": value["stdout"], "stderr": value["stderr"]}


def load_failed_compare(path: Path, expected_output: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    value = read_json(path)
    require(value.get("schema") == COMMAND_SCHEMA and value.get("label") == "compare", "original compare receipt schema/label mismatch")
    require(value.get("exit_code") == 1, "original compare receipt is not the recorded comparator failure")
    expect_descriptor(value.get("stdout"), "original compare stdout")
    expect_descriptor(value.get("stderr"), "original compare stderr")
    argv = value.get("argv")
    require(isinstance(argv, list) and str(ORIGINAL_COMPARATOR) in argv, "original compare command binding mismatch")
    require("--output" in argv and Path(argv[argv.index("--output") + 1]).resolve() == expected_output.resolve(), "original compare output binding mismatch")
    stderr_path = Path(value["stderr"]["path"])
    stderr = stderr_path.read_text(encoding="utf-8")
    require("TypeError: 'int' object is not iterable" in stderr, "original compare failure is not the known summary normalization bug")
    return value, {"receipt": descriptor(path), "stdout": value["stdout"], "stderr": value["stderr"]}


def environment() -> tuple[dict[str, str], dict[str, str]]:
    values = {
        key: os.environ[key]
        for key in ("PATH", "HOME", "RUSTUP_HOME", "CARGO_HOME", "USER", "LOGNAME", "TERM")
        if key in os.environ
    }
    values.update({"RUSTUP_TOOLCHAIN": TOOLCHAIN, "LC_ALL": "C", "LANG": "C", "TZ": "UTC", "PYTHONHASHSEED": "0"})
    recorded = {key: values[key] for key in ("RUSTUP_TOOLCHAIN", "LC_ALL", "LANG", "TZ", "PYTHONHASHSEED")}
    for key in ("RUSTFLAGS", "RUSTC_WRAPPER", "CARGO_BUILD_RUSTC_WRAPPER", "CARGO_ENCODED_RUSTFLAGS", "CARGO_PROFILE_RELEASE_LTO", "CARGO_PROFILE_DEV_LTO"):
        values.pop(key, None)
    return values, recorded


def run_corrected_compare(corpus: Path, baseline: Path, candidate: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    output = corpus / "comparison-v2.json"
    stdout = corpus / "compare-v2.stdout.txt"
    stderr = corpus / "compare-v2.stderr.txt"
    receipt_path = corpus / "compare-v2.receipt.json"
    require(not output.exists() and not stdout.exists() and not stderr.exists() and not receipt_path.exists(), "corrected comparator artifacts already exist")
    command = [
        sys.executable,
        "-B",
        str(CORRECTED_COMPARATOR),
        "--baseline",
        str(baseline.resolve()),
        "--candidate",
        str(candidate.resolve()),
        "--output",
        str(output.resolve()),
    ]
    env, recorded_env = environment()
    started = utc_now()
    with stdout.open("x", encoding="utf-8") as out, stderr.open("x", encoding="utf-8") as err:
        result = subprocess.run(command, cwd=HERE, env=env, stdout=out, stderr=err, text=True, check=False)
    finished = utc_now()
    receipt = {
        "schema": COMMAND_SCHEMA,
        "label": "compare-v2",
        "argv": command,
        "cwd": str(HERE),
        "environment": recorded_env,
        "started_utc": started,
        "finished_utc": finished,
        "exit_code": result.returncode,
        "stdout": descriptor(stdout),
        "stderr": descriptor(stderr),
        "status": "pass" if result.returncode == 0 else "mismatch",
    }
    write_once(receipt_path, receipt)
    value = read_json(output) if output.exists() else None
    return receipt, {
        "receipt": descriptor(receipt_path),
        "stdout": receipt["stdout"],
        "stderr": receipt["stderr"],
        "output": descriptor(output) if output.exists() else None,
        "value": value,
    }


def args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--corpus-root", type=Path, default=DEFAULT_CORPUS)
    parser.add_argument("--baseline-build", type=Path, default=DEFAULT_BASELINE_BUILD)
    parser.add_argument("--candidate-build", type=Path, default=DEFAULT_CANDIDATE_BUILD)
    return parser.parse_args()


def main() -> int:
    options = args()
    corpus = options.corpus_root.resolve()
    require(corpus.is_dir() and not corpus.is_symlink(), f"corpus root is missing: {corpus}")
    baseline_jsonl = corpus / "baseline.jsonl"
    candidate_jsonl = corpus / "candidate.jsonl"
    census_before_path = corpus / "census-before.jsonl"
    census_after_path = corpus / "census-after.jsonl"
    baseline_build, baseline_build_info = load_build(options.baseline_build.resolve(), "baseline")
    candidate_build, candidate_build_info = load_build(options.candidate_build.resolve(), "candidate")
    require(baseline_build_info["lock"]["sha256"] == candidate_build_info["lock"]["sha256"], "baseline/candidate probe lock differs")
    baseline_stats, baseline_jsonl_info = load_probe_jsonl(baseline_jsonl, "baseline")
    candidate_stats, candidate_jsonl_info = load_probe_jsonl(candidate_jsonl, "candidate")
    require(baseline_stats == candidate_stats, "baseline/candidate corpus record counts or paths differ")
    census_before, census_before_info = load_census(census_before_path)
    census_after, census_after_info = load_census(census_after_path)
    require(all(census_before_info["descriptor"][key] == census_after_info["descriptor"][key] for key in ("bytes", "sha256", "mode")), "census changed between before and after receipts")
    require(census_before["input_paths"] == baseline_stats["input_paths"], "census paths do not match probe manifest paths")
    require(census_after["input_paths"] == baseline_stats["input_paths"], "after-census paths do not match probe manifest paths")

    baseline_probe, baseline_probe_info = load_command(
        corpus / "baseline-probe.receipt.json",
        "baseline-probe",
        baseline_jsonl,
        Path(baseline_build_info["binary"]["path"]),
    )
    candidate_probe, candidate_probe_info = load_command(
        corpus / "candidate-probe.receipt.json",
        "candidate-probe",
        candidate_jsonl,
        Path(candidate_build_info["binary"]["path"]),
    )
    census_before_command, census_before_command_info = load_command(
        corpus / "census-before.receipt.json",
        "census-before",
        census_before_path,
        output_flag="--output",
    )
    census_after_command, census_after_command_info = load_command(
        corpus / "census-after.receipt.json",
        "census-after",
        census_after_path,
        output_flag="--output",
    )
    original_compare_output = corpus / "comparison.json"
    original_compare, original_compare_info = load_failed_compare(corpus / "compare.receipt.json", original_compare_output)
    require(not original_compare_output.exists(), "unexpected original comparison output appeared")

    corrected_receipt, corrected_info = run_corrected_compare(corpus, baseline_jsonl, candidate_jsonl)
    comparison_value = corrected_info["value"]
    comparison_status = comparison_value.get("status") if isinstance(comparison_value, dict) else None
    comparison_pass = corrected_receipt["exit_code"] == 0 and comparison_status == "pass"
    census_equal = all(census_before_info["descriptor"][key] == census_after_info["descriptor"][key] for key in ("bytes", "sha256", "mode"))
    status = "pass" if census_equal and comparison_pass else "mismatch"
    marker = corpus / ".owner-litchi-0894-probe-run"
    run_receipt = {
        "schema": RUN_SCHEMA,
        "status": status,
        "started_utc": census_before_command.get("started_utc", baseline_probe.get("started_utc")),
        "finished_utc": corrected_receipt["finished_utc"],
        "marker": descriptor(marker),
        "baseline_build": baseline_build_info["receipt"],
        "candidate_build": candidate_build_info["receipt"],
        "input_roots": [str((ROOT / "test-data").resolve())],
        "helpers": {
            "input_census.py": descriptor(INPUT_CENSUS),
            "compare.py": descriptor(ORIGINAL_COMPARATOR),
            "compare-v2.py": descriptor(CORRECTED_COMPARATOR),
            "run.py": descriptor(ORIGINAL_RUNNER),
            "corpus_finalize.py": descriptor(Path(__file__).resolve()),
        },
        "census_before": census_before_info["descriptor"],
        "census_after": census_after_info["descriptor"],
        "census_equal": census_equal,
        "baseline_jsonl": baseline_jsonl_info["descriptor"],
        "candidate_jsonl": candidate_jsonl_info["descriptor"],
        "comparison": corrected_info["output"],
        "comparison_status": comparison_status,
        "commands": {
            "census_before": census_before_command_info["receipt"],
            "baseline": baseline_probe_info["receipt"],
            "candidate": candidate_probe_info["receipt"],
            "census_after": census_after_command_info["receipt"],
            "compare_original_failed": original_compare_info["receipt"],
            "compare_v2": corrected_info["receipt"],
        },
        "environment": corrected_receipt["environment"],
        "integrity": {
            "expected_counts": EXPECTED,
            "baseline_counts": baseline_stats,
            "candidate_counts": candidate_stats,
            "counts_equal": baseline_stats == candidate_stats,
            "census_before": census_before,
            "census_after": census_after,
        },
        "recovery": {
            "kind": "known_original_comparator_summary_scalar_bug",
            "reason": "The original compare.py applied input normalization to the summary record, whose legs field is an integer; both probes and both census passes completed successfully before that host-only failure.",
            "original_compare_receipt": original_compare_info["receipt"],
            "original_compare_stdout": original_compare_info["stdout"],
            "original_compare_stderr": original_compare_info["stderr"],
            "corrected_comparator": corrected_info["receipt"],
            "original_run_json_present": (corpus / "run.json").exists(),
        },
        "performance_claim": None,
    }
    write_once(corpus / "run.json", run_receipt)
    require(status == "pass", "corrected observable comparison did not pass")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
