#!/usr/bin/env python3
"""Run both built 0894 probes over one hash-bound input manifest.

The two binaries receive exactly the same absolute input roots in the same
order.  A census is taken before and after both runs.  Every subprocess has
create-only command, stdout, stderr, and receipt files.  The driver does not
build, retry, skip, or mutate an input; comparison status is retained even
when the comparator reports a baseline/candidate mismatch.
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


PROBE = Path(__file__).resolve().parent
ROOT = PROBE.parents[4]
INPUT_CENSUS = PROBE / "input_census.py"
COMPARATOR = PROBE / "compare.py"
TOOLCHAIN = "1.95.0"


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
    mode = stat.S_IMODE(path.stat().st_mode)
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
        "mode": mode,
        "executable": bool(mode & stat.S_IXUSR),
    }


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def read_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as stream:
        return json.load(stream)


def load_build(path: Path, role: str) -> dict[str, Any]:
    value = read_json(path)
    require(value.get("schema") == "litchi.xlsx-observable-probe-build-0894.v1", f"{role} build schema mismatch")
    require(value.get("status") == "pass" and value.get("role") == role, f"{role} build is not passing")
    require(value.get("toolchain") == TOOLCHAIN and value.get("profile") == "dev", f"{role} build configuration mismatch")
    binary = value.get("binary")
    require(isinstance(binary, dict), f"{role} binary descriptor missing")
    binary_path = Path(str(binary.get("path", "")))
    require(binary_path.is_file() and not binary_path.is_symlink(), f"{role} probe binary missing")
    require(descriptor(binary_path) == binary, f"{role} probe binary descriptor changed")
    return value


def environment(run_root: Path) -> tuple[dict[str, str], dict[str, str]]:
    tmp = run_root / "tmp"
    tmp.mkdir(parents=True, exist_ok=True)
    inherited = {
        key: os.environ[key]
        for key in ("PATH", "HOME", "RUSTUP_HOME", "CARGO_HOME", "USER", "LOGNAME", "TERM")
        if key in os.environ
    }
    values = inherited
    values.update(
        {
            "RUSTUP_TOOLCHAIN": TOOLCHAIN,
            "LC_ALL": "C",
            "LANG": "C",
            "TZ": "UTC",
            "PYTHONHASHSEED": "0",
            "TMPDIR": str(tmp.resolve()),
        }
    )
    for key in (
        "RUSTFLAGS",
        "RUSTC_WRAPPER",
        "CARGO_BUILD_RUSTC_WRAPPER",
        "CARGO_ENCODED_RUSTFLAGS",
        "CARGO_PROFILE_RELEASE_LTO",
        "CARGO_PROFILE_DEV_LTO",
    ):
        values.pop(key, None)
    recorded = {key: values[key] for key in ("RUSTUP_TOOLCHAIN", "LC_ALL", "LANG", "TZ", "PYTHONHASHSEED", "TMPDIR")}
    return values, recorded


def invoke(
    label: str,
    argv: list[str],
    cwd: Path,
    env: dict[str, str],
    recorded_env: dict[str, str],
    root: Path,
    allow_exit: tuple[int, ...] = (0,),
) -> dict[str, Any]:
    stdout = root / f"{label}.stdout.txt"
    stderr = root / f"{label}.stderr.txt"
    receipt_path = root / f"{label}.receipt.json"
    require(not stdout.exists() and not stderr.exists() and not receipt_path.exists(), f"owned run artifact already exists for {label}")
    started = utc_now()
    with stdout.open("x", encoding="utf-8") as out, stderr.open("x", encoding="utf-8") as err:
        result = subprocess.run(argv, cwd=cwd, env=env, stdout=out, stderr=err, text=True, check=False)
    finished = utc_now()
    receipt = {
        "schema": "litchi.xlsx-observable-probe-command-receipt-0894.v1",
        "label": label,
        "argv": argv,
        "cwd": str(cwd),
        "environment": recorded_env,
        "started_utc": started,
        "finished_utc": finished,
        "exit_code": result.returncode,
        "stdout": descriptor(stdout),
        "stderr": descriptor(stderr),
        "status": "pass" if result.returncode in allow_exit else "fail",
    }
    write_once(receipt_path, receipt)
    require(result.returncode in allow_exit, f"{label} failed ({result.returncode}); see {stderr}")
    return {"receipt": descriptor(receipt_path), "stdout": descriptor(stdout), "stderr": descriptor(stderr), "value": receipt}


def args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-build", type=Path, required=True)
    parser.add_argument("--candidate-build", type=Path, required=True)
    parser.add_argument("--input-root", type=Path, action="append", required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    return parser.parse_args()


def main() -> int:
    options = args()
    baseline_path = options.baseline_build.resolve()
    candidate_path = options.candidate_build.resolve()
    baseline = load_build(baseline_path, "baseline")
    candidate = load_build(candidate_path, "candidate")
    roots = [path.resolve() for path in options.input_root]
    require(all(path.exists() for path in roots), "one or more input roots are missing")
    run_root = options.output_root.resolve()
    require(not run_root.exists(), f"owned probe run root already exists: {run_root}")
    run_root.mkdir(parents=True)
    marker = run_root / ".owner-litchi-0894-probe-run"
    marker.write_text("litchi 0894 observable probe run\n", encoding="utf-8")
    env, recorded_env = environment(run_root)

    before_path = run_root / "census-before.jsonl"
    before = invoke(
        "census-before",
        [sys.executable, "-B", str(INPUT_CENSUS), *map(str, roots), "--output", str(before_path)],
        PROBE,
        env,
        recorded_env,
        run_root,
    )
    baseline_jsonl = run_root / "baseline.jsonl"
    baseline_run = invoke(
        "baseline-probe",
        [str(Path(baseline["binary"]["path"])), "manifest", *map(str, roots), "--jsonl", str(baseline_jsonl)],
        PROBE,
        env,
        recorded_env,
        run_root,
    )
    candidate_jsonl = run_root / "candidate.jsonl"
    candidate_run = invoke(
        "candidate-probe",
        [str(Path(candidate["binary"]["path"])), "manifest", *map(str, roots), "--jsonl", str(candidate_jsonl)],
        PROBE,
        env,
        recorded_env,
        run_root,
    )
    after_path = run_root / "census-after.jsonl"
    after = invoke(
        "census-after",
        [sys.executable, "-B", str(INPUT_CENSUS), *map(str, roots), "--output", str(after_path)],
        PROBE,
        env,
        recorded_env,
        run_root,
    )
    census_equal = sha(before_path) == sha(after_path)
    comparison_path = run_root / "comparison.json"
    comparison = invoke(
        "compare",
        [
            sys.executable,
            "-B",
            str(COMPARATOR),
            "--baseline",
            str(baseline_jsonl),
            "--candidate",
            str(candidate_jsonl),
            "--output",
            str(comparison_path),
        ],
        PROBE,
        env,
        recorded_env,
        run_root,
        allow_exit=(0, 1),
    )
    comparison_value = read_json(comparison_path)
    status = "pass" if census_equal and comparison_value.get("status") == "pass" else "mismatch"
    receipt = {
        "schema": "litchi.xlsx-observable-probe-run-0894.v1",
        "status": status,
        "started_utc": before["value"]["started_utc"],
        "finished_utc": comparison["value"]["finished_utc"],
        "marker": descriptor(marker),
        "baseline_build": descriptor(baseline_path),
        "candidate_build": descriptor(candidate_path),
        "input_roots": [str(path) for path in roots],
        "helpers": {
            "input_census.py": descriptor(INPUT_CENSUS),
            "compare.py": descriptor(COMPARATOR),
            "run.py": descriptor(Path(__file__).resolve()),
        },
        "census_before": descriptor(before_path),
        "census_after": descriptor(after_path),
        "census_equal": census_equal,
        "baseline_jsonl": descriptor(baseline_jsonl),
        "candidate_jsonl": descriptor(candidate_jsonl),
        "comparison": descriptor(comparison_path),
        "comparison_status": comparison_value.get("status"),
        "commands": {
            key: value["receipt"]
            for key, value in {
                "census_before": before,
                "baseline": baseline_run,
                "candidate": candidate_run,
                "census_after": after,
                "compare": comparison,
            }.items()
        },
        "environment": recorded_env,
        "performance_claim": None,
    }
    write_once(run_root / "run.json", receipt)
    require(status == "pass", "baseline/candidate observable probe comparison did not pass")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
