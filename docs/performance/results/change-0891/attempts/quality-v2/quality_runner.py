#!/usr/bin/env python3
"""Run the frozen 0891 XLSX quality gates once, with create-only custody."""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import time
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
TARGET = ROOT / "target/litchi-0891-quality"
TARGET_MARKER = TARGET / ".owner-0891-quality"
MARKER_TEXT = "litchi 0891 XLSX quality target\n"
MANIFEST = ROOT / "Cargo.toml"
LOCK = ROOT / "Cargo.lock"
CARGO = Path("/home/zhuhe/.rustup/toolchains/1.95.0-x86_64-unknown-linux-gnu/bin/cargo")
FREEZE = PACKET / "candidate-freeze.json"
ATTEMPT = PACKET / "attempts/quality-v1"
LOGS = ATTEMPT / "logs"
RECEIPT = ATTEMPT / "quality-receipt.json"
START = ATTEMPT / "quality-start.json"

FROZEN_PATHS = (
    "crates/litchi-xlsx/src/raw/worksheet/x14ac.rs",
    "crates/litchi-xlsx/src/raw/worksheet/x14ac/reuse_tests.rs",
)

ENV_KEYS = (
    "CARGO_BUILD_JOBS",
    "CARGO_INCREMENTAL",
    "CARGO_NET_OFFLINE",
    "CARGO_PROFILE_DEV_DEBUG",
    "CARGO_PROFILE_TEST_DEBUG",
    "CARGO_TARGET_DIR",
    "CARGO_HOME",
    "HOME",
    "LANG",
    "LC_ALL",
    "PATH",
    "RUSTDOCFLAGS",
    "RUSTFLAGS",
    "RUSTUP_HOME",
    "RUSTUP_TOOLCHAIN",
    "TMPDIR",
    "TZ",
)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def identity(path: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"expected regular file without symlink: {path}")
    metadata = path.stat()
    return {
        "path": str(path),
        "bytes": metadata.st_size,
        "mode": oct(stat.S_IMODE(metadata.st_mode)),
        "sha256": sha256(path),
    }


def write_exclusive(path: Path, value: dict[str, Any]) -> None:
    if path.exists() or path.is_symlink():
        raise RuntimeError(f"create-only receipt already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as destination:
        json.dump(value, destination, indent=2, sort_keys=True)
        destination.write("\n")
        destination.flush()
        os.fsync(destination.fileno())
    directory = os.open(path.parent, os.O_RDONLY | getattr(os, "O_DIRECTORY", 0))
    try:
        os.fsync(directory)
    finally:
        os.close(directory)


def read_json(path: Path) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"required JSON receipt is missing or symlinked: {path}")
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise RuntimeError(f"expected JSON object: {path}")
    return value


def freeze_files(freeze: dict[str, Any]) -> list[dict[str, Any]]:
    candidate_paths = freeze.get("candidate_paths")
    if candidate_paths is not None and candidate_paths != list(FROZEN_PATHS):
        raise RuntimeError("candidate freeze paths are not the two reviewed x14ac files in order")
    files = freeze.get("files")
    if not isinstance(files, list) or len(files) != len(FROZEN_PATHS):
        raise RuntimeError("candidate freeze must contain exactly two file descriptors")
    by_path: dict[str, dict[str, Any]] = {}
    for item in files:
        if not isinstance(item, dict) or not isinstance(item.get("path"), str):
            raise RuntimeError("candidate freeze file descriptor is incomplete")
        if item["path"] in by_path:
            raise RuntimeError(f"duplicate frozen source file: {item['path']}")
        by_path[item["path"]] = item
    if tuple(by_path) != FROZEN_PATHS:
        raise RuntimeError("candidate freeze source paths are not the two reviewed x14ac files")
    result: list[dict[str, Any]] = []
    for relative in FROZEN_PATHS:
        item = by_path[relative]
        path = ROOT / relative
        actual = identity(path)
        for field in ("bytes", "sha256"):
            if item.get(field) != actual[field]:
                raise RuntimeError(f"frozen source {relative} does not match live source: {field}")
        result.append({"path": relative, "expected": item, "actual": actual})
    return result


def source_snapshot() -> list[dict[str, Any]]:
    return [identity(ROOT / relative) for relative in FROZEN_PATHS]


def environment() -> dict[str, str]:
    cargo_bin = str(CARGO.parent)
    return {
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
        "CARGO_NET_OFFLINE": "true",
        "CARGO_PROFILE_DEV_DEBUG": "0",
        "CARGO_PROFILE_TEST_DEBUG": "0",
        "CARGO_TARGET_DIR": str(TARGET),
        "CARGO_HOME": "/home/zhuhe/.cargo",
        "HOME": "/home/zhuhe",
        "LANG": "C",
        "LC_ALL": "C",
        "PATH": f"{cargo_bin}:/usr/bin:/bin",
        "RUSTDOCFLAGS": "-D warnings -C debuginfo=0",
        "RUSTFLAGS": "-C debuginfo=0",
        "RUSTUP_HOME": "/home/zhuhe/.rustup",
        "RUSTUP_TOOLCHAIN": "1.95.0-x86_64-unknown-linux-gnu",
        "TMPDIR": "/tmp",
        "TZ": "UTC",
    }


def commands() -> list[tuple[str, list[str]]]:
    manifest = str(MANIFEST)
    common = ["--manifest-path", manifest, "--locked", "--offline", "-j", "2"]
    return [
        (
            "focused_x14ac_reuse_tests",
            [str(CARGO), "test", *common, "-p", "litchi-xlsx", "--lib", "raw::worksheet::x14ac::reuse_tests", "--", "--test-threads=1"],
        ),
        (
            "xlsx_full_lib_tests",
            [str(CARGO), "test", *common, "-p", "litchi-xlsx", "--lib", "--", "--test-threads=1"],
        ),
        (
            "xlsx_all_features_check",
            [str(CARGO), "check", *common, "-p", "litchi-xlsx", "--lib", "--all-features"],
        ),
        (
            "xlsx_clippy",
            [str(CARGO), "clippy", *common, "-p", "litchi-xlsx", "--lib", "--no-deps", "--", "-D", "warnings"],
        ),
        (
            "xlsx_rustdoc",
            [str(CARGO), "rustdoc", *common, "-p", "litchi-xlsx", "--lib", "--", "-D", "warnings"],
        ),
        (
            "xlsx_fmt",
            [str(CARGO), "fmt", "--manifest-path", manifest, "-p", "litchi-xlsx", "--", "--check"],
        ),
        (
            "crate_boundaries",
            ["python3", "-B", "tools/check_crate_boundaries.py"],
        ),
    ]


def run_command(name: str, argv: list[str], env: dict[str, str]) -> dict[str, Any]:
    log = LOGS / f"{name}.log"
    started = utc_now()
    started_ns = time.monotonic_ns()
    launch_error: str | None = None
    returncode: int | None = None
    with log.open("xb") as output:
        output.write(f"cwd={ROOT}\nargv={json.dumps(argv, separators=(',', ':'))}\n\n".encode())
        output.flush()
        try:
            result = subprocess.run(argv, cwd=ROOT, env=env, stdout=output, stderr=subprocess.STDOUT, check=False)
            returncode = result.returncode
        except OSError as error:
            launch_error = repr(error)
    test_result_matches: list[int] = []
    if name in {"focused_x14ac_reuse_tests", "xlsx_full_lib_tests"}:
        text = log.read_text(encoding="utf-8", errors="replace")
        test_result_matches = [int(value) for value in re.findall(r"test result: (?:ok|FAILED)\.\s+(\d+) passed;", text)]
    result: dict[str, Any] = {
        "name": name,
        "argv": argv,
        "cwd": str(ROOT),
        "environment": {key: env[key] for key in ENV_KEYS},
        "started_utc": started,
        "finished_utc": utc_now(),
        "elapsed_ns": time.monotonic_ns() - started_ns,
        "returncode": returncode,
        "pass": returncode == 0 and (name not in {"focused_x14ac_reuse_tests", "xlsx_full_lib_tests"} or sum(test_result_matches) > 0),
        "log": str(log.relative_to(PACKET)),
        "log_identity": identity(log),
    }
    if name in {"focused_x14ac_reuse_tests", "xlsx_full_lib_tests"}:
        result["passed_count"] = sum(test_result_matches)
        result["test_result_matches"] = test_result_matches
        if returncode == 0 and not test_result_matches:
            result["failure_reason"] = "no Cargo test result line"
        elif returncode == 0 and sum(test_result_matches) == 0:
            result["failure_reason"] = "Cargo test passed zero tests"
    if launch_error is not None:
        result["launch_error"] = launch_error
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run", action="store_true", help="run the frozen quality gates")
    parser.add_argument("--attempt", default="quality-v2", help="create-only attempt name under attempts/")
    parser.add_argument("--freeze", default="candidate-freeze-v2.json", help="packet-local candidate freeze JSON")
    args = parser.parse_args()
    if not args.run:
        print(json.dumps({"status": "prepared", "attempt": args.attempt, "freeze": args.freeze, "message": "waiting for approved candidate freeze; no Cargo command launched"}))
        return 0

    global ATTEMPT, LOGS, RECEIPT, START, FREEZE
    attempt_name = Path(args.attempt)
    if attempt_name.is_absolute() or attempt_name.name != args.attempt or args.attempt in {"", ".", ".."}:
        raise SystemExit("--attempt must be one packet-local create-only name")
    freeze_path = Path(args.freeze)
    if not freeze_path.is_absolute():
        freeze_path = PACKET / freeze_path
    freeze_path = freeze_path.resolve()
    if PACKET not in freeze_path.parents or freeze_path == PACKET:
        raise SystemExit("--freeze must resolve inside the 0891 packet")
    ATTEMPT = PACKET / "attempts" / attempt_name
    LOGS = ATTEMPT / "logs"
    RECEIPT = ATTEMPT / "quality-receipt.json"
    START = ATTEMPT / "quality-start.json"
    FREEZE = freeze_path
    if ATTEMPT.exists():
        raise SystemExit(f"refusing to reuse existing quality attempt: {ATTEMPT}")
    if TARGET.is_symlink() or not TARGET.is_dir() or not TARGET_MARKER.is_file() or TARGET_MARKER.read_text() != MARKER_TEXT:
        raise SystemExit("owned 0891 quality target marker is missing or mismatched")
    if not CARGO.is_file() or not os.access(CARGO, os.X_OK):
        raise SystemExit(f"Rust 1.95 Cargo is unavailable: {CARGO}")
    freeze = read_json(FREEZE)
    frozen = freeze_files(freeze)
    lock_before = identity(LOCK)
    freeze_before = identity(FREEZE)
    source_before = source_snapshot()
    env = environment()
    planned = [{"name": name, "argv": argv} for name, argv in commands()]
    ATTEMPT.mkdir(parents=True)
    LOGS.mkdir()
    runner_copy = ATTEMPT / "quality_runner.py"
    with runner_copy.open("xb") as output:
        output.write(Path(__file__).read_bytes())
    start_record = {
        "schema": "litchi.xlsx-x14ac-quality-start-0891.v1",
        "status": "running",
        "purpose": "frozen_xlsx_x14ac_quality_gates",
        "performance_claim": False,
        "started_utc": utc_now(),
        "cwd": str(ROOT),
        "manifest": str(MANIFEST),
        "target_dir": str(TARGET),
        "target_marker": identity(TARGET_MARKER),
        "freeze": {"path": str(FREEZE.relative_to(PACKET)), "identity": freeze_before, "schema": freeze.get("schema")},
        "frozen_files": frozen,
        "source_before": source_before,
        "root_lock_before": lock_before,
        "environment": {key: env[key] for key in ENV_KEYS},
        "planned_commands": planned,
        "runner": identity(runner_copy),
    }
    write_exclusive(START, start_record)

    commands_run: list[dict[str, Any]] = []
    failures: list[str] = []
    skipped: list[str] = []
    for index, (name, argv) in enumerate(commands()):
        command = run_command(name, argv, env)
        commands_run.append(command)
        if not command["pass"]:
            failures.append(name)
            skipped.extend(item[0] for item in commands()[index + 1 :])
            break

    source_after = source_snapshot()
    frozen_after = [{"path": item["path"], "actual": actual} for item, actual in zip(frozen, source_after, strict=True)]
    lock_after = identity(LOCK)
    freeze_after = identity(FREEZE)
    if source_after != source_before:
        failures.append("source_changed_during_quality")
    if lock_after != lock_before:
        failures.append("root_Cargo.lock_changed_during_quality")
    if freeze_after != freeze_before:
        failures.append("candidate_freeze_changed_during_quality")
    result = {
        "schema": "litchi.xlsx-x14ac-quality-0891.v1",
        "status": "pass" if not failures else "fail",
        "purpose": "frozen_xlsx_x14ac_quality_gates",
        "performance_claim": False,
        "started_utc": start_record["started_utc"],
        "finished_utc": utc_now(),
        "cwd": str(ROOT),
        "manifest": str(MANIFEST),
        "target_dir": str(TARGET),
        "target_marker": identity(TARGET_MARKER),
        "freeze": {"path": str(FREEZE.relative_to(PACKET)), "identity_before": freeze_before, "identity_after": freeze_after, "schema": freeze.get("schema")},
        "frozen_files_before": frozen,
        "frozen_files_after": frozen_after,
        "source_before": source_before,
        "source_after": source_after,
        "root_lock_before": lock_before,
        "root_lock_after": lock_after,
        "environment": {key: env[key] for key in ENV_KEYS},
        "commands": commands_run,
        "skipped_after_first_failure": skipped,
        "failures": failures,
        "stop_after_first_failure": True,
        "runner": identity(runner_copy),
    }
    write_exclusive(RECEIPT, result)
    print(json.dumps({"status": result["status"], "commands": [(item["name"], item["returncode"], item["pass"]) for item in commands_run], "failures": failures, "receipt": str(RECEIPT)}, sort_keys=True))
    return 0 if result["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
