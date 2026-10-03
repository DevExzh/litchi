#!/usr/bin/env python3
"""Run the 0890 continuation gates after the sealed owner-gate attempt."""

from __future__ import annotations

import datetime as dt
import hashlib
import json
from pathlib import Path
import stat
import subprocess
import time
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
SOURCE = Path("/tmp/litchi-0890-correctness-source")
TARGET = ROOT / "target/litchi-0890-correctness"
TARGET_MARKER = TARGET / ".owner-0890-target"
SOURCE_MARKER = SOURCE / ".owner-0890-source"
CARGO = Path("/home/zhuhe/.rustup/toolchains/1.95.0-x86_64-unknown-linux-gnu/bin/cargo")
ROOT_MANIFEST = ROOT / "Cargo.toml"
SOURCE_MANIFEST = SOURCE / "Cargo.toml"
BASELINE_RECEIPT = PACKET / "attempts/baseline-f790c2228c-retry1/baseline-receipt.json"
CANDIDATE_RECEIPT = PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/candidate-receipt.json"
CANDIDATE_FREEZE = PACKET / "candidate-freeze-v2.json"
LOCK_CUSTODY = PACKET / "attempts/baseline-f790c2228c-retry1/root-lock-custody.json"
PRIOR_OWNER_RECEIPT = PACKET / "attempts/owner-gates-v2/owner-gates-receipt.json"
ATTEMPT_NAME = "owner-continuation-v2"
ATTEMPT = PACKET / "attempts" / ATTEMPT_NAME
LOGS = ATTEMPT / "logs"
CANDIDATE_PATHS = (
    "crates/litchi-ooxml-common/src/mce/codec.rs",
    "crates/litchi-ooxml-common/src/mce/tests.rs",
)
ENV_KEYS = (
    "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL", "CARGO_NET_OFFLINE",
    "CARGO_PROFILE_DEV_DEBUG", "CARGO_PROFILE_TEST_DEBUG", "CARGO_HOME",
    "HOME", "CARGO_TARGET_DIR", "LANG", "LC_ALL", "PATH", "RUSTDOCFLAGS",
    "RUSTFLAGS", "RUSTUP_HOME", "TMPDIR", "TZ",
)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def identity(path: Path) -> dict[str, Any]:
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "mode": oct(stat.S_IMODE(path.stat().st_mode)),
        "sha256": sha(path),
    }


def write_json(path: Path, value: dict[str, Any]) -> None:
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")
    temporary.replace(path)


def environment() -> dict[str, str]:
    return {
        "CARGO_HOME": "/home/zhuhe/.cargo", "CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0",
        "CARGO_NET_OFFLINE": "true", "CARGO_PROFILE_DEV_DEBUG": "0", "CARGO_PROFILE_TEST_DEBUG": "0",
        "CARGO_TARGET_DIR": str(TARGET), "HOME": "/home/zhuhe", "LANG": "C", "LC_ALL": "C",
        "PATH": str(CARGO.parent) + ":/usr/bin:/bin", "RUSTDOCFLAGS": "-C debuginfo=0",
        "RUSTFLAGS": "-C debuginfo=0", "RUSTUP_HOME": "/home/zhuhe/.rustup", "TMPDIR": "/tmp", "TZ": "UTC",
    }


def expected_source_files(baseline: dict[str, Any]) -> list[dict[str, Any]]:
    source = baseline.get("source")
    if not isinstance(source, dict) or not isinstance(source.get("files"), list):
        raise RuntimeError("baseline source manifest is missing")
    return source["files"]


def source_manifest(expected: list[dict[str, Any]]) -> dict[str, Any]:
    expected_paths = {entry["path"] for entry in expected}
    files: list[dict[str, Any]] = []
    for name in sorted(expected_paths):
        path = SOURCE / name
        if not path.is_file():
            raise RuntimeError(f"isolated source file is missing: {name}")
        files.append({
            "path": name,
            "bytes": path.stat().st_size,
            "sha256": sha(path),
        })
    actual_paths = {
        path.relative_to(SOURCE).as_posix()
        for path in SOURCE.rglob("*")
        if path.is_file() and path != SOURCE_MARKER
    }
    unexpected = sorted(actual_paths - expected_paths)
    missing = sorted(expected_paths - actual_paths)
    return {
        "root": str(SOURCE),
        "file_count": len(files),
        "files": files,
        "unexpected_files": unexpected,
        "missing_files": missing,
    }


def run(
    name: str,
    argv: list[str],
    cwd: Path,
    env: dict[str, str],
    receipt: dict[str, Any],
    receipt_path: Path,
) -> dict[str, Any]:
    log = LOGS / f"{name}.log"
    started = utc_now()
    started_ns = time.monotonic_ns()
    returncode: int | None = None
    launch_error: str | None = None
    with log.open("wb") as output:
        output.write((f"cwd={cwd}\nargv={json.dumps(argv, separators=(',', ':'))}\n\n").encode())
        try:
            result = subprocess.run(argv, cwd=cwd, env=env, stdout=output, stderr=subprocess.STDOUT, check=False)
            returncode = result.returncode
        except OSError as error:
            launch_error = repr(error)
    command: dict[str, Any] = {
        "name": name,
        "argv": argv,
        "cwd": str(cwd),
        "expected": "pass",
        "returncode": returncode,
        "started_utc": started,
        "finished_utc": utc_now(),
        "elapsed_ns": time.monotonic_ns() - started_ns,
        "log": str(log.relative_to(PACKET)),
        "log_bytes": log.stat().st_size,
        "log_sha256": sha(log),
        "pass": returncode == 0,
    }
    if launch_error is not None:
        command["launch_error"] = launch_error
    receipt.setdefault("commands", []).append(command)
    receipt["last_command"] = command
    write_json(receipt_path, receipt)
    return command


def main() -> int:
    if ATTEMPT.exists():
        raise SystemExit(f"create-only attempt already exists: {ATTEMPT}")
    if not SOURCE.is_dir() or not SOURCE_MARKER.is_file():
        raise SystemExit("owned isolated source is missing")
    if not TARGET.is_dir() or not TARGET_MARKER.is_file():
        raise SystemExit("owned correctness target is missing")
    baseline = json.loads(BASELINE_RECEIPT.read_text())
    candidate = json.loads(CANDIDATE_RECEIPT.read_text())
    freeze = json.loads(CANDIDATE_FREEZE.read_text())
    prior_owner = json.loads(PRIOR_OWNER_RECEIPT.read_text())
    if baseline.get("status") != "pass" or candidate.get("status") != "pass":
        raise SystemExit("baseline and candidate diagnostic receipts must pass")
    if prior_owner.get("status") != "fail" or prior_owner.get("failures") != ["fmt_check"]:
        raise SystemExit("prior owner receipt must preserve the sole known fmt_check failure")
    if freeze.get("base_commit") != "f790c2228cbd9e89c2892ba3ffaeb49802fac43b":
        raise SystemExit("candidate freeze base mismatch")
    expected = expected_source_files(baseline)
    expected_by_path = {entry["path"]: entry for entry in expected}
    initial_manifest = source_manifest(expected)
    if set(expected_by_path) != {entry["path"] for entry in initial_manifest["files"]}:
        raise SystemExit("unexpected baseline source manifest shape")
    for path in CANDIDATE_PATHS:
        current = identity(SOURCE / path)
        frozen = next(entry for entry in freeze["files"] if entry["path"] == path)
        if current["bytes"] != frozen["bytes"] or current["sha256"] != frozen["sha256"]:
            raise SystemExit(f"isolated candidate source differs from freeze: {path}")
    source_before = initial_manifest
    if source_before["unexpected_files"] or source_before["missing_files"]:
        raise SystemExit("isolated source file inventory is not exact")

    ATTEMPT.mkdir(parents=True)
    LOGS.mkdir()
    runner_copy = ATTEMPT / Path(__file__).name
    runner_copy.write_bytes(Path(__file__).read_bytes())
    env = environment()
    current_binary = TARGET / "debug/xml_attribute_bounds"
    if not current_binary.is_file():
        raise SystemExit("candidate diagnostic binary is missing before continuation")
    if identity(current_binary) != candidate.get("candidate_binary"):
        raise SystemExit("candidate diagnostic binary differs before continuation")
    unrelated = ROOT / "crates/litchi/examples/bench_word_parse.rs"
    prior_fmt = PACKET / "attempts/owner-gates-v2/logs/fmt_check.log"
    receipt: dict[str, Any] = {
        "schema": "litchi.mce-0890-owner-continuation.v1",
        "purpose": "isolated_format_and_unfiltered_owner_correctness_gates",
        "performance_claim": False,
        "status": "running",
        "started_utc": utc_now(),
        "root": str(ROOT),
        "source_root": str(SOURCE),
        "target_dir": str(TARGET),
        "target_marker": str(TARGET_MARKER),
        "manifest": str(ROOT_MANIFEST),
        "environment": {key: env.get(key, "") for key in ENV_KEYS},
        "baseline_receipt": {"path": str(BASELINE_RECEIPT.relative_to(PACKET)), "sha256": sha(BASELINE_RECEIPT)},
        "candidate_receipt": {"path": str(CANDIDATE_RECEIPT.relative_to(PACKET)), "sha256": sha(CANDIDATE_RECEIPT)},
        "candidate_freeze": {"path": str(CANDIDATE_FREEZE.relative_to(PACKET)), "sha256": sha(CANDIDATE_FREEZE)},
        "root_lock_custody": {"path": str(LOCK_CUSTODY.relative_to(PACKET)), "sha256": sha(LOCK_CUSTODY)},
        "prior_owner_attempt": {"path": str(PRIOR_OWNER_RECEIPT.relative_to(PACKET)), "sha256": sha(PRIOR_OWNER_RECEIPT)},
        "prior_global_fmt_analysis": {
            "receipt": str(PRIOR_OWNER_RECEIPT.relative_to(PACKET)),
            "log": str(prior_fmt.relative_to(PACKET)),
            "log_sha256": sha(prior_fmt),
            "unrelated_untracked_file": identity(unrelated),
            "cause": "root-wide fmt check failed only on the unrelated untracked benchmark source; the frozen tracked candidate is checked independently below",
        },
        "source_before": source_before,
        "candidate_binary_before": identity(current_binary),
        "runner": {"path": str(runner_copy.relative_to(PACKET)), "bytes": runner_copy.stat().st_size, "sha256": sha(runner_copy)},
        "commands": [],
    }
    receipt_path = ATTEMPT / "owner-continuation-receipt.json"
    write_json(receipt_path, receipt)
    commands = [
        (
            "isolated_workspace_fmt",
            [str(CARGO), "fmt", "--manifest-path", str(SOURCE_MANIFEST), "--all", "--", "--check"],
            SOURCE,
        ),
        (
            "root_owner_fmt",
            [str(CARGO), "fmt", "--manifest-path", str(ROOT_MANIFEST), "--package", "litchi-ooxml-common", "--", "--check"],
            ROOT,
        ),
        (
            "mce_owner_unfiltered_tests",
            [str(CARGO), "test", "--manifest-path", str(ROOT_MANIFEST), "--locked", "--offline", "-p", "litchi-ooxml-common", "--lib", "--", "--test-threads=1"],
            ROOT,
        ),
    ]
    for name, argv, cwd in commands:
        command = run(name, argv, cwd, env, receipt, receipt_path)
        if not command["pass"]:
            receipt.setdefault("failures", []).append(name)
    receipt["source_after"] = source_manifest(expected)
    if receipt["source_before"] != receipt["source_after"]:
        receipt.setdefault("failures", []).append("source_changed_during_continuation")
    if receipt["source_after"]["unexpected_files"] or receipt["source_after"]["missing_files"]:
        receipt.setdefault("failures", []).append("source_inventory_changed_during_continuation")
    if current_binary.is_file():
        receipt["candidate_binary_after"] = identity(current_binary)
        if receipt["candidate_binary_after"] != candidate.get("candidate_binary"):
            receipt.setdefault("failures", []).append("diagnostic_binary_changed_during_continuation")
    else:
        receipt.setdefault("failures", []).append("diagnostic_binary_missing_after_continuation")
    receipt["finished_utc"] = utc_now()
    receipt["status"] = "pass" if not receipt.get("failures") else "fail"
    write_json(receipt_path, receipt)
    print(json.dumps({
        "status": receipt["status"],
        "receipt": str(receipt_path),
        "commands": [(entry["name"], entry["returncode"], entry["pass"]) for entry in receipt["commands"]],
        "failures": receipt.get("failures", []),
    }, sort_keys=True))
    return 0 if receipt["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
