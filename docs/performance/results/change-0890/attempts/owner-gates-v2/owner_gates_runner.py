#!/usr/bin/env python3
"""Run 0890 owner, shared-format, and workspace correctness gates sequentially."""

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
TARGET = ROOT / "target/litchi-0890-correctness"
TARGET_MARKER = TARGET / ".owner-0890-target"
CARGO = Path("/home/zhuhe/.rustup/toolchains/1.95.0-x86_64-unknown-linux-gnu/bin/cargo")
MANIFEST = ROOT / "Cargo.toml"
ATTEMPT_NAME = "owner-gates-v2"
ATTEMPT = PACKET / "attempts" / ATTEMPT_NAME
LOGS = ATTEMPT / "logs"
CANDIDATE_RECEIPT = PACKET / "attempts/candidate-diagnostic-f790c2228c-v2-retry1/candidate-receipt.json"
BASELINE_RECEIPT = PACKET / "attempts/baseline-f790c2228c-retry1/baseline-receipt.json"
FREEZE = PACKET / "candidate-freeze-v2.json"
LOCK_CUSTODY = PACKET / "attempts/baseline-f790c2228c-retry1/root-lock-custody.json"
SOURCE_PATHS = ("Cargo.toml", "clippy.toml", "deny.toml", "rust-toolchain.toml", "rustfmt.toml", ".cargo", "crates", "tools/perf-baseline")
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
    return {"path": str(path), "bytes": path.stat().st_size, "mode": oct(stat.S_IMODE(path.stat().st_mode)), "sha256": sha(path)}


def write_json(path: Path, value: dict[str, Any]) -> None:
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")
    temporary.replace(path)


def source_manifest() -> dict[str, Any]:
    listing = subprocess.run(["git", "ls-files", "-z", "--", *SOURCE_PATHS], cwd=ROOT, stdout=subprocess.PIPE, check=True).stdout
    paths = sorted(item.decode() for item in listing.split(b"\0") if item)
    files = []
    for name in paths:
        path = ROOT / name
        files.append({"path": name, "bytes": path.stat().st_size, "sha256": sha(path)})
    return {"paths": SOURCE_PATHS, "file_count": len(files), "files": files}


def environment() -> dict[str, str]:
    return {
        "CARGO_HOME": "/home/zhuhe/.cargo", "CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0",
        "CARGO_NET_OFFLINE": "true", "CARGO_PROFILE_DEV_DEBUG": "0", "CARGO_PROFILE_TEST_DEBUG": "0",
        "CARGO_TARGET_DIR": str(TARGET), "HOME": "/home/zhuhe", "LANG": "C", "LC_ALL": "C",
        "PATH": str(CARGO.parent) + ":/usr/bin:/bin", "RUSTDOCFLAGS": "-C debuginfo=0",
        "RUSTFLAGS": "-C debuginfo=0", "RUSTUP_HOME": "/home/zhuhe/.rustup", "TMPDIR": "/tmp", "TZ": "UTC",
    }


def run(name: str, argv: list[str], env: dict[str, str], receipt: dict[str, Any], receipt_path: Path) -> dict[str, Any]:
    log = LOGS / f"{name}.log"
    started = utc_now()
    start_ns = time.monotonic_ns()
    returncode: int | None = None
    launch_error: str | None = None
    with log.open("wb") as output:
        output.write((f"cwd={ROOT}\nargv={json.dumps(argv, separators=(',', ':'))}\n\n").encode())
        try:
            result = subprocess.run(argv, cwd=ROOT, env=env, stdout=output, stderr=subprocess.STDOUT, check=False)
            returncode = result.returncode
        except OSError as error:
            launch_error = repr(error)
    command: dict[str, Any] = {
        "name": name, "argv": argv, "cwd": str(ROOT), "expected": "pass", "returncode": returncode,
        "started_utc": started, "finished_utc": utc_now(), "elapsed_ns": time.monotonic_ns() - start_ns,
        "log": str(log.relative_to(PACKET)), "log_bytes": log.stat().st_size, "log_sha256": sha(log),
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
    if not TARGET_MARKER.is_file() or TARGET_MARKER.read_text() != "litchi 0890 correctness target\n":
        raise SystemExit("owned correctness target marker is missing")
    candidate = json.loads(CANDIDATE_RECEIPT.read_text())
    baseline = json.loads(BASELINE_RECEIPT.read_text())
    if candidate.get("status") != "pass" or baseline.get("status") != "pass":
        raise SystemExit("diagnostic receipts must pass before owner gates")
    ATTEMPT.mkdir(parents=True)
    LOGS.mkdir()
    copy = ATTEMPT / Path(__file__).name
    copy.write_bytes(Path(__file__).read_bytes())
    env = environment()
    receipt: dict[str, Any] = {
        "schema": "litchi.mce-0890-owner-gates.v1",
        "purpose": "owner_and_shared_format_correctness_gates",
        "performance_claim": False,
        "status": "running",
        "started_utc": utc_now(),
        "cwd": str(ROOT),
        "target_dir": str(TARGET),
        "target_marker": str(TARGET_MARKER),
        "manifest": str(MANIFEST),
        "environment": {key: env.get(key, "") for key in ENV_KEYS},
        "candidate_receipt": {"path": str(CANDIDATE_RECEIPT.relative_to(PACKET)), "sha256": sha(CANDIDATE_RECEIPT)},
        "baseline_receipt": {"path": str(BASELINE_RECEIPT.relative_to(PACKET)), "sha256": sha(BASELINE_RECEIPT)},
        "candidate_freeze": {"path": str(FREEZE.relative_to(PACKET)), "sha256": sha(FREEZE)},
        "root_lock_custody": {"path": str(LOCK_CUSTODY.relative_to(PACKET)), "sha256": sha(LOCK_CUSTODY)},
        "source_before": source_manifest(),
        "commands": [],
        "runner": {"path": str(copy.relative_to(PACKET)), "bytes": copy.stat().st_size, "sha256": sha(copy)},
    }
    receipt_path = ATTEMPT / "owner-gates-receipt.json"
    write_json(receipt_path, receipt)
    common = ["--manifest-path", str(MANIFEST), "--locked", "--offline"]
    commands = [
        ("mce_owner_tests", [str(CARGO), "test", *common, "-p", "litchi-ooxml-common", "--lib", "mce", "--", "--test-threads=1"]),
        ("xlsx_shared_tests", [str(CARGO), "test", *common, "-p", "litchi-xlsx", "--lib", "--", "--test-threads=1"]),
        ("docx_shared_tests", [str(CARGO), "test", *common, "-p", "litchi-docx", "--lib", "--", "--test-threads=1"]),
        ("pptx_shared_tests", [str(CARGO), "test", *common, "-p", "litchi-pptx", "--lib", "--", "--test-threads=1"]),
        ("fmt_check", [str(CARGO), "fmt", "--manifest-path", str(MANIFEST), "--all", "--", "--check"]),
        ("mce_owner_clippy", [str(CARGO), "clippy", *common, "-p", "litchi-ooxml-common", "--lib", "--no-deps", "--", "-D", "warnings"]),
        ("mce_owner_rustdoc", [str(CARGO), "rustdoc", *common, "-p", "litchi-ooxml-common", "--lib", "--", "-D", "warnings"]),
        ("crate_boundary_check", ["python3", "-B", "tools/check_crate_boundaries.py"]),
    ]
    for name, argv in commands:
        command = run(name, argv, env, receipt, receipt_path)
        if not command["pass"]:
            receipt.setdefault("failures", []).append(name)

    receipt["source_after"] = source_manifest()
    if receipt["source_before"] != receipt["source_after"]:
        receipt.setdefault("failures", []).append("source_changed_during_owner_gates")
    current_binary = ROOT / "target/litchi-0890-correctness/debug/xml_attribute_bounds"
    if current_binary.is_file():
        receipt["diagnostic_binary_after"] = identity(current_binary)
        if receipt["diagnostic_binary_after"] != candidate.get("candidate_binary"):
            receipt.setdefault("failures", []).append("diagnostic_binary_changed_during_owner_gates")
    else:
        receipt.setdefault("failures", []).append("diagnostic_binary_missing_after_owner_gates")
    receipt["finished_utc"] = utc_now()
    receipt["status"] = "pass" if not receipt.get("failures") else "fail"
    write_json(receipt_path, receipt)
    print(json.dumps({"status": receipt["status"], "receipt": str(receipt_path), "commands": [(x["name"], x["returncode"], x["pass"]) for x in receipt["commands"]], "failures": receipt.get("failures", [])}, sort_keys=True))
    return 0 if receipt["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
