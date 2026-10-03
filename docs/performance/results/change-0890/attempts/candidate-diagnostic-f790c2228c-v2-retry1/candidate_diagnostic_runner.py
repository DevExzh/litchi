#!/usr/bin/env python3
"""Build and qualify the frozen 0890 candidate diagnostic source."""

from __future__ import annotations

import datetime as dt
import gzip
import hashlib
import json
from pathlib import Path
import shutil
import stat
import subprocess
import time
from typing import Any


REPO = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
SOURCE = Path("/tmp/litchi-0890-correctness-source")
TARGET = REPO / "target/litchi-0890-correctness"
TARGET_MARKER = TARGET / ".owner-0890-target"
MANIFEST = SOURCE / "tools/perf-baseline/Cargo.toml"
CARGO = Path("/home/zhuhe/.rustup/toolchains/1.95.0-x86_64-unknown-linux-gnu/bin/cargo")
BINARY = TARGET / "debug/xml_attribute_bounds"
BASELINE_RECEIPT = PACKET / "attempts/baseline-f790c2228c-retry1/baseline-receipt.json"
OVERLAY_RECEIPT = PACKET / "candidate-diagnostic-overlay.json"
ATTEMPT_NAME = "candidate-diagnostic-f790c2228c-v2-retry1"
ATTEMPT = PACKET / "attempts" / ATTEMPT_NAME
LOGS = ATTEMPT / "logs"
RUNTIME = ATTEMPT / "runtime"
FIXTURES = [
    REPO / "test-data/ooxml/xlsx/StructuredRefs-lots-with-lookups.xlsx",
    REPO / "test-data/ooxml/docx/drawing.docx",
    REPO / "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx",
]


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


def env() -> dict[str, str]:
    return {
        "CARGO_HOME": "/home/zhuhe/.cargo",
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
        "CARGO_NET_OFFLINE": "true",
        "CARGO_PROFILE_DEV_DEBUG": "0",
        "CARGO_PROFILE_TEST_DEBUG": "0",
        "CARGO_TARGET_DIR": str(TARGET),
        "HOME": "/home/zhuhe",
        "LANG": "C",
        "LC_ALL": "C",
        "PATH": str(CARGO.parent) + ":/usr/bin:/bin",
        "RUSTDOCFLAGS": "-C debuginfo=0",
        "RUSTFLAGS": "-C debuginfo=0",
        "RUSTUP_HOME": "/home/zhuhe/.rustup",
        "TMPDIR": "/tmp",
        "TZ": "UTC",
    }


def run_command(name: str, argv: list[str], environment: dict[str, str], receipt: dict[str, Any], receipt_path: Path, *, cwd: Path = SOURCE) -> dict[str, Any]:
    log = LOGS / f"{name}.log"
    started = utc_now()
    started_ns = time.monotonic_ns()
    launch_error: str | None = None
    returncode: int | None = None
    with log.open("wb") as output:
        output.write((f"cwd={cwd}\nargv={json.dumps(argv, separators=(',', ':'))}\n\n").encode())
        try:
            result = subprocess.run(argv, cwd=cwd, env=environment, stdout=output, stderr=subprocess.STDOUT, check=False)
            returncode = result.returncode
        except OSError as error:
            launch_error = repr(error)
    command = {
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


def differential_argv(output: Path) -> list[str]:
    return [
        str(BINARY), "differential", "--json", str(output),
        "--generated", "20000", "--aliasing", "20000",
        *(str(path) for path in FIXTURES),
    ]


def compress_report(source: Path, destination: Path) -> dict[str, Any]:
    source_size = source.stat().st_size
    source_sha = sha(source)
    with source.open("rb") as input_file, destination.open("wb") as output_file:
        with gzip.GzipFile(filename="", mode="wb", fileobj=output_file, compresslevel=9, mtime=0) as compressed:
            shutil.copyfileobj(input_file, compressed, length=1024 * 1024)
    digest = hashlib.sha256()
    size = 0
    with gzip.open(destination, "rb") as input_file:
        for block in iter(lambda: input_file.read(1024 * 1024), b""):
            digest.update(block)
            size += len(block)
    if (size, digest.hexdigest()) != (source_size, source_sha):
        raise RuntimeError(f"gzip round-trip failed: {destination}")
    return {
        "path": str(destination.relative_to(PACKET)),
        "bytes": destination.stat().st_size,
        "sha256": sha(destination),
        "decompressed_bytes": source_size,
        "decompressed_sha256": source_sha,
    }


def main() -> int:
    if ATTEMPT.exists():
        raise SystemExit(f"create-only attempt already exists: {ATTEMPT}")
    if not SOURCE.is_dir() or not TARGET.is_dir() or not TARGET_MARKER.is_file():
        raise SystemExit("owned source/target roots are missing")
    overlay = json.loads(OVERLAY_RECEIPT.read_text())
    if overlay.get("status") != "pass" or not overlay.get("unchanged_non_candidate_sources"):
        raise SystemExit("candidate diagnostic overlay receipt is not passing")
    baseline = json.loads(BASELINE_RECEIPT.read_text())
    if baseline.get("status") != "pass":
        raise SystemExit("baseline receipt is not passing")
    for fixture in FIXTURES:
        if not fixture.is_file():
            raise SystemExit(f"fixture is missing: {fixture}")

    ATTEMPT.mkdir(parents=True)
    LOGS.mkdir()
    RUNTIME.mkdir()
    copy = ATTEMPT / Path(__file__).name
    shutil.copy2(__file__, copy)
    environment = env()
    receipt: dict[str, Any] = {
        "schema": "litchi.mce-0890-correctness-candidate.v1",
        "purpose": "candidate_differential_and_focused_test_qualification",
        "performance_claim": False,
        "status": "running",
        "started_utc": utc_now(),
        "source_root": str(SOURCE),
        "target_dir": str(TARGET),
        "target_marker": str(TARGET_MARKER),
        "manifest": str(MANIFEST),
        "toolchain": {"cargo": str(CARGO)},
        "environment": environment,
        "overlay_receipt": {"path": str(OVERLAY_RECEIPT.relative_to(PACKET)), "sha256": sha(OVERLAY_RECEIPT)},
        "baseline_receipt": {"path": str(BASELINE_RECEIPT.relative_to(PACKET)), "sha256": sha(BASELINE_RECEIPT)},
        "fixtures": [identity(path) for path in FIXTURES],
        "commands": [],
        "runner": {"path": str(copy.relative_to(PACKET)), "bytes": copy.stat().st_size, "sha256": sha(copy)},
    }
    receipt_path = ATTEMPT / "candidate-receipt.json"
    try:
        write_json(receipt_path, receipt)
        common = ["--manifest-path", str(MANIFEST), "--locked", "--offline"]
        build = run_command("cargo-build-candidate", [str(CARGO), "build", *common, "--bin", "xml_attribute_bounds"], environment, receipt, receipt_path)
        if not build["pass"] or not BINARY.is_file():
            receipt.setdefault("failures", []).append("candidate_build")
        else:
            receipt["baseline_binary"] = baseline.get("binary")
            receipt["candidate_binary"] = identity(BINARY)
            receipt["binary_differs_from_baseline"] = receipt["candidate_binary"] != baseline.get("binary")
            write_json(receipt_path, receipt)
            test_manifest = REPO / "tools/perf-baseline/Cargo.toml"
            test = run_command("cargo-test-diagnostic", [str(CARGO), "test", "--manifest-path", str(test_manifest), "--locked", "--offline", "--bin", "xml_attribute_bounds", "--", "--test-threads=1"], environment, receipt, receipt_path, cwd=REPO)
            if not test["pass"]:
                receipt.setdefault("failures", []).append("candidate_diagnostic_tests")
            if test["pass"]:
                first = RUNTIME / "differential-first.json"
                second = RUNTIME / "differential-second.json"
                first_record = run_command("differential-candidate-a", differential_argv(first), environment, receipt, receipt_path)
                second_record = run_command("differential-candidate-b", differential_argv(second), environment, receipt, receipt_path)
                if not first_record["pass"] or not second_record["pass"]:
                    receipt.setdefault("failures", []).append("candidate_differential")
                elif not first.is_file() or not second.is_file():
                    receipt.setdefault("failures", []).append("candidate_differential_output_missing")
                else:
                    first_sha = sha(first)
                    second_sha = sha(second)
                    receipt["differential"] = {
                        "generated": 20000,
                        "aliasing": 20000,
                        "packages": 3,
                        "selected_members": 67,
                        "first_bytes": first.stat().st_size,
                        "first_sha256": first_sha,
                        "second_bytes": second.stat().st_size,
                        "second_sha256": second_sha,
                        "byte_identical": first_sha == second_sha and first.read_bytes() == second.read_bytes(),
                        "compressed_first": compress_report(first, ATTEMPT / "differential-first.json.gz"),
                        "compressed_second": compress_report(second, ATTEMPT / "differential-second.json.gz"),
                    }
                    first.unlink()
                    second.unlink()
                    if not receipt["differential"]["byte_identical"]:
                        receipt.setdefault("failures", []).append("candidate_differential_repeat_mismatch")
    except Exception as error:
        receipt.setdefault("failures", []).append(repr(error))
    receipt["finished_utc"] = utc_now()
    receipt["status"] = "pass" if not receipt.get("failures") else "fail"
    write_json(receipt_path, receipt)
    print(json.dumps({"status": receipt["status"], "receipt": str(receipt_path), "binary": receipt.get("candidate_binary"), "differential": receipt.get("differential")}, sort_keys=True))
    return 0 if receipt["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
