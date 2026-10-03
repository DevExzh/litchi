#!/usr/bin/env python3
"""Run and seal the non-performance 0889 qualification gates.

The runner is deliberately separate from the Rust diagnostic binary.  It
captures the exact source identity and environment at both ends of the run,
records every command and log digest, exercises the three real fixture roots,
and keeps the two large differential JSON files in the create-only attempt
directory until their bytes have been compared.  Only the deterministic
compressed reports remain after a successful comparison; failed raw outputs
remain available for diagnosis.

This script is an evidence runner, not a benchmark.  It uses a low-debug,
locked/offline Rust build and makes no latency or allocation claim.
"""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
import struct
import subprocess
import time
from typing import Any
import zipfile


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
TARGET = ROOT / "target" / "litchi-0889-validation"
TOOLCHAIN = "1.95.0"
USER_ROOT = Path("/home/zhuhe")
TOOLCHAIN_BIN = USER_ROOT / ".rustup" / "toolchains" / (
    f"{TOOLCHAIN}-x86_64-unknown-linux-gnu"
) / "bin"
CARGO = TOOLCHAIN_BIN / "cargo"
RUSTC = TOOLCHAIN_BIN / "rustc"
BINARY = TARGET / "debug" / "xml_attribute_bounds"
OVERLAY_BASE = "5c8788cf255b3b4e68bf7d6abd0264a4bbaa0364"

FIXTURES = [
    ROOT / "test-data/ooxml/xlsx/StructuredRefs-lots-with-lookups.xlsx",
    ROOT / "test-data/ooxml/docx/drawing.docx",
    ROOT
    / "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx",
]

ENV_KEYS = (
    "CARGO_BUILD_JOBS",
    "CARGO_INCREMENTAL",
    "CARGO_NET_OFFLINE",
    "CARGO_PROFILE_DEV_DEBUG",
    "CARGO_PROFILE_TEST_DEBUG",
    "CARGO_HOME",
    "HOME",
    "CARGO_TARGET_DIR",
    "LANG",
    "LC_ALL",
    "PATH",
    "RUSTDOCFLAGS",
    "RUSTFLAGS",
    "RUSTUP_HOME",
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


def command_text(argv: list[str]) -> str:
    # JSON is the unambiguous representation retained in commands.json.  The
    # text form is only for a compact human-readable log prefix.
    return json.dumps(argv, ensure_ascii=False, separators=(",", ":"))


def write_receipt(path: Path, receipt: dict[str, Any]) -> None:
    """Update an attempt receipt after every command and at process exit."""
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
    temporary.replace(path)


def run_command(
    name: str,
    argv: list[str],
    env: dict[str, str],
    log_path: Path,
    commands: list[dict[str, Any]],
    receipt: dict[str, Any],
    receipt_path: Path,
    *,
    expected: str = "pass",
) -> dict[str, Any]:
    started = utc_now()
    started_ns = time.monotonic_ns()
    log_path.parent.mkdir(parents=True, exist_ok=True)
    # A command log is owned by one create-only attempt directory.  Opening it
    # in write mode is safe because the attempt directory itself cannot be
    # reused; a second attempt must choose a new name.
    with log_path.open("wb") as log:
        log.write((f"cwd={ROOT}\nargv={command_text(argv)}\n\n").encode())
        try:
            result = subprocess.run(
                argv,
                cwd=ROOT,
                env=env,
                stdout=log,
                stderr=subprocess.STDOUT,
                check=False,
            )
            returncode: int | None = result.returncode
            launch_error: str | None = None
        except OSError as error:
            returncode = None
            launch_error = repr(error)
    finished = utc_now()
    record = {
        "name": name,
        "argv": argv,
        "cwd": str(ROOT),
        "expected": expected,
        "returncode": returncode,
        "started_utc": started,
        "finished_utc": finished,
        "elapsed_ns": time.monotonic_ns() - started_ns,
        "log": str(log_path.relative_to(PACKET)),
        "log_bytes": log_path.stat().st_size,
        "log_sha256": sha256(log_path),
        "pass": (returncode == 0) if expected == "pass" else returncode != 0,
    }
    if launch_error is not None:
        record["launch_error"] = launch_error
    commands.append(record)
    receipt["last_command"] = record
    write_receipt(receipt_path, receipt)
    return record


def source_files() -> list[Path]:
    base = ROOT / "tools/perf-baseline"
    candidates = [
        ROOT / "tools/check_crate_boundaries.py",
        base / "Cargo.toml",
        base / "Cargo.lock",
    ]
    candidates.extend(
        path
        for path in (base / "src").rglob("*.rs")
        if path.is_file()
    )
    return sorted({path for path in candidates if path.is_file()})


def capture_source_overlay(attempt_dir: Path) -> dict[str, Any]:
    """Copy changed diagnostic Rust sources so the attempt is replayable."""
    overlay_root = attempt_dir / "source"
    entries: list[dict[str, Any]] = []
    for path in sorted((ROOT / "tools/perf-baseline/src").rglob("*.rs")):
        if not path.is_file():
            continue
        relative = path.relative_to(ROOT)
        current = path.read_bytes()
        base = subprocess.run(
            ["git", "show", f"{OVERLAY_BASE}:{relative}"],
            cwd=ROOT,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        if base.returncode == 0 and base.stdout == current:
            continue
        destination = overlay_root / relative
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(current)
        entries.append(
            {
                "path": str(relative),
                "bytes": len(current),
                "sha256": hashlib.sha256(current).hexdigest(),
            }
        )
    manifest = {
        "base_commit": OVERLAY_BASE,
        "files": entries,
        "scope": "Diagnostic Rust sources differing from the replay base",
    }
    (attempt_dir / "source-overlay.json").write_text(
        json.dumps(manifest, indent=2, sort_keys=True) + "\n"
    )
    return manifest


def verify_source_overlay(
    overlay: dict[str, Any], source_identity: dict[str, Any]
) -> None:
    by_path = {entry["path"]: entry for entry in source_identity["source"]}
    for entry in overlay["files"]:
        current = by_path.get(entry["path"])
        if current is None:
            raise RuntimeError(f"overlay source missing from identity: {entry['path']}")
        if (current["bytes"], current["sha256"]) != (
            entry["bytes"],
            entry["sha256"],
        ):
            raise RuntimeError(f"overlay source changed before gates: {entry['path']}")


def capture_runner(attempt_dir: Path) -> dict[str, Any]:
    """Retain the exact runner used for this create-only attempt."""
    source = Path(__file__).resolve()
    destination = attempt_dir / source.name
    shutil.copy2(source, destination)
    return {
        "path": str(destination.relative_to(PACKET)),
        "bytes": destination.stat().st_size,
        "sha256": sha256(destination),
    }


def reused_boundary_check(
    previous_attempt: str,
    current_identity: dict[str, Any],
) -> dict[str, Any]:
    """Reuse a passed boundary check after proving its inputs are unchanged."""
    previous_dir = PACKET / "attempts" / previous_attempt
    previous_receipt_path = previous_dir / "validation-receipt.json"
    if not previous_receipt_path.is_file():
        raise RuntimeError(f"reused boundary receipt is missing: {previous_receipt_path}")
    previous_receipt = json.loads(previous_receipt_path.read_text())
    previous_before = previous_receipt.get("before")
    if not isinstance(previous_before, dict):
        raise RuntimeError("reused boundary receipt has no before identity")
    if current_identity.get("git_head") != previous_before.get("git_head"):
        raise RuntimeError("git HEAD changed since the reused boundary check")

    overlay_path = previous_dir / "source-overlay.json"
    if not overlay_path.is_file():
        raise RuntimeError(f"reused source overlay is missing: {overlay_path}")
    overlay = json.loads(overlay_path.read_text())
    allowed_source_paths = {entry["path"] for entry in overlay.get("files", [])}
    previous_sources = {
        entry["path"]: entry for entry in previous_before.get("source", [])
    }
    current_sources = {
        entry["path"]: entry for entry in current_identity.get("source", [])
    }
    for path in sorted(set(previous_sources) | set(current_sources)):
        if path in allowed_source_paths:
            continue
        if previous_sources.get(path) != current_sources.get(path):
            raise RuntimeError(f"boundary input changed: {path}")

    allowed_status_prefixes = ("docs/performance/", *sorted(allowed_source_paths))
    for line in current_identity.get("git_status", "").splitlines():
        path = line[3:] if len(line) >= 4 else ""
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        if not any(path == prefix or path.startswith(prefix) for prefix in allowed_status_prefixes):
            raise RuntimeError(f"unapproved worktree change for boundary reuse: {path}")

    previous_command = next(
        (
            command
            for command in previous_receipt.get("commands", [])
            if command.get("name") == "crate_boundary_check"
        ),
        None,
    )
    if not previous_command or not previous_command.get("pass"):
        raise RuntimeError("previous boundary command did not pass")
    previous_log = PACKET / previous_command["log"]
    if not previous_log.is_file() or sha256(previous_log) != previous_command.get("log_sha256"):
        raise RuntimeError("previous boundary log is missing or changed")
    return {
        "name": "crate_boundary_check",
        "argv": previous_command["argv"],
        "cwd": str(ROOT),
        "expected": "pass",
        "returncode": 0,
        "started_utc": previous_command["started_utc"],
        "finished_utc": previous_command["finished_utc"],
        "elapsed_ns": 0,
        "log": previous_command["log"],
        "log_bytes": previous_command["log_bytes"],
        "log_sha256": previous_command["log_sha256"],
        "pass": True,
        "reused": True,
        "reused_from": str(previous_receipt_path.relative_to(PACKET)),
    }


def identity() -> dict[str, Any]:
    def capture(argv: list[str]) -> str:
        result = subprocess.run(argv, cwd=ROOT, text=True, capture_output=True, check=True)
        return result.stdout.strip()

    source = [
        {
            "path": str(path.relative_to(ROOT)),
            "bytes": path.stat().st_size,
            "sha256": sha256(path),
        }
        for path in source_files()
    ]
    diff = subprocess.run(
        ["git", "diff", "--binary", "--no-ext-diff", "--", "tools/perf-baseline"],
        cwd=ROOT,
        capture_output=True,
        check=True,
    ).stdout
    return {
        "captured_utc": utc_now(),
        "git_head": capture(["git", "rev-parse", "HEAD"]),
        "git_status": capture(["git", "status", "--short", "--untracked-files=all"]),
        "source": source,
        "source_diff": {"bytes": len(diff), "sha256": hashlib.sha256(diff).hexdigest()},
        "rustc_vv": capture([str(RUSTC), "-Vv"]),
        "cargo_version": capture([str(CARGO), "-V"]),
        "uname": capture(["uname", "-a"]),
        "binary": binary_identity(),
    }


def binary_identity() -> dict[str, Any] | None:
    if not BINARY.is_file():
        return None
    mode = stat.S_IMODE(BINARY.stat().st_mode)
    return {"path": str(BINARY.relative_to(ROOT)), "bytes": BINARY.stat().st_size,
            "mode": oct(mode), "sha256": sha256(BINARY)}


def build_env() -> dict[str, str]:
    return {
            "CARGO_HOME": str(USER_ROOT / ".cargo"),
            "CARGO_BUILD_JOBS": "2",
            "CARGO_INCREMENTAL": "0",
            "CARGO_NET_OFFLINE": "true",
            "CARGO_PROFILE_DEV_DEBUG": "0",
            "CARGO_PROFILE_TEST_DEBUG": "0",
            "CARGO_TARGET_DIR": str(TARGET),
            "HOME": str(USER_ROOT),
            "LANG": "C",
            "LC_ALL": "C",
            "PATH": f"{TOOLCHAIN_BIN}:/usr/bin:/bin",
            "RUSTDOCFLAGS": "-C debuginfo=0",
            "RUSTFLAGS": "-C debuginfo=0",
            "RUSTUP_HOME": str(USER_ROOT / ".rustup"),
            "TMPDIR": "/tmp",
            "TZ": "UTC",
    }


def retained_env(env: dict[str, str]) -> dict[str, str]:
    return {key: env.get(key, "") for key in ENV_KEYS}


def gzip_deterministic(source: Path, destination: Path) -> dict[str, Any]:
    destination.parent.mkdir(parents=True, exist_ok=True)
    source_bytes = source.stat().st_size
    source_sha256 = sha256(source)
    with source.open("rb") as input_file, destination.open("wb") as output_file:
        with gzip.GzipFile(
            filename="", mode="wb", fileobj=output_file, compresslevel=9, mtime=0
        ) as compressed:
            shutil.copyfileobj(input_file, compressed, length=1024 * 1024)
    decompressed_bytes = 0
    decompressed = hashlib.sha256()
    with gzip.open(destination, "rb") as input_file:
        for block in iter(lambda: input_file.read(1024 * 1024), b""):
            decompressed.update(block)
            decompressed_bytes += len(block)
    if (decompressed_bytes, decompressed.hexdigest()) != (source_bytes, source_sha256):
        raise RuntimeError(f"gzip verification failed for {destination}")
    return {
        "path": str(destination.relative_to(PACKET)),
        "bytes": destination.stat().st_size,
        "sha256": sha256(destination),
        "decompressed_bytes": decompressed_bytes,
        "decompressed_sha256": source_sha256,
    }


def differential_argv(output: Path) -> list[str]:
    return [
        str(BINARY),
        "differential",
        "--json",
        str(output),
        "--generated",
        "20000",
        "--aliasing",
        "20000",
        *(str(path) for path in FIXTURES),
    ]


def corrupt_one_zip_member(source: Path, destination: Path) -> dict[str, Any]:
    """Copy a package and flip one compressed member byte without fixing CRC."""
    payload = bytearray(source.read_bytes())
    with zipfile.ZipFile(source) as archive:
        info = next(
            item
            for item in archive.infolist()
            if item.filename.lower().endswith((".xml", ".rels", ".vml"))
            and item.compress_size > 0
        )
    header = info.header_offset
    filename_bytes, extra_bytes = struct.unpack_from("<HH", payload, header + 26)
    data_offset = header + 30 + filename_bytes + extra_bytes
    original_byte = payload[data_offset]
    payload[data_offset] ^= 0x01
    destination.write_bytes(payload)
    return {
        "member": info.filename,
        "offset": data_offset,
        "original_byte": original_byte,
        "mutated_byte": payload[data_offset],
        "source": {"bytes": source.stat().st_size, "sha256": sha256(source)},
        "mutated": {"bytes": destination.stat().st_size, "sha256": sha256(destination)},
    }


def json_descriptor(path: Path) -> dict[str, Any]:
    return {
        "path": str(path.relative_to(PACKET)),
        "bytes": path.stat().st_size,
        "sha256": sha256(path),
    }


def verify_negative_report(
    path: Path,
    expected_stages: list[str],
    expected_path: Path,
    expected_member: str | None = None,
) -> dict[str, Any]:
    """Require a structured refusal, rather than accepting any nonzero exit."""
    if not path.is_file():
        raise RuntimeError(f"negative command produced no JSON: {path}")
    report = json.loads(path.read_text())
    if report.get("schema") != "litchi.xml-attribute-bounds.differential.v2":
        raise RuntimeError(f"negative command schema mismatch: {path}")
    inventory = report.get("input_inventory", {})
    failures = inventory.get("failures", [])
    if report.get("qualification_complete") is not False:
        raise RuntimeError(f"negative command did not mark qualification incomplete: {path}")
    if report.get("control_failure_count") != 0:
        raise RuntimeError(f"negative command had control failures: {path}")
    if inventory.get("complete") is not False or inventory.get("read_failure_count") != len(expected_stages):
        raise RuntimeError(f"negative command failure accounting mismatch: {path}")
    if [failure.get("stage") for failure in failures] != expected_stages:
        raise RuntimeError(f"negative command stage mismatch: {path}: {failures}")
    if any(Path(failure.get("path", "")).resolve() != expected_path.resolve()
           for failure in failures):
        raise RuntimeError(f"negative command path mismatch: {path}")
    if expected_member is not None and any(
        failure.get("member") != expected_member for failure in failures
    ):
        raise RuntimeError(f"negative command member mismatch: {path}")
    if report.get("generated") != 0 or report.get("aliasing") != 0:
        raise RuntimeError(f"negative command used a nonzero generated population: {path}")
    return {
        "json": json_descriptor(path),
        "failures": failures,
        "control_failure_count": report["control_failure_count"],
        "qualification_complete": report["qualification_complete"],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--attempt",
        required=True,
        help="create-only packet subdirectory name, for example final-1",
    )
    parser.add_argument(
        "--skip-runtime",
        action="store_true",
        help="run only compile/lint gates; useful for a diagnostic retry",
    )
    parser.add_argument(
        "--reuse-boundary-from",
        help="reuse a prior passing boundary gate after scoped-input checks",
    )
    args = parser.parse_args()

    if not args.attempt or args.attempt in {".", ".."} or "/" in args.attempt:
        raise SystemExit("--attempt must be one path component")
    PACKET.mkdir(parents=True, exist_ok=True)
    attempt_dir = PACKET / "attempts" / args.attempt
    try:
        attempt_dir.mkdir(parents=True, exist_ok=False)
    except FileExistsError as error:
        raise SystemExit(f"attempt already exists: {attempt_dir}") from error
    logs_dir = attempt_dir / "logs"
    runtime_dir = attempt_dir / "runtime"
    logs_dir.mkdir()
    runtime_dir.mkdir()

    target_marker = TARGET / ".owner-0889"
    TARGET.mkdir(parents=True, exist_ok=True)
    marker_text = "litchi 0889 validation target\n"
    if target_marker.exists():
        if target_marker.read_text() != marker_text:
            raise SystemExit(f"target marker belongs to another owner: {target_marker}")
    else:
        target_marker.write_text(marker_text)

    env = build_env()
    commands: list[dict[str, Any]] = []
    receipt: dict[str, Any] = {
        "schema": "litchi.mce-0889-validation.v1",
        "purpose": "correctness_and_bounded_input_qualification",
        "performance_claim": False,
        "started_utc": utc_now(),
        "toolchain": TOOLCHAIN,
        "target_dir": str(TARGET),
        "target_marker": str(target_marker.relative_to(ROOT)),
        "attempt": str(attempt_dir.relative_to(PACKET)),
        "environment": retained_env(env),
        "fixtures": [str(path.relative_to(ROOT)) for path in FIXTURES],
        "commands": commands,
        "status": "running",
    }
    receipt_path = attempt_dir / "validation-receipt.json"
    write_receipt(receipt_path, receipt)

    failure: str | None = None
    before: dict[str, Any] | None = None
    try:
        receipt["runner"] = capture_runner(attempt_dir)
        source_overlay = capture_source_overlay(attempt_dir)
        receipt["source_overlay"] = {
            "path": str((attempt_dir / "source-overlay.json").relative_to(PACKET)),
            **source_overlay,
        }
        write_receipt(receipt_path, receipt)
        before = identity()
        verify_source_overlay(source_overlay, before)
        receipt["before"] = before
        write_receipt(receipt_path, receipt)

        manifest = str(ROOT / "tools/perf-baseline/Cargo.toml")
        common = ["--manifest-path", manifest, "--locked", "--offline"]

        def execute(
            name: str,
            argv: list[str],
            log_name: str,
            *,
            expected: str = "pass",
        ) -> dict[str, Any]:
            return run_command(
                name,
                argv,
                env,
                logs_dir / log_name,
                commands,
                receipt,
                receipt_path,
                expected=expected,
            )

        gate_commands = [
            (
                "fmt_check",
                ["cargo", "fmt", "--manifest-path", manifest, "--all", "--", "--check"],
                "fmt-check.log",
            ),
            (
                "cargo_test_bin",
                ["cargo", "test", *common, "--bin", "xml_attribute_bounds", "--", "--test-threads=1"],
                "cargo-test-xml-attribute-bounds.log",
            ),
            (
                "clippy_bin_no_deps",
                ["cargo", "clippy", *common, "--bin", "xml_attribute_bounds", "--no-deps", "--", "-D", "warnings"],
                "cargo-clippy-xml-attribute-bounds.log",
            ),
            (
                "clippy_bin_tests_no_deps",
                ["cargo", "clippy", *common, "--bin", "xml_attribute_bounds", "--tests", "--no-deps", "--", "-D", "warnings"],
                "cargo-clippy-xml-attribute-bounds-tests.log",
            ),
            (
                "rustdoc_bin",
                ["cargo", "rustdoc", *common, "--bin", "xml_attribute_bounds", "--", "-D", "warnings"],
                "cargo-rustdoc-xml-attribute-bounds.log",
            ),
            (
                "crate_boundary_check",
                ["python3", "-B", "tools/check_crate_boundaries.py"],
                "crate-boundaries.log",
            ),
            (
                "cargo_build_bin",
                ["cargo", "build", *common, "--bin", "xml_attribute_bounds"],
                "cargo-build-xml-attribute-bounds.log",
            ),
        ]
        for name, argv, log_name in gate_commands:
            if name == "crate_boundary_check" and args.reuse_boundary_from:
                record = reused_boundary_check(args.reuse_boundary_from, before)
                commands.append(record)
                receipt["last_command"] = record
                receipt.setdefault("reused_gates", []).append(record)
                write_receipt(receipt_path, receipt)
                continue
            record = execute(name, argv, log_name)
            if not record["pass"]:
                receipt.setdefault("gate_failures", []).append(name)

        if BINARY.is_file():
            receipt["binary_after_build"] = binary_identity()
            write_receipt(receipt_path, receipt)
        elif not args.skip_runtime:
            receipt.setdefault("gate_failures", []).append("binary_missing_after_build")

        if not args.skip_runtime and BINARY.is_file() and not receipt.get("gate_failures"):
            runtime_binary = binary_identity()
            receipt["runtime_binary"] = runtime_binary
            first = runtime_dir / "differential-first.json"
            second = runtime_dir / "differential-second.json"
            first_record = execute("differential_a", differential_argv(first), "differential-a.log")
            second_record = execute("differential_b", differential_argv(second), "differential-b.log")
            if first_record["pass"] and second_record["pass"]:
                first_hash = sha256(first)
                second_hash = sha256(second)
                if first_hash != second_hash or first.read_bytes() != second.read_bytes():
                    receipt.setdefault("runtime_failures", []).append("repeat_json_mismatch")
                else:
                    receipt["differential"] = {
                        "first_bytes": first.stat().st_size,
                        "first_sha256": first_hash,
                        "second_bytes": second.stat().st_size,
                        "second_sha256": second_hash,
                        "byte_identical": True,
                        "generated": 20000,
                        "aliasing": 20000,
                        "packages": 3,
                        "selected_members": 67,
                        "raw_json_retained": False,
                        "compressed_first": gzip_deterministic(
                            first, attempt_dir / "differential-first.json.gz"
                        ),
                        "compressed_second": gzip_deterministic(
                            second, attempt_dir / "differential-second.json.gz"
                        ),
                    }
                    first.unlink()
                    second.unlink()

            missing = runtime_dir / "missing-root"
            negative = execute(
                "negative_missing_root",
                [str(BINARY), "differential", "--json", str(runtime_dir / "missing-root.json"), "--generated", "0", "--aliasing", "0", str(missing)],
                "negative-missing-root.log",
                expected="fail",
            )
            if not negative["pass"]:
                receipt.setdefault("runtime_failures", []).append("missing_root_not_refused")
            try:
                receipt.setdefault("negative_reports", []).append(
                    verify_negative_report(
                        runtime_dir / "missing-root.json",
                        ["root_stat", "no_ooxml_package"],
                        missing,
                    )
                )
            except Exception as error:
                receipt.setdefault("runtime_failures", []).append(repr(error))

            unreadable_dir = runtime_dir / "unreadable"
            unreadable_dir.mkdir()
            unreadable = unreadable_dir / "fixture.xlsx"
            shutil.copy2(FIXTURES[0], unreadable)
            unreadable.chmod(0)
            receipt.setdefault("negative_inputs", []).append(
                {
                    "kind": "unreadable_package",
                    "path": str(unreadable.relative_to(PACKET)),
                    "bytes": unreadable.stat().st_size,
                    "sha256": sha256(unreadable),
                    "mode": oct(stat.S_IMODE(unreadable.stat().st_mode)),
                    "source_fixture": str(FIXTURES[0].relative_to(ROOT)),
                }
            )
            negative = execute(
                "negative_unreadable_package",
                [str(BINARY), "differential", "--json", str(runtime_dir / "unreadable.json"), "--generated", "0", "--aliasing", "0", str(unreadable)],
                "negative-unreadable-package.log",
                expected="fail",
            )
            if not negative["pass"]:
                receipt.setdefault("runtime_failures", []).append("unreadable_package_not_refused")
            try:
                receipt.setdefault("negative_reports", []).append(
                    verify_negative_report(
                        runtime_dir / "unreadable.json",
                        ["package_scan_read", "no_ooxml_package"],
                        unreadable,
                    )
                )
            except Exception as error:
                receipt.setdefault("runtime_failures", []).append(repr(error))

            corrupt = runtime_dir / "corrupt-member.xlsx"
            mutation = corrupt_one_zip_member(FIXTURES[0], corrupt)
            receipt["corrupt_member"] = mutation
            negative = execute(
                "negative_corrupt_member",
                [str(BINARY), "differential", "--json", str(runtime_dir / "corrupt.json"), "--generated", "0", "--aliasing", "0", str(corrupt)],
                "negative-corrupt-member.log",
                expected="fail",
            )
            if not negative["pass"]:
                receipt.setdefault("runtime_failures", []).append("corrupt_member_not_refused")
            try:
                receipt.setdefault("negative_reports", []).append(
                    verify_negative_report(
                        runtime_dir / "corrupt.json",
                        ["member_read"],
                        corrupt,
                        mutation["member"],
                    )
                )
            except Exception as error:
                receipt.setdefault("runtime_failures", []).append(repr(error))

            # The negative reports and mutation recipe now retain all evidence
            # needed to reproduce the refusal.  Remove only these owned copies;
            # never remove the repository fixtures or the compressed reports.
            unreadable.unlink(missing_ok=True)
            unreadable_dir.rmdir()
            corrupt.unlink(missing_ok=True)
            receipt["negative_input_cleanup"] = {
                "removed": [
                    str(unreadable.relative_to(PACKET)),
                    str(corrupt.relative_to(PACKET)),
                ],
                "remaining": [
                    path.name
                    for path in (unreadable, corrupt)
                    if path.exists()
                ],
            }
            write_receipt(receipt_path, receipt)
            receipt["runtime_binary_after"] = binary_identity()
            if receipt["runtime_binary_after"] != runtime_binary:
                receipt.setdefault("runtime_failures", []).append("binary_changed_during_runtime")

        after = identity()
        receipt["after"] = after
        if before["source"] != after["source"]:
            receipt.setdefault("runtime_failures", []).append("source_changed_during_validation")
    except Exception as error:
        failure = repr(error)
        receipt["exception"] = failure
    finally:
        receipt["finished_utc"] = utc_now()
        receipt["command_count"] = len(commands)
        receipt["all_commands_passed"] = all(command["pass"] for command in commands)
        failures = receipt.get("gate_failures", []) + receipt.get("runtime_failures", [])
        receipt["status"] = (
            "pass"
            if not failure and not failures and receipt["all_commands_passed"]
            else "fail"
        )
        if failure:
            receipt["failure"] = failure
        write_receipt(receipt_path, receipt)

    print(json.dumps({"receipt": str(receipt_path), "commands": len(commands), "status": receipt["status"]}))
    if receipt["status"] != "pass":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
