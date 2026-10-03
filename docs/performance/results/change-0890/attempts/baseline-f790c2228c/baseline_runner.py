#!/usr/bin/env python3
"""Build and capture the 0890 MCE baseline from the f790c2228c source tree.

This is a correctness/equality capture, not a performance benchmark.  The
source tree is a bounded archive of the committed base: it contains every
tracked Cargo/build input below ``crates`` and ``tools`` plus the workspace
configuration, while the repository's committed evidence, media, and fixture
trees are deliberately excluded.  The three runtime fixtures stay at their
original absolute paths so a later candidate capture can compare the report
bytes directly.
"""

from __future__ import annotations

import datetime as dt
import gzip
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
import subprocess
import time
from typing import Any


REPO = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
BASE_COMMIT = "f790c2228cbd9e89c2892ba3ffaeb49802fac43b"
SOURCE = Path("/tmp/litchi-0890-correctness-source")
TARGET = REPO / "target/litchi-0890-correctness"
TARGET_MARKER = TARGET / ".owner-0890-target"
BASELINE_COPY_ROOT = Path("/tmp/litchi-0890-correctness-baseline")
BASELINE_COPY = BASELINE_COPY_ROOT / "xml_attribute_bounds"
BASELINE_MARKER = BASELINE_COPY_ROOT / ".owner-0890-baseline"
MANIFEST = SOURCE / "tools/perf-baseline/Cargo.toml"
CARGO = Path("/home/zhuhe/.rustup/toolchains/1.95.0-x86_64-unknown-linux-gnu/bin/cargo")
RUSTC = Path("/home/zhuhe/.rustup/toolchains/1.95.0-x86_64-unknown-linux-gnu/bin/rustc")
BINARY = TARGET / "debug/xml_attribute_bounds"
ATTEMPT_NAME = "baseline-f790c2228c"
ATTEMPT = PACKET / "attempts" / ATTEMPT_NAME
LOGS = ATTEMPT / "logs"
RUNTIME = ATTEMPT / "runtime"

FIXTURES = [
    REPO / "test-data/ooxml/xlsx/StructuredRefs-lots-with-lookups.xlsx",
    REPO / "test-data/ooxml/docx/drawing.docx",
    REPO / "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx",
]

SOURCE_PATHS = ("Cargo.toml", "clippy.toml", "deny.toml", "rust-toolchain.toml", "rustfmt.toml", ".cargo", "crates", "tools")
SOURCE_EXCLUDED_PREFIXES = ("docs/", "test-data/", "media/", ".github/")
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


def write_json(path: Path, value: dict[str, Any]) -> None:
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")
    temporary.replace(path)


def command_text(argv: list[str]) -> str:
    return json.dumps(argv, ensure_ascii=False, separators=(",", ":"))


def source_blob_sha1(data: bytes) -> str:
    return hashlib.sha1(b"blob " + str(len(data)).encode() + b"\0" + data).hexdigest()


def source_manifest() -> dict[str, Any]:
    listing = subprocess.run(
        ["git", "ls-tree", "-r", "-z", BASE_COMMIT, "--", *SOURCE_PATHS],
        cwd=REPO,
        stdout=subprocess.PIPE,
        check=True,
    ).stdout
    expected: dict[str, str] = {}
    for record in listing.split(b"\0"):
        if not record:
            continue
        metadata, path_bytes = record.split(b"\t", 1)
        fields = metadata.split()
        if len(fields) != 3 or fields[1] != b"blob":
            raise RuntimeError(f"unexpected source tree row: {record!r}")
        path = path_bytes.decode()
        if path.startswith(SOURCE_EXCLUDED_PREFIXES):
            raise RuntimeError(f"source selection unexpectedly includes excluded path: {path}")
        expected[path] = fields[2].decode()

    actual = {
        str(path.relative_to(SOURCE))
        for path in SOURCE.rglob("*")
        if path.is_file() or path.is_symlink()
    }
    allowed_extra = {".owner-0890-source"}
    missing = sorted(set(expected) - actual)
    extra = sorted(actual - set(expected) - allowed_extra)
    if missing or extra:
        raise RuntimeError(f"source file set mismatch: missing={missing[:8]} extra={extra[:8]}")
    entries: list[dict[str, Any]] = []
    for path in sorted(expected):
        source_path = SOURCE / path
        if source_path.is_symlink():
            raise RuntimeError(f"source archive unexpectedly contains symlink: {path}")
        data = source_path.read_bytes()
        blob_sha1 = source_blob_sha1(data)
        if blob_sha1 != expected[path]:
            raise RuntimeError(f"source blob differs from base: {path}")
        entries.append(
            {
                "path": path,
                "bytes": len(data),
                "mode": oct(stat.S_IMODE(source_path.stat().st_mode)),
                "git_blob_sha1": blob_sha1,
                "sha256": hashlib.sha256(data).hexdigest(),
            }
        )
    marker = SOURCE / ".owner-0890-source"
    if marker.read_text() != "litchi 0890 correctness source archive\n":
        raise RuntimeError(f"source ownership marker mismatch: {marker}")
    return {
        "base_commit": BASE_COMMIT,
        "root": str(SOURCE),
        "selection": list(SOURCE_PATHS),
        "excluded_top_level_prefixes": list(SOURCE_EXCLUDED_PREFIXES),
        "files": entries,
        "file_count": len(entries),
        "marker": {"path": str(marker), "sha256": sha256(marker)},
    }


def identity(path: Path) -> dict[str, Any] | None:
    if not path.is_file():
        return None
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "mode": oct(stat.S_IMODE(path.stat().st_mode)),
        "sha256": sha256(path),
    }


def environment() -> dict[str, str]:
    target = str(TARGET)
    toolchain_bin = str(CARGO.parent)
    return {
        "CARGO_HOME": "/home/zhuhe/.cargo",
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
        "CARGO_NET_OFFLINE": "true",
        "CARGO_PROFILE_DEV_DEBUG": "0",
        "CARGO_PROFILE_TEST_DEBUG": "0",
        "CARGO_TARGET_DIR": target,
        "HOME": "/home/zhuhe",
        "LANG": "C",
        "LC_ALL": "C",
        "PATH": toolchain_bin + ":/usr/bin:/bin",
        "RUSTDOCFLAGS": "-C debuginfo=0",
        "RUSTFLAGS": "-C debuginfo=0",
        "RUSTUP_HOME": "/home/zhuhe/.rustup",
        "TMPDIR": "/tmp",
        "TZ": "UTC",
    }


def retained_env(env: dict[str, str]) -> dict[str, str]:
    return {key: env.get(key, "") for key in ENV_KEYS}


def run_command(
    name: str,
    argv: list[str],
    env: dict[str, str],
    receipt: dict[str, Any],
    receipt_path: Path,
    *,
    expected: str = "pass",
) -> dict[str, Any]:
    log_path = LOGS / f"{name}.log"
    started = utc_now()
    started_ns = time.monotonic_ns()
    launch_error: str | None = None
    returncode: int | None = None
    LOGS.mkdir(parents=True, exist_ok=True)
    with log_path.open("wb") as log:
        log.write(f"cwd={SOURCE}\nargv={command_text(argv)}\n\n".encode())
        try:
            result = subprocess.run(
                argv,
                cwd=SOURCE,
                env=env,
                stdout=log,
                stderr=subprocess.STDOUT,
                check=False,
            )
            returncode = result.returncode
        except OSError as error:
            launch_error = repr(error)
    record: dict[str, Any] = {
        "name": name,
        "argv": argv,
        "cwd": str(SOURCE),
        "expected": expected,
        "returncode": returncode,
        "started_utc": started,
        "finished_utc": utc_now(),
        "elapsed_ns": time.monotonic_ns() - started_ns,
        "log": str(log_path.relative_to(PACKET)),
        "log_bytes": log_path.stat().st_size,
        "log_sha256": sha256(log_path),
        "pass": (returncode == 0) if expected == "pass" else returncode != 0,
    }
    if launch_error is not None:
        record["launch_error"] = launch_error
    receipt.setdefault("commands", []).append(record)
    receipt["last_command"] = record
    write_json(receipt_path, receipt)
    return record


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


def compress_report(source: Path, destination: Path) -> dict[str, Any]:
    source_size = source.stat().st_size
    source_sha = sha256(source)
    with source.open("rb") as input_file, destination.open("wb") as output_file:
        with gzip.GzipFile(filename="", mode="wb", fileobj=output_file, compresslevel=9, mtime=0) as compressed:
            shutil.copyfileobj(input_file, compressed, length=1024 * 1024)
    decompressed_size = 0
    decompressed = hashlib.sha256()
    with gzip.open(destination, "rb") as input_file:
        for block in iter(lambda: input_file.read(1024 * 1024), b""):
            decompressed.update(block)
            decompressed_size += len(block)
    if (decompressed_size, decompressed.hexdigest()) != (source_size, source_sha):
        raise RuntimeError(f"compressed report failed round-trip verification: {destination}")
    return {
        "path": str(destination.relative_to(PACKET)),
        "bytes": destination.stat().st_size,
        "sha256": sha256(destination),
        "decompressed_bytes": source_size,
        "decompressed_sha256": source_sha,
    }


def fixture_identity(path: Path) -> dict[str, Any]:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def main() -> int:
    if ATTEMPT.exists():
        raise SystemExit(f"create-only attempt already exists: {ATTEMPT}")
    if not SOURCE.is_dir() or not TARGET.is_dir():
        raise SystemExit("owned source/target roots are missing")
    if not TARGET_MARKER.is_file() or TARGET_MARKER.read_text() != "litchi 0890 correctness target\n":
        raise SystemExit(f"target marker mismatch: {TARGET_MARKER}")
    if not BASELINE_MARKER.is_file() or BASELINE_MARKER.read_text() != "litchi 0890 baseline executable custody\n":
        raise SystemExit(f"baseline custody marker mismatch: {BASELINE_MARKER}")
    if BASELINE_COPY.exists():
        raise SystemExit(f"baseline copy already exists: {BASELINE_COPY}")
    for fixture in FIXTURES:
        if not fixture.is_file():
            raise SystemExit(f"fixture missing: {fixture}")

    ATTEMPT.mkdir(parents=True)
    RUNTIME.mkdir()
    script_copy = ATTEMPT / Path(__file__).name
    shutil.copy2(__file__, script_copy)
    env = environment()
    receipt: dict[str, Any] = {
        "schema": "litchi.mce-0890-correctness-baseline.v1",
        "purpose": "baseline_correctness_and_candidate_equality_reference",
        "performance_claim": False,
        "status": "running",
        "started_utc": utc_now(),
        "base_commit": BASE_COMMIT,
        "source": None,
        "source_root": str(SOURCE),
        "target_dir": str(TARGET),
        "target_marker": str(TARGET_MARKER),
        "baseline_copy_root": str(BASELINE_COPY_ROOT),
        "baseline_copy_marker": str(BASELINE_MARKER),
        "manifest": str(MANIFEST),
        "toolchain": {"cargo": str(CARGO), "rustc": str(RUSTC)},
        "environment": retained_env(env),
        "fixtures": [fixture_identity(path) for path in FIXTURES],
        "commands": [],
        "runner": {"path": str(script_copy.relative_to(PACKET)), "bytes": script_copy.stat().st_size, "sha256": sha256(script_copy)},
        "repo_identity": {
            "git_head": subprocess.run(["git", "rev-parse", "HEAD"], cwd=REPO, text=True, capture_output=True, check=True).stdout.rstrip("\n"),
            "git_status": subprocess.run(["git", "status", "--short", "--untracked-files=all"], cwd=REPO, text=True, capture_output=True, check=True).stdout.rstrip("\n"),
        },
    }
    receipt_path = ATTEMPT / "baseline-receipt.json"
    try:
        receipt["source"] = source_manifest()
        write_json(receipt_path, receipt)
        common = ["--manifest-path", str(MANIFEST), "--locked", "--offline"]
        build = run_command(
            "cargo-build-baseline",
            [str(CARGO), "build", *common, "--bin", "xml_attribute_bounds"],
            env,
            receipt,
            receipt_path,
        )
        if not build["pass"] or not BINARY.is_file():
            receipt.setdefault("failures", []).append("baseline_build")
        else:
            receipt["binary"] = identity(BINARY)
            shutil.copy2(BINARY, BASELINE_COPY)
            receipt["baseline_copy"] = identity(BASELINE_COPY)
            write_json(receipt_path, receipt)
            first = RUNTIME / "differential-first.json"
            second = RUNTIME / "differential-second.json"
            first_record = run_command("differential-baseline-a", differential_argv(first), env, receipt, receipt_path)
            second_record = run_command("differential-baseline-b", differential_argv(second), env, receipt, receipt_path)
            if not first_record["pass"] or not second_record["pass"]:
                receipt.setdefault("failures", []).append("baseline_differential")
            elif not first.is_file() or not second.is_file():
                receipt.setdefault("failures", []).append("baseline_differential_output_missing")
            else:
                first_sha = sha256(first)
                second_sha = sha256(second)
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
                    receipt.setdefault("failures", []).append("baseline_differential_repeat_mismatch")
    except Exception as error:
        receipt.setdefault("failures", []).append(repr(error))
    receipt["finished_utc"] = utc_now()
    receipt["status"] = "pass" if not receipt.get("failures") else "fail"
    write_json(receipt_path, receipt)
    print(json.dumps({"status": receipt["status"], "receipt": str(receipt_path), "binary": receipt.get("binary"), "differential": receipt.get("differential")}, sort_keys=True))
    return 0 if receipt["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
