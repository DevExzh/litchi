#!/usr/bin/env python3
"""Build the isolated 0925 harness with complete standalone custody.

This wrapper is intentionally inert until the caller sets
``LITCHI_PERF_0925_BUILD_APPROVED=1``. It never writes into the checkout or a
previous target directory, and it records the exact compiler, lock, profile,
environment, log, and binary identity needed by ``comparison-capture.py``.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


REPO_ROOT = Path(__file__).resolve().parents[4]
BASE_REVISION = "aa077c030819c3019288e350f34959666ab7204f"
MAX_ARTIFACT_BYTES = 1 << 20


class BuildError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise BuildError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def external(path: Path) -> Path:
    path = path.resolve()
    try:
        path.relative_to(REPO_ROOT)
    except ValueError:
        return path
    raise BuildError(f"build output/log path must be external to checkout: {path}")


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def run_capture(command: list[str], *, log: Path, cwd: Path) -> tuple[int, str]:
    started = utc_now()
    with log.open("w", encoding="utf-8") as stream:
        stream.write("$ " + " ".join(command) + "\n")
        stream.write(f"started_utc={started}\n")
        stream.flush()
        completed = subprocess.run(command, cwd=str(cwd), stdout=stream, stderr=subprocess.STDOUT,
                                   text=True, check=False)
        ended = utc_now()
        stream.write(f"ended_utc={ended}\nexit_code={completed.returncode}\n")
    return completed.returncode, started


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-root", type=Path, required=True)
    parser.add_argument("--target-dir", type=Path, required=True)
    parser.add_argument("--build-log", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--source-receipt", type=Path, required=True)
    parser.add_argument("--full-source-listing", type=Path, required=True)
    parser.add_argument("--source-revision", default=BASE_REVISION)
    args = parser.parse_args()
    try:
        require(os.environ.get("LITCHI_PERF_0925_BUILD_APPROVED") == "1",
                "LITCHI_PERF_0925_BUILD_APPROVED=1 is required to build")
        source = external(args.source_root)
        target = external(args.target_dir)
        log = external(args.build_log)
        receipt_path = external(args.output)
        source_receipt_path = external(args.source_receipt)
        source_listing_path = external(args.full_source_listing)
        require(args.source_revision == BASE_REVISION,
                f"unexpected source revision: {args.source_revision}")
        require(source.is_dir(), f"source root is not a directory: {source}")
        manifest = source / "tools/perf-baseline/Cargo.toml"
        lock = source / "tools/perf-baseline/Cargo.lock"
        root_lock = source / "Cargo.lock"
        rustfmt = source / "rustfmt.toml"
        for path in (manifest, lock, root_lock, rustfmt):
            require(path.is_file(), f"source custody file is missing: {path}")
        require(source_receipt_path.is_file() and not source_receipt_path.is_symlink(),
                f"source receipt is missing: {source_receipt_path}")
        require(source_listing_path.is_file() and not source_listing_path.is_symlink(),
                f"source listing is missing: {source_listing_path}")
        try:
            source_receipt = json.loads(source_receipt_path.read_text(encoding="utf-8"))
        except (OSError, UnicodeError, ValueError) as error:
            raise BuildError(f"source receipt is not valid JSON: {source_receipt_path}") from error
        require(isinstance(source_receipt, dict), "source receipt is not an object")
        receipt_revision = source_receipt.get("base_revision") or source_receipt.get("source_revision")
        require(receipt_revision == BASE_REVISION, "source receipt base revision changed")
        source_files = source_receipt.get("source_files") or source_receipt.get("candidate_source_files")
        require(isinstance(source_files, list) and source_files,
                "source receipt has no exact source file list")
        require(not target.exists(), f"refusing to reuse target directory: {target}")
        target.mkdir(parents=True)
        log.parent.mkdir(parents=True, exist_ok=True)
        rustc_version = subprocess.check_output(["rustc", "+1.95.0", "-vV"], text=True)
        cargo_version = subprocess.check_output(["cargo", "+1.95.0", "-V"], text=True)
        command = ["cargo", "+1.95.0", "build", "--offline", "-j2", "--locked", "--release",
                   "--manifest-path", str(manifest), "--bin", "litchi-perf-baseline",
                   "--target-dir", str(target)]
        return_code, started = run_capture(command, log=log, cwd=source)
        ended = utc_now()
        binary = target / "release/litchi-perf-baseline"
        require(return_code == 0 and binary.is_file(),
                f"standalone build failed (exit {return_code}); retain {log}")
        binary_descriptor = descriptor(binary)
        info = target / ".rustc_info.json"
        receipt: dict[str, Any] = {
            "schema": "litchi.performance.0925.provider-build-receipt.v1",
            "created_utc": utc_now(),
            "base_revision": BASE_REVISION,
            "source_root": str(source),
            "source_custody": {
                "source_receipt": descriptor(source_receipt_path),
                "full_source_listing": descriptor(source_listing_path),
            },
            "profile": "release",
            "command": " ".join(command),
            "command_argv": command,
            "toolchain": {"rustc": rustc_version, "cargo": cargo_version},
            "environment": {
                key: os.environ[key]
                for key in sorted(os.environ)
                if key.startswith(("CARGO_PROFILE_", "CARGO_BUILD_", "RUSTFLAGS", "RUSTC_WRAPPER"))
            },
            "locks": {
                "tools_performance_cargo_lock": {"path": str(lock), "bytes": lock.stat().st_size,
                                                   "sha256": sha256(lock)},
                "root_cargo_lock": {"path": str(root_lock), "bytes": root_lock.stat().st_size,
                                     "sha256": sha256(root_lock)},
                "rustfmt_toml": {"path": str(rustfmt), "bytes": rustfmt.stat().st_size,
                                  "sha256": sha256(rustfmt)},
            },
            "binary": binary_descriptor,
            "rustc_info": descriptor(info) if info.is_file() else None,
            "build_log": descriptor(log),
            "started_utc": started,
            "ended_utc": ended,
            "exit_code": return_code,
        }
        data = (json.dumps(receipt, indent=2, sort_keys=True) + "\n").encode()
        require(len(data) < MAX_ARTIFACT_BYTES, "build receipt is at least 1 MiB")
        require(not receipt_path.exists(), f"refusing to replace retained receipt: {receipt_path}")
        receipt_path.parent.mkdir(parents=True, exist_ok=True)
        receipt_path.write_bytes(data)
        print(json.dumps({"status": "built", "receipt": str(receipt_path),
                          "binary": binary_descriptor}, sort_keys=True))
    except (BuildError, OSError, subprocess.SubprocessError, ValueError) as error:
        print(f"0925 build custody failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
