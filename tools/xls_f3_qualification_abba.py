#!/usr/bin/env python3
"""Run the native XLS F3 probe in strict A1/B1/B2/A2 order.

The probe binaries are built separately.  This driver only launches them,
captures each leg outside the checkout, validates the CSV output and writes a
small JSON receipt containing source, binary, driver and host identities.
There are no retries or reordered legs: a failed child or an output digest
mismatch aborts the batch.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


SCHEMA = "litchi.xls.f3.abba.v1"
ROLES = ("A1", "B1", "B2", "A2")
ROUTES = ("write-only", "first-read", "read-after-write")
CORPORA = ("number", "rk-mulrk")
HEX_DIGEST_LENGTH = 64


class RunnerError(RuntimeError):
    """A fail-closed runner or child protocol error."""


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> dict[str, Any]:
    resolved = path.expanduser().resolve(strict=True)
    if not resolved.is_file():
        raise RunnerError(f"binary/script is not a regular file: {resolved}")
    digest = hashlib.sha256()
    size = 0
    with resolved.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
            size += len(chunk)
    return {"path": str(resolved), "bytes": size, "sha256": digest.hexdigest()}


def canonical_json(value: Any) -> bytes:
    return json.dumps(
        value,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=True,
        allow_nan=False,
    ).encode("utf-8")


def run_capture(command: list[str]) -> str:
    try:
        completed = subprocess.run(
            command,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="strict",
        )
    except (OSError, subprocess.CalledProcessError, UnicodeError) as error:
        raise RunnerError(f"command failed: {' '.join(command)}: {error}") from error
    return completed.stdout.strip()


def hash_git_diff(source: Path) -> str:
    process = subprocess.Popen(
        ["git", "-C", str(source), "diff", "--no-ext-diff", "--binary", "HEAD", "--"],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    assert process.stdout is not None
    digest = hashlib.sha256()
    while chunk := process.stdout.read(1024 * 1024):
        digest.update(chunk)
    stderr = process.stderr.read() if process.stderr is not None else b""
    returncode = process.wait()
    if returncode != 0:
        raise RunnerError(
            f"cannot hash git diff for {source}: {stderr.decode('utf-8', 'replace')}"
        )
    return digest.hexdigest()


def source_identity(path: Path) -> dict[str, Any]:
    source = path.expanduser().resolve(strict=True)
    if not source.is_dir():
        raise RunnerError(f"source snapshot is not a directory: {source}")
    revision = run_capture(["git", "-C", str(source), "rev-parse", "HEAD"])
    status = run_capture(
        ["git", "-C", str(source), "status", "--porcelain=v1", "--untracked-files=all"]
    )
    diff_sha256 = hash_git_diff(source)
    identity = {
        "path": str(source),
        "revision": revision,
        "worktree_clean": not status,
        "status_sha256": sha256_bytes(status.encode("utf-8")),
        "diff_sha256": diff_sha256,
    }
    identity["source_sha256"] = sha256_bytes(canonical_json(identity))
    return identity


def host_identity() -> dict[str, Any]:
    uname = platform.uname()._asdict()
    affinity = sorted(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else None
    cpuinfo = Path("/proc/cpuinfo")
    meminfo = Path("/proc/meminfo")
    cpuinfo_sha256 = sha256_file(cpuinfo)["sha256"] if cpuinfo.is_file() else None
    meminfo_sha256 = sha256_file(meminfo)["sha256"] if meminfo.is_file() else None
    facts = {
        "uname": uname,
        "python": sys.version,
        "cpu_count": os.cpu_count(),
        "affinity": affinity,
        "cpuinfo_sha256": cpuinfo_sha256,
        "meminfo_sha256": meminfo_sha256,
    }
    facts["host_sha256"] = sha256_bytes(canonical_json(facts))
    return facts


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--control-binary", type=Path, required=True)
    parser.add_argument("--candidate-binary", type=Path, required=True)
    parser.add_argument("--control-source", type=Path, required=True)
    parser.add_argument("--candidate-source", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--cpu", type=int)
    parser.add_argument("--warmup", type=int, default=20)
    parser.add_argument("--samples", type=int, default=40)
    parser.add_argument("--batch", type=int, required=True)
    parser.add_argument("--timeout-seconds", type=float, default=3600.0)
    parser.add_argument("--route", choices=ROUTES, action="append", dest="routes")
    parser.add_argument("--corpus", choices=CORPORA, action="append", dest="corpora")
    return parser.parse_args()


def validate_options(options: argparse.Namespace) -> tuple[list[str], list[str]]:
    if options.warmup < 0 or options.samples <= 0 or options.batch <= 0:
        raise RunnerError("warmup must be nonnegative; samples and batch must be positive")
    if options.timeout_seconds <= 0:
        raise RunnerError("timeout must be positive")
    if options.cpu is not None and options.cpu < 0:
        raise RunnerError("CPU affinity must be nonnegative")
    if options.cpu is not None and shutil.which("taskset") is None:
        raise RunnerError("--cpu requires taskset")
    if options.control_binary.resolve() == options.candidate_binary.resolve():
        raise RunnerError("control and candidate binaries must be distinct paths")
    return options.routes or list(ROUTES), options.corpora or list(CORPORA)


def parse_probe_output(
    stdout_path: Path,
    route: str,
    corpus: str,
    samples: int,
    batch: int,
) -> dict[str, Any]:
    raw = stdout_path.read_bytes()
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        raise RunnerError(f"{stdout_path} is not UTF-8: {error}") from error
    rows = list(csv.DictReader(line for line in text.splitlines() if line.strip()))
    required = {
        "route",
        "corpus",
        "sample",
        "nanoseconds",
        "edit_ns",
        "set_ns",
        "commit_ns",
        "write_ns",
        "read_ns",
        "output_bytes",
        "output_sha256",
        "batch",
    }
    if not rows:
        raise RunnerError(f"{stdout_path} has no CSV rows")
    if set(rows[0]) != required:
        raise RunnerError(f"{stdout_path} has unexpected CSV columns: {set(rows[0])}")
    if len(rows) != samples:
        raise RunnerError(f"{stdout_path} has {len(rows)} rows, expected {samples}")
    digests: set[str] = set()
    output_bytes: set[int] = set()
    parsed_samples: list[int] = []
    for row in rows:
        if row["route"] != route or row["corpus"] != corpus:
            raise RunnerError(f"{stdout_path} contains a route/corpus mismatch")
        try:
            sample = int(row["sample"])
            numeric = {
                name: int(row[name])
                for name in (
                    "nanoseconds",
                    "edit_ns",
                    "set_ns",
                    "commit_ns",
                    "write_ns",
                    "read_ns",
                    "output_bytes",
                    "batch",
                )
            }
        except ValueError as error:
            raise RunnerError(f"{stdout_path} contains a non-integer CSV value") from error
        if sample < 0 or any(value < 0 for value in numeric.values()):
            raise RunnerError(f"{stdout_path} contains a negative measurement")
        if numeric["batch"] != batch:
            raise RunnerError(f"{stdout_path} reports batch {numeric['batch']}, expected {batch}")
        digest = row["output_sha256"]
        if len(digest) != HEX_DIGEST_LENGTH or any(
            character not in "0123456789abcdef" for character in digest
        ):
            raise RunnerError(f"{stdout_path} contains an invalid output digest")
        parsed_samples.append(sample)
        digests.add(digest)
        output_bytes.add(numeric["output_bytes"])
    if parsed_samples != list(range(samples)):
        raise RunnerError(f"{stdout_path} sample indexes are not contiguous")
    if len(digests) != 1 or len(output_bytes) != 1:
        raise RunnerError(f"{stdout_path} output identity is not stable")
    return {
        "rows": len(rows),
        "output_sha256": next(iter(digests)),
        "output_bytes": next(iter(output_bytes)),
        "stdout_sha256": sha256_bytes(raw),
    }


def run_leg(
    *,
    role: str,
    binary: Path,
    route: str,
    corpus: str,
    options: argparse.Namespace,
    output_dir: Path,
) -> dict[str, Any]:
    stem = f"{route}-{corpus}-{role}"
    stdout_path = output_dir / f"{stem}.stdout.csv"
    stderr_path = output_dir / f"{stem}.stderr.log"
    command = [
        str(binary.resolve()),
        "--route",
        route,
        "--corpus",
        corpus,
        "--warmup",
        str(options.warmup),
        "--samples",
        str(options.samples),
        "--batch",
        str(options.batch),
    ]
    if options.cpu is not None:
        command = ["taskset", "-c", str(options.cpu), *command]
    try:
        with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
            completed = subprocess.run(
                command,
                stdout=stdout,
                stderr=stderr,
                timeout=options.timeout_seconds,
                check=False,
            )
    except (OSError, subprocess.TimeoutExpired) as error:
        raise RunnerError(f"{role} child failed to run: {error}") from error
    if completed.returncode != 0:
        raise RunnerError(f"{role} child returned {completed.returncode}: {stderr_path}")
    parsed = parse_probe_output(stdout_path, route, corpus, options.samples, options.batch)
    parsed.update(
        {
            "role": role,
            "route": route,
            "corpus": corpus,
            "command": command,
            "returncode": completed.returncode,
            "stdout": sha256_file(stdout_path),
            "stderr": sha256_file(stderr_path),
        }
    )
    return parsed


def main() -> int:
    options = parse_args()
    try:
        routes, corpora = validate_options(options)
        control_binary = sha256_file(options.control_binary)
        candidate_binary = sha256_file(options.candidate_binary)
        if control_binary["sha256"] == candidate_binary["sha256"]:
            raise RunnerError("control and candidate binaries have the same SHA-256")
        control_source = source_identity(options.control_source)
        candidate_source = source_identity(options.candidate_source)
        driver = sha256_file(Path(__file__))
        host = host_identity()
        output_dir = options.output_dir.expanduser().resolve()
        output_dir.mkdir(parents=True, exist_ok=True)
        legs: list[dict[str, Any]] = []
        expected_outputs: dict[tuple[str, str], str] = {}
        for route in routes:
            for corpus in corpora:
                binaries = {
                    "A1": options.control_binary,
                    "B1": options.candidate_binary,
                    "B2": options.candidate_binary,
                    "A2": options.control_binary,
                }
                for role in ROLES:
                    leg = run_leg(
                        role=role,
                        binary=binaries[role],
                        route=route,
                        corpus=corpus,
                        options=options,
                        output_dir=output_dir,
                    )
                    legs.append(leg)
                selected = [
                    leg["output_sha256"]
                    for leg in legs
                    if leg["route"] == route and leg["corpus"] == corpus
                ]
                if len(set(selected)) != 1:
                    raise RunnerError(f"{route}/{corpus} output digest differs across ABBA legs")
                expected_outputs[(route, corpus)] = selected[0]
        receipt = {
            "schema": SCHEMA,
            "created_utc": datetime.now(timezone.utc).isoformat(),
            "order": list(ROLES),
            "routes": routes,
            "corpora": corpora,
            "configuration": {
                "warmup": options.warmup,
                "samples": options.samples,
                "batch": options.batch,
                "cpu": options.cpu,
                "timeout_seconds": options.timeout_seconds,
            },
            "sources": {"control": control_source, "candidate": candidate_source},
            "binaries": {"control": control_binary, "candidate": candidate_binary},
            "driver": driver,
            "host": host,
            "outputs": {
                f"{route}/{corpus}": digest
                for (route, corpus), digest in sorted(expected_outputs.items())
            },
            "legs": legs,
        }
        receipt_path = output_dir / "xls-f3-abba-receipt.json"
        receipt_path.write_bytes(canonical_json(receipt) + b"\n")
        print(receipt_path)
        return 0
    except RunnerError as error:
        print(f"xls F3 ABBA: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
