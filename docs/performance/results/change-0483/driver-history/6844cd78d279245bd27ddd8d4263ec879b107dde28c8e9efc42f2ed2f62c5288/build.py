#!/usr/bin/env python3
"""Build and custody one normal and one allocator DOCX tail executable.

This driver is intentionally inert until the coordinator invokes it. Every
invocation requires an explicit attempt token; accepted and developmental
builds therefore receive distinct validation, binary and source identities.
The coordinator must run one build driver at a time: the gate serializes each
Cargo invocation, while the copy and custody step immediately after that gate
must stay in the same externally serialized build window so another build
cannot replace Cargo's shared release output before it is copied.
"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path
import subprocess
import sys

from common import ENV, ROOT, REPO, TEMP, meta, now, read, sha, write


BINARY_NAME = "docx_bounded_tail_append_compare"


def parse() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--attempt",
        required=True,
        help="path-safe custody token, for example dev-01 or accepted",
    )
    return parser.parse_args()


def check_token(value: str) -> str:
    if (
        not value
        or value in {".", ".."}
        or "/" in value
        or "\\" in value
        or any(character.isspace() for character in value)
    ):
        raise SystemExit("--attempt must be a non-empty path-safe token")
    return value


def build_one(attempt: str, instrumentation: str) -> dict[str, object]:
    revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO, text=True).strip()
    command = [
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        "tools/perf-baseline/Cargo.toml",
        "--bin",
        BINARY_NAME,
    ]
    if instrumentation == "allocator":
        command.extend(["--features", "allocator-metrics"])
    label = f"build-{instrumentation}"
    gate_command = [
        sys.executable,
        "-B",
        str(ROOT / "gate.py"),
        "--attempt",
        attempt,
        label,
        *command,
    ]
    subprocess.run(gate_command, cwd=REPO, env=ENV, check=True)
    if subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO, text=True).strip() != revision:
        raise RuntimeError("Git revision changed during build")
    gate_path = ROOT / "validation" / f"{label}-{attempt}.json"
    gate = read(gate_path)
    if gate.get("exit_code") != 0 or gate.get("source_unchanged") is not True:
        raise RuntimeError(f"{gate_path}: build gate did not pass")
    origin = REPO / "tools/perf-baseline" / "target" / "release" / BINARY_NAME
    if not origin.is_file():
        raise RuntimeError(f"missing build output: {origin}")
    destination = TEMP / attempt / instrumentation / BINARY_NAME
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        raise RuntimeError(f"refusing to replace copied binary: {destination}")
    origin_meta = meta(origin)
    shutil.copy2(origin, destination)
    if meta(destination) != origin_meta:
        raise RuntimeError("copied binary metadata differs from Cargo output")
    gate_relative = gate_path.relative_to(ROOT).as_posix()
    build = dict(
        schema="docx-tail-append-build-v1",
        attempt=attempt,
        instrumentation=instrumentation,
        git_revision=revision,
        binary_name=BINARY_NAME,
        command=command,
        gate_label=f"{label}-{attempt}",
        gate_path=gate_relative,
        gate_sha256=sha(gate_path),
        gate=gate,
        source_manifest_sha256=gate["source_after"]["sha256"],
        copied_utc=now(),
        original_binary={"path": str(origin), **origin_meta},
        binary={"path": str(destination), **origin_meta},
    )
    build_path = ROOT / "builds" / f"{instrumentation}-{attempt}.json"
    write(build_path, build)
    return {
        "path": str(destination),
        **origin_meta,
        "build_path": build_path.relative_to(ROOT).as_posix(),
        "build_sha256": sha(build_path),
        "source_manifest_sha256": gate["source_after"]["sha256"],
    }


def main() -> None:
    options = parse()
    attempt = check_token(options.attempt)
    binaries = {
        instrumentation: build_one(attempt, instrumentation)
        for instrumentation in ("normal", "allocator")
    }
    if binaries["normal"]["source_manifest_sha256"] != binaries["allocator"]["source_manifest_sha256"]:
        raise RuntimeError("normal and allocator builds have different source manifests")
    revisions = {read(ROOT / value["build_path"])["git_revision"] for value in binaries.values()}
    if len(revisions) != 1:
        raise RuntimeError("normal and allocator builds have different Git revisions")
    write(ROOT / "builds" / f"binaries-{attempt}.json", {
        "schema": "docx-tail-append-binaries-v1",
        "attempt": attempt,
        "binary_name": BINARY_NAME,
        "git_revision": revisions.pop(),
        "source_manifest_sha256": binaries["normal"]["source_manifest_sha256"],
        "binaries": binaries,
    })


if __name__ == "__main__":
    main()
